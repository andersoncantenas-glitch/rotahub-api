import 'package:flutter/material.dart';

import '../models/app_config.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';
import 'programacao_create_page.dart';

class RascunhoPage extends StatefulWidget {
  const RascunhoPage({
    super.key,
    required this.config,
    required this.api,
    required this.onProgramCreated,
  });

  final AppConfig config;
  final VendedorApiService api;
  final Future<void> Function() onProgramCreated;

  @override
  State<RascunhoPage> createState() => RascunhoPageState();
}

class RascunhoPageState extends State<RascunhoPage> {
  final TextEditingController _buscaCtrl = TextEditingController();
  bool _loading = true;
  bool _movingToPreProgramacao = false;
  bool _savingPreProgramacao = false;
  String? _error;
  String _statusFiltro = 'TODAS';
  String _vendedorFiltro = 'TODOS';
  List<Map<String, dynamic>> _itens = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _preProgramacoes = <Map<String, dynamic>>[];
  String? _preProgramacaoAtualId;
  String _preProgramacaoAtualTitulo = '';
  final Set<String> _selecionadosOrigem = <String>{};
  final List<String> _preProgramacaoIds = <String>[];

  @override
  void initState() {
    super.initState();
    reload();
  }

  @override
  void dispose() {
    _buscaCtrl.dispose();
    super.dispose();
  }

  Future<void> reload() async {
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final data = await widget.api.listarRascunhoVendas();
      final preProgramacoes =
          await widget.api.listarPreProgramacoes(status: 'ABERTA', limit: 120);
      final idsValidos = data
          .map((item) => (item['id'] ?? '').toString().trim())
          .where((item) => item.isNotEmpty)
          .toSet();
      final vendedores = data
          .map((item) => _vendorName(item))
          .where((item) => item.isNotEmpty)
          .toSet();

      if (!mounted) return;
      setState(() {
        _itens = data;
        _preProgramacoes = preProgramacoes;
        _preProgramacaoIds.removeWhere((id) => !idsValidos.contains(id));
        _selecionadosOrigem.removeWhere((id) => !idsValidos.contains(id));
        final preAtualExiste = _preProgramacaoAtualId != null &&
            preProgramacoes.any(
              (item) => _text(item['id']) == _preProgramacaoAtualId,
            );
        if (!preAtualExiste) {
          _preProgramacaoAtualId = null;
          _preProgramacaoAtualTitulo = '';
        }
        if (_vendedorFiltro != 'TODOS' && !vendedores.contains(_vendedorFiltro)) {
          _vendedorFiltro = vendedores.isEmpty ? 'TODOS' : vendedores.first;
        }
        _loading = false;
      });
    } catch (error) {
      if (!mounted) return;
      setState(() {
        _error = error.toString();
        _loading = false;
      });
    }
  }

  String _text(dynamic value) => (value ?? '').toString().trim();

  String _upper(dynamic value) => _text(value).toUpperCase();

  String _vendorName(Map<String, dynamic> item) {
    final origem = _upper(item['vendedor_origem']);
    if (origem.isNotEmpty) return origem;
    return _upper(item['vendedor_cadastro']);
  }

  int _int(dynamic value) {
    if (value is int) return value;
    return int.tryParse(_text(value)) ?? 0;
  }

  double _double(dynamic value) {
    if (value is double) return value;
    if (value is int) return value.toDouble();
    return double.tryParse(_text(value).replaceAll(',', '.')) ?? 0.0;
  }

  String _formatMoney(dynamic value) {
    final number = _double(value);
    return number.toStringAsFixed(2).replaceAll('.', ',');
  }

  Map<String, int> get _contagemPorVendedor {
    final counts = <String, int>{};
    for (final item in _itens) {
      final vendedor = _vendorName(item);
      if (vendedor.isEmpty) continue;
      counts[vendedor] = (counts[vendedor] ?? 0) + 1;
    }
    return counts;
  }

  List<String> get _vendedoresDisponiveis {
    final vendedores = _contagemPorVendedor.keys.toList()..sort();
    return vendedores;
  }

  List<Map<String, dynamic>> get _itensFiltrados {
    final term = _buscaCtrl.text.trim().toUpperCase();
    return _itens.where((item) {
      final status = _upper(item['status']);
      final vendedor = _vendorName(item);
      final haystack = <String>[
        _text(item['cod_cliente']),
        _text(item['nome_cliente']),
        _text(item['cidade']),
        _vendorName(item),
        _text(item['vendedor_cadastro']),
      ].join(' | ').toUpperCase();

      final matchStatus = _statusFiltro == 'TODAS' || status == _statusFiltro;
      final matchVendedor =
          _vendedorFiltro == 'TODOS' || vendedor == _vendedorFiltro;
      final matchTerm = term.isEmpty || haystack.contains(term);
      return matchStatus && matchVendedor && matchTerm;
    }).toList();
  }

  List<Map<String, dynamic>> get _itensOrigem {
    final staged = _preProgramacaoIds.toSet();
    return _itensFiltrados
        .where((item) => !staged.contains(_text(item['id'])))
        .toList();
  }

  Map<String, List<Map<String, dynamic>>> get _itensOrigemAgrupados {
    final grouped = <String, List<Map<String, dynamic>>>{};
    for (final item in _itensOrigem) {
      final vendedor = _vendorName(item).isEmpty ? 'SEM VENDEDOR' : _vendorName(item);
      grouped.putIfAbsent(vendedor, () => <Map<String, dynamic>>[]).add(item);
    }
    final keys = grouped.keys.toList()..sort();
    return <String, List<Map<String, dynamic>>>{
      for (final key in keys) key: grouped[key]!,
    };
  }

  List<Map<String, dynamic>> get _itensPreProgramacao {
    final byId = <String, Map<String, dynamic>>{
      for (final item in _itens) _text(item['id']): item,
    };
    return _preProgramacaoIds
        .map((id) => byId[id])
        .whereType<Map<String, dynamic>>()
        .toList();
  }

  int get _pendentes =>
      _itens.where((item) => _upper(item['status']) == 'PENDENTE').length;

  int get _finalizadas =>
      _itens.where((item) => _upper(item['status']) == 'FINALIZADA').length;

  int get _prePendentes => _itensPreProgramacao
      .where((item) => _upper(item['status']) == 'PENDENTE')
      .length;

  int get _preFinalizadas => _itensPreProgramacao
      .where((item) => _upper(item['status']) == 'FINALIZADA')
      .length;

  int get _preTotalCaixas => _itensPreProgramacao.fold<int>(
        0,
        (acc, item) => acc + _int(item['caixas']),
      );

  double get _preTotalPreco => _itensPreProgramacao.fold<double>(
        0,
        (acc, item) => acc + _double(item['preco']),
      );

  bool get _preProgramacaoPronta =>
      _itensPreProgramacao.isNotEmpty && _prePendentes == 0;

  String _tituloPadraoPreProgramacao() {
    final base = _vendedorFiltro == 'TODOS' ? 'GERAL' : _vendedorFiltro;
    final data = DateTime.now();
    final dia = data.day.toString().padLeft(2, '0');
    final mes = data.month.toString().padLeft(2, '0');
    final hora = data.hour.toString().padLeft(2, '0');
    final minuto = data.minute.toString().padLeft(2, '0');
    return 'PRE-$base-$dia$mes-$hora$minuto';
  }

  Future<void> _toggleStatus(Map<String, dynamic> item) async {
    final id = _text(item['id']);
    final statusAtual = _upper(item['status']);
    final next = statusAtual == 'FINALIZADA' ? 'PENDENTE' : 'FINALIZADA';
    try {
      await widget.api.atualizarVendaRascunho(id, status: next);
      await reload();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao atualizar status: $error')),
      );
    }
  }

  Future<void> _removerItem(Map<String, dynamic> item) async {
    final id = _text(item['id']);
    try {
      await widget.api.removerVendaRascunho(id);
      if (!mounted) return;
      setState(() {
        _preProgramacaoIds.remove(id);
        _selecionadosOrigem.remove(id);
      });
      await reload();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao excluir venda: $error')),
      );
    }
  }

  Future<void> _editarItem(Map<String, dynamic> item) async {
    final caixasCtrl =
        TextEditingController(text: _text(item['caixas']).isEmpty ? '0' : _text(item['caixas']));
    final precoCtrl = TextEditingController(
      text: _formatMoney(item['preco']),
    );
    final obsCtrl =
        TextEditingController(text: _text(item['observacao']));
    String status = _upper(item['status']).isEmpty ? 'PENDENTE' : _upper(item['status']);
    String? localError;

    final result = await showDialog<bool>(
      context: context,
      builder: (context) {
        return StatefulBuilder(
          builder: (context, setLocal) {
            return AlertDialog(
              title: const Text('Editar venda do rascunho'),
              content: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 520),
                child: SingleChildScrollView(
                  child: Column(
                    mainAxisSize: MainAxisSize.min,
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        '${item['cod_cliente']} - ${item['nome_cliente']}',
                        style: const TextStyle(fontWeight: FontWeight.w800),
                      ),
                      const SizedBox(height: 12),
                      DropdownButtonFormField<String>(
                        initialValue: status,
                        decoration: const InputDecoration(labelText: 'Status'),
                        items: const [
                          DropdownMenuItem(
                            value: 'PENDENTE',
                            child: Text('PENDENTE'),
                          ),
                          DropdownMenuItem(
                            value: 'FINALIZADA',
                            child: Text('FINALIZADA'),
                          ),
                        ],
                        onChanged: (value) => status = value ?? 'PENDENTE',
                      ),
                      const SizedBox(height: 12),
                      TextField(
                        controller: caixasCtrl,
                        keyboardType: TextInputType.number,
                        decoration: const InputDecoration(labelText: 'Caixas'),
                      ),
                      const SizedBox(height: 12),
                      TextField(
                        controller: precoCtrl,
                        keyboardType:
                            const TextInputType.numberWithOptions(decimal: true),
                        decoration: const InputDecoration(labelText: 'Preco'),
                      ),
                      const SizedBox(height: 12),
                      TextField(
                        controller: obsCtrl,
                        maxLines: 2,
                        decoration: const InputDecoration(
                          labelText: 'Observacao',
                        ),
                      ),
                      if (localError != null) ...[
                        const SizedBox(height: 12),
                        Text(
                          localError!,
                          style: const TextStyle(
                            color: VendorUiColors.danger,
                            fontWeight: FontWeight.w700,
                          ),
                        ),
                      ],
                    ],
                  ),
                ),
              ),
              actions: [
                TextButton(
                  onPressed: () => Navigator.of(context).pop(false),
                  child: const Text('Cancelar'),
                ),
                ElevatedButton(
                  onPressed: () async {
                    final caixas = int.tryParse(caixasCtrl.text.trim()) ?? 0;
                    final preco = double.tryParse(
                          precoCtrl.text.trim().replaceAll(',', '.'),
                        ) ??
                        0.0;
                    if (caixas <= 0 || preco <= 0) {
                      setLocal(() {
                        localError =
                            'Informe caixas e preco validos para a venda.';
                      });
                      return;
                    }
                    try {
                      await widget.api.atualizarVendaRascunho(
                        _text(item['id']),
                        caixas: caixas,
                        preco: preco,
                        observacao: obsCtrl.text.trim(),
                        status: status,
                      );
                    } catch (error) {
                      if (!context.mounted) return;
                      setLocal(() {
                        localError = error.toString();
                      });
                      return;
                    }
                    if (!context.mounted) return;
                    Navigator.of(context).pop(true);
                  },
                  child: const Text('Salvar'),
                ),
              ],
            );
          },
        );
      },
    );

    caixasCtrl.dispose();
    precoCtrl.dispose();
    obsCtrl.dispose();

    if (result == true) {
      await reload();
    }
  }

  void _toggleOrigemSelection(Map<String, dynamic> item, bool selected) {
    final id = _text(item['id']);
    setState(() {
      if (selected) {
        _selecionadosOrigem.add(id);
      } else {
        _selecionadosOrigem.remove(id);
      }
    });
  }

  void _adicionarItemAPreProgramacao(Map<String, dynamic> item) {
    final id = _text(item['id']);
    if (id.isEmpty || _preProgramacaoIds.contains(id)) return;
    setState(() {
      _preProgramacaoIds.add(id);
      _selecionadosOrigem.remove(id);
    });
  }

  Future<void> _adicionarSelecionadosAPreProgramacao() async {
    if (_selecionadosOrigem.isEmpty || _movingToPreProgramacao) return;
    setState(() => _movingToPreProgramacao = true);
    try {
      final novosIds = _itensOrigem
          .map((item) => _text(item['id']))
          .where((id) => _selecionadosOrigem.contains(id))
          .where((id) => !_preProgramacaoIds.contains(id))
          .toList();
      if (!mounted) return;
      setState(() {
        _preProgramacaoIds.addAll(novosIds);
        _selecionadosOrigem.clear();
      });
    } finally {
      if (mounted) {
        setState(() => _movingToPreProgramacao = false);
      }
    }
  }

  void _removerDaPreProgramacao(Map<String, dynamic> item) {
    final id = _text(item['id']);
    setState(() {
      _preProgramacaoIds.remove(id);
    });
  }

  void _novaPreProgramacao() {
    setState(() {
      _preProgramacaoAtualId = null;
      _preProgramacaoAtualTitulo = '';
      _preProgramacaoIds.clear();
      _selecionadosOrigem.clear();
    });
  }

  Future<void> _salvarPreProgramacaoAtual() async {
    if (_preProgramacaoIds.isEmpty || _savingPreProgramacao) return;
    final tituloCtrl = TextEditingController(
      text: _preProgramacaoAtualTitulo.isEmpty
          ? _tituloPadraoPreProgramacao()
          : _preProgramacaoAtualTitulo,
    );
    final obsCtrl = TextEditingController();
    String? localError;

    final confirmed = await showDialog<bool>(
      context: context,
      builder: (context) {
        return StatefulBuilder(
          builder: (context, setLocal) {
            return AlertDialog(
              title: const Text('Salvar pre-programacao'),
              content: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 520),
                child: SingleChildScrollView(
                  child: Column(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      TextField(
                        controller: tituloCtrl,
                        decoration: const InputDecoration(
                          labelText: 'Titulo da pre-programacao',
                        ),
                      ),
                      const SizedBox(height: 12),
                      TextField(
                        controller: obsCtrl,
                        maxLines: 2,
                        decoration: const InputDecoration(
                          labelText: 'Observacao (opcional)',
                        ),
                      ),
                      if (localError != null) ...[
                        const SizedBox(height: 12),
                        Text(
                          localError!,
                          style: const TextStyle(
                            color: VendorUiColors.danger,
                            fontWeight: FontWeight.w700,
                          ),
                        ),
                      ],
                    ],
                  ),
                ),
              ),
              actions: [
                TextButton(
                  onPressed: () => Navigator.of(context).pop(false),
                  child: const Text('Cancelar'),
                ),
                ElevatedButton(
                  onPressed: () {
                    if (tituloCtrl.text.trim().isEmpty) {
                      setLocal(() => localError = 'Informe um titulo.');
                      return;
                    }
                    Navigator.of(context).pop(true);
                  },
                  child: const Text('Salvar'),
                ),
              ],
            );
          },
        );
      },
    );

    if (confirmed != true) {
      tituloCtrl.dispose();
      obsCtrl.dispose();
      return;
    }

    setState(() => _savingPreProgramacao = true);
    try {
      final result = await widget.api.salvarPreProgramacao(
        id: _preProgramacaoAtualId,
        itemIds: _preProgramacaoIds,
        titulo: tituloCtrl.text.trim(),
        observacao: obsCtrl.text.trim(),
        status: 'ABERTA',
      );
      if (!mounted) return;
      setState(() {
        _preProgramacaoAtualId = _text(result['id']);
        _preProgramacaoAtualTitulo = tituloCtrl.text.trim();
      });
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            'Pre-programacao ${_preProgramacaoAtualTitulo.isEmpty ? _preProgramacaoAtualId : _preProgramacaoAtualTitulo} salva.',
          ),
        ),
      );
      await reload();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao salvar pre-programacao: $error')),
      );
    } finally {
      tituloCtrl.dispose();
      obsCtrl.dispose();
      if (mounted) {
        setState(() => _savingPreProgramacao = false);
      }
    }
  }

  Future<void> _abrirPreProgramacaoSalva(Map<String, dynamic> pre) async {
    final target = _text(pre['id']);
    if (target.isEmpty) return;
    try {
      final detalhe = await widget.api.detalhePreProgramacao(target);
      final ids = (_asIdList(detalhe['item_ids']));
      if (!mounted) return;
      setState(() {
        _preProgramacaoAtualId = target;
        _preProgramacaoAtualTitulo = _text(pre['titulo']);
        _preProgramacaoIds
          ..clear()
          ..addAll(ids);
        _selecionadosOrigem.clear();
      });
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            'Pre-programacao ${_preProgramacaoAtualTitulo.isEmpty ? target : _preProgramacaoAtualTitulo} carregada.',
          ),
        ),
      );
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao abrir pre-programacao: $error')),
      );
    }
  }

  Future<void> _removerPreProgramacaoSalva(Map<String, dynamic> pre) async {
    final target = _text(pre['id']);
    if (target.isEmpty) return;
    try {
      await widget.api.removerPreProgramacao(target);
      if (!mounted) return;
      if (_preProgramacaoAtualId == target) {
        _novaPreProgramacao();
      }
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            'Pre-programacao ${_text(pre['titulo']).isEmpty ? target : _text(pre['titulo'])} removida.',
          ),
        ),
      );
      await reload();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao remover pre-programacao: $error')),
      );
    }
  }

  List<String> _asIdList(dynamic value) {
    if (value is! List) return <String>[];
    return value
        .map((item) => item.toString().trim())
        .where((item) => item.isNotEmpty)
        .toList();
  }

  void _moverNaPreProgramacao(int index, int delta) {
    final nextIndex = index + delta;
    if (index < 0 || index >= _preProgramacaoIds.length) return;
    if (nextIndex < 0 || nextIndex >= _preProgramacaoIds.length) return;
    setState(() {
      final item = _preProgramacaoIds.removeAt(index);
      _preProgramacaoIds.insert(nextIndex, item);
    });
  }

  Future<void> _abrirProgramacao() async {
    final itens = _itensPreProgramacao;
    if (itens.isEmpty) return;
    if (_prePendentes > 0) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text(
            'Finalize todos os pedidos da pre-programacao antes de abrir a programacao final.',
          ),
        ),
      );
      return;
    }

    final codigo = await Navigator.of(context).push<String>(
      MaterialPageRoute<String>(
        builder: (_) => ProgramacaoCreatePage(
          config: widget.config,
          api: widget.api,
          itens: itens,
        ),
      ),
    );
    if (codigo == null || !mounted) return;
    final preIdConsumida = _preProgramacaoAtualId;
    if (preIdConsumida != null && preIdConsumida.trim().isNotEmpty) {
      try {
        await widget.api.removerPreProgramacao(preIdConsumida);
      } catch (_) {
        // A programacao principal ja foi criada; falha ao limpar a pre-programacao
        // nao deve interromper o fluxo do usuario.
      }
    }
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('Programacao $codigo criada com sucesso.')),
    );
    setState(() {
      _preProgramacaoIds.clear();
      _preProgramacaoAtualId = null;
      _preProgramacaoAtualTitulo = '';
    });
    await reload();
    await widget.onProgramCreated();
  }

  Widget _buildToolbar() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Rascunho do dia',
            subtitle:
                'Os pedidos entram por vendedor e podem ser separados em uma pre-programacao editavel antes da programacao oficial.',
            icon: Icons.inventory_2_outlined,
            trailing: StatusBadge(
              label: '${_itensPreProgramacao.length} na pre-programacao',
              color: _itensPreProgramacao.isEmpty
                  ? Colors.blueGrey.shade700
                  : VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _buscaCtrl,
            onChanged: (_) => setState(() {}),
            decoration: const InputDecoration(
              labelText: 'Buscar por cliente, codigo, vendedor ou cidade',
              prefixIcon: Icon(Icons.search),
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              FilterChip(
                label: const Text('Todas'),
                selected: _statusFiltro == 'TODAS',
                onSelected: (_) => setState(() => _statusFiltro = 'TODAS'),
              ),
              FilterChip(
                label: const Text('Pendentes'),
                selected: _statusFiltro == 'PENDENTE',
                onSelected: (_) => setState(() => _statusFiltro = 'PENDENTE'),
              ),
              FilterChip(
                label: const Text('Finalizadas'),
                selected: _statusFiltro == 'FINALIZADA',
                onSelected: (_) => setState(() => _statusFiltro = 'FINALIZADA'),
              ),
            ],
          ),
          const SizedBox(height: 8),
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            child: Row(
              children: [
                FilterChip(
                  label: const Text('Todos vendedores'),
                  selected: _vendedorFiltro == 'TODOS',
                  onSelected: (_) => setState(() => _vendedorFiltro = 'TODOS'),
                ),
                const SizedBox(width: 8),
                ..._vendedoresDisponiveis.map((vendedor) {
                  final total = _contagemPorVendedor[vendedor] ?? 0;
                  return Padding(
                    padding: const EdgeInsets.only(right: 8),
                    child: FilterChip(
                      label: Text('$vendedor ($total)'),
                      selected: _vendedorFiltro == vendedor,
                      onSelected: (_) =>
                          setState(() => _vendedorFiltro = vendedor),
                    ),
                  );
                }),
              ],
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Pendentes: $_pendentes',
                color: VendorUiColors.warning,
              ),
              StatusBadge(
                label: 'Finalizadas: $_finalizadas',
                color: VendorUiColors.success,
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildPreProgramacaoPanel() {
    final itens = _itensPreProgramacao;
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Pre-programacao em montagem',
            subtitle:
                'Junte aqui pedidos de vendedores diferentes, ajuste os dados finais e so depois abra a programacao definitiva.',
            icon: Icons.table_chart_outlined,
            trailing: StatusBadge(
              label: '${itens.length} item(ns)',
              color: itens.isEmpty
                  ? Colors.blueGrey.shade700
                  : VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              if (_preProgramacaoAtualTitulo.isNotEmpty)
                StatusBadge(
                  label: _preProgramacaoAtualTitulo,
                  color: VendorUiColors.primary,
                ),
              OutlinedButton.icon(
                onPressed: _savingPreProgramacao || itens.isEmpty
                    ? null
                    : _salvarPreProgramacaoAtual,
                icon: _savingPreProgramacao
                    ? const SizedBox(
                        width: 16,
                        height: 16,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    : const Icon(Icons.save_outlined),
                label: const Text('Salvar pre-programacao'),
              ),
              TextButton.icon(
                onPressed: itens.isEmpty && _preProgramacaoAtualId == null
                    ? null
                    : _novaPreProgramacao,
                icon: const Icon(Icons.add_box_outlined),
                label: const Text('Nova'),
              ),
            ],
          ),
          const SizedBox(height: 12),
          if (itens.isEmpty)
            const AppPanel(
              padding: EdgeInsets.all(12),
              backgroundColor: VendorUiColors.surfaceAlt,
              child: Text(
                'Selecione vendas na lista dos vendedores e envie para esta area. Aqui e onde a pre-programacao fica montada antes da programacao final.',
                style: TextStyle(
                  color: VendorUiColors.heading,
                  fontWeight: FontWeight.w700,
                ),
              ),
            )
          else ...[
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: [
                StatusBadge(
                  label: 'Pendentes: $_prePendentes',
                  color: VendorUiColors.warning,
                ),
                StatusBadge(
                  label: 'Finalizadas: $_preFinalizadas',
                  color: VendorUiColors.success,
                ),
                StatusBadge(
                  label: 'Caixas: $_preTotalCaixas',
                  color: Colors.blueGrey.shade700,
                ),
                StatusBadge(
                  label: 'Preco: R\$ ${_preTotalPreco.toStringAsFixed(2).replaceAll('.', ',')}',
                  color: VendorUiColors.primary,
                ),
              ],
            ),
            if (_prePendentes > 0) ...[
              const SizedBox(height: 12),
              AppPanel(
                padding: const EdgeInsets.all(12),
                backgroundColor:
                    VendorUiColors.warning.withValues(alpha: 0.10),
                child: const Text(
                  'A programacao final so abre quando todos os pedidos desta pre-programacao estiverem como FINALIZADA.',
                  style: TextStyle(
                    color: VendorUiColors.warning,
                    fontWeight: FontWeight.w800,
                  ),
                ),
              ),
            ],
            const SizedBox(height: 12),
            ...itens.asMap().entries.map(
                  (entry) => _buildPreProgramacaoItemCard(
                    entry.value,
                    entry.key,
                    itens.length,
                  ),
                ),
          ],
          const SizedBox(height: 12),
          ElevatedButton.icon(
            onPressed: _preProgramacaoPronta ? _abrirProgramacao : null,
            icon: const Icon(Icons.route),
            label: const Text('Enviar para programacao'),
          ),
        ],
      ),
    );
  }

  Widget _buildPreProgramacoesSalvasPanel() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Pre-programacoes salvas',
            subtitle:
                'Montagens compartilhadas que podem ser abertas de novo por qualquer vendedor.',
            icon: Icons.folder_open,
          ),
          const SizedBox(height: 12),
          if (_preProgramacoes.isEmpty)
            const AppPanel(
              padding: EdgeInsets.all(12),
              backgroundColor: VendorUiColors.surfaceAlt,
              child: Text(
                'Nenhuma pre-programacao salva ainda.',
                style: TextStyle(
                  color: VendorUiColors.heading,
                  fontWeight: FontWeight.w700,
                ),
              ),
            )
          else
            ..._preProgramacoes.map((pre) {
              final isAtual = _preProgramacaoAtualId == _text(pre['id']);
              final vendedores = (pre['vendedores'] is List)
                  ? (pre['vendedores'] as List)
                      .map((item) => item.toString().trim())
                      .where((item) => item.isNotEmpty)
                      .join(', ')
                  : '';
              return AppPanel(
                margin: const EdgeInsets.only(bottom: 10),
                padding: const EdgeInsets.all(12),
                backgroundColor: isAtual
                    ? VendorUiColors.primary.withValues(alpha: 0.06)
                    : VendorUiColors.surfaceAlt,
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                _text(pre['titulo']).isEmpty
                                    ? _text(pre['id'])
                                    : _text(pre['titulo']),
                                style: const TextStyle(
                                  color: VendorUiColors.heading,
                                  fontWeight: FontWeight.w900,
                                ),
                              ),
                              const SizedBox(height: 4),
                              Text(
                                [
                                  'Itens: ${pre['itens_total'] ?? 0}',
                                  'Pendentes: ${pre['itens_pendentes'] ?? 0}',
                                  'Finalizadas: ${pre['itens_finalizadas'] ?? 0}',
                                ].join(' | '),
                                style: const TextStyle(
                                  color: VendorUiColors.muted,
                                  fontWeight: FontWeight.w700,
                                  fontSize: 12,
                                ),
                              ),
                              if (vendedores.isNotEmpty) ...[
                                const SizedBox(height: 4),
                                Text(
                                  'Vendedores: $vendedores',
                                  style: const TextStyle(
                                    color: VendorUiColors.muted,
                                    fontWeight: FontWeight.w700,
                                    fontSize: 12,
                                  ),
                                ),
                              ],
                            ],
                          ),
                        ),
                        if (isAtual)
                          const StatusBadge(
                            label: 'ATUAL',
                            color: VendorUiColors.primary,
                          ),
                      ],
                    ),
                    const SizedBox(height: 8),
                    Wrap(
                      spacing: 4,
                      runSpacing: 4,
                      children: [
                        TextButton.icon(
                          onPressed: () => _abrirPreProgramacaoSalva(pre),
                          icon: const Icon(Icons.folder_shared_outlined),
                          label: const Text('Abrir'),
                        ),
                        TextButton.icon(
                          onPressed: () => _removerPreProgramacaoSalva(pre),
                          icon: const Icon(Icons.delete_outline),
                          label: const Text('Excluir'),
                        ),
                      ],
                    ),
                  ],
                ),
              );
            }),
        ],
      ),
    );
  }

  Widget _buildOrigemPanel() {
    final itens = _itensOrigem;
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: _vendedorFiltro == 'TODOS'
                ? 'Pedidos de todos os vendedores'
                : 'Pedidos de $_vendedorFiltro',
            subtitle:
                'Escolha os pedidos que vao compor a proxima pre-programacao. Eles podem ser de qualquer vendedor.',
            icon: Icons.people_alt_outlined,
            trailing: StatusBadge(
              label: '${_selecionadosOrigem.length} selecionado(s)',
              color: _selecionadosOrigem.isEmpty
                  ? Colors.blueGrey.shade700
                  : VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 12),
          ElevatedButton.icon(
            onPressed: _selecionadosOrigem.isEmpty || _movingToPreProgramacao
                ? null
                : _adicionarSelecionadosAPreProgramacao,
            icon: _movingToPreProgramacao
                ? const SizedBox(
                    width: 16,
                    height: 16,
                    child: CircularProgressIndicator(strokeWidth: 2),
                  )
                : const Icon(Icons.playlist_add),
            label: const Text('Levar selecionadas para pre-programacao'),
          ),
          const SizedBox(height: 12),
          if (itens.isEmpty)
            const AppPanel(
              padding: EdgeInsets.all(12),
              backgroundColor: VendorUiColors.surfaceAlt,
              child: Text(
                'Nenhuma venda disponivel para os filtros atuais. Ajuste a busca, troque o vendedor ou finalize mais pedidos.',
                style: TextStyle(
                  color: VendorUiColors.heading,
                  fontWeight: FontWeight.w700,
                ),
              ),
            )
          else
            ..._itensOrigemAgrupados.entries.map(
              (entry) => _buildVendedorSection(
                entry.key,
                entry.value,
              ),
            ),
        ],
      ),
    );
  }

  Widget _buildVendedorSection(
    String vendedor,
    List<Map<String, dynamic>> itens,
  ) {
    final pendentes =
        itens.where((item) => _upper(item['status']) == 'PENDENTE').length;
    final finalizadas =
        itens.where((item) => _upper(item['status']) == 'FINALIZADA').length;

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      padding: EdgeInsets.zero,
      child: ExpansionTile(
        initiallyExpanded: _vendedorFiltro != 'TODOS' || itens.length <= 3,
        tilePadding: const EdgeInsets.fromLTRB(16, 12, 16, 12),
        childrenPadding: const EdgeInsets.fromLTRB(16, 0, 16, 12),
        title: Text(
          vendedor,
          style: const TextStyle(
            color: VendorUiColors.heading,
            fontWeight: FontWeight.w900,
          ),
        ),
        subtitle: Text(
          '${itens.length} pedido(s) | Pendentes: $pendentes | Finalizadas: $finalizadas',
          style: const TextStyle(
            color: VendorUiColors.muted,
            fontWeight: FontWeight.w700,
            fontSize: 12,
          ),
        ),
        children: itens.map(_buildOrigemItemCard).toList(),
      ),
    );
  }

  Widget _buildOrigemItemCard(Map<String, dynamic> item) {
    final id = _text(item['id']);
    final status = _upper(item['status']);
    final selected = _selecionadosOrigem.contains(id);
    final warningCode = _text(item['alerta_codigo_programacao']);

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      backgroundColor:
          selected ? VendorUiColors.primary.withValues(alpha: 0.06) : Colors.white,
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Checkbox(
                value: selected,
                onChanged: (value) => _toggleOrigemSelection(item, value ?? false),
              ),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      '${item['cod_cliente']} - ${item['nome_cliente']}',
                      style: const TextStyle(
                        color: VendorUiColors.heading,
                        fontWeight: FontWeight.w800,
                      ),
                    ),
                    const SizedBox(height: 4),
                    Text(
                      [
                        _text(item['cidade']),
                        _text(item['bairro']),
                        'Vendedor: ${_vendorName(item)}',
                      ].where((part) => part.isNotEmpty).join(' | '),
                      style: const TextStyle(
                        color: VendorUiColors.muted,
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                  ],
                ),
              ),
              StatusBadge(
                label: status,
                color: status == 'FINALIZADA'
                    ? VendorUiColors.success
                    : VendorUiColors.warning,
              ),
            ],
          ),
          const SizedBox(height: 10),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Caixas: ${item['caixas'] ?? 0}',
                color: Colors.blueGrey.shade700,
              ),
              StatusBadge(
                label: 'Preco: R\$ ${_formatMoney(item['preco'])}',
                color: VendorUiColors.primary,
              ),
            ],
          ),
          if (_text(item['observacao']).isNotEmpty) ...[
            const SizedBox(height: 8),
            Text(
              'Obs: ${item['observacao']}',
              style: const TextStyle(fontWeight: FontWeight.w700),
            ),
          ],
          if (warningCode.isNotEmpty) ...[
            const SizedBox(height: 8),
            Text(
              'Aviso: cliente em ${item['alerta_status_rota']} na programacao $warningCode.',
              style: const TextStyle(
                color: VendorUiColors.warning,
                fontWeight: FontWeight.w700,
              ),
            ),
          ],
          const SizedBox(height: 10),
          Wrap(
            spacing: 4,
            runSpacing: 4,
            children: [
              TextButton.icon(
                onPressed: () => _adicionarItemAPreProgramacao(item),
                icon: const Icon(Icons.arrow_upward),
                label: const Text('Levar para pre-programacao'),
              ),
              TextButton.icon(
                onPressed: () => _toggleStatus(item),
                icon: Icon(
                  status == 'FINALIZADA'
                      ? Icons.rotate_left_outlined
                      : Icons.task_alt,
                ),
                label: Text(
                  status == 'FINALIZADA'
                      ? 'Voltar para pendente'
                      : 'Finalizar',
                ),
              ),
              TextButton.icon(
                onPressed: () => _editarItem(item),
                icon: const Icon(Icons.edit_outlined),
                label: const Text('Editar'),
              ),
              TextButton.icon(
                onPressed: () => _removerItem(item),
                icon: const Icon(Icons.delete_outline),
                label: const Text('Excluir'),
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildPreProgramacaoItemCard(
    Map<String, dynamic> item,
    int index,
    int total,
  ) {
    final status = _upper(item['status']);

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      backgroundColor: VendorUiColors.primary.withValues(alpha: 0.04),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      '${item['cod_cliente']} - ${item['nome_cliente']}',
                      style: const TextStyle(
                        color: VendorUiColors.heading,
                        fontWeight: FontWeight.w800,
                      ),
                    ),
                    const SizedBox(height: 4),
                    Text(
                      [
                        'Vendedor: ${_vendorName(item)}',
                        _text(item['cidade']),
                        _text(item['bairro']),
                      ].where((part) => part.isNotEmpty).join(' | '),
                      style: const TextStyle(
                        color: VendorUiColors.muted,
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                  ],
                ),
              ),
              StatusBadge(
                label: status,
                color: status == 'FINALIZADA'
                    ? VendorUiColors.success
                    : VendorUiColors.warning,
              ),
            ],
          ),
          const SizedBox(height: 10),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Ordem ${index + 1}',
                color: VendorUiColors.primary,
              ),
              StatusBadge(
                label: 'Caixas: ${item['caixas'] ?? 0}',
                color: Colors.blueGrey.shade700,
              ),
              StatusBadge(
                label: 'Preco: R\$ ${_formatMoney(item['preco'])}',
                color: VendorUiColors.primary,
              ),
            ],
          ),
          if (_text(item['observacao']).isNotEmpty) ...[
            const SizedBox(height: 8),
            Text(
              'Obs: ${item['observacao']}',
              style: const TextStyle(fontWeight: FontWeight.w700),
            ),
          ],
          const SizedBox(height: 10),
          Wrap(
            spacing: 4,
            runSpacing: 4,
            children: [
              TextButton.icon(
                onPressed: index == 0 ? null : () => _moverNaPreProgramacao(index, -1),
                icon: const Icon(Icons.arrow_upward),
                label: const Text('Subir'),
              ),
              TextButton.icon(
                onPressed:
                    index == total - 1 ? null : () => _moverNaPreProgramacao(index, 1),
                icon: const Icon(Icons.arrow_downward),
                label: const Text('Descer'),
              ),
              TextButton.icon(
                onPressed: () => _toggleStatus(item),
                icon: Icon(
                  status == 'FINALIZADA'
                      ? Icons.rotate_left_outlined
                      : Icons.task_alt,
                ),
                label: Text(
                  status == 'FINALIZADA'
                      ? 'Voltar para pendente'
                      : 'Finalizar',
                ),
              ),
              TextButton.icon(
                onPressed: () => _editarItem(item),
                icon: const Icon(Icons.edit_outlined),
                label: const Text('Editar'),
              ),
              TextButton.icon(
                onPressed: () => _removerDaPreProgramacao(item),
                icon: const Icon(Icons.undo_outlined),
                label: const Text('Tirar da pre-programacao'),
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildScrollableContent() {
    return RefreshIndicator(
      onRefresh: reload,
      child: ListView(
        padding: const EdgeInsets.fromLTRB(16, 8, 16, 24),
        children: [
          _buildToolbar(),
          const SizedBox(height: 12),
          _buildPreProgramacoesSalvasPanel(),
          const SizedBox(height: 12),
          _buildPreProgramacaoPanel(),
          const SizedBox(height: 12),
          _buildOrigemPanel(),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    if (_loading) {
      return const Center(child: CircularProgressIndicator());
    }
    if (_error != null) {
      return ListView(
        padding: const EdgeInsets.all(16),
        children: [
          AppPanel(
            child: Text(
              _error!,
              style: const TextStyle(
                color: VendorUiColors.danger,
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
        ],
      );
    }
    return _buildScrollableContent();
  }
}
