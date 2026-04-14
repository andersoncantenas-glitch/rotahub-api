import 'package:flutter/material.dart';

import '../models/app_config.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';

class ProgramacaoCreatePage extends StatefulWidget {
  const ProgramacaoCreatePage({
    super.key,
    required this.config,
    required this.api,
    required this.itens,
  });

  final AppConfig config;
  final VendedorApiService api;
  final List<Map<String, dynamic>> itens;

  @override
  State<ProgramacaoCreatePage> createState() => _ProgramacaoCreatePageState();
}

class _ProgramacaoCreatePageState extends State<ProgramacaoCreatePage> {
  final _formKey = GlobalKey<FormState>();
  final TextEditingController _localCarregamentoCtrl = TextEditingController();
  final TextEditingController _adiantamentoCtrl =
      TextEditingController(text: '0');
  final TextEditingController _estimativaCtrl = TextEditingController();
  final TextEditingController _obsCtrl = TextEditingController();

  DateTime _dataSelecionada = DateTime.now();
  bool _loadingRefs = true;
  bool _saving = false;
  List<Map<String, dynamic>> _motoristas = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _veiculos = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _ajudantes = <Map<String, dynamic>>[];
  late List<Map<String, dynamic>> _itensProgramacao;
  final Set<int> _ajudantesSelecionados = <int>{};
  Map<String, dynamic>? _motoristaSelecionado;
  Map<String, dynamic>? _veiculoSelecionado;
  String _localRota = 'SERRA';
  String _tipoOperacao = 'FOB';

  String _text(dynamic value) => (value ?? '').toString().trim();

  int _int(dynamic value) {
    if (value is int) return value;
    return int.tryParse(_text(value)) ?? 0;
  }

  double _double(dynamic value) {
    if (value is double) return value;
    if (value is int) return value.toDouble();
    return double.tryParse(_text(value).replaceAll(',', '.')) ?? 0.0;
  }

  String _vendorName(Map<String, dynamic> item) {
    final origem = _text(item['vendedor_origem']).toUpperCase();
    if (origem.isNotEmpty) return origem;
    return _text(item['vendedor_cadastro']).toUpperCase();
  }

  int get _totalCaixas {
    return _itensProgramacao.fold<int>(
      0,
      (acc, item) => acc + _int(item['caixas']),
    );
  }

  double get _totalPreco {
    return _itensProgramacao.fold<double>(
      0,
      (acc, item) => acc + _double(item['preco']),
    );
  }

  Map<String, Map<String, num>> get _resumoPorVendedor {
    final resumo = <String, Map<String, num>>{};
    for (final item in _itensProgramacao) {
      final vendedor = _vendorName(item).isEmpty ? 'SEM VENDEDOR' : _vendorName(item);
      final linha = resumo.putIfAbsent(
        vendedor,
        () => <String, num>{'clientes': 0, 'caixas': 0, 'preco': 0},
      );
      linha['clientes'] = (linha['clientes'] ?? 0) + 1;
      linha['caixas'] = (linha['caixas'] ?? 0) + _int(item['caixas']);
      linha['preco'] = (linha['preco'] ?? 0) + _double(item['preco']);
    }
    return resumo;
  }

  @override
  void initState() {
    super.initState();
    _itensProgramacao = List<Map<String, dynamic>>.from(widget.itens);
    _estimativaCtrl.text = _totalCaixas.toString();
    _loadReferences();
  }

  @override
  void dispose() {
    _localCarregamentoCtrl.dispose();
    _adiantamentoCtrl.dispose();
    _estimativaCtrl.dispose();
    _obsCtrl.dispose();
    super.dispose();
  }

  String _formatDate(DateTime value) {
    final month = value.month.toString().padLeft(2, '0');
    final day = value.day.toString().padLeft(2, '0');
    return '${value.year}-$month-$day';
  }

  Future<void> _loadReferences() async {
    setState(() => _loadingRefs = true);
    try {
      final motoristas = await widget.api.listarMotoristas();
      final veiculos = await widget.api.listarVeiculos();
      final ajudantes = await widget.api.listarAjudantes();
      if (!mounted) return;
      setState(() {
        _motoristas = motoristas;
        _veiculos = veiculos;
        _ajudantes = ajudantes;
        _motoristaSelecionado = motoristas.isNotEmpty ? motoristas.first : null;
        _veiculoSelecionado = veiculos.isNotEmpty ? veiculos.first : null;
        _loadingRefs = false;
      });
    } catch (error) {
      if (!mounted) return;
      setState(() => _loadingRefs = false);
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao carregar referencias: $error')),
      );
    }
  }

  Future<void> _pickDate() async {
    final picked = await showDatePicker(
      context: context,
      initialDate: _dataSelecionada,
      firstDate: DateTime.now().subtract(const Duration(days: 7)),
      lastDate: DateTime.now().add(const Duration(days: 365)),
    );
    if (picked == null) return;
    setState(() => _dataSelecionada = picked);
  }

  void _toggleAjudante(Map<String, dynamic> ajudante, bool selected) {
    final id = (ajudante['id'] as num?)?.toInt() ?? 0;
    setState(() {
      if (selected) {
        if (_ajudantesSelecionados.length < 2) {
          _ajudantesSelecionados.add(id);
        }
      } else {
        _ajudantesSelecionados.remove(id);
      }
    });
  }

  void _moverItem(int index, int delta) {
    final nextIndex = index + delta;
    if (index < 0 || index >= _itensProgramacao.length) return;
    if (nextIndex < 0 || nextIndex >= _itensProgramacao.length) return;
    setState(() {
      final item = _itensProgramacao.removeAt(index);
      _itensProgramacao.insert(nextIndex, item);
      if (_tipoOperacao == 'FOB') {
        _estimativaCtrl.text = _totalCaixas.toString();
      }
    });
  }

  Future<void> _submit() async {
    if (!_formKey.currentState!.validate()) return;
    if (_motoristaSelecionado == null || _veiculoSelecionado == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Selecione motorista e veiculo.')),
      );
      return;
    }
    if (_ajudantesSelecionados.length != 2) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Selecione exatamente 2 ajudantes.')),
      );
      return;
    }

    final estimativa = double.tryParse(
          _estimativaCtrl.text.trim().replaceAll(',', '.'),
        ) ??
        0.0;
    if (_tipoOperacao == 'FOB' && estimativa <= 0) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Informe as caixas estimadas para FOB.')),
      );
      return;
    }
    if (_tipoOperacao == 'CIF' && estimativa <= 0) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Informe o KG estimado para CIF.')),
      );
      return;
    }

    setState(() => _saving = true);
    try {
      final codigo = await widget.api.gerarCodigoProgramacao();
      final ajudantes = _ajudantes
          .where(
            (item) => _ajudantesSelecionados.contains((item['id'] as num?)?.toInt() ?? 0),
          )
          .toList();
      final adiantamento = double.tryParse(
            _adiantamentoCtrl.text.trim().replaceAll(',', '.'),
          ) ??
          0.0;

      await widget.api.criarProgramacaoOficial(
        codigoProgramacao: codigo,
        dataProgramada: _dataSelecionada,
        motorista: _motoristaSelecionado!,
        veiculo: _veiculoSelecionado!,
        ajudantes: ajudantes,
        localRota: _localRota,
        localCarregamento: _localCarregamentoCtrl.text.trim(),
        tipoOperacao: _tipoOperacao,
        adiantamento: adiantamento,
        kgEstimado: _tipoOperacao == 'CIF' ? estimativa : 0.0,
        caixasEstimado: _tipoOperacao == 'FOB' ? estimativa.round() : 0,
        itens: _itensProgramacao,
        usuarioCriacao: widget.config.vendedorPadrao,
        observacao: _obsCtrl.text.trim(),
      );
      final ids = _itensProgramacao
          .map((item) => (item['id'] ?? '').toString())
          .where((item) => item.trim().isNotEmpty)
          .toList();
      await widget.api.removerVendasRascunhoPorIds(ids);

      if (!mounted) return;
      Navigator.of(context).pop<String>(codigo);
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao gerar programacao: $error')),
      );
    } finally {
      if (mounted) {
        setState(() => _saving = false);
      }
    }
  }

  Widget _buildResumo() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Resumo da selecao',
            subtitle:
                'Itens finalizados do rascunho que vao compor a programacao oficial.',
            icon: Icons.fact_check_outlined,
            trailing: StatusBadge(
              label: '${_itensProgramacao.length} itens',
              color: VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 14),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Caixas: $_totalCaixas',
                color: Colors.blueGrey.shade700,
              ),
              StatusBadge(
                label: 'Preco: R\$ ${_totalPreco.toStringAsFixed(2).replaceAll('.', ',')}',
                color: VendorUiColors.success,
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildResumoPorVendedor() {
    final entries = _resumoPorVendedor.entries.toList()
      ..sort((a, b) => a.key.compareTo(b.key));
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Composicao por vendedor',
            subtitle:
                'Resumo dos pedidos separados por vendedor antes da programacao definitiva.',
            icon: Icons.groups,
          ),
          const SizedBox(height: 12),
          ...entries.map((entry) {
            final dados = entry.value;
            return AppPanel(
              margin: const EdgeInsets.only(bottom: 10),
              padding: const EdgeInsets.all(12),
              backgroundColor: VendorUiColors.surfaceAlt,
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    entry.key,
                    style: const TextStyle(
                      color: VendorUiColors.heading,
                      fontWeight: FontWeight.w900,
                    ),
                  ),
                  const SizedBox(height: 8),
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      StatusBadge(
                        label: 'Clientes: ${dados['clientes']?.toInt() ?? 0}',
                        color: VendorUiColors.primary,
                      ),
                      StatusBadge(
                        label: 'Caixas: ${dados['caixas']?.toInt() ?? 0}',
                        color: Colors.blueGrey.shade700,
                      ),
                      StatusBadge(
                        label:
                            'Preco: R\$ ${((dados['preco'] ?? 0).toDouble()).toStringAsFixed(2).replaceAll('.', ',')}',
                        color: VendorUiColors.success,
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

  Widget _buildForm() {
    return AppPanel(
      child: Form(
        key: _formKey,
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const PanelHeader(
              title: 'Programacao definitiva',
              subtitle:
                  'Preencha os dados finais da rota no mesmo fluxo operacional online do desktop.',
              icon: Icons.route,
            ),
            const SizedBox(height: 16),
            ListTile(
              contentPadding: EdgeInsets.zero,
              title: const Text('Data da programacao'),
              subtitle: Text(_formatDate(_dataSelecionada)),
              trailing: OutlinedButton(
                onPressed: _pickDate,
                child: const Text('Alterar'),
              ),
            ),
            const SizedBox(height: 8),
            DropdownButtonFormField<Map<String, dynamic>>(
              initialValue: _motoristaSelecionado,
              decoration: const InputDecoration(labelText: 'Motorista'),
              items: _motoristas
                  .map(
                    (item) => DropdownMenuItem<Map<String, dynamic>>(
                      value: item,
                      child: Text('${item['codigo']} - ${item['nome']}'),
                    ),
                  )
                  .toList(),
              onChanged: (value) => setState(() => _motoristaSelecionado = value),
            ),
            const SizedBox(height: 16),
            DropdownButtonFormField<Map<String, dynamic>>(
              initialValue: _veiculoSelecionado,
              decoration: const InputDecoration(labelText: 'Veiculo'),
              items: _veiculos
                  .map(
                    (item) => DropdownMenuItem<Map<String, dynamic>>(
                      value: item,
                      child: Text(
                        '${item['placa']} ${item['modelo'] ?? ''}'.trim(),
                      ),
                    ),
                  )
                  .toList(),
              onChanged: (value) => setState(() => _veiculoSelecionado = value),
            ),
            const SizedBox(height: 16),
            Text('Ajudantes', style: Theme.of(context).textTheme.titleMedium),
            const SizedBox(height: 8),
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: _ajudantes.map((ajudante) {
                final selected = _ajudantesSelecionados
                    .contains((ajudante['id'] as num?)?.toInt() ?? 0);
                return FilterChip(
                  label: Text((ajudante['nome'] ?? '-').toString()),
                  selected: selected,
                  onSelected: (value) => _toggleAjudante(ajudante, value),
                );
              }).toList(),
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                Expanded(
                  child: DropdownButtonFormField<String>(
                    initialValue: _localRota,
                    decoration: const InputDecoration(labelText: 'Local da rota'),
                    items: const [
                      DropdownMenuItem(value: 'SERRA', child: Text('SERRA')),
                      DropdownMenuItem(value: 'SERTAO', child: Text('SERTAO')),
                    ],
                    onChanged: (value) =>
                        setState(() => _localRota = value ?? 'SERRA'),
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: DropdownButtonFormField<String>(
                    initialValue: _tipoOperacao,
                    decoration: const InputDecoration(labelText: 'Tipo de carregamento'),
                    items: const [
                      DropdownMenuItem(value: 'FOB', child: Text('FOB')),
                      DropdownMenuItem(value: 'CIF', child: Text('CIF')),
                    ],
                    onChanged: (value) {
                      setState(() {
                        _tipoOperacao = value ?? 'FOB';
                        _estimativaCtrl.text = _tipoOperacao == 'FOB'
                            ? _totalCaixas.toString()
                            : '';
                      });
                    },
                  ),
                ),
              ],
            ),
            const SizedBox(height: 16),
            TextFormField(
              controller: _localCarregamentoCtrl,
              decoration: const InputDecoration(
                labelText: 'Local de carregamento',
              ),
              validator: (value) {
                if ((value ?? '').trim().isEmpty) {
                  return 'Informe o local de carregamento.';
                }
                return null;
              },
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                Expanded(
                  child: TextFormField(
                    controller: _adiantamentoCtrl,
                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                    decoration: const InputDecoration(
                      labelText: 'Adiantamento',
                    ),
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: TextFormField(
                    controller: _estimativaCtrl,
                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                    decoration: InputDecoration(
                      labelText: _tipoOperacao == 'FOB'
                          ? 'Caixas estimadas'
                          : 'KG estimado',
                    ),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 16),
            TextFormField(
              controller: _obsCtrl,
              maxLines: 2,
              decoration: const InputDecoration(
                labelText: 'Observacao geral (opcional)',
              ),
            ),
            const SizedBox(height: 20),
            ElevatedButton.icon(
              onPressed: _saving ? null : _submit,
              icon: _saving
                  ? const SizedBox(
                      width: 16,
                      height: 16,
                      child: CircularProgressIndicator(strokeWidth: 2),
                    )
                  : const Icon(Icons.save),
              label: const Text('Gerar programacao'),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildItens() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Itens selecionados',
            subtitle:
                'Clientes que serao enviados para a programacao oficial e para o app do motorista, respeitando a ordem da pre-programacao.',
            icon: Icons.list_alt,
          ),
          const SizedBox(height: 12),
          ..._itensProgramacao.asMap().entries.map((entry) {
            final index = entry.key;
            final item = entry.value;
            return AppPanel(
              margin: const EdgeInsets.only(bottom: 10),
              padding: const EdgeInsets.all(12),
              backgroundColor: VendorUiColors.surfaceAlt,
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      StatusBadge(
                        label: 'Ordem ${index + 1}',
                        color: VendorUiColors.primary,
                      ),
                      TextButton.icon(
                        onPressed: index == 0 ? null : () => _moverItem(index, -1),
                        icon: const Icon(Icons.arrow_upward),
                        label: const Text('Subir'),
                      ),
                      TextButton.icon(
                        onPressed: index == _itensProgramacao.length - 1
                            ? null
                            : () => _moverItem(index, 1),
                        icon: const Icon(Icons.arrow_downward),
                        label: const Text('Descer'),
                      ),
                    ],
                  ),
                  const SizedBox(height: 8),
                  Text(
                    '${item['cod_cliente']} - ${item['nome_cliente']}',
                    style: const TextStyle(
                      color: VendorUiColors.heading,
                      fontWeight: FontWeight.w800,
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    'Origem: ${_vendorName(item).isEmpty ? '-' : _vendorName(item)} | Caixas: ${_int(item['caixas'])} | Preco: R\$ ${_double(item['preco']).toStringAsFixed(2).replaceAll('.', ',')}',
                    style: const TextStyle(
                      color: VendorUiColors.muted,
                      fontWeight: FontWeight.w700,
                    ),
                  ),
                ],
              ),
            );
          }),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Nova programacao'),
      ),
      body: _loadingRefs
          ? const Center(child: CircularProgressIndicator())
          : RefreshIndicator(
              onRefresh: _loadReferences,
              child: ListView(
                padding: const EdgeInsets.fromLTRB(16, 8, 16, 24),
                children: [
                  _buildResumo(),
                  const SizedBox(height: 12),
                  _buildResumoPorVendedor(),
                  const SizedBox(height: 12),
                  _buildForm(),
                  const SizedBox(height: 12),
                  _buildItens(),
                ],
              ),
            ),
    );
  }
}
