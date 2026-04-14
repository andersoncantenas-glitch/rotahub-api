import 'package:flutter/material.dart';

import '../models/app_config.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';

class ClientesPage extends StatefulWidget {
  const ClientesPage({
    super.key,
    required this.config,
    required this.api,
    required this.onDraftChanged,
  });

  final AppConfig config;
  final VendedorApiService api;
  final Future<void> Function() onDraftChanged;

  @override
  State<ClientesPage> createState() => ClientesPageState();
}

class ClientesPageState extends State<ClientesPage> {
  final TextEditingController _buscaCtrl = TextEditingController();
  final TextEditingController _vendedorCtrl = TextEditingController();
  final TextEditingController _cidadeCtrl = TextEditingController();
  bool _loading = true;
  bool _sending = false;
  String? _error;
  List<Map<String, dynamic>> _clientes = <Map<String, dynamic>>[];
  final Set<String> _selecionados = <String>{};

  @override
  void initState() {
    super.initState();
    reload();
  }

  @override
  void dispose() {
    _buscaCtrl.dispose();
    _vendedorCtrl.dispose();
    _cidadeCtrl.dispose();
    super.dispose();
  }

  void _limparFiltros() {
    _buscaCtrl.clear();
    _vendedorCtrl.clear();
    _cidadeCtrl.clear();
  }

  Future<void> reload() async {
    setState(() {
      _loading = true;
      _error = null;
    });

    try {
      final data = await widget.api.buscarClientes(
        _buscaCtrl.text,
        vendedor: _vendedorCtrl.text,
        cidade: _cidadeCtrl.text,
        limit: 1000,
      );
      if (!mounted) return;
      setState(() {
        _clientes = data;
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

  void _toggleCliente(String codCliente, bool selected) {
    setState(() {
      if (selected) {
        _selecionados.add(codCliente);
      } else {
        _selecionados.remove(codCliente);
      }
    });
  }

  Future<_DraftInputResult?> _showDraftDialog(
    List<Map<String, dynamic>> clientes,
    Map<String, Map<String, dynamic>> alertas,
  ) async {
    final caixasCtrl = TextEditingController(text: '1');
    final precoCtrl = TextEditingController();
    final obsCtrl = TextEditingController();
    String? localError;

    final result = await showDialog<_DraftInputResult>(
      context: context,
      builder: (context) {
        return StatefulBuilder(
          builder: (context, setLocal) {
            return AlertDialog(
              title: const Text('Enviar para rascunho'),
              content: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 520),
                child: SingleChildScrollView(
                  child: Column(
                    mainAxisSize: MainAxisSize.min,
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        '${clientes.length} cliente(s) selecionado(s). Defina caixas e preco para enviar ao rascunho.',
                        style: const TextStyle(fontWeight: FontWeight.w700),
                      ),
                      const SizedBox(height: 16),
                      TextField(
                        controller: caixasCtrl,
                        keyboardType: TextInputType.number,
                        decoration: const InputDecoration(
                          labelText: 'Quantidade de caixas',
                        ),
                      ),
                      const SizedBox(height: 12),
                      TextField(
                        controller: precoCtrl,
                        keyboardType: const TextInputType.numberWithOptions(decimal: true),
                        decoration: const InputDecoration(
                          labelText: 'Preco',
                          hintText: 'Ex.: 150,00',
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
                      if (alertas.isNotEmpty) ...[
                        const SizedBox(height: 16),
                        AppPanel(
                          padding: const EdgeInsets.all(12),
                          backgroundColor:
                              VendorUiColors.warning.withValues(alpha: 0.10),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              const Text(
                                'Aviso: ha clientes em programacoes EM_ENTREGAS. O envio segue liberado.',
                                style: TextStyle(
                                  color: VendorUiColors.warning,
                                  fontWeight: FontWeight.w800,
                                ),
                              ),
                              const SizedBox(height: 8),
                              ...alertas.entries.map((entry) {
                                final info = entry.value;
                                return Padding(
                                  padding: const EdgeInsets.only(bottom: 6),
                                  child: Text(
                                    '${entry.key}: ${info['codigo_programacao'] ?? '-'} ${info['motorista'] ?? ''}'.trim(),
                                    style: const TextStyle(
                                      color: VendorUiColors.warning,
                                      fontWeight: FontWeight.w700,
                                    ),
                                  ),
                                );
                              }),
                            ],
                          ),
                        ),
                      ],
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
                  onPressed: () => Navigator.of(context).pop(),
                  child: const Text('Cancelar'),
                ),
                ElevatedButton(
                  onPressed: () {
                    final caixas = int.tryParse(caixasCtrl.text.trim()) ?? 0;
                    final preco = double.tryParse(
                          precoCtrl.text.trim().replaceAll(',', '.'),
                        ) ??
                        0.0;
                    if (caixas <= 0) {
                      setLocal(() => localError = 'Informe a quantidade de caixas.');
                      return;
                    }
                    if (preco <= 0) {
                      setLocal(() => localError = 'Informe o preco da venda.');
                      return;
                    }
                    Navigator.of(context).pop(
                      _DraftInputResult(
                        caixas: caixas,
                        preco: preco,
                        observacao: obsCtrl.text.trim(),
                      ),
                    );
                  },
                  child: const Text('Enviar'),
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
    return result;
  }

  Future<void> _enviarParaRascunho() async {
    if (_selecionados.isEmpty || _sending) return;
    final clientes = _clientes
        .where((item) => _selecionados.contains((item['cod_cliente'] ?? '').toString().trim().toUpperCase()))
        .toList();
    if (clientes.isEmpty) return;

    setState(() => _sending = true);
    Map<String, Map<String, dynamic>> alertas = <String, Map<String, dynamic>>{};
    try {
      final cods = clientes.map((item) => (item['cod_cliente'] ?? '').toString()).toList();
      alertas = await widget.api.carregarAlertasClientesEmEntrega(cods);
    } catch (_) {
      alertas = <String, Map<String, dynamic>>{};
    } finally {
      if (mounted) {
        setState(() => _sending = false);
      }
    }

    if (!mounted) return;
    final result = await _showDraftDialog(clientes, alertas);
    if (result == null) return;

    setState(() => _sending = true);
    try {
      await widget.api.adicionarVendasAoRascunho(
        clientes: clientes,
        vendedorOrigem: widget.config.vendedorPadrao,
        caixas: result.caixas,
        preco: result.preco,
        observacao: result.observacao,
        alertas: alertas,
      );
      if (!mounted) return;
      setState(() {
        _selecionados.clear();
      });
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('${clientes.length} venda(s) enviadas ao rascunho.'),
        ),
      );
      await widget.onDraftChanged();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao enviar para o rascunho: $error')),
      );
    } finally {
      if (mounted) {
        setState(() => _sending = false);
      }
    }
  }

  Widget _buildToolbar() {
    return AppPanel(
      margin: const EdgeInsets.fromLTRB(16, 8, 16, 0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Clientes da base',
            subtitle:
                'Selecione clientes da base online do desktop, informe caixas e preco, e envie para o rascunho compartilhado.',
            icon: Icons.groups_2_outlined,
            trailing: StatusBadge(
              label: '${_selecionados.length} selecionado(s)',
              color: _selecionados.isEmpty
                  ? Colors.blueGrey.shade700
                  : VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _buscaCtrl,
            decoration: const InputDecoration(
              labelText: 'Buscar por nome ou codigo',
              prefixIcon: Icon(Icons.search),
            ),
            onSubmitted: (_) => reload(),
          ),
          const SizedBox(height: 12),
          Row(
            children: [
              Expanded(
                child: TextField(
                  controller: _vendedorCtrl,
                  decoration: const InputDecoration(
                    labelText: 'Filtro por vendedor',
                  ),
                  onSubmitted: (_) => reload(),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: TextField(
                  controller: _cidadeCtrl,
                  decoration: const InputDecoration(
                    labelText: 'Filtro por cidade',
                  ),
                  onSubmitted: (_) => reload(),
                ),
              ),
            ],
          ),
          const SizedBox(height: 12),
          Row(
            children: [
              ElevatedButton.icon(
                onPressed: _loading ? null : reload,
                icon: const Icon(Icons.refresh),
                label: const Text('Atualizar lista'),
              ),
              const SizedBox(width: 8),
              OutlinedButton.icon(
                onPressed: _loading
                    ? null
                    : () {
                        _limparFiltros();
                        reload();
                      },
                icon: const Icon(Icons.unfold_more),
                label: const Text('Listar todos'),
              ),
              const SizedBox(width: 8),
              Text(
                'Resultados: ${_clientes.length}',
                style: Theme.of(context).textTheme.bodySmall,
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildClienteCard(Map<String, dynamic> cliente) {
    final cod = (cliente['cod_cliente'] ?? '').toString().trim().toUpperCase();
    final selected = _selecionados.contains(cod);

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      backgroundColor:
          selected ? VendorUiColors.primary.withValues(alpha: 0.06) : Colors.white,
      child: InkWell(
        onTap: () => _toggleCliente(cod, !selected),
        borderRadius: BorderRadius.circular(14),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Checkbox(
              value: selected,
              onChanged: (value) => _toggleCliente(cod, value ?? false),
            ),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    '$cod - ${(cliente['nome_cliente'] ?? '-').toString().trim()}',
                    style: const TextStyle(
                      color: VendorUiColors.heading,
                      fontWeight: FontWeight.w800,
                    ),
                  ),
                  const SizedBox(height: 6),
                  Text(
                    [
                      (cliente['cidade'] ?? '').toString().trim(),
                      (cliente['bairro'] ?? '').toString().trim(),
                      if ((cliente['vendedor'] ?? '').toString().trim().isNotEmpty)
                        'Vendedor: ${cliente['vendedor']}',
                    ].where((item) => item.isNotEmpty).join(' | '),
                    style: const TextStyle(
                      color: VendorUiColors.muted,
                      fontSize: 12,
                      fontWeight: FontWeight.w700,
                    ),
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildBody() {
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
    if (_clientes.isEmpty) {
      return ListView(
        padding: const EdgeInsets.all(16),
        children: const [
          AppPanel(
            child: Text(
              'Nenhum cliente encontrado para os filtros atuais.',
              style: TextStyle(
                color: VendorUiColors.heading,
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
        ],
      );
    }
    return ListView.builder(
      padding: const EdgeInsets.fromLTRB(16, 12, 16, 88),
      itemCount: _clientes.length,
      itemBuilder: (context, index) => _buildClienteCard(_clientes[index]),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Stack(
      children: [
        Column(
          children: [
            _buildToolbar(),
            Expanded(
              child: RefreshIndicator(
                onRefresh: reload,
                child: _buildBody(),
              ),
            ),
          ],
        ),
        if (_selecionados.isNotEmpty)
          Positioned(
            left: 16,
            right: 16,
            bottom: 16,
            child: SafeArea(
              child: ElevatedButton.icon(
                onPressed: _sending ? null : _enviarParaRascunho,
                icon: _sending
                    ? const SizedBox(
                        width: 16,
                        height: 16,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    : const Icon(Icons.send_to_mobile_outlined),
                label: const Text('Enviar para rascunho'),
              ),
            ),
          ),
      ],
    );
  }
}

class _DraftInputResult {
  const _DraftInputResult({
    required this.caixas,
    required this.preco,
    required this.observacao,
  });

  final int caixas;
  final double preco;
  final String observacao;
}
