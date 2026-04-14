import 'package:flutter/material.dart';

import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';

class ProgramacaoDetalhePage extends StatefulWidget {
  const ProgramacaoDetalhePage({
    super.key,
    required this.api,
    required this.codigoProgramacao,
  });

  final VendedorApiService api;
  final String codigoProgramacao;

  @override
  State<ProgramacaoDetalhePage> createState() => _ProgramacaoDetalhePageState();
}

class _ProgramacaoDetalhePageState extends State<ProgramacaoDetalhePage> {
  final TextEditingController _buscaCtrl = TextEditingController();
  bool _loading = true;
  String? _error;
  String _statusFiltro = 'TODOS';
  Map<String, dynamic> _rota = <String, dynamic>{};
  List<Map<String, dynamic>> _clientes = <Map<String, dynamic>>[];

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    _buscaCtrl.dispose();
    super.dispose();
  }

  Future<void> _load() async {
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final data = await widget.api.detalheProgramacaoOficial(
        widget.codigoProgramacao,
      );
      if (!mounted) return;
      setState(() {
        _rota = Map<String, dynamic>.from(data['rota'] ?? <String, dynamic>{});
        _clientes = ((data['clientes'] ?? <dynamic>[]) as List)
            .map<Map<String, dynamic>>(
              (item) => Map<String, dynamic>.from(item as Map),
            )
            .toList();
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

  Color _statusColor(String value) {
    final status = value.trim().toUpperCase();
    if (status == 'EM_ENTREGAS' || status == 'EM ENTREGAS') {
      return Colors.blue.shade700;
    }
    if (status == 'EM_ROTA' || status == 'EM ROTA') {
      return Colors.indigo.shade700;
    }
    if (status == 'FINALIZADA' || status == 'FINALIZADO') {
      return Colors.grey.shade700;
    }
    return VendorUiColors.success;
  }

  Color _pedidoStatusColor(String value) {
    final status = value.trim().toUpperCase();
    switch (status) {
      case 'ENTREGUE':
        return VendorUiColors.success;
      case 'ALTERADO':
        return Colors.indigo.shade700;
      case 'CANCELADO':
      case 'CANCELADA':
        return VendorUiColors.danger;
      case 'PENDENTE':
      default:
        return VendorUiColors.warning;
    }
  }

  String _statusRota() {
    return ((_rota['status_operacional'] ?? _rota['status'] ?? 'ATIVA')
            .toString())
        .trim()
        .toUpperCase();
  }

  bool _edicaoLiberadaNoVendedor() {
    final status = _statusRota();
    return status.isEmpty ||
        status == 'ATIVA' ||
        status == 'ABERTA' ||
        status == 'PENDENTE' ||
        status == 'PROGRAMADA';
  }

  String _pedidoStatus(Map<String, dynamic> item) {
    final status = (item['status_pedido'] ?? 'PENDENTE').toString().trim();
    return status.isEmpty ? 'PENDENTE' : status.toUpperCase();
  }

  bool _matchesPedidoFiltro(String status) {
    if (_statusFiltro == 'TODOS') return true;
    if (_statusFiltro == 'CANCELADO') {
      return status == 'CANCELADO' || status == 'CANCELADA';
    }
    return status == _statusFiltro;
  }

  int _toInt(dynamic value) {
    if (value is int) return value;
    return int.tryParse((value ?? '').toString().trim()) ?? 0;
  }

  double _toDouble(dynamic value) {
    if (value is double) return value;
    if (value is int) return value.toDouble();
    return double.tryParse(
          (value ?? '').toString().trim().replaceAll(',', '.'),
        ) ??
        0.0;
  }

  String _money(dynamic value) =>
      _toDouble(value).toStringAsFixed(2).replaceAll('.', ',');

  List<Map<String, dynamic>> get _clientesVisiveis {
    final term = _buscaCtrl.text.trim().toUpperCase();
    return _clientes.where((item) {
      final status = _pedidoStatus(item);
      final haystack = <String>[
        (item['cod_cliente'] ?? '').toString(),
        (item['nome_cliente'] ?? '').toString(),
        (item['vendedor'] ?? '').toString(),
        (item['pedido'] ?? '').toString(),
      ].join(' | ').toUpperCase();
      final matchStatus = _matchesPedidoFiltro(status);
      final matchBusca = term.isEmpty || haystack.contains(term);
      return matchStatus && matchBusca;
    }).toList();
  }

  int _countPedidos(String status) =>
      _clientes.where((item) => _pedidoStatus(item) == status).length;

  @override
  Widget build(BuildContext context) {
    final statusRota = _statusRota();
    final edicaoLiberada = _edicaoLiberadaNoVendedor();
    return Scaffold(
      appBar: AppBar(title: Text(widget.codigoProgramacao)),
      body: RefreshIndicator(
        onRefresh: _load,
        child: _loading
            ? ListView(
                children: const [
                  SizedBox(height: 180),
                  Center(child: CircularProgressIndicator()),
                ],
              )
            : _error != null
                ? ListView(
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
                  )
                : ListView(
                    padding: const EdgeInsets.all(16),
                    children: [
                      AppPanel(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            PanelHeader(
                              title: (_rota['codigo_programacao'] ?? '-')
                                  .toString(),
                              subtitle:
                                  'Resumo da programacao oficial enviada ao ecossistema do desktop e motorista.',
                              icon: Icons.route,
                              trailing: StatusBadge(
                                label: statusRota,
                                color: _statusColor(statusRota),
                              ),
                            ),
                            const SizedBox(height: 16),
                            Wrap(
                              spacing: 8,
                              runSpacing: 8,
                              children: [
                                StatusBadge(
                                  label: edicaoLiberada
                                      ? 'Edicao liberada no vendedor'
                                      : 'Edicao bloqueada no vendedor',
                                  color: edicaoLiberada
                                      ? VendorUiColors.success
                                      : VendorUiColors.warning,
                                ),
                                StatusBadge(
                                  label: 'Clientes: ${_clientes.length}',
                                  color: VendorUiColors.primary,
                                ),
                                StatusBadge(
                                  label: 'Entregues: ${_countPedidos('ENTREGUE')}',
                                  color: VendorUiColors.success,
                                ),
                                StatusBadge(
                                  label: 'Pendentes: ${_countPedidos('PENDENTE')}',
                                  color: VendorUiColors.warning,
                                ),
                                StatusBadge(
                                  label: 'Alterados: ${_countPedidos('ALTERADO')}',
                                  color: Colors.indigo.shade700,
                                ),
                                StatusBadge(
                                  label:
                                      'Cancelados: ${_countPedidos('CANCELADO') + _countPedidos('CANCELADA')}',
                                  color: VendorUiColors.danger,
                                ),
                              ],
                            ),
                            const SizedBox(height: 12),
                            AppPanel(
                              padding: const EdgeInsets.all(12),
                              backgroundColor: edicaoLiberada
                                  ? VendorUiColors.success.withValues(alpha: 0.08)
                                  : VendorUiColors.warning.withValues(alpha: 0.10),
                              child: Text(
                                edicaoLiberada
                                    ? 'Enquanto a programacao estiver ATIVA, ela permanece elegivel para ajuste pelo vendedor.'
                                    : 'Depois que a rota entra em execucao, as alteracoes passam a acontecer apenas no app do motorista.',
                                style: TextStyle(
                                  color: edicaoLiberada
                                      ? VendorUiColors.success
                                      : VendorUiColors.warning,
                                  fontWeight: FontWeight.w800,
                                ),
                              ),
                            ),
                            const SizedBox(height: 16),
                            AppPanel(
                              padding: const EdgeInsets.all(14),
                              backgroundColor: VendorUiColors.surfaceAlt,
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  InfoRow(
                                    label: 'Motorista',
                                    value: (_rota['motorista'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Veiculo',
                                    value: (_rota['veiculo'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Equipe',
                                    value: (_rota['equipe'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Local rota',
                                    value: (_rota['local_rota'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Carregamento',
                                    value:
                                        (_rota['local_carregamento'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Tipo',
                                    value:
                                        (_rota['tipo_operacao'] ?? _rota['tipo_estimativa'] ?? '')
                                            .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Total caixas',
                                    value: (_rota['total_caixas'] ?? '').toString(),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 12),
                      AppPanel(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            PanelHeader(
                              title: 'Pedidos da programacao',
                              subtitle:
                                  'Acompanhe o que ja foi entregue ou alterado pelo motorista.',
                              icon: Icons.list_alt,
                              trailing: StatusBadge(
                                label: '${_clientesVisiveis.length} visiveis',
                                color: VendorUiColors.primary,
                              ),
                            ),
                            const SizedBox(height: 12),
                            TextField(
                              controller: _buscaCtrl,
                              onChanged: (_) => setState(() {}),
                              decoration: const InputDecoration(
                                labelText:
                                    'Buscar por cliente, codigo, vendedor ou pedido',
                                prefixIcon: Icon(Icons.search),
                              ),
                            ),
                            const SizedBox(height: 12),
                            SingleChildScrollView(
                              scrollDirection: Axis.horizontal,
                              child: Row(
                                children: [
                                  FilterChip(
                                    label: const Text('Todos'),
                                    selected: _statusFiltro == 'TODOS',
                                    onSelected: (_) =>
                                        setState(() => _statusFiltro = 'TODOS'),
                                  ),
                                  const SizedBox(width: 8),
                                  FilterChip(
                                    label: const Text('PENDENTE'),
                                    selected: _statusFiltro == 'PENDENTE',
                                    onSelected: (_) => setState(
                                      () => _statusFiltro = 'PENDENTE',
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  FilterChip(
                                    label: const Text('ENTREGUE'),
                                    selected: _statusFiltro == 'ENTREGUE',
                                    onSelected: (_) => setState(
                                      () => _statusFiltro = 'ENTREGUE',
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  FilterChip(
                                    label: const Text('ALTERADO'),
                                    selected: _statusFiltro == 'ALTERADO',
                                    onSelected: (_) => setState(
                                      () => _statusFiltro = 'ALTERADO',
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  FilterChip(
                                    label: const Text('CANCELADO'),
                                    selected: _statusFiltro == 'CANCELADO',
                                    onSelected: (_) => setState(
                                      () => _statusFiltro = 'CANCELADO',
                                    ),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 8),
                      ..._clientesVisiveis.map((item) {
                        final statusPedido = _pedidoStatus(item);
                        final caixasBase = _toInt(item['qnt_caixas']);
                        final caixasAtual = _toInt(
                          item['caixas_atual'] ?? item['qnt_caixas'],
                        );
                        final precoBase = _toDouble(item['preco']);
                        final precoAtual = _toDouble(
                          item['preco_atual'] ?? item['preco'],
                        );
                        return AppPanel(
                          margin: const EdgeInsets.only(bottom: 10),
                          backgroundColor: VendorUiColors.surfaceAlt,
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Expanded(
                                    child: Column(
                                      crossAxisAlignment:
                                          CrossAxisAlignment.start,
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
                                            if ((item['vendedor'] ?? '')
                                                .toString()
                                                .trim()
                                                .isNotEmpty)
                                              'Vendedor: ${item['vendedor']}',
                                            if ((item['pedido'] ?? '')
                                                .toString()
                                                .trim()
                                                .isNotEmpty)
                                              'Pedido: ${item['pedido']}',
                                          ].join(' | '),
                                          style: const TextStyle(
                                            color: VendorUiColors.muted,
                                            fontWeight: FontWeight.w700,
                                          ),
                                        ),
                                      ],
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  StatusBadge(
                                    label: statusPedido,
                                    color: _pedidoStatusColor(statusPedido),
                                  ),
                                ],
                              ),
                              const SizedBox(height: 8),
                              Wrap(
                                spacing: 8,
                                runSpacing: 8,
                                children: [
                                  StatusBadge(
                                    label: 'Caixas base: $caixasBase',
                                    color: Colors.blueGrey.shade700,
                                  ),
                                  StatusBadge(
                                    label: 'Caixas atual: $caixasAtual',
                                    color: caixasAtual != caixasBase
                                        ? Colors.indigo.shade700
                                        : VendorUiColors.primary,
                                  ),
                                  StatusBadge(
                                    label:
                                        'Preco atual: R\$ ${_money(precoAtual)}',
                                    color: precoAtual != precoBase
                                        ? Colors.indigo.shade700
                                        : VendorUiColors.success,
                                  ),
                                ],
                              ),
                              if ((item['alteracao_tipo'] ?? '')
                                  .toString()
                                  .trim()
                                  .isNotEmpty) ...[
                                const SizedBox(height: 8),
                                Text(
                                  'Alteracao: ${(item['alteracao_tipo'] ?? '').toString().trim()}',
                                  style: const TextStyle(
                                    color: VendorUiColors.heading,
                                    fontWeight: FontWeight.w800,
                                  ),
                                ),
                              ],
                              if ((item['alteracao_detalhe'] ?? '')
                                  .toString()
                                  .trim()
                                  .isNotEmpty) ...[
                                const SizedBox(height: 4),
                                Text(
                                  'Detalhe: ${(item['alteracao_detalhe'] ?? '').toString().trim()}',
                                  style: const TextStyle(
                                    color: VendorUiColors.muted,
                                    fontWeight: FontWeight.w700,
                                  ),
                                ),
                              ],
                              if ((item['alterado_em'] ?? '')
                                  .toString()
                                  .trim()
                                  .isNotEmpty) ...[
                                const SizedBox(height: 4),
                                Text(
                                  'Ultima movimentacao: ${(item['alterado_em'] ?? '').toString().trim()}',
                                  style: const TextStyle(
                                    color: VendorUiColors.muted,
                                    fontWeight: FontWeight.w700,
                                  ),
                                ),
                              ],
                            ],
                          ),
                        );
                      }),
                    ],
                  ),
      ),
    );
  }
}
