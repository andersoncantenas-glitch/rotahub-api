import 'package:flutter/material.dart';

import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';

class AvulsaDetalhePage extends StatefulWidget {
  const AvulsaDetalhePage({
    super.key,
    required this.api,
    required this.codigoAvulsa,
  });

  final VendedorApiService api;
  final String codigoAvulsa;

  @override
  State<AvulsaDetalhePage> createState() => _AvulsaDetalhePageState();
}

class _AvulsaDetalhePageState extends State<AvulsaDetalhePage> {
  bool _loading = true;
  String? _error;
  Map<String, dynamic> _avulsa = <String, dynamic>{};
  List<Map<String, dynamic>> _itens = <Map<String, dynamic>>[];

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    setState(() {
      _loading = true;
      _error = null;
    });

    try {
      final data = await widget.api.detalheAvulsa(widget.codigoAvulsa);
      if (!mounted) return;
      setState(() {
        _avulsa = Map<String, dynamic>.from(
            (data['avulsa'] ?? <String, dynamic>{}) as Map);
        _itens = ((data['itens'] ?? <dynamic>[]) as List)
            .map<Map<String, dynamic>>(
                (item) => Map<String, dynamic>.from(item as Map))
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
    switch (value.trim().toUpperCase()) {
      case 'CONCILIADA':
        return Colors.green.shade700;
      case 'CONCILIADA_PARCIAL':
        return Colors.orange.shade800;
      case 'FILA_OFFLINE':
        return Colors.deepOrange.shade700;
      default:
        return Colors.blueGrey.shade700;
    }
  }

  String _formatMoney(dynamic value) {
    final raw = (value ?? '').toString().trim().replaceAll(',', '.');
    final parsed = double.tryParse(raw);
    if (parsed == null) return (value ?? '').toString().trim();
    final fixed = parsed.toStringAsFixed(2).split('.');
    return 'R\$ ${fixed[0]},${fixed[1]}';
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.codigoAvulsa),
      ),
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
                              title:
                                  (_avulsa['codigo_avulsa'] ?? '').toString(),
                              subtitle:
                                  'Resumo operacional da avulsa e dados de conciliacao.',
                              icon: Icons.route,
                              trailing: StatusBadge(
                                label: ((_avulsa['status'] ?? 'SEM STATUS')
                                        .toString())
                                    .trim()
                                    .toUpperCase(),
                                color: _statusColor(
                                  (_avulsa['status'] ?? '').toString(),
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
                                    label: 'Data',
                                    value:
                                        (_avulsa['data_programada'] ?? '')
                                            .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Motorista',
                                    value: [
                                      (_avulsa['motorista_codigo'] ?? '')
                                          .toString(),
                                      (_avulsa['motorista_nome'] ?? '')
                                          .toString(),
                                    ]
                                        .where(
                                            (part) => part.trim().isNotEmpty)
                                        .join(' - '),
                                  ),
                                  InfoRow(
                                    label: 'Veiculo',
                                    value:
                                        (_avulsa['veiculo'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Equipe',
                                    value: (_avulsa['equipe'] ?? '').toString(),
                                  ),
                                  InfoRow(
                                    label: 'Local',
                                    value: (_avulsa['local_rota'] ?? '')
                                        .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Criado por',
                                    value: (_avulsa['criado_por'] ?? '')
                                        .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Criado em',
                                    value: (_avulsa['criado_em'] ?? '')
                                        .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Prog. oficial',
                                    value: (_avulsa[
                                                'programacao_oficial_codigo'] ??
                                            '')
                                        .toString(),
                                  ),
                                  InfoRow(
                                    label: 'Observacao',
                                    value:
                                        (_avulsa['observacao'] ?? '').toString(),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 8),
                      Text(
                        'Clientes (${_itens.length})',
                        style: Theme.of(context).textTheme.titleMedium,
                      ),
                      const SizedBox(height: 8),
                      ..._itens.map(
                        (item) {
                          final chips = <String>[
                            if ((item['status_item'] ?? '')
                                .toString()
                                .trim()
                                .isNotEmpty)
                              'Status: ${item['status_item']}',
                            if ((item['pedido'] ?? '').toString().trim().isNotEmpty)
                              'Pedido: ${item['pedido']}',
                            if ((item['nf'] ?? '').toString().trim().isNotEmpty)
                              'NF: ${item['nf']}',
                            if ((item['caixas'] ?? '').toString().trim().isNotEmpty &&
                                (item['caixas'] ?? '').toString().trim() != '0')
                              'Caixas: ${item['caixas']}',
                            if ((item['preco'] ?? '').toString().trim().isNotEmpty &&
                                (item['preco'] ?? '').toString().trim() != '0' &&
                                (item['preco'] ?? '').toString().trim() != '0.0')
                              'Preco: ${_formatMoney(item['preco'])}',
                          ];

                          return AppPanel(
                            margin: const EdgeInsets.only(bottom: 12),
                            backgroundColor: VendorUiColors.surfaceAlt,
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  '${item['ordem'] ?? '-'} | ${item['cod_cliente'] ?? '-'} - ${item['nome_cliente'] ?? '-'}',
                                  style: const TextStyle(
                                    color: VendorUiColors.heading,
                                    fontWeight: FontWeight.w800,
                                  ),
                                ),
                                const SizedBox(height: 8),
                                if ((item['cidade'] ?? '')
                                        .toString()
                                        .trim()
                                        .isNotEmpty ||
                                    (item['bairro'] ?? '')
                                        .toString()
                                        .trim()
                                        .isNotEmpty)
                                  Text(
                                    [
                                      (item['cidade'] ?? '').toString(),
                                      (item['bairro'] ?? '').toString(),
                                    ]
                                        .where(
                                            (part) => part.trim().isNotEmpty)
                                        .join(' | '),
                                    style: const TextStyle(
                                      color: VendorUiColors.muted,
                                      fontWeight: FontWeight.w700,
                                    ),
                                  ),
                                if ((item['observacao'] ?? '')
                                    .toString()
                                    .trim()
                                    .isNotEmpty) ...[
                                  const SizedBox(height: 6),
                                  Text('Obs: ${item['observacao']}'),
                                ],
                                if (chips.isNotEmpty) ...[
                                  const SizedBox(height: 10),
                                  Wrap(
                                    spacing: 8,
                                    runSpacing: 8,
                                    children: chips
                                        .map(
                                          (chip) => Container(
                                            padding: const EdgeInsets.symmetric(
                                              horizontal: 10,
                                              vertical: 6,
                                            ),
                                            decoration: BoxDecoration(
                                              color: Colors.white,
                                              borderRadius:
                                                  BorderRadius.circular(999),
                                              border: Border.all(
                                                color: VendorUiColors.border,
                                              ),
                                            ),
                                            child: Text(
                                              chip,
                                              style: const TextStyle(
                                                color: VendorUiColors.heading,
                                                fontWeight: FontWeight.w700,
                                              ),
                                            ),
                                          ),
                                        )
                                        .toList(),
                                  ),
                                ],
                                if ((item['updated_at'] ?? '')
                                    .toString()
                                    .trim()
                                    .isNotEmpty) ...[
                                  const SizedBox(height: 10),
                                  Text(
                                    'Atualizado em: ${item['updated_at']}',
                                    style:
                                        Theme.of(context).textTheme.bodySmall,
                                  ),
                                ],
                              ],
                            ),
                          );
                        },
                      ),
                    ],
                  ),
      ),
    );
  }
}
