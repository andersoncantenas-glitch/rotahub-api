import 'package:flutter/material.dart';

import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';
import 'avulsa_detalhe_page.dart';

class AvulsasPage extends StatefulWidget {
  const AvulsasPage({
    super.key,
    required this.api,
  });

  final VendedorApiService api;

  @override
  State<AvulsasPage> createState() => AvulsasPageState();
}

class AvulsasPageState extends State<AvulsasPage> {
  final TextEditingController _searchCtrl = TextEditingController();
  bool _loading = true;
  String? _error;
  String _status = '';
  DateTime? _dataDe;
  DateTime? _dataAte;
  List<Map<String, dynamic>> _items = <Map<String, dynamic>>[];

  @override
  void initState() {
    super.initState();
    reload();
  }

  @override
  void dispose() {
    _searchCtrl.dispose();
    super.dispose();
  }

  String _formatDate(DateTime value) {
    final month = value.month.toString().padLeft(2, '0');
    final day = value.day.toString().padLeft(2, '0');
    return '${value.year}-$month-$day';
  }

  Future<void> _pickDate({required bool isStart}) async {
    final now = DateTime.now();
    final picked = await showDatePicker(
      context: context,
      initialDate: isStart ? (_dataDe ?? now) : (_dataAte ?? _dataDe ?? now),
      firstDate: DateTime(now.year - 1),
      lastDate: DateTime(now.year + 2),
    );
    if (picked == null) return;
    setState(() {
      if (isStart) {
        _dataDe = picked;
        if (_dataAte != null && _dataAte!.isBefore(picked)) {
          _dataAte = picked;
        }
      } else {
        _dataAte = picked;
      }
    });
    await reload();
  }

  void _clearFilters() {
    setState(() {
      _status = '';
      _dataDe = null;
      _dataAte = null;
      _searchCtrl.clear();
    });
    reload();
  }

  Future<void> reload() async {
    setState(() {
      _loading = true;
      _error = null;
    });

    try {
      final data = await widget.api.listarAvulsas(
        status: _status,
        dataDe: _dataDe == null ? '' : _formatDate(_dataDe!),
        dataAte: _dataAte == null ? '' : _formatDate(_dataAte!),
      );
      if (!mounted) return;
      setState(() {
        _items = data;
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

  List<Map<String, dynamic>> get _visibleItems {
    final term = _searchCtrl.text.trim().toUpperCase();
    if (term.isEmpty) return _items;
    return _items.where((item) {
      final haystack = <String>[
        (item['codigo_avulsa'] ?? '').toString(),
        (item['motorista_codigo'] ?? '').toString(),
        (item['motorista_nome'] ?? '').toString(),
        (item['veiculo'] ?? '').toString(),
        (item['equipe'] ?? '').toString(),
        (item['local_rota'] ?? '').toString(),
        (item['programacao_oficial_codigo'] ?? '').toString(),
        (item['criado_por'] ?? '').toString(),
      ].join(' | ').toUpperCase();
      return haystack.contains(term);
    }).toList();
  }

  Color _statusColor(String value) {
    switch (value.trim().toUpperCase()) {
      case 'CONCILIADA':
        return VendorUiColors.success;
      case 'CONCILIADA_PARCIAL':
        return VendorUiColors.warning;
      case 'FILA_OFFLINE':
        return Colors.deepOrange.shade700;
      default:
        return Colors.blueGrey.shade700;
    }
  }

  Widget _buildFilterChip(String value, String label) {
    final selected = _status == value;
    return FilterChip(
      selected: selected,
      label: Text(label),
      onSelected: (_) {
        setState(() {
          _status = value;
        });
        reload();
      },
    );
  }

  Widget _buildToolbar() {
    return AppPanel(
      margin: const EdgeInsets.fromLTRB(16, 8, 16, 0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Avulsas registradas',
            subtitle:
                'Lista operacional com busca, filtros e status de conciliacao.',
            icon: Icons.list_alt,
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _searchCtrl,
            onChanged: (_) => setState(() {}),
            decoration: const InputDecoration(
              labelText: 'Buscar por codigo, motorista, veiculo ou rota',
              prefixIcon: Icon(Icons.search),
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              OutlinedButton.icon(
                onPressed: () => _pickDate(isStart: true),
                icon: const Icon(Icons.event_available),
                label: Text(
                  _dataDe == null
                      ? 'Data inicial'
                      : 'De: ${_formatDate(_dataDe!)}',
                ),
              ),
              OutlinedButton.icon(
                onPressed: () => _pickDate(isStart: false),
                icon: const Icon(Icons.event),
                label: Text(
                  _dataAte == null
                      ? 'Data final'
                      : 'Ate: ${_formatDate(_dataAte!)}',
                ),
              ),
              TextButton.icon(
                onPressed: _clearFilters,
                icon: const Icon(Icons.filter_alt_off),
                label: const Text('Limpar filtros'),
              ),
            ],
          ),
          const SizedBox(height: 12),
          Text(
            'Resultados: ${_visibleItems.length}',
            style: Theme.of(context).textTheme.bodySmall,
          ),
        ],
      ),
    );
  }

  Widget _buildEmpty(String message, {Color? color}) {
    return ListView(
      padding: const EdgeInsets.all(16),
      children: [
        AppPanel(
          child: Text(
            message,
            style: TextStyle(
              color: color ?? VendorUiColors.heading,
              fontWeight: FontWeight.w700,
            ),
          ),
        ),
      ],
    );
  }

  Future<void> _openDetails(Map<String, dynamic> item) async {
    await Navigator.of(context).push(
      MaterialPageRoute<void>(
        builder: (_) => AvulsaDetalhePage(
          api: widget.api,
          codigoAvulsa: (item['codigo_avulsa'] ?? '').toString().trim(),
        ),
      ),
    );
    await reload();
  }

  Widget _buildItemCard(Map<String, dynamic> item) {
    final status = (item['status'] ?? '').toString().trim().toUpperCase();
    final oficial =
        (item['programacao_oficial_codigo'] ?? '').toString().trim();

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      backgroundColor: VendorUiColors.surfaceAlt,
      child: InkWell(
        onTap: () => _openDetails(item),
        borderRadius: BorderRadius.circular(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Container(
                  width: 42,
                  height: 42,
                  decoration: BoxDecoration(
                    color: VendorUiColors.primary,
                    borderRadius: BorderRadius.circular(12),
                  ),
                  child: const Icon(Icons.local_shipping, color: Colors.white),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        (item['codigo_avulsa'] ?? '-').toString(),
                        style: const TextStyle(
                          color: VendorUiColors.heading,
                          fontWeight: FontWeight.w900,
                          fontSize: 13,
                        ),
                      ),
                      const SizedBox(height: 4),
                      Text(
                        (item['motorista_nome'] ?? '-').toString(),
                        style: const TextStyle(
                          color: VendorUiColors.heading,
                          fontWeight: FontWeight.w700,
                        ),
                      ),
                      const SizedBox(height: 4),
                      Text(
                        [
                          (item['data_programada'] ?? '').toString(),
                          (item['veiculo'] ?? '').toString(),
                          (item['local_rota'] ?? '').toString(),
                          if (oficial.isNotEmpty) 'Oficial: $oficial',
                        ].where((part) => part.trim().isNotEmpty).join(' | '),
                        style: const TextStyle(
                          color: VendorUiColors.muted,
                          fontSize: 12,
                          fontWeight: FontWeight.w700,
                        ),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 8),
                StatusBadge(
                  label: status.isEmpty ? 'SEM STATUS' : status,
                  color: _statusColor(status),
                ),
              ],
            ),
            const SizedBox(height: 12),
            Row(
              children: [
                Expanded(
                  child: Text(
                    'Equipe: ${(item['equipe'] ?? '-').toString()}',
                    style: const TextStyle(
                      color: VendorUiColors.muted,
                      fontSize: 12,
                      fontWeight: FontWeight.w700,
                    ),
                  ),
                ),
                const SizedBox(width: 8),
                const Icon(
                  Icons.open_in_new,
                  size: 18,
                  color: VendorUiColors.muted,
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final visibleItems = _visibleItems;
    return Column(
      children: [
        _buildToolbar(),
        SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          padding: const EdgeInsets.fromLTRB(16, 8, 16, 0),
          child: Row(
            children: [
              _buildFilterChip('', 'Todas'),
              const SizedBox(width: 8),
              _buildFilterChip('AVULSA_ATIVA', 'Ativas'),
              const SizedBox(width: 8),
              _buildFilterChip('CONCILIADA', 'Conciliadas'),
              const SizedBox(width: 8),
              _buildFilterChip('CONCILIADA_PARCIAL', 'Parciais'),
              const SizedBox(width: 8),
              _buildFilterChip('FILA_OFFLINE', 'Fila offline'),
            ],
          ),
        ),
        Expanded(
          child: RefreshIndicator(
            onRefresh: reload,
            child: _loading
                ? ListView(
                    children: const [
                      SizedBox(height: 180),
                      Center(child: CircularProgressIndicator()),
                    ],
                  )
                : _error != null
                    ? _buildEmpty(_error!, color: VendorUiColors.danger)
                    : _items.isEmpty
                        ? _buildEmpty('Nenhuma avulsa encontrada.')
                        : visibleItems.isEmpty
                            ? _buildEmpty(
                                'Nenhuma avulsa encontrada com os filtros atuais.',
                              )
                            : ListView.builder(
                                padding:
                                    const EdgeInsets.fromLTRB(16, 12, 16, 24),
                                itemCount: visibleItems.length,
                                itemBuilder: (context, index) {
                                  return _buildItemCard(visibleItems[index]);
                                },
                              ),
          ),
        ),
      ],
    );
  }
}
