import 'package:flutter/material.dart';

import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';
import 'programacao_detalhe_page.dart';

class ProgramacoesPage extends StatefulWidget {
  const ProgramacoesPage({
    super.key,
    required this.api,
  });

  final VendedorApiService api;

  @override
  State<ProgramacoesPage> createState() => ProgramacoesPageState();
}

class ProgramacoesPageState extends State<ProgramacoesPage> {
  final TextEditingController _buscaCtrl = TextEditingController();
  bool _loading = true;
  String? _error;
  String _modo = 'ativas';
  String _statusFiltro = 'TODOS';
  List<Map<String, dynamic>> _items = <Map<String, dynamic>>[];

  @override
  void dispose() {
    _buscaCtrl.dispose();
    super.dispose();
  }

  @override
  void initState() {
    super.initState();
    reload();
  }

  Future<void> reload() async {
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final data = await widget.api.listarProgramacoesOficiais(modo: _modo);
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

  Color _statusColor(Map<String, dynamic> item) {
    final status = _status(item);
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

  String _status(Map<String, dynamic> item) {
    return ((item['status_operacional'] ?? item['status'] ?? '').toString())
        .trim()
        .toUpperCase();
  }

  bool _edicaoLiberada(Map<String, dynamic> item) {
    final status = _status(item);
    return status.isEmpty ||
        status == 'ATIVA' ||
        status == 'ABERTA' ||
        status == 'PENDENTE' ||
        status == 'PROGRAMADA';
  }

  bool _matchesStatusFiltro(String status) {
    if (_statusFiltro == 'TODOS') return true;
    if (_statusFiltro == 'EM_ROTA') {
      return status == 'EM_ROTA' ||
          status == 'EM ROTA' ||
          status == 'INICIADA' ||
          status == 'CARREGADA';
    }
    if (_statusFiltro == 'EM_ENTREGAS') {
      return status == 'EM_ENTREGAS' || status == 'EM ENTREGAS';
    }
    if (_statusFiltro == 'FINALIZADA') {
      return status == 'FINALIZADA' || status == 'FINALIZADO';
    }
    return status == _statusFiltro;
  }

  List<Map<String, dynamic>> get _visibleItems {
    final term = _buscaCtrl.text.trim().toUpperCase();
    return _items.where((item) {
      final status = _status(item);
      final haystack = <String>[
        (item['codigo_programacao'] ?? '').toString(),
        (item['motorista'] ?? '').toString(),
        (item['veiculo'] ?? '').toString(),
        (item['equipe'] ?? '').toString(),
      ].join(' | ').toUpperCase();
      final matchStatus = _matchesStatusFiltro(status);
      final matchBusca = term.isEmpty || haystack.contains(term);
      return matchStatus && matchBusca;
    }).toList();
  }

  int _countByStatus(String status) =>
      _items.where((item) => _status(item) == status).length;

  Widget _buildToolbar() {
    return AppPanel(
      margin: const EdgeInsets.fromLTRB(16, 8, 16, 0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Rotas ativas e programacoes',
            subtitle:
                'Acompanhamento do status operacional vindo do app do motorista e do desktop.',
            icon: Icons.assignment_turned_in_outlined,
            trailing: StatusBadge(
              label: '${_visibleItems.length} registros',
              color: VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _buscaCtrl,
            onChanged: (_) => setState(() {}),
            decoration: const InputDecoration(
              labelText: 'Buscar por codigo, motorista, veiculo ou equipe',
              prefixIcon: Icon(Icons.search),
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              FilterChip(
                label: const Text('Ativas'),
                selected: _modo == 'ativas',
                onSelected: (_) {
                  setState(() => _modo = 'ativas');
                  reload();
                },
              ),
              FilterChip(
                label: const Text('Todas'),
                selected: _modo == 'todas',
                onSelected: (_) {
                  setState(() => _modo = 'todas');
                  reload();
                },
              ),
            ],
          ),
          const SizedBox(height: 12),
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            child: Row(
              children: [
                FilterChip(
                  label: const Text('Todos status'),
                  selected: _statusFiltro == 'TODOS',
                  onSelected: (_) => setState(() => _statusFiltro = 'TODOS'),
                ),
                const SizedBox(width: 8),
                FilterChip(
                  label: const Text('ATIVA'),
                  selected: _statusFiltro == 'ATIVA',
                  onSelected: (_) => setState(() => _statusFiltro = 'ATIVA'),
                ),
                const SizedBox(width: 8),
                FilterChip(
                  label: const Text('EM ROTA'),
                  selected: _statusFiltro == 'EM_ROTA',
                  onSelected: (_) => setState(() => _statusFiltro = 'EM_ROTA'),
                ),
                const SizedBox(width: 8),
                FilterChip(
                  label: const Text('EM ENTREGAS'),
                  selected: _statusFiltro == 'EM_ENTREGAS',
                  onSelected: (_) =>
                      setState(() => _statusFiltro = 'EM_ENTREGAS'),
                ),
                const SizedBox(width: 8),
                FilterChip(
                  label: const Text('FINALIZADA'),
                  selected: _statusFiltro == 'FINALIZADA',
                  onSelected: (_) =>
                      setState(() => _statusFiltro = 'FINALIZADA'),
                ),
              ],
            ),
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Ativas: ${_countByStatus('ATIVA')}',
                color: VendorUiColors.success,
              ),
              StatusBadge(
                label: 'Em rota: ${_countByStatus('EM_ROTA') + _countByStatus('EM ROTA') + _countByStatus('INICIADA') + _countByStatus('CARREGADA')}',
                color: Colors.indigo.shade700,
              ),
              StatusBadge(
                label:
                    'Em entregas: ${_countByStatus('EM_ENTREGAS') + _countByStatus('EM ENTREGAS')}',
                color: Colors.blue.shade700,
              ),
              StatusBadge(
                label:
                    'Finalizadas: ${_countByStatus('FINALIZADA') + _countByStatus('FINALIZADO')}',
                color: Colors.grey.shade700,
              ),
            ],
          ),
        ],
      ),
    );
  }

  Widget _buildItemCard(Map<String, dynamic> item) {
    final status = _status(item);
    final edicaoLiberada = _edicaoLiberada(item);

    return AppPanel(
      margin: const EdgeInsets.only(bottom: 12),
      backgroundColor: VendorUiColors.surfaceAlt,
      child: InkWell(
        onTap: () async {
          await Navigator.of(context).push(
            MaterialPageRoute<void>(
              builder: (_) => ProgramacaoDetalhePage(
                api: widget.api,
                codigoProgramacao:
                    (item['codigo_programacao'] ?? '').toString().trim(),
              ),
            ),
          );
          await reload();
        },
        borderRadius: BorderRadius.circular(14),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Container(
              width: 42,
              height: 42,
              decoration: BoxDecoration(
                color: VendorUiColors.primary,
                borderRadius: BorderRadius.circular(12),
              ),
              child: const Icon(Icons.route, color: Colors.white),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    (item['codigo_programacao'] ?? '-').toString(),
                    style: const TextStyle(
                      color: VendorUiColors.heading,
                      fontWeight: FontWeight.w900,
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    [
                      (item['motorista'] ?? '').toString().trim(),
                      (item['veiculo'] ?? '').toString().trim(),
                      (item['data_criacao'] ?? '').toString().trim(),
                    ].where((part) => part.isNotEmpty).join(' | '),
                    style: const TextStyle(
                      color: VendorUiColors.muted,
                      fontWeight: FontWeight.w700,
                    ),
                  ),
                  const SizedBox(height: 8),
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      StatusBadge(
                        label: edicaoLiberada
                            ? 'Edicao liberada'
                            : 'Edicao bloqueada',
                        color: edicaoLiberada
                            ? VendorUiColors.success
                            : VendorUiColors.warning,
                      ),
                    ],
                  ),
                ],
              ),
            ),
            const SizedBox(width: 8),
            StatusBadge(
              label: status.isEmpty ? 'ATIVA' : status,
              color: _statusColor(item),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildBody() {
    final visibleItems = _visibleItems;
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
    if (visibleItems.isEmpty) {
      return ListView(
        padding: const EdgeInsets.all(16),
        children: [
          AppPanel(
            child: Text(
              _items.isEmpty
                  ? 'Nenhuma programacao encontrada para o filtro atual.'
                  : 'Nenhuma programacao corresponde aos filtros aplicados.',
              style: const TextStyle(
                color: VendorUiColors.heading,
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
        ],
      );
    }
    return ListView.builder(
      padding: const EdgeInsets.fromLTRB(16, 12, 16, 24),
      itemCount: visibleItems.length,
      itemBuilder: (context, index) => _buildItemCard(visibleItems[index]),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        _buildToolbar(),
        Expanded(
          child: RefreshIndicator(
            onRefresh: reload,
            child: _buildBody(),
          ),
        ),
      ],
    );
  }
}
