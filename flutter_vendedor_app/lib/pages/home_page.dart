import 'package:flutter/material.dart';

import '../models/app_config.dart';
import '../models/vendedor_session.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';
import 'avulsas_page.dart';
import 'clientes_page.dart';
import 'config_page.dart';
import 'nova_avulsa_page.dart';
import 'programacoes_page.dart';
import 'rascunho_page.dart';

class HomePage extends StatefulWidget {
  const HomePage({
    super.key,
    required this.config,
    required this.session,
    required this.api,
    required this.onRefreshShell,
    required this.onLogout,
  });

  final AppConfig config;
  final VendedorSession session;
  final VendedorApiService api;
  final Future<void> Function() onRefreshShell;
  final Future<void> Function() onLogout;

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  final GlobalKey<RascunhoPageState> _rascunhoKey =
      GlobalKey<RascunhoPageState>();
  final GlobalKey<ProgramacoesPageState> _programacoesKey =
      GlobalKey<ProgramacoesPageState>();
  final GlobalKey<AvulsasPageState> _avulsasKey =
      GlobalKey<AvulsasPageState>();
  int _selectedIndex = 0;
  bool _refreshingBase = false;
  bool _syncingQueue = false;
  bool _serverReachable = false;
  int _pendingOfflineQueue = 0;

  @override
  void initState() {
    super.initState();
    _refreshOperationalState(silent: true);
  }

  Future<void> _refreshOperationalState({
    bool silent = false,
  }) async {
    final online = await widget.api.isOnlineServerReachable();
    final pending = await widget.api.pendingQueueCount();

    if (!mounted) return;
    setState(() {
      _serverReachable = online;
      _pendingOfflineQueue = pending;
    });

    if (online && pending > 0) {
      await _syncOfflineQueue(silent: silent);
      return;
    }

    if (!silent && mounted) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            online
                ? 'Conexao com o servidor validada.'
                : 'Servidor indisponivel para o fluxo online do vendedor.',
          ),
        ),
      );
    }
  }

  Future<void> _syncOfflineQueue({
    bool silent = false,
  }) async {
    if (_syncingQueue) return;

    setState(() => _syncingQueue = true);
    try {
      final online = await widget.api.isOnlineServerReachable();
      final pendingBefore = await widget.api.pendingQueueCount();
      if (!online) {
        if (!mounted) return;
        setState(() {
          _serverReachable = false;
          _pendingOfflineQueue = pendingBefore;
        });
        if (!silent) {
          final message = pendingBefore > 0
              ? 'Servidor offline. Fila aguardando sincronizacao: $pendingBefore item(ns).'
              : 'Servidor offline no momento.';
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: Text(message)),
          );
        }
        return;
      }

      final report = await widget.api.flushPendingQueue();
      if (!mounted) return;
      setState(() {
        _serverReachable = true;
        _pendingOfflineQueue = report.pending;
      });

      if (report.sent > 0) {
        await _avulsasKey.currentState?.reload();
        if (!mounted) return;
      }

      if (!silent || report.sent > 0 || report.failed > 0) {
        final message = report.sent > 0 && report.failed == 0
            ? 'Fila offline sincronizada: ${report.sent} envio(s).'
            : report.sent > 0
                ? 'Fila parcial: ${report.sent} envio(s), ${report.failed} falha(s), ${report.pending} pendente(s).'
                : report.failed > 0
                    ? 'Nao foi possivel sincronizar a fila offline. Pendentes: ${report.pending}.'
                    : 'Fila offline sem pendencias.';
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text(message)),
        );
      }
    } finally {
      if (mounted) {
        setState(() => _syncingQueue = false);
      }
    }
  }

  Future<void> _refreshBase() async {
    setState(() => _refreshingBase = true);
    try {
      await widget.api.bootstrapFirstOnlineSync(
        vendedorPadrao: widget.config.vendedorPadrao,
        cidadePadrao: widget.config.cidadePadrao,
      );
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('Cadastros sincronizados com o servidor.'),
        ),
      );
      await _refreshOperationalState(silent: true);
      await _rascunhoKey.currentState?.reload();
      await _programacoesKey.currentState?.reload();
      await _avulsasKey.currentState?.reload();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao atualizar base: $error')),
      );
    } finally {
      if (mounted) {
        setState(() => _refreshingBase = false);
      }
    }
  }

  Future<void> _openConfig() async {
    await Navigator.of(context).push<AppConfig>(
      MaterialPageRoute<AppConfig>(
        builder: (_) => ConfigPage(initialConfig: widget.config),
      ),
    );
    await widget.onRefreshShell();
    await _refreshOperationalState(silent: true);
  }

  Future<void> _handleDraftChanged() async {
    await _rascunhoKey.currentState?.reload();
  }

  Future<void> _handleProgramCreated() async {
    await _rascunhoKey.currentState?.reload();
    await _programacoesKey.currentState?.reload();
    await _refreshOperationalState();
    if (!mounted) return;
    setState(() => _selectedIndex = 2);
  }

  Future<void> _openNovaAvulsa() async {
    await Navigator.of(context).push<void>(
      MaterialPageRoute<void>(
        builder: (_) => NovaAvulsaPage(
          config: widget.config,
          api: widget.api,
          onCreated: () async {
            await _avulsasKey.currentState?.reload();
            await _refreshOperationalState(silent: true);
          },
        ),
      ),
    );
    await _avulsasKey.currentState?.reload();
    await _refreshOperationalState(silent: true);
  }

  Widget _buildHeader() {
    return AppPanel(
      margin: const EdgeInsets.fromLTRB(16, 16, 16, 8),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Painel operacional do vendedor',
            subtitle:
                'Fluxo operacional com clientes da base, avulsas, rascunho compartilhado e programacao oficial.',
            icon: Icons.storefront,
          ),
          const SizedBox(height: 14),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: _serverReachable ? 'Online' : 'Offline',
                color: _serverReachable
                    ? VendorUiColors.success
                    : VendorUiColors.danger,
              ),
              StatusBadge(
                label: _syncingQueue
                    ? 'Sincronizando fila...'
                    : 'Fila offline: $_pendingOfflineQueue',
                color: _pendingOfflineQueue > 0
                    ? VendorUiColors.warning
                    : Colors.blueGrey.shade700,
              ),
            ],
          ),
          if (!_serverReachable) ...[
            const SizedBox(height: 14),
            AppPanel(
              padding: const EdgeInsets.all(12),
              backgroundColor: VendorUiColors.warning.withValues(alpha: 0.10),
              child: const Text(
                'Clientes, rascunho e programacoes do vendedor operam online para manter o vinculo com o desktop e o app do motorista.',
                style: TextStyle(
                  color: VendorUiColors.warning,
                  fontWeight: FontWeight.w800,
                ),
              ),
            ),
          ],
          const SizedBox(height: 14),
          AppPanel(
            padding: const EdgeInsets.all(14),
            backgroundColor: VendorUiColors.surfaceAlt,
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  widget.session.nome,
                  style: const TextStyle(
                    color: VendorUiColors.heading,
                    fontWeight: FontWeight.w900,
                    fontSize: 14,
                  ),
                ),
                const SizedBox(height: 8),
                InfoRow(
                  label: 'Codigo',
                  value: widget.session.codigo,
                ),
                InfoRow(
                  label: 'API',
                  value: widget.config.normalizedBaseUrl,
                ),
                if (widget.config.cidadePadrao.trim().isNotEmpty)
                  InfoRow(
                    label: 'Cidade base',
                    value: widget.config.cidadePadrao,
                  ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  PreferredSizeWidget _buildAppBar() {
    return AppBar(
      title: const Text('RotaHub Vendedor'),
      actions: [
        IconButton(
          onPressed: _refreshingBase ? null : _refreshBase,
          icon: _refreshingBase
              ? const SizedBox(
                  width: 18,
                  height: 18,
                  child: CircularProgressIndicator(strokeWidth: 2),
                )
              : const Icon(Icons.cloud_sync_outlined),
          tooltip: 'Sincronizar cadastros',
        ),
        if (_selectedIndex == 3)
          IconButton(
            onPressed: _syncingQueue
                ? null
                : () {
                    _syncOfflineQueue();
                  },
            icon: _syncingQueue
                ? const SizedBox(
                    width: 18,
                    height: 18,
                    child: CircularProgressIndicator(strokeWidth: 2),
                  )
                : const Icon(Icons.sync_outlined),
            tooltip: 'Sincronizar fila offline',
          ),
        if (_selectedIndex == 3)
          IconButton(
            onPressed: () {
              _openNovaAvulsa();
            },
            icon: const Icon(Icons.add_road_outlined),
            tooltip: 'Nova avulsa',
          ),
        IconButton(
          onPressed: _openConfig,
          icon: const Icon(Icons.settings_outlined),
          tooltip: 'Configuracao',
        ),
        IconButton(
          onPressed: widget.onLogout,
          icon: const Icon(Icons.logout),
          tooltip: 'Sair',
        ),
      ],
    );
  }

  Widget _buildBody() {
    return Column(
      children: [
        _buildHeader(),
        Expanded(
          child: IndexedStack(
            index: _selectedIndex,
            children: [
              ClientesPage(
                config: widget.config,
                api: widget.api,
                onDraftChanged: _handleDraftChanged,
              ),
              RascunhoPage(
                key: _rascunhoKey,
                config: widget.config,
                api: widget.api,
                onProgramCreated: _handleProgramCreated,
              ),
              ProgramacoesPage(
                key: _programacoesKey,
                api: widget.api,
              ),
              AvulsasPage(
                key: _avulsasKey,
                api: widget.api,
              ),
            ],
          ),
        ),
      ],
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: _buildAppBar(),
      body: _buildBody(),
      bottomNavigationBar: NavigationBar(
        selectedIndex: _selectedIndex,
        onDestinationSelected: (index) {
          setState(() => _selectedIndex = index);
        },
        destinations: const [
          NavigationDestination(
            icon: Icon(Icons.groups_outlined),
            selectedIcon: Icon(Icons.groups),
            label: 'Clientes',
          ),
          NavigationDestination(
            icon: Icon(Icons.inventory_2_outlined),
            selectedIcon: Icon(Icons.inventory_2),
            label: 'Rascunho',
          ),
          NavigationDestination(
            icon: Icon(Icons.assignment_turned_in_outlined),
            selectedIcon: Icon(Icons.assignment_turned_in),
            label: 'Programacoes',
          ),
          NavigationDestination(
            icon: Icon(Icons.local_shipping_outlined),
            selectedIcon: Icon(Icons.local_shipping),
            label: 'Avulsas',
          ),
        ],
      ),
    );
  }
}
