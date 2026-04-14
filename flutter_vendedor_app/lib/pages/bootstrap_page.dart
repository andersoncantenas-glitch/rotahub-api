import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';

import '../core/app_config_store.dart';
import '../core/default_app_config.dart';
import '../core/shared_prefs_sync_store.dart';
import '../core/vendedor_session_store.dart';
import '../models/app_config.dart';
import '../models/vendedor_session.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';
import 'config_page.dart';
import 'home_page.dart';

class BootstrapPage extends StatefulWidget {
  const BootstrapPage({super.key});

  @override
  State<BootstrapPage> createState() => _BootstrapPageState();
}

class _BootstrapPageState extends State<BootstrapPage> {
  bool _loading = true;
  bool _loggingIn = false;
  bool _checkingServer = false;
  bool _showSenha = false;
  String? _error;
  AppConfig? _config;
  VendedorApiService? _api;
  VendedorSession? _session;
  bool _serverReachable = false;
  final TextEditingController _loginCtrl = TextEditingController();
  final TextEditingController _senhaCtrl = TextEditingController();

  Future<VendedorApiService> _buildApi(AppConfig config) async {
    final prefs = await SharedPreferences.getInstance();
    final store = SharedPrefsSyncStore(prefs);
    return VendedorApiService(
      baseUrl: config.normalizedBaseUrl,
      desktopSecret: config.desktopSecret.trim(),
      store: store,
    );
  }

  @override
  void initState() {
    super.initState();
    _reload();
  }

  @override
  void dispose() {
    _loginCtrl.dispose();
    _senhaCtrl.dispose();
    super.dispose();
  }

  Future<void> _reload({bool silent = false}) async {
    if (!silent) {
      setState(() {
        _loading = true;
        _error = null;
      });
    } else if (mounted) {
      setState(() {
        _checkingServer = true;
        _error = null;
      });
    }

    try {
      final config = await AppConfigStore.load() ?? buildDefaultAppConfig();
      final api = await _buildApi(config);
      final serverReachable = await api.isOnlineServerReachable();
      var session = await VendedorSessionStore.load();
      final codigoConfigurado = config.vendedorPadrao.trim().toUpperCase();
      if (session != null &&
          codigoConfigurado.isNotEmpty &&
          session.codigo.trim().toUpperCase() != codigoConfigurado) {
        await VendedorSessionStore.clear();
        session = null;
      }

      final loginPrefill = session?.nome.trim().isNotEmpty == true
          ? session!.nome
          : config.vendedorLogin;
      if (_loginCtrl.text.trim().isEmpty && loginPrefill.trim().isNotEmpty) {
        _loginCtrl.text = loginPrefill;
      }

      if (!mounted) return;
      setState(() {
        _config = config;
        _api = api;
        _session = session;
        _serverReachable = serverReachable;
        _loading = false;
        _checkingServer = false;
      });
    } catch (error) {
      if (!mounted) return;
      setState(() {
        _error = error.toString();
        _loading = false;
        _checkingServer = false;
      });
    }
  }

  Future<void> _openConfig() async {
    await Navigator.of(context).push<AppConfig>(
      MaterialPageRoute<AppConfig>(
        builder: (_) => ConfigPage(initialConfig: _config ?? buildDefaultAppConfig()),
      ),
    );
    await _reload();
  }

  Future<void> _loginVendedor() async {
    final config = _config;
    final api = _api;
    final identificador = _loginCtrl.text.trim();
    final senha = _senhaCtrl.text.trim();
    if (config == null || api == null) return;

    if (identificador.isEmpty) {
      setState(() {
        _error = 'Informe o nome do vendedor.';
      });
      return;
    }
    if (senha.isEmpty) {
      setState(() {
        _error = 'Informe a senha do vendedor.';
      });
      return;
    }

    FocusScope.of(context).unfocus();
    setState(() {
      _loggingIn = true;
      _error = null;
    });

    try {
      final payload = await api.autenticarVendedor(
        identificador: identificador,
        senha: senha,
      );
      final session = VendedorSession(
        token: (payload['token'] ?? '').toString(),
        codigo: (payload['codigo'] ?? '').toString(),
        nome: (payload['nome'] ?? identificador).toString(),
      );
      await VendedorSessionStore.save(session);
      await AppConfigStore.save(
        config.copyWith(
          vendedorPadrao: session.codigo,
          vendedorLogin: session.nome,
        ),
      );
      await VendedorSessionStore.save(session);
      _senhaCtrl.clear();
      try {
        await api.bootstrapFirstOnlineSync(
          vendedorPadrao: session.codigo,
          cidadePadrao: config.cidadePadrao,
        );
      } catch (_) {
        // O app trabalha online; a sincronizacao inicial serve como aquecimento
        // local, mas nao deve bloquear o acesso apos o login valido.
      }
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Login realizado: ${session.nome}.')),
      );
      await _reload();
    } catch (error) {
      if (!mounted) return;
      setState(() {
        _error = error.toString();
      });
    } finally {
      if (mounted) {
        setState(() {
          _loggingIn = false;
        });
      }
    }
  }

  Future<void> _logoutVendedor() async {
    final nomeAtual = _session?.nome ?? _loginCtrl.text.trim();
    await VendedorSessionStore.clear();
    _loginCtrl.text = nomeAtual;
    _senhaCtrl.clear();
    await _reload();
  }

  Widget _buildScrollableShell({
    required Widget child,
    double maxWidth = 560,
  }) {
    return SafeArea(
      child: LayoutBuilder(
        builder: (context, constraints) {
          return SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: ConstrainedBox(
              constraints: BoxConstraints(
                minHeight: constraints.maxHeight - 48,
              ),
              child: Center(
                child: ConstrainedBox(
                  constraints: BoxConstraints(maxWidth: maxWidth),
                  child: child,
                ),
              ),
            ),
          );
        },
      ),
    );
  }

  PreferredSizeWidget _buildLoginAppBar() {
    return AppBar(
      title: const Text('Login do vendedor'),
      actions: [
        IconButton(
          onPressed: _checkingServer ? null : () => _reload(silent: true),
          icon: _checkingServer
              ? const SizedBox(
                  width: 18,
                  height: 18,
                  child: CircularProgressIndicator(strokeWidth: 2),
                )
              : const Icon(Icons.wifi_tethering),
          tooltip: 'Status da API',
        ),
        IconButton(
          onPressed: _openConfig,
          icon: const Icon(Icons.settings),
          tooltip: 'Configuracao',
        ),
      ],
    );
  }

  Widget _buildErrorCard() {
    if (_error == null) return const SizedBox.shrink();
    return AppPanel(
      padding: const EdgeInsets.all(14),
      backgroundColor: Colors.red.withValues(alpha: 0.06),
      child: Text(
        _error!,
        style: const TextStyle(
          color: VendorUiColors.danger,
          fontWeight: FontWeight.w700,
        ),
      ),
    );
  }

  Widget _buildLoginBody() {
    final config = _config ?? buildDefaultAppConfig();
    return _buildScrollableShell(
      maxWidth: 520,
      child: AppPanel(
        padding: const EdgeInsets.all(24),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const PanelHeader(
              title: 'Acesso do vendedor',
              subtitle:
                  'Use o mesmo cadastro criado no desktop. O app ja abre com a conexao padrao configurada.',
              icon: Icons.lock_outline,
            ),
            const SizedBox(height: 18),
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: [
                StatusBadge(
                  label: _serverReachable ? 'Servidor online' : 'Servidor indisponivel',
                  color: _serverReachable
                      ? VendorUiColors.success
                      : VendorUiColors.danger,
                ),
              ],
            ),
            const SizedBox(height: 16),
            TextField(
              controller: _loginCtrl,
              enabled: !_loggingIn,
              textCapitalization: TextCapitalization.words,
              textInputAction: TextInputAction.next,
              decoration: const InputDecoration(
                labelText: 'Nome do vendedor',
                hintText: 'Pedro, Roberto, Escritorio...',
                prefixIcon: Icon(Icons.person_outline),
              ),
            ),
            const SizedBox(height: 12),
            TextField(
              controller: _senhaCtrl,
              enabled: !_loggingIn,
              obscureText: !_showSenha,
              textInputAction: TextInputAction.done,
              onSubmitted: (_) => _loggingIn ? null : _loginVendedor(),
              decoration: InputDecoration(
                labelText: 'Senha',
                prefixIcon: const Icon(Icons.password_outlined),
                suffixIcon: IconButton(
                  onPressed: () {
                    setState(() {
                      _showSenha = !_showSenha;
                    });
                  },
                  icon: Icon(
                    _showSenha ? Icons.visibility_off : Icons.visibility,
                  ),
                  tooltip: _showSenha ? 'Ocultar senha' : 'Mostrar senha',
                ),
              ),
            ),
            const SizedBox(height: 16),
            _buildErrorCard(),
            if (_error != null) const SizedBox(height: 16),
            SizedBox(
              width: double.infinity,
              child: ElevatedButton.icon(
                onPressed: _loggingIn ? null : _loginVendedor,
                icon: _loggingIn
                    ? const SizedBox(
                        width: 16,
                        height: 16,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    : const Icon(Icons.login),
                label: const Text('Entrar'),
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
                    label: 'API',
                    value: config.normalizedBaseUrl,
                  ),
                  if (config.cidadePadrao.trim().isNotEmpty)
                    InfoRow(
                      label: 'Cidade base',
                      value: config.cidadePadrao,
                    ),
                ],
              ),
            ),
            const SizedBox(height: 12),
            TextButton(
              onPressed: _checkingServer ? null : () => _reload(silent: true),
              child: const Text('Atualizar status da API'),
            ),
          ],
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    if (_loading) {
      return const Scaffold(
        body: Center(child: CircularProgressIndicator()),
      );
    }

    if (_session == null || _config == null || _api == null) {
      return Scaffold(
        appBar: _buildLoginAppBar(),
        body: _buildLoginBody(),
      );
    }

    return HomePage(
      key: ValueKey(
        '${_config!.normalizedBaseUrl}|${_config!.vendedorPadrao}|${_config!.cidadePadrao}|${_session!.codigo}',
      ),
      config: _config!,
      session: _session!,
      api: _api!,
      onRefreshShell: _reload,
      onLogout: _logoutVendedor,
    );
  }
}
