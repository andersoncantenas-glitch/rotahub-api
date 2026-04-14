import 'package:flutter/material.dart';

import '../core/app_config_store.dart';
import '../models/app_config.dart';
import '../ui/app_visuals.dart';

class ConfigPage extends StatefulWidget {
  const ConfigPage({
    super.key,
    this.initialConfig,
  });

  final AppConfig? initialConfig;

  @override
  State<ConfigPage> createState() => _ConfigPageState();
}

class _ConfigPageState extends State<ConfigPage> {
  final _formKey = GlobalKey<FormState>();

  late final TextEditingController _baseUrlCtrl;
  late final TextEditingController _secretCtrl;
  late final TextEditingController _vendedorLoginCtrl;
  late final TextEditingController _cidadeCtrl;
  bool _saving = false;

  @override
  void initState() {
    super.initState();
    _baseUrlCtrl = TextEditingController(
        text: widget.initialConfig?.normalizedBaseUrl ?? '');
    _secretCtrl =
        TextEditingController(text: widget.initialConfig?.desktopSecret ?? '');
    _vendedorLoginCtrl =
        TextEditingController(text: widget.initialConfig?.vendedorLogin ?? '');
    _cidadeCtrl =
        TextEditingController(text: widget.initialConfig?.cidadePadrao ?? '');
  }

  @override
  void dispose() {
    _baseUrlCtrl.dispose();
    _secretCtrl.dispose();
    _vendedorLoginCtrl.dispose();
    _cidadeCtrl.dispose();
    super.dispose();
  }

  Future<void> _save() async {
    if (!_formKey.currentState!.validate()) return;

    setState(() {
      _saving = true;
    });

    final config = AppConfig(
      baseUrl: _baseUrlCtrl.text.trim(),
      desktopSecret: _secretCtrl.text.trim(),
      vendedorPadrao: widget.initialConfig?.vendedorPadrao ?? '',
      vendedorLogin: _vendedorLoginCtrl.text.trim(),
      cidadePadrao: _cidadeCtrl.text.trim().toUpperCase(),
    );

    try {
      await AppConfigStore.save(config);
      if (!mounted) return;
      Navigator.of(context).pop(config);
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao salvar: $error')),
      );
    } finally {
      if (mounted) {
        setState(() {
          _saving = false;
        });
      }
    }
  }

  Future<void> _clear() async {
    await AppConfigStore.clear();
    if (!mounted) return;
    Navigator.of(context).pop();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Configuracao'),
        actions: [
          if (widget.initialConfig != null)
            IconButton(
              onPressed: _saving ? null : _clear,
              icon: const Icon(Icons.delete_outline),
              tooltip: 'Restaurar padrao',
            ),
        ],
      ),
      body: SafeArea(
        child: SingleChildScrollView(
          padding: const EdgeInsets.all(16),
          child: Center(
            child: ConstrainedBox(
              constraints: const BoxConstraints(maxWidth: 640),
              child: Form(
                key: _formKey,
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    AppPanel(
                      padding: const EdgeInsets.all(20),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const PanelHeader(
                            title: 'Conexao do app vendedor',
                            subtitle:
                                'Use a mesma API e o mesmo X-Desktop-Secret do desktop/main.py. Altere esses dados apenas quando mudar o ambiente compartilhado.',
                            icon: Icons.settings_ethernet,
                          ),
                          const SizedBox(height: 20),
                          TextFormField(
                            controller: _baseUrlCtrl,
                            keyboardType: TextInputType.url,
                            decoration: const InputDecoration(
                              labelText: 'URL da API',
                              hintText: 'http://192.168.0.10:8000',
                              helperText:
                                  'Deve ser a mesma baseUrl usada pelo desktop.',
                            ),
                            validator: (value) {
                              final text = (value ?? '').trim();
                              if (text.isEmpty) {
                                return 'Informe a URL da API.';
                              }
                              if (!text.startsWith('http://') &&
                                  !text.startsWith('https://')) {
                                return 'Use http:// ou https://';
                              }
                              return null;
                            },
                          ),
                          const SizedBox(height: 16),
                          TextFormField(
                            controller: _secretCtrl,
                            decoration: const InputDecoration(
                              labelText: 'X-Desktop-Secret',
                              helperText:
                                  'Use exatamente o mesmo valor configurado em ROTA_SECRET.',
                            ),
                            validator: (value) {
                              if ((value ?? '').trim().isEmpty) {
                                return 'Informe o secret do backend.';
                              }
                              return null;
                            },
                          ),
                          const SizedBox(height: 16),
                          TextFormField(
                            controller: _vendedorLoginCtrl,
                            decoration: const InputDecoration(
                              labelText: 'Nome de login (opcional)',
                              hintText: 'Pedro, Roberto, Escritorio...',
                            ),
                          ),
                          const SizedBox(height: 16),
                          TextFormField(
                            controller: _cidadeCtrl,
                            decoration: const InputDecoration(
                              labelText: 'Cidade base (opcional)',
                            ),
                          ),
                          const SizedBox(height: 24),
                          ElevatedButton.icon(
                            onPressed: _saving ? null : _save,
                            icon: _saving
                                ? const SizedBox(
                                    width: 16,
                                    height: 16,
                                    child: CircularProgressIndicator(
                                      strokeWidth: 2,
                                    ),
                                  )
                                : const Icon(Icons.save),
                            label: const Text('Salvar configuracao'),
                          ),
                        ],
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ),
      ),
    );
  }
}
