import 'dart:convert';

import 'package:shared_preferences/shared_preferences.dart';

import '../models/app_config.dart';

class AppConfigStore {
  static const String _kConfigKey = 'vendedor_app_config';
  static const String _kFirstOnlineSyncDone = 'first_online_sync_done';
  static const List<String> _kResetKeys = <String>[
    _kFirstOnlineSyncDone,
    'vendedor_auth_session',
    'offline_mutation_queue',
    'ref_motoristas',
    'ref_veiculos',
    'ref_ajudantes',
    'ref_clientes',
    'cache_avulsas',
    'local_draft_avulsas',
  ];

  static Future<void> _clearRuntimeCache(SharedPreferences prefs) async {
    final keys = prefs.getKeys().toList(growable: false);
    for (final key in keys) {
      if (_kResetKeys.contains(key) || key.startsWith('cache_avulsa_')) {
        await prefs.remove(key);
      }
    }
  }

  static Future<AppConfig?> load() async {
    final prefs = await SharedPreferences.getInstance();
    final raw = prefs.getString(_kConfigKey);
    if (raw == null || raw.trim().isEmpty) return null;
    final decoded = jsonDecode(raw);
    if (decoded is! Map) return null;
    final map = decoded.map((key, value) => MapEntry(key.toString(), value));
    return AppConfig.fromJson(map);
  }

  static Future<void> save(AppConfig config) async {
    final prefs = await SharedPreferences.getInstance();
    final current = await load();
    final next = config.copyWith(
      baseUrl: config.normalizedBaseUrl,
      vendedorPadrao: config.vendedorPadrao.trim().toUpperCase(),
      vendedorLogin: config.vendedorLogin.trim(),
      cidadePadrao: config.cidadePadrao.trim().toUpperCase(),
    );
    final resetBootstrap = current == null ||
        current.normalizedBaseUrl != next.normalizedBaseUrl ||
        current.desktopSecret.trim() != next.desktopSecret.trim() ||
        current.vendedorPadrao.trim().toUpperCase() !=
            next.vendedorPadrao.trim().toUpperCase() ||
        current.vendedorLogin.trim() != next.vendedorLogin.trim() ||
        current.cidadePadrao.trim().toUpperCase() !=
            next.cidadePadrao.trim().toUpperCase();

    await prefs.setString(_kConfigKey, jsonEncode(next.toJson()));
    if (resetBootstrap) {
      await _clearRuntimeCache(prefs);
    }
  }

  static Future<void> clear() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.remove(_kConfigKey);
    await _clearRuntimeCache(prefs);
  }
}
