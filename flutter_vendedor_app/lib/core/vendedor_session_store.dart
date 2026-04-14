import 'dart:convert';

import 'package:shared_preferences/shared_preferences.dart';

import '../models/vendedor_session.dart';

class VendedorSessionStore {
  static const String key = 'vendedor_auth_session';

  static Future<VendedorSession?> load() async {
    final prefs = await SharedPreferences.getInstance();
    final raw = prefs.getString(key);
    if (raw == null || raw.trim().isEmpty) return null;
    final decoded = jsonDecode(raw);
    if (decoded is! Map) return null;
    final map = decoded.map((key, value) => MapEntry(key.toString(), value));
    final session = VendedorSession.fromJson(map);
    return session.isValid ? session : null;
  }

  static Future<void> save(VendedorSession session) async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setString(key, jsonEncode(session.toJson()));
  }

  static Future<void> clear() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.remove(key);
  }
}
