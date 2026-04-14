import 'dart:convert';

import 'package:shared_preferences/shared_preferences.dart';

abstract class LocalSyncStore {
  Future<void> putString(String key, String value);
  Future<String?> getString(String key);
  Future<void> putJson(String key, Object value);
  Future<dynamic> getJson(String key);
  Future<void> remove(String key);
}

class SharedPrefsSyncStore implements LocalSyncStore {
  SharedPrefsSyncStore(this._prefs);

  final SharedPreferences _prefs;

  @override
  Future<void> putString(String key, String value) async {
    await _prefs.setString(key, value);
  }

  @override
  Future<String?> getString(String key) async {
    return _prefs.getString(key);
  }

  @override
  Future<void> putJson(String key, Object value) async {
    await _prefs.setString(key, jsonEncode(value));
  }

  @override
  Future<dynamic> getJson(String key) async {
    final raw = _prefs.getString(key);
    if (raw == null || raw.trim().isEmpty) return null;
    return jsonDecode(raw);
  }

  @override
  Future<void> remove(String key) async {
    await _prefs.remove(key);
  }
}
