import 'package:flutter/foundation.dart';

import '../models/app_config.dart';

AppConfig buildDefaultAppConfig() {
  return AppConfig(
    baseUrl: _resolveBaseUrl(),
    desktopSecret: _resolveDesktopSecret(),
    vendedorPadrao: '',
    vendedorLogin: _resolveDefaultLogin(),
    cidadePadrao: _resolveDefaultCity(),
  );
}

String _resolveBaseUrl() {
  const vendorBaseUrl = String.fromEnvironment('VENDOR_API_BASE_URL');
  const sharedBaseUrl = String.fromEnvironment('ROTA_SERVER_URL');
  final trimmed = vendorBaseUrl.trim().isNotEmpty
      ? vendorBaseUrl.trim()
      : sharedBaseUrl.trim();
  if (trimmed.isNotEmpty) {
    return trimmed;
  }

  if (kDebugMode &&
      !kIsWeb &&
      defaultTargetPlatform == TargetPlatform.android) {
    return 'http://10.0.2.2:8000';
  }

  return 'https://rotahub-api.onrender.com';
}

String _resolveDesktopSecret() {
  const vendorSecret = String.fromEnvironment('VENDOR_DESKTOP_SECRET');
  const sharedSecret = String.fromEnvironment(
    'ROTA_SECRET',
    defaultValue: 'rota-secreta',
  );
  return vendorSecret.trim().isNotEmpty
      ? vendorSecret.trim()
      : sharedSecret.trim();
}

String _resolveDefaultLogin() {
  const vendorLogin = String.fromEnvironment(
    'VENDOR_LOGIN_NAME',
    defaultValue: '',
  );
  const sharedLogin = String.fromEnvironment(
    'ROTA_VENDEDOR_LOGIN',
    defaultValue: '',
  );
  return vendorLogin.trim().isNotEmpty
      ? vendorLogin.trim()
      : sharedLogin.trim();
}

String _resolveDefaultCity() {
  const vendorCity = String.fromEnvironment(
    'VENDOR_CIDADE_BASE',
    defaultValue: '',
  );
  const sharedCity = String.fromEnvironment(
    'ROTA_CIDADE_BASE',
    defaultValue: '',
  );
  final city = vendorCity.trim().isNotEmpty
      ? vendorCity.trim()
      : sharedCity.trim();
  return city.toUpperCase();
}
