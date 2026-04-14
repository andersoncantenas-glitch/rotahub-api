class AppConfig {
  const AppConfig({
    required this.baseUrl,
    required this.desktopSecret,
    required this.vendedorPadrao,
    required this.vendedorLogin,
    required this.cidadePadrao,
  });

  final String baseUrl;
  final String desktopSecret;
  final String vendedorPadrao;
  final String vendedorLogin;
  final String cidadePadrao;

  String get normalizedBaseUrl => baseUrl.trim().replaceAll(RegExp(r'/+$'), '');

  bool get isComplete =>
      normalizedBaseUrl.isNotEmpty &&
      desktopSecret.trim().isNotEmpty;

  AppConfig copyWith({
    String? baseUrl,
    String? desktopSecret,
    String? vendedorPadrao,
    String? vendedorLogin,
    String? cidadePadrao,
  }) {
    return AppConfig(
      baseUrl: baseUrl ?? this.baseUrl,
      desktopSecret: desktopSecret ?? this.desktopSecret,
      vendedorPadrao: vendedorPadrao ?? this.vendedorPadrao,
      vendedorLogin: vendedorLogin ?? this.vendedorLogin,
      cidadePadrao: cidadePadrao ?? this.cidadePadrao,
    );
  }

  Map<String, dynamic> toJson() => <String, dynamic>{
        'base_url': normalizedBaseUrl,
        'desktop_secret': desktopSecret.trim(),
        'vendedor_padrao': vendedorPadrao.trim().toUpperCase(),
        'vendedor_login': vendedorLogin.trim(),
        'cidade_padrao': cidadePadrao.trim().toUpperCase(),
      };

  factory AppConfig.fromJson(Map<String, dynamic> json) {
    return AppConfig(
      baseUrl: (json['base_url'] ?? '').toString(),
      desktopSecret: (json['desktop_secret'] ?? '').toString(),
      vendedorPadrao: (json['vendedor_padrao'] ?? '').toString(),
      vendedorLogin: (json['vendedor_login'] ?? '').toString(),
      cidadePadrao: (json['cidade_padrao'] ?? '').toString(),
    );
  }
}
