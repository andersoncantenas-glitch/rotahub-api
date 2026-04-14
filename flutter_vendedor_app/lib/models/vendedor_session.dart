class VendedorSession {
  const VendedorSession({
    required this.token,
    required this.codigo,
    required this.nome,
  });

  final String token;
  final String codigo;
  final String nome;

  bool get isValid =>
      token.trim().isNotEmpty &&
      codigo.trim().isNotEmpty &&
      nome.trim().isNotEmpty;

  Map<String, dynamic> toJson() => <String, dynamic>{
        'token': token.trim(),
        'codigo': codigo.trim().toUpperCase(),
        'nome': nome.trim().toUpperCase(),
      };

  factory VendedorSession.fromJson(Map<String, dynamic> json) {
    return VendedorSession(
      token: (json['token'] ?? '').toString(),
      codigo: (json['codigo'] ?? '').toString(),
      nome: (json['nome'] ?? '').toString(),
    );
  }
}
