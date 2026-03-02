// Exemplo de service para o novo App Vendedor (projeto Flutter separado).
// Ajuste caminhos/imports conforme sua estrutura.

import 'dart:convert';
import 'package:http/http.dart' as http;

class VendedorApiService {
  VendedorApiService({
    required this.baseUrl,
    required this.desktopSecret,
    this.timeoutSeconds = 25,
  });

  final String baseUrl;
  final String desktopSecret;
  final int timeoutSeconds;

  Uri _uri(String path) {
    final p = path.startsWith('/') ? path : '/$path';
    return Uri.parse('${baseUrl.trim().replaceAll(RegExp(r"/+$"), "")}$p');
  }

  Map<String, String> _headers() => {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'X-Desktop-Secret': desktopSecret,
      };

  Future<dynamic> _decode(http.Response r) async {
    if (r.body.trim().isEmpty) return {};
    return jsonDecode(r.body);
  }

  Future<dynamic> _get(String path) async {
    final r = await http
        .get(_uri(path), headers: _headers())
        .timeout(Duration(seconds: timeoutSeconds));
    if (r.statusCode < 200 || r.statusCode >= 300) {
      throw Exception('GET $path falhou (${r.statusCode}): ${r.body}');
    }
    return _decode(r);
  }

  Future<dynamic> _post(String path, Map<String, dynamic> body) async {
    final r = await http
        .post(_uri(path), headers: _headers(), body: jsonEncode(body))
        .timeout(Duration(seconds: timeoutSeconds));
    if (r.statusCode < 200 || r.statusCode >= 300) {
      throw Exception('POST $path falhou (${r.statusCode}): ${r.body}');
    }
    return _decode(r);
  }

  // Apoio de cadastro
  Future<List<dynamic>> listarMotoristas() async {
    final d = await _get('/desktop/cadastros/motoristas');
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> listarVeiculos() async {
    final d = await _get('/desktop/cadastros/veiculos');
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> listarAjudantes() async {
    final d = await _get('/desktop/cadastros/ajudantes');
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> buscarClientes(
    String q, {
    String vendedor = '',
    String cidade = '',
    String ordem = 'nome',
    int limit = 300,
  }) async {
    final qp = Uri.encodeQueryComponent(q.trim());
    final vend = Uri.encodeQueryComponent(vendedor.trim());
    final cid = Uri.encodeQueryComponent(cidade.trim());
    final ord = Uri.encodeQueryComponent(ordem.trim().isEmpty ? 'nome' : ordem.trim());
    final lim = limit <= 0 ? 300 : limit;
    final d = await _get(
      '/desktop/clientes/base?q=$qp&vendedor=$vend&cidade=$cid&ordem=$ord&limit=$lim',
    );
    return (d is List) ? d : <dynamic>[];
  }

  // Avulsas
  Future<Map<String, dynamic>> criarAvulsa(Map<String, dynamic> payload) async {
    final d = await _post('/desktop/avulsas', payload);
    return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
  }

  Future<List<dynamic>> listarAvulsas({
    String status = '',
    String dataDe = '',
    String dataAte = '',
  }) async {
    final q = <String, String>{};
    if (status.trim().isNotEmpty) q['status'] = status.trim();
    if (dataDe.trim().isNotEmpty) q['data_de'] = dataDe.trim();
    if (dataAte.trim().isNotEmpty) q['data_ate'] = dataAte.trim();
    q['limit'] = '200';
    final qs = q.entries.map((e) => '${e.key}=${Uri.encodeQueryComponent(e.value)}').join('&');
    final d = await _get('/desktop/avulsas?$qs');
    return (d is List) ? d : <dynamic>[];
  }

  Future<Map<String, dynamic>> detalheAvulsa(String codigoAvulsa) async {
    final d = await _get('/desktop/avulsas/${codigoAvulsa.trim()}');
    return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
  }

  Future<Map<String, dynamic>> conciliarAvulsa({
    required String codigoAvulsa,
    required String codigoProgramacaoOficial,
    required String usuario,
  }) async {
    final d = await _post('/desktop/avulsas/${codigoAvulsa.trim()}/conciliar', {
      'codigo_programacao_oficial': codigoProgramacaoOficial.trim(),
      'usuario': usuario.trim(),
    });
    return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
  }
}
