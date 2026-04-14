// Service Flutter de referencia para o APP_VENDEDOR_API.
//
// Regra arquitetural obrigatoria:
// - `main.py`, `flutter_application_1` (projeto atual `flutter_vendedor_app`)
//   e este service precisam falar com a mesma API central.
// - A fonte de verdade compartilhada e `api_server.py`.
// - Para endpoints `desktop/*`, use a mesma `baseUrl` e o mesmo
//   `X-Desktop-Secret` do desktop.
// - Para `auth/vendedor/login` e `vendedor/*`, use token `Bearer`.
//
// Este arquivo deve seguir o contrato do service real em:
// `flutter_vendedor_app/lib/services/vendedor_api_service.dart`
//
// Regras de operacao:
// 1) Primeiro acesso obrigatoriamente online para baixar base e marcar app habilitado.
// 2) Depois disso, app opera offline e coloca mutacoes em fila local.
// 3) Ao reconectar, a fila e reenviada para o mesmo servidor/API central.
//
// Este arquivo e agnostico de storage/rede reativa.
// Integre LocalSyncStore com Sqflite/Hive/Isar e dispare flushPendingQueue()
// quando o app detectar reconexao (ex.: connectivity_plus).

import 'dart:convert';
import 'dart:math';
import 'package:http/http.dart' as http;

const String _kFirstOnlineSyncDone = 'first_online_sync_done';
const String _kQueueKey = 'offline_mutation_queue';
const String _kRefMotoristas = 'ref_motoristas';
const String _kRefVeiculos = 'ref_veiculos';
const String _kRefAjudantes = 'ref_ajudantes';
const String _kRefClientes = 'ref_clientes';

abstract class LocalSyncStore {
  Future<void> putString(String key, String value);
  Future<String?> getString(String key);
  Future<void> putJson(String key, Object value);
  Future<dynamic> getJson(String key);
}

class OfflineMutation {
  OfflineMutation({
    required this.id,
    required this.method,
    required this.path,
    required this.body,
    required this.createdAtIso,
    required this.retries,
  });

  final String id;
  final String method;
  final String path;
  final Map<String, dynamic> body;
  final String createdAtIso;
  final int retries;

  Map<String, dynamic> toJson() => <String, dynamic>{
        'id': id,
        'method': method,
        'path': path,
        'body': body,
        'created_at': createdAtIso,
        'retries': retries,
      };

  static OfflineMutation fromJson(Map<String, dynamic> json) {
    return OfflineMutation(
      id: (json['id'] ?? '').toString(),
      method: (json['method'] ?? 'POST').toString().toUpperCase(),
      path: (json['path'] ?? '').toString(),
      body: (json['body'] is Map<String, dynamic>)
          ? (json['body'] as Map<String, dynamic>)
          : <String, dynamic>{},
      createdAtIso: (json['created_at'] ?? '').toString(),
      retries: int.tryParse((json['retries'] ?? '0').toString()) ?? 0,
    );
  }

  OfflineMutation withRetries(int nextRetries) => OfflineMutation(
        id: id,
        method: method,
        path: path,
        body: body,
        createdAtIso: createdAtIso,
        retries: nextRetries,
      );
}

class FlushReport {
  FlushReport({
    required this.sent,
    required this.failed,
    required this.pending,
  });

  final int sent;
  final int failed;
  final int pending;
}

class VendedorApiService {
  VendedorApiService({
    required this.baseUrl,
    required this.desktopSecret,
    required this.store,
    this.timeoutSeconds = 25,
    this.maxRetries = 12,
  });

  final String baseUrl;
  final String desktopSecret;
  final LocalSyncStore store;
  final int timeoutSeconds;
  final int maxRetries;

  Uri _uri(String path) {
    final p = path.startsWith('/') ? path : '/$path';
    return Uri.parse('${baseUrl.trim().replaceAll(RegExp(r'/+$'), '')}$p');
  }

  Map<String, String> _headers() => <String, String>{
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'X-Desktop-Secret': desktopSecret,
      };

  Future<dynamic> _decode(http.Response r) async {
    if (r.body.trim().isEmpty) return <String, dynamic>{};
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

  Future<dynamic> _sendMutation(OfflineMutation m) async {
    if (m.method == 'POST') return _post(m.path, m.body);
    throw Exception('Metodo nao suportado na fila offline: ${m.method}');
  }

  String _newMutationId() {
    final now = DateTime.now().microsecondsSinceEpoch;
    final rnd = Random().nextInt(1 << 31);
    return '${now}_$rnd';
  }

  Future<List<OfflineMutation>> _loadQueue() async {
    final raw = await store.getJson(_kQueueKey);
    if (raw is! List) return <OfflineMutation>[];
    final out = <OfflineMutation>[];
    for (final item in raw) {
      if (item is Map<String, dynamic>) out.add(OfflineMutation.fromJson(item));
    }
    return out;
  }

  Future<void> _saveQueue(List<OfflineMutation> q) async {
    await store.putJson(_kQueueKey, q.map((e) => e.toJson()).toList());
  }

  Future<void> _enqueueMutation({
    required String method,
    required String path,
    required Map<String, dynamic> body,
  }) async {
    final q = await _loadQueue();
    q.add(
      OfflineMutation(
        id: _newMutationId(),
        method: method.toUpperCase(),
        path: path,
        body: body,
        createdAtIso: DateTime.now().toUtc().toIso8601String(),
        retries: 0,
      ),
    );
    await _saveQueue(q);
  }

  Future<bool> hasFirstOnlineSyncDone() async {
    final v = await store.getString(_kFirstOnlineSyncDone);
    return (v ?? '') == '1';
  }

  Future<bool> isOnlineServerReachable() async {
    try {
      await _get('/ping');
      return true;
    } catch (_) {
      return false;
    }
  }

  // Primeiro acesso obrigatorio online:
  // baixa base e salva cache local para uso offline.
  Future<void> bootstrapFirstOnlineSync({
    required String vendedorPadrao,
    String cidadePadrao = '',
  }) async {
    final online = await isOnlineServerReachable();
    if (!online) {
      throw Exception(
        'Primeiro acesso exige conexao com servidor. '
        'Conecte a internet/rede e tente novamente.',
      );
    }

    final motoristas = await listarMotoristasOnline();
    final veiculos = await listarVeiculosOnline();
    final ajudantes = await listarAjudantesOnline();
    final clientes = await buscarClientesOnline(
      '',
      vendedor: '',
      cidade: '',
      ordem: 'nome',
      limit: 1000,
    );

    await store.putJson(_kRefMotoristas, motoristas);
    await store.putJson(_kRefVeiculos, veiculos);
    await store.putJson(_kRefAjudantes, ajudantes);
    await store.putJson(_kRefClientes, clientes);
    await store.putString(_kFirstOnlineSyncDone, '1');
  }

  // Flush manual/automatico ao detectar reconexao.
  Future<FlushReport> flushPendingQueue() async {
    final canSend = await isOnlineServerReachable();
    final q = await _loadQueue();
    if (!canSend || q.isEmpty) {
      return FlushReport(sent: 0, failed: 0, pending: q.length);
    }

    final keep = <OfflineMutation>[];
    var sent = 0;
    var failed = 0;

    for (final m in q) {
      try {
        await _sendMutation(m);
        sent += 1;
      } catch (_) {
        final next = m.withRetries(m.retries + 1);
        if (next.retries < maxRetries) {
          keep.add(next);
        }
        failed += 1;
      }
    }

    await _saveQueue(keep);
    return FlushReport(sent: sent, failed: failed, pending: keep.length);
  }

  // -------------------------
  // Leitura online + cache local
  // -------------------------
  Future<List<dynamic>> listarMotoristasOnline() async {
    final d = await _get('/desktop/cadastros/motoristas');
    final out = (d is List) ? d : <dynamic>[];
    await store.putJson(_kRefMotoristas, out);
    return out;
  }

  Future<List<dynamic>> listarVeiculosOnline() async {
    final d = await _get('/desktop/cadastros/veiculos');
    final out = (d is List) ? d : <dynamic>[];
    await store.putJson(_kRefVeiculos, out);
    return out;
  }

  Future<List<dynamic>> listarAjudantesOnline() async {
    final d = await _get('/desktop/cadastros/ajudantes');
    final out = (d is List) ? d : <dynamic>[];
    await store.putJson(_kRefAjudantes, out);
    return out;
  }

  Future<List<dynamic>> buscarClientesOnline(
    String q, {
    String vendedor = '',
    String cidade = '',
    String ordem = 'nome',
    int limit = 300,
  }) async {
    final qp = Uri.encodeQueryComponent(q.trim());
    final vend = Uri.encodeQueryComponent(vendedor.trim());
    final cid = Uri.encodeQueryComponent(cidade.trim());
    final ord = Uri.encodeQueryComponent(
      ordem.trim().isEmpty ? 'nome' : ordem.trim(),
    );
    final lim = limit <= 0 ? 300 : limit;
    final d = await _get(
      '/desktop/clientes/base?q=$qp&vendedor=$vend&cidade=$cid&ordem=$ord&limit=$lim',
    );
    final out = (d is List) ? d : <dynamic>[];
    await store.putJson(_kRefClientes, out);
    return out;
  }

  Future<List<dynamic>> listarMotoristas() async {
    if (await isOnlineServerReachable()) {
      return listarMotoristasOnline();
    }
    final d = await store.getJson(_kRefMotoristas);
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> listarVeiculos() async {
    if (await isOnlineServerReachable()) {
      return listarVeiculosOnline();
    }
    final d = await store.getJson(_kRefVeiculos);
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> listarAjudantes() async {
    if (await isOnlineServerReachable()) {
      return listarAjudantesOnline();
    }
    final d = await store.getJson(_kRefAjudantes);
    return (d is List) ? d : <dynamic>[];
  }

  Future<List<dynamic>> buscarClientes(
    String q, {
    String vendedor = '',
    String cidade = '',
    String ordem = 'nome',
    int limit = 300,
  }) async {
    if (await isOnlineServerReachable()) {
      return buscarClientesOnline(
        q,
        vendedor: vendedor,
        cidade: cidade,
        ordem: ordem,
        limit: limit,
      );
    }
    final d = await store.getJson(_kRefClientes);
    final all = (d is List) ? d : <dynamic>[];
    if (q.trim().isEmpty) return all;
    final qUp = q.trim().toUpperCase();
    return all.where((e) {
      final m = (e is Map) ? e : <String, dynamic>{};
      final cod = (m['cod_cliente'] ?? '').toString().toUpperCase();
      final nome = (m['nome_cliente'] ?? '').toString().toUpperCase();
      return cod.contains(qUp) || nome.contains(qUp);
    }).toList();
  }

  // -------------------------
  // Escrita offline-first
  // -------------------------
  Future<Map<String, dynamic>> criarAvulsa(Map<String, dynamic> payload) async {
    if (await isOnlineServerReachable()) {
      final d = await _post('/desktop/avulsas', payload);
      return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
    }
    await _enqueueMutation(
      method: 'POST',
      path: '/desktop/avulsas',
      body: payload,
    );
    return <String, dynamic>{
      'ok': true,
      'queued_offline': true,
      'message': 'Sem conexao. Avulsa registrada na fila para sincronizar depois.',
    };
  }

  Future<Map<String, dynamic>> conciliarAvulsa({
    required String codigoAvulsa,
    required String codigoProgramacaoOficial,
    required String usuario,
  }) async {
    final path = '/desktop/avulsas/${codigoAvulsa.trim()}/conciliar';
    final body = <String, dynamic>{
      'codigo_programacao_oficial': codigoProgramacaoOficial.trim(),
      'usuario': usuario.trim(),
    };

    if (await isOnlineServerReachable()) {
      final d = await _post(path, body);
      return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
    }
    await _enqueueMutation(method: 'POST', path: path, body: body);
    return <String, dynamic>{
      'ok': true,
      'queued_offline': true,
      'message':
          'Sem conexao. Conciliacao registrada na fila para sincronizar depois.',
    };
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
    final qs = q.entries
        .map((e) => '${e.key}=${Uri.encodeQueryComponent(e.value)}')
        .join('&');
    final d = await _get('/desktop/avulsas?$qs');
    return (d is List) ? d : <dynamic>[];
  }

  Future<Map<String, dynamic>> detalheAvulsa(String codigoAvulsa) async {
    final d = await _get('/desktop/avulsas/${codigoAvulsa.trim()}');
    return (d is Map<String, dynamic>) ? d : <String, dynamic>{};
  }
}
