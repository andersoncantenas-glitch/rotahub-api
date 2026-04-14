import 'dart:convert';
import 'dart:math';

import 'package:http/http.dart' as http;

import '../core/shared_prefs_sync_store.dart';

const String firstOnlineSyncDoneKey = 'first_online_sync_done';
const String _kQueueKey = 'offline_mutation_queue';
const String _kRefMotoristas = 'ref_motoristas';
const String _kRefVeiculos = 'ref_veiculos';
const String _kRefAjudantes = 'ref_ajudantes';
const String _kRefClientes = 'ref_clientes';
const String _kCacheAvulsas = 'cache_avulsas';
const String _kCacheAvulsaPrefix = 'cache_avulsa_';
const String _kLocalDraftAvulsas = 'local_draft_avulsas';
const String _kSalesDraftItems = 'sales_draft_items_v1';
const String _kVendedorSessionKey = 'vendedor_auth_session';

Map<String, dynamic> _asStringMap(dynamic value) {
  if (value is Map<String, dynamic>) return Map<String, dynamic>.from(value);
  if (value is Map) {
    return value.map((key, data) => MapEntry(key.toString(), data));
  }
  return <String, dynamic>{};
}

List<Map<String, dynamic>> _asMapList(dynamic value) {
  if (value is! List) return <Map<String, dynamic>>[];
  return value.map<Map<String, dynamic>>((item) => _asStringMap(item)).toList();
}

String _upperValue(dynamic value) =>
    (value ?? '').toString().trim().toUpperCase();

String _textValue(dynamic value) => (value ?? '').toString().trim();

int _intValue(dynamic value) {
  if (value == null) return 0;
  if (value is int) return value;
  final raw = value.toString().trim();
  if (raw.isEmpty) return 0;
  return int.tryParse(raw) ?? 0;
}

double _doubleValue(dynamic value) {
  if (value == null) return 0;
  if (value is double) return value;
  if (value is int) return value.toDouble();
  final raw = value.toString().trim().replaceAll(',', '.');
  if (raw.isEmpty) return 0;
  return double.tryParse(raw) ?? 0;
}

DateTime? _parseDateOnly(dynamic value) {
  final raw = (value ?? '').toString().trim();
  if (raw.isEmpty) return null;
  try {
    if (RegExp(r'^\d{4}-\d{2}-\d{2}$').hasMatch(raw)) {
      return DateTime.parse('${raw}T00:00:00');
    }
    if (RegExp(r'^\d{2}/\d{2}/\d{4}$').hasMatch(raw)) {
      final parts = raw.split('/');
      return DateTime(
        int.parse(parts[2]),
        int.parse(parts[1]),
        int.parse(parts[0]),
      );
    }
    final parsed = DateTime.tryParse(raw.replaceFirst(' ', 'T'));
    if (parsed == null) return null;
    return DateTime(parsed.year, parsed.month, parsed.day);
  } catch (_) {
    return null;
  }
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
      body: _asStringMap(json['body']),
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
    this.allowOfflineFallback = false,
  });

  final String baseUrl;
  final String desktopSecret;
  final LocalSyncStore store;
  final int timeoutSeconds;
  final int maxRetries;
  final bool allowOfflineFallback;

  Uri _uri(String path) {
    final cleanBase = baseUrl.trim().replaceAll(RegExp(r'/+$'), '');
    final cleanPath = path.startsWith('/') ? path : '/$path';
    return Uri.parse('$cleanBase$cleanPath');
  }

  Map<String, String> _headers() => <String, String>{
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'X-Desktop-Secret': desktopSecret,
      };

  Future<dynamic> _decode(http.Response response) async {
    if (response.body.trim().isEmpty) return <String, dynamic>{};
    return jsonDecode(response.body);
  }

  Future<dynamic> _get(String path) async {
    final response = await http
        .get(_uri(path), headers: _headers())
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'GET $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<dynamic> _post(String path, Map<String, dynamic> body) async {
    final response = await http
        .post(
          _uri(path),
          headers: _headers(),
          body: jsonEncode(body),
        )
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'POST $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<Map<String, String>> _vendedorHeaders() async {
    final session = _asStringMap(await store.getJson(_kVendedorSessionKey));
    final token = _textValue(session['token']);
    if (token.isEmpty) {
      throw Exception('Sessao do vendedor ausente. Faca login novamente.');
    }
    return <String, String>{
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Authorization': 'Bearer $token',
    };
  }

  Future<dynamic> _vendorGet(String path) async {
    final response = await http
        .get(_uri(path), headers: await _vendedorHeaders())
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'GET $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<dynamic> _vendorPost(String path, Map<String, dynamic> body) async {
    final response = await http
        .post(
          _uri(path),
          headers: await _vendedorHeaders(),
          body: jsonEncode(body),
        )
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'POST $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<dynamic> _vendorPatch(String path, Map<String, dynamic> body) async {
    final response = await http
        .patch(
          _uri(path),
          headers: await _vendedorHeaders(),
          body: jsonEncode(body),
        )
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'PATCH $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<dynamic> _vendorDelete(String path) async {
    final response = await http
        .delete(_uri(path), headers: await _vendedorHeaders())
        .timeout(Duration(seconds: timeoutSeconds));
    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception(
          'DELETE $path falhou (${response.statusCode}): ${response.body}');
    }
    return _decode(response);
  }

  Future<Map<String, dynamic>> autenticarVendedor({
    required String identificador,
    required String senha,
  }) async {
    final payload = await _post(
      '/auth/vendedor/login',
      <String, dynamic>{
        'codigo': identificador.trim(),
        'senha': senha.trim(),
      },
    );
    return _asStringMap(payload);
  }

  Future<dynamic> _sendMutation(OfflineMutation mutation) async {
    if (mutation.method == 'POST') {
      return _post(mutation.path, mutation.body);
    }
    throw Exception('Metodo nao suportado na fila offline: ${mutation.method}');
  }

  String _newMutationId() {
    final now = DateTime.now().microsecondsSinceEpoch;
    final rnd = Random().nextInt(1 << 31);
    return '${now}_$rnd';
  }

  String _newLocalId(String prefix) {
    final now = DateTime.now().microsecondsSinceEpoch;
    final rnd = Random().nextInt(1 << 20).toRadixString(16).toUpperCase();
    return '$prefix$now$rnd';
  }

  Future<List<OfflineMutation>> _loadQueue() async {
    final raw = await store.getJson(_kQueueKey);
    if (raw is! List) return <OfflineMutation>[];
    return raw
        .map<OfflineMutation>(
            (item) => OfflineMutation.fromJson(_asStringMap(item)))
        .toList();
  }

  Future<void> _saveQueue(List<OfflineMutation> queue) async {
    await store.putJson(
        _kQueueKey, queue.map((item) => item.toJson()).toList());
  }

  Future<String> _enqueueMutation({
    required String method,
    required String path,
    required Map<String, dynamic> body,
  }) async {
    final queue = await _loadQueue();
    final mutationId = _newMutationId();
    queue.add(
      OfflineMutation(
        id: mutationId,
        method: method.toUpperCase(),
        path: path,
        body: body,
        createdAtIso: DateTime.now().toUtc().toIso8601String(),
        retries: 0,
      ),
    );
    await _saveQueue(queue);
    return mutationId;
  }

  Future<List<Map<String, dynamic>>> _loadLocalDraftAvulsas() async {
    return _asMapList(await store.getJson(_kLocalDraftAvulsas));
  }

  Future<List<Map<String, dynamic>>> _loadSalesDraftItems() async {
    final items = _asMapList(await store.getJson(_kSalesDraftItems));
    items.sort((a, b) {
      final updatedA = _textValue(a['updated_at']);
      final updatedB = _textValue(b['updated_at']);
      return updatedB.compareTo(updatedA);
    });
    return items;
  }

  Future<void> _saveSalesDraftItems(List<Map<String, dynamic>> items) async {
    await store.putJson(_kSalesDraftItems, items);
  }

  Future<void> _saveLocalDraftAvulsas(List<Map<String, dynamic>> drafts) async {
    await store.putJson(_kLocalDraftAvulsas, drafts);
  }

  String _offlineCodeFromMutationId(String mutationId) {
    final base = mutationId.split('_').first;
    return 'OFFLINE-$base';
  }

  Map<String, dynamic> _buildDraftHeader(
    String mutationId,
    Map<String, dynamic> payload,
  ) {
    return <String, dynamic>{
      'offline_id': mutationId,
      'codigo_avulsa': _offlineCodeFromMutationId(mutationId),
      'data_programada': (payload['data_programada'] ?? '').toString(),
      'status': 'FILA_OFFLINE',
      'motorista_codigo': _upperValue(payload['motorista_codigo']),
      'motorista_nome': _upperValue(payload['motorista_nome']),
      'veiculo': _upperValue(payload['veiculo']),
      'equipe': _upperValue(payload['equipe']),
      'local_rota': _upperValue(payload['local_rota']),
      'programacao_oficial_codigo': '',
      'criado_por': _upperValue(payload['criado_por']),
      'criado_em': DateTime.now().toIso8601String(),
      'observacao': (payload['observacao'] ?? '').toString().trim(),
      'itens': _asMapList(payload['itens']),
    };
  }

  Future<void> _upsertDraftAvulsa(
    String mutationId,
    Map<String, dynamic> payload,
  ) async {
    final drafts = await _loadLocalDraftAvulsas();
    drafts.removeWhere(
        (draft) => (draft['offline_id'] ?? '').toString() == mutationId);
    drafts.insert(0, _buildDraftHeader(mutationId, payload));
    await _saveLocalDraftAvulsas(drafts);
  }

  Future<void> _removeDraftAvulsaByMutationId(String mutationId) async {
    final drafts = await _loadLocalDraftAvulsas();
    drafts.removeWhere(
        (draft) => (draft['offline_id'] ?? '').toString() == mutationId);
    await _saveLocalDraftAvulsas(drafts);
  }

  Future<Map<String, dynamic>?> _findDraftAvulsaByCode(
      String codigoAvulsa) async {
    final code = codigoAvulsa.trim().toUpperCase();
    final drafts = await _loadLocalDraftAvulsas();
    for (final draft in drafts) {
      if (_upperValue(draft['codigo_avulsa']) == code) {
        return draft;
      }
    }
    return null;
  }

  Future<void> _mergeClientCache(List<Map<String, dynamic>> incoming) async {
    final existing = _asMapList(await store.getJson(_kRefClientes));
    final merged = <String, Map<String, dynamic>>{};

    for (final item in <Map<String, dynamic>>[...existing, ...incoming]) {
      final cod = _upperValue(item['cod_cliente']);
      final nome = _upperValue(item['nome_cliente']);
      final cidade = _upperValue(item['cidade']);
      final key = cod.isNotEmpty ? cod : '$nome|$cidade';
      merged[key] = item;
    }

    final list = merged.values.toList()
      ..sort((a, b) {
        final nomeA = _upperValue(a['nome_cliente']);
        final nomeB = _upperValue(b['nome_cliente']);
        return nomeA.compareTo(nomeB);
      });
    await store.putJson(_kRefClientes, list);
  }

  int _compareAvulsaRecency(
    Map<String, dynamic> a,
    Map<String, dynamic> b,
  ) {
    final dateA = _parseDateOnly(a['data_programada']);
    final dateB = _parseDateOnly(b['data_programada']);
    if (dateA != null && dateB != null) {
      final cmp = dateB.compareTo(dateA);
      if (cmp != 0) return cmp;
    } else if (dateA != null || dateB != null) {
      return dateA == null ? 1 : -1;
    }

    final createdA = _parseDateOnly(a['criado_em']);
    final createdB = _parseDateOnly(b['criado_em']);
    if (createdA != null && createdB != null) {
      final cmp = createdB.compareTo(createdA);
      if (cmp != 0) return cmp;
    } else if (createdA != null || createdB != null) {
      return createdA == null ? 1 : -1;
    }

    return _upperValue(a['codigo_avulsa']).compareTo(_upperValue(b['codigo_avulsa']));
  }

  bool _matchesDateRange(
    Map<String, dynamic> item, {
    String dataDe = '',
    String dataAte = '',
  }) {
    final date = _parseDateOnly(item['data_programada']);
    final from = _parseDateOnly(dataDe);
    final to = _parseDateOnly(dataAte);
    if (from == null && to == null) return true;
    if (date == null) return false;
    if (from != null && date.isBefore(from)) return false;
    if (to != null && date.isAfter(to)) return false;
    return true;
  }

  Future<void> _upsertAvulsasCache(
    List<Map<String, dynamic>> incoming, {
    required bool replaceAll,
  }) async {
    final merged = <String, Map<String, dynamic>>{};
    if (!replaceAll) {
      for (final item in _asMapList(await store.getJson(_kCacheAvulsas))) {
        final code = _upperValue(item['codigo_avulsa']);
        if (code.isNotEmpty) merged[code] = _toAvulsaResumo(item);
      }
    }

    for (final item in incoming) {
      final code = _upperValue(item['codigo_avulsa']);
      if (code.isEmpty) continue;
      merged[code] = _toAvulsaResumo(item);
    }

    final list = merged.values.toList()..sort(_compareAvulsaRecency);
    await store.putJson(_kCacheAvulsas, list);
  }

  Map<String, dynamic> _toAvulsaResumo(Map<String, dynamic> avulsa) {
    return <String, dynamic>{
      'codigo_avulsa': (avulsa['codigo_avulsa'] ?? '').toString(),
      'data_programada': (avulsa['data_programada'] ?? '').toString(),
      'status': _upperValue(avulsa['status']),
      'motorista_codigo': _upperValue(avulsa['motorista_codigo']),
      'motorista_nome': _upperValue(avulsa['motorista_nome']),
      'veiculo': _upperValue(avulsa['veiculo']),
      'equipe': _upperValue(avulsa['equipe']),
      'local_rota': _upperValue(avulsa['local_rota']),
      'programacao_oficial_codigo':
          _upperValue(avulsa['programacao_oficial_codigo']),
      'criado_por': _upperValue(avulsa['criado_por']),
      'criado_em': (avulsa['criado_em'] ?? '').toString(),
      'observacao': (avulsa['observacao'] ?? '').toString(),
    };
  }

  Future<void> _cacheCreatedAvulsa(
    Map<String, dynamic> payload,
    Map<String, dynamic> response,
  ) async {
    final codigo = (response['codigo_avulsa'] ?? '').toString().trim();
    if (codigo.isEmpty) return;

    final header = _toAvulsaResumo(
      <String, dynamic>{
        ...payload,
        ...response,
        'codigo_avulsa': codigo,
        'status': 'AVULSA_ATIVA',
        'criado_em': DateTime.now().toIso8601String(),
      },
    );

    await store.putJson(
      '$_kCacheAvulsaPrefix${codigo.toUpperCase()}',
      <String, dynamic>{
        'avulsa': header,
        'itens': _asMapList(payload['itens']),
      },
    );

    final cachedList = _asMapList(await store.getJson(_kCacheAvulsas));
    cachedList.removeWhere(
      (item) => _upperValue(item['codigo_avulsa']) == codigo.toUpperCase(),
    );
    cachedList.insert(0, header);
    await store.putJson(_kCacheAvulsas, cachedList);
  }

  Future<List<Map<String, dynamic>>> _mergeAvulsasWithDrafts({
    required List<Map<String, dynamic>> serverItems,
    String status = '',
    String dataDe = '',
    String dataAte = '',
  }) async {
    final normalizedStatus = _upperValue(status);
    final drafts = await _loadLocalDraftAvulsas();
    final filteredDrafts = normalizedStatus.isEmpty
        ? drafts
        : drafts
            .where((draft) => _upperValue(draft['status']) == normalizedStatus)
            .toList();
    final rangeDrafts = filteredDrafts
        .where((draft) => _matchesDateRange(draft, dataDe: dataDe, dataAte: dataAte))
        .toList();
    final filteredServer = normalizedStatus.isEmpty
        ? serverItems
        : serverItems
            .where((item) => _upperValue(item['status']) == normalizedStatus)
            .toList();
    final rangeServer = filteredServer
        .where((item) => _matchesDateRange(item, dataDe: dataDe, dataAte: dataAte))
        .toList();

    final combined = <String, Map<String, dynamic>>{};
    for (final item in rangeServer.map(_toAvulsaResumo)) {
      final code = _upperValue(item['codigo_avulsa']);
      if (code.isNotEmpty) combined[code] = item;
    }
    for (final item in rangeDrafts.map(_toAvulsaResumo)) {
      final code = _upperValue(item['codigo_avulsa']);
      if (code.isNotEmpty) combined[code] = item;
    }

    final out = combined.values.toList()..sort(_compareAvulsaRecency);
    return out;
  }

  Future<bool> hasFirstOnlineSyncDone() async {
    final value = await store.getString(firstOnlineSyncDoneKey);
    return (value ?? '') == '1';
  }

  Future<int> pendingQueueCount() async {
    final queue = await _loadQueue();
    return queue.length;
  }

  Future<bool> isOnlineServerReachable() async {
    try {
      await _get('/ping');
      return true;
    } catch (_) {
      return false;
    }
  }

  Exception _sharedFlowOnlineException(String action) {
    return Exception(
      '$action exige conexao com o servidor. '
      'O app vendedor trabalha online para manter o vinculo com o desktop e o app do motorista.',
    );
  }

  Exception _sharedCacheUnavailableException(String action) {
    return Exception(
      '$action exige uma sincronizacao online inicial para preencher o cache local.',
    );
  }

  Future<List<Map<String, dynamic>>> _loadCachedReferenceList(
    String key, {
    required String action,
  }) async {
    final cached = _asMapList(await store.getJson(key));
    if (cached.isNotEmpty) {
      return cached;
    }
    throw _sharedCacheUnavailableException(action);
  }

  Future<bool> _shouldUseOnlineSharedFlow(String action) async {
    final online = await isOnlineServerReachable();
    if (online) return true;
    if (allowOfflineFallback) return false;
    throw _sharedFlowOnlineException(action);
  }

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
    await store.putString(firstOnlineSyncDoneKey, '1');
  }

  Future<FlushReport> flushPendingQueue() async {
    final canSend = await isOnlineServerReachable();
    final queue = await _loadQueue();
    if (!canSend || queue.isEmpty) {
      return FlushReport(sent: 0, failed: 0, pending: queue.length);
    }

    final keep = <OfflineMutation>[];
    var sent = 0;
    var failed = 0;

    for (final mutation in queue) {
      try {
        final response = await _sendMutation(mutation);
        if (mutation.path == '/desktop/avulsas') {
          await _removeDraftAvulsaByMutationId(mutation.id);
          await _cacheCreatedAvulsa(mutation.body, _asStringMap(response));
        }
        sent += 1;
      } catch (_) {
        final retries = min(mutation.retries + 1, maxRetries);
        keep.add(mutation.withRetries(retries));
        failed += 1;
      }
    }

    await _saveQueue(keep);
    return FlushReport(sent: sent, failed: failed, pending: keep.length);
  }

  Future<List<Map<String, dynamic>>> listarMotoristasOnline() async {
    final data = await _get('/desktop/cadastros/motoristas');
    final out = _asMapList(data);
    await store.putJson(_kRefMotoristas, out);
    return out;
  }

  Future<List<Map<String, dynamic>>> listarVeiculosOnline() async {
    final data = await _get('/desktop/cadastros/veiculos');
    final out = _asMapList(data);
    await store.putJson(_kRefVeiculos, out);
    return out;
  }

  Future<List<Map<String, dynamic>>> listarAjudantesOnline() async {
    final data = await _get('/desktop/cadastros/ajudantes');
    final out = _asMapList(data);
    await store.putJson(_kRefAjudantes, out);
    return out;
  }

  Future<List<Map<String, dynamic>>> buscarClientesOnline(
    String q, {
    String vendedor = '',
    String cidade = '',
    String ordem = 'nome',
    int limit = 300,
  }) async {
    final qp = Uri.encodeQueryComponent(q.trim());
    final vend = Uri.encodeQueryComponent(vendedor.trim());
    final cid = Uri.encodeQueryComponent(cidade.trim());
    final ord =
        Uri.encodeQueryComponent(ordem.trim().isEmpty ? 'nome' : ordem.trim());
    final lim = limit <= 0 ? 300 : limit;
    final data = await _get(
      '/desktop/clientes/base?q=$qp&vendedor=$vend&cidade=$cid&ordem=$ord&limit=$lim',
    );
    final out = _asMapList(data);
    await _mergeClientCache(out);
    return out;
  }

  Future<List<Map<String, dynamic>>> listarMotoristas({
    bool preferCacheWhenOffline = false,
  }) async {
    final online = await isOnlineServerReachable();
    if (online) {
      return listarMotoristasOnline();
    }
    if (preferCacheWhenOffline || allowOfflineFallback) {
      return _loadCachedReferenceList(
        _kRefMotoristas,
        action: 'Carregar motoristas',
      );
    }
    throw _sharedFlowOnlineException('Carregar motoristas');
  }

  Future<List<Map<String, dynamic>>> listarVeiculos({
    bool preferCacheWhenOffline = false,
  }) async {
    final online = await isOnlineServerReachable();
    if (online) {
      return listarVeiculosOnline();
    }
    if (preferCacheWhenOffline || allowOfflineFallback) {
      return _loadCachedReferenceList(
        _kRefVeiculos,
        action: 'Carregar veiculos',
      );
    }
    throw _sharedFlowOnlineException('Carregar veiculos');
  }

  Future<List<Map<String, dynamic>>> listarAjudantes({
    bool preferCacheWhenOffline = false,
  }) async {
    final online = await isOnlineServerReachable();
    if (online) {
      return listarAjudantesOnline();
    }
    if (preferCacheWhenOffline || allowOfflineFallback) {
      return _loadCachedReferenceList(
        _kRefAjudantes,
        action: 'Carregar ajudantes',
      );
    }
    throw _sharedFlowOnlineException('Carregar ajudantes');
  }

  Future<List<Map<String, dynamic>>> buscarClientes(
    String q, {
    String vendedor = '',
    String cidade = '',
    String ordem = 'nome',
    int limit = 300,
    bool preferCacheWhenOffline = false,
  }) async {
    final online = await isOnlineServerReachable();
    if (online) {
      return buscarClientesOnline(
        q,
        vendedor: vendedor,
        cidade: cidade,
        ordem: ordem,
        limit: limit,
      );
    }

    if (!preferCacheWhenOffline && !allowOfflineFallback) {
      throw _sharedFlowOnlineException('Consultar clientes');
    }

    final all = _asMapList(await store.getJson(_kRefClientes));
    if (all.isEmpty) {
      throw _sharedCacheUnavailableException('Consultar clientes');
    }
    final term = _upperValue(q);
    final vend = _upperValue(vendedor);
    final cid = _upperValue(cidade);
    final filtered = all.where((item) {
      final codCliente = _upperValue(item['cod_cliente']);
      final nomeCliente = _upperValue(item['nome_cliente']);
      final cidadeCliente = _upperValue(item['cidade']);
      final vendedorCliente = _upperValue(item['vendedor']);

      final matchQ = term.isEmpty ||
          codCliente.contains(term) ||
          nomeCliente.contains(term) ||
          cidadeCliente.contains(term);
      final matchVend = vend.isEmpty || vendedorCliente.contains(vend);
      final matchCid = cid.isEmpty || cidadeCliente.contains(cid);
      return matchQ && matchVend && matchCid;
    }).toList();

    if (ordem.trim().toLowerCase() == 'codigo') {
      filtered.sort(
        (a, b) => _upperValue(a['cod_cliente'])
            .compareTo(_upperValue(b['cod_cliente'])),
      );
    } else {
      filtered.sort(
        (a, b) => _upperValue(a['nome_cliente'])
            .compareTo(_upperValue(b['nome_cliente'])),
      );
    }
    return filtered.take(limit).toList();
  }

  Future<List<Map<String, dynamic>>> listarRascunhoVendas() async {
    if (await _shouldUseOnlineSharedFlow('Carregar o rascunho compartilhado')) {
      final data = await _vendorGet('/vendedor/rascunho?limit=1000');
      return _asMapList(data);
    }
    return _loadSalesDraftItems();
  }

  Future<void> adicionarVendasAoRascunho({
    required List<Map<String, dynamic>> clientes,
    required String vendedorOrigem,
    required int caixas,
    required double preco,
    String observacao = '',
    Map<String, Map<String, dynamic>> alertas = const <String, Map<String, dynamic>>{},
  }) async {
    final now = DateTime.now().toIso8601String();
    final novos = <Map<String, dynamic>>[];

    for (final cliente in clientes) {
      final codCliente = _upperValue(cliente['cod_cliente']);
      final nomeCliente = _textValue(cliente['nome_cliente']);
      if (codCliente.isEmpty || nomeCliente.isEmpty) {
        continue;
      }
      final alerta = alertas[codCliente] ?? <String, dynamic>{};
      novos.add(
        <String, dynamic>{
          'id': _newLocalId('DRV-'),
          'cod_cliente': codCliente,
          'nome_cliente': nomeCliente,
          'cidade': _upperValue(cliente['cidade']),
          'bairro': _upperValue(cliente['bairro']),
          'endereco': _textValue(cliente['endereco']),
          'vendedor_cadastro': _upperValue(cliente['vendedor']),
          'vendedor_origem': _upperValue(vendedorOrigem),
          'preco': preco,
          'caixas': caixas,
          'status': 'PENDENTE',
          'observacao': observacao.trim(),
          'criado_em': now,
          'updated_at': now,
          'alerta_codigo_programacao':
              _upperValue(alerta['codigo_programacao']),
          'alerta_status_rota': _upperValue(alerta['status']),
        },
      );
    }

    if (novos.isEmpty) return;
    if (await _shouldUseOnlineSharedFlow('Enviar vendas para o rascunho')) {
      await _vendorPost(
        '/vendedor/rascunho/itens',
        <String, dynamic>{'itens': novos},
      );
      return;
    }

    final items = await _loadSalesDraftItems();
    items.insertAll(0, novos);
    await _saveSalesDraftItems(items);
  }

  Future<void> atualizarVendaRascunho(
    String id, {
    int? caixas,
    double? preco,
    String? observacao,
    String? status,
  }) async {
    final body = <String, dynamic>{};
    if (caixas != null) body['caixas'] = caixas;
    if (preco != null) body['preco'] = preco;
    if (observacao != null) body['observacao'] = observacao.trim();
    if (status != null && status.trim().isNotEmpty) {
      body['status'] = _upperValue(status);
    }
    if (body.isEmpty) return;

    if (await _shouldUseOnlineSharedFlow('Atualizar venda do rascunho')) {
      await _vendorPatch(
        '/vendedor/rascunho/${Uri.encodeComponent(id.trim())}',
        body,
      );
      return;
    }
    final items = await _loadSalesDraftItems();
    final target = id.trim();
    final now = DateTime.now().toIso8601String();

    for (final item in items) {
      if (_textValue(item['id']) != target) continue;
      if (caixas != null) item['caixas'] = caixas;
      if (preco != null) item['preco'] = preco;
      if (observacao != null) item['observacao'] = observacao.trim();
      if (status != null && status.trim().isNotEmpty) {
        item['status'] = _upperValue(status);
      }
      item['updated_at'] = now;
      break;
    }

    await _saveSalesDraftItems(items);
  }

  Future<void> removerVendaRascunho(String id) async {
    if (await _shouldUseOnlineSharedFlow('Excluir venda do rascunho')) {
      await _vendorDelete('/vendedor/rascunho/${Uri.encodeComponent(id.trim())}');
      return;
    }
    final items = await _loadSalesDraftItems();
    items.removeWhere((item) => _textValue(item['id']) == id.trim());
    await _saveSalesDraftItems(items);
  }

  Future<void> removerVendasRascunhoPorIds(List<String> ids) async {
    final normalized = ids.map((item) => item.trim()).where((item) => item.isNotEmpty).toSet();
    if (normalized.isEmpty) return;
    if (await _shouldUseOnlineSharedFlow('Remover vendas do rascunho')) {
      await _vendorPost(
        '/vendedor/rascunho/remover-em-lote',
        <String, dynamic>{'ids': normalized.toList()},
      );
      return;
    }
    final items = await _loadSalesDraftItems();
    items.removeWhere((item) => normalized.contains(_textValue(item['id'])));
    await _saveSalesDraftItems(items);
  }

  Future<List<Map<String, dynamic>>> listarPreProgramacoes({
    String status = 'ABERTA',
    int limit = 100,
  }) async {
    await _shouldUseOnlineSharedFlow('Consultar pre-programacoes');
    final data = await _vendorGet(
      '/vendedor/pre-programacoes?status=${Uri.encodeQueryComponent(status.trim().isEmpty ? 'ABERTA' : status.trim())}&limit=$limit',
    );
    return _asMapList(data);
  }

  Future<Map<String, dynamic>> detalhePreProgramacao(String id) async {
    final target = id.trim();
    if (target.isEmpty) {
      throw Exception('Pre-programacao invalida.');
    }
    await _shouldUseOnlineSharedFlow('Consultar detalhe da pre-programacao');
    final data = await _vendorGet(
      '/vendedor/pre-programacoes/${Uri.encodeComponent(target)}',
    );
    return _asStringMap(data);
  }

  Future<Map<String, dynamic>> salvarPreProgramacao({
    String? id,
    required List<String> itemIds,
    String titulo = '',
    String observacao = '',
    String status = 'ABERTA',
  }) async {
    await _shouldUseOnlineSharedFlow('Salvar pre-programacao');
    final cleanedIds = itemIds
        .map((item) => item.trim())
        .where((item) => item.isNotEmpty)
        .toList();
    final data = await _vendorPost(
      '/vendedor/pre-programacoes/upsert',
      <String, dynamic>{
        if ((id ?? '').trim().isNotEmpty) 'id': id!.trim(),
        'titulo': titulo.trim(),
        'observacao': observacao.trim(),
        'status': _upperValue(status),
        'item_ids': cleanedIds,
      },
    );
    return _asStringMap(data);
  }

  Future<void> removerPreProgramacao(String id) async {
    final target = id.trim();
    if (target.isEmpty) return;
    await _shouldUseOnlineSharedFlow('Remover pre-programacao');
    await _vendorDelete('/vendedor/pre-programacoes/${Uri.encodeComponent(target)}');
  }

  Future<Map<String, Map<String, dynamic>>> carregarAlertasClientesEmEntrega(
    List<String> codClientes,
  ) async {
    final cods = codClientes.map(_upperValue).where((item) => item.isNotEmpty).toSet().toList();
    if (cods.isEmpty) return <String, Map<String, dynamic>>{};
    if (!await isOnlineServerReachable()) {
      return <String, Map<String, dynamic>>{};
    }

    final programacoes = await listarProgramacoesOficiais(modo: 'ativas', limit: 120);
    final warnings = <String, Map<String, dynamic>>{};

    for (final programacao in programacoes) {
      final status = _upperValue(programacao['status_operacional']).isNotEmpty
          ? _upperValue(programacao['status_operacional'])
          : _upperValue(programacao['status']);
      if (status != 'EM_ENTREGAS' && status != 'EM ENTREGAS') {
        continue;
      }

      final codigo = _upperValue(programacao['codigo_programacao']);
      if (codigo.isEmpty) continue;

      try {
        final detalhe = _asStringMap(await _get('/desktop/rotas/$codigo'));
        final clientes = _asMapList(detalhe['clientes']);
        for (final cliente in clientes) {
          final codCliente = _upperValue(cliente['cod_cliente']);
          if (!cods.contains(codCliente) || warnings.containsKey(codCliente)) {
            continue;
          }
          warnings[codCliente] = <String, dynamic>{
            'codigo_programacao': codigo,
            'status': status,
            'motorista': _textValue(programacao['motorista']),
          };
        }
      } catch (_) {
        // Aviso apenas complementar; falhas aqui nao bloqueiam o fluxo.
      }
    }

    return warnings;
  }

  Future<List<Map<String, dynamic>>> listarProgramacoesOficiais({
    String modo = 'ativas',
    int limit = 200,
  }) async {
    await _shouldUseOnlineSharedFlow('Consultar programacoes oficiais');
    final data = await _get(
      '/desktop/programacoes?modo=${Uri.encodeQueryComponent(modo.trim().isEmpty ? 'ativas' : modo.trim())}&limit=$limit',
    );
    if (data is List) {
      return _asMapList(data);
    }
    final payload = _asStringMap(data);
    final programacoes = _asMapList(payload['programacoes']);
    if (programacoes.isNotEmpty) {
      return programacoes;
    }
    return _asMapList(payload['rows']);
  }

  Future<Map<String, dynamic>> detalheProgramacaoOficial(
    String codigoProgramacao,
  ) async {
    final code = _upperValue(codigoProgramacao);
    if (code.isEmpty) {
      throw Exception('Codigo da programacao invalido.');
    }
    await _shouldUseOnlineSharedFlow('Consultar detalhe da programacao');
    return _asStringMap(await _get('/desktop/rotas/$code'));
  }

  Future<String> gerarCodigoProgramacao() async {
    final rows = await listarProgramacoesOficiais(modo: 'todas', limit: 500);
    final year = DateTime.now().year.toString();
    final prefix = 'PG$year';
    var suffix = 1;

    for (final row in rows) {
      final code = _upperValue(row['codigo_programacao']);
      if (!code.startsWith(prefix)) continue;
      final tail = code.substring(prefix.length);
      final digits = tail.replaceAll(RegExp(r'[^0-9]'), '');
      if (digits.isEmpty) continue;
      final candidate = int.tryParse(digits);
      if (candidate != null && candidate >= suffix) {
        suffix = candidate + 1;
      }
    }

    return '$prefix${suffix.toString().padLeft(2, '0')}';
  }

  Future<Map<String, dynamic>> criarProgramacaoOficial({
    required String codigoProgramacao,
    required DateTime dataProgramada,
    required Map<String, dynamic> motorista,
    required Map<String, dynamic> veiculo,
    required List<Map<String, dynamic>> ajudantes,
    required String localRota,
    required String localCarregamento,
    required String tipoOperacao,
    required double adiantamento,
    required double kgEstimado,
    required int caixasEstimado,
    required List<Map<String, dynamic>> itens,
    required String usuarioCriacao,
    String observacao = '',
  }) async {
    if (!await isOnlineServerReachable()) {
      throw Exception(
        'Criacao de programacao oficial exige conexao para vincular com o desktop e o app do motorista.',
      );
    }

    final tipoNormalizado = _upperValue(tipoOperacao) == 'FOB' ? 'CX' : 'KG';
    final equipeTxt = ajudantes
        .map((item) => _upperValue(item['nome']))
        .where((item) => item.isNotEmpty)
        .join('|');
    final totalCaixas = itens.fold<int>(
      0,
      (acc, item) => acc + _intValue(item['caixas']),
    );

    final payload = <String, dynamic>{
      'codigo_programacao': _upperValue(codigoProgramacao),
      'data_criacao': dataProgramada.toIso8601String(),
      'motorista': _upperValue(motorista['nome']),
      'motorista_id': _intValue(motorista['id']),
      'motorista_codigo': _upperValue(motorista['codigo']),
      'veiculo': _upperValue(veiculo['placa']),
      'equipe': equipeTxt,
      'kg_estimado': tipoNormalizado == 'KG' ? kgEstimado : 0.0,
      'tipo_estimativa': tipoNormalizado,
      'caixas_estimado': tipoNormalizado == 'CX' ? caixasEstimado : 0,
      'status': 'ATIVA',
      'local_rota': _upperValue(localRota),
      'local_carregamento': _upperValue(localCarregamento),
      'adiantamento': adiantamento,
      'total_caixas': totalCaixas,
      'quilos': tipoNormalizado == 'KG' ? kgEstimado : 0.0,
      'usuario_criacao': _upperValue(usuarioCriacao),
      'usuario_ultima_edicao': _upperValue(usuarioCriacao),
      'itens': itens.map((item) {
        return <String, dynamic>{
          'cod_cliente': _upperValue(item['cod_cliente']),
          'nome_cliente': _textValue(item['nome_cliente']),
          'qnt_caixas': _intValue(item['caixas']),
          'kg': 0.0,
          'preco': _doubleValue(item['preco']),
          'endereco': _textValue(item['endereco']),
          'vendedor': _upperValue(item['vendedor_origem']).isNotEmpty
              ? _upperValue(item['vendedor_origem'])
              : _upperValue(item['vendedor_cadastro']),
          'pedido': '',
          'produto': '',
          'obs': [
            _textValue(item['observacao']),
            observacao.trim(),
          ].where((part) => part.isNotEmpty).join(' | '),
        };
      }).toList(),
    };

    return _asStringMap(await _post('/desktop/rotas/upsert', payload));
  }

  Future<Map<String, dynamic>> criarAvulsa(Map<String, dynamic> payload) async {
    if (await isOnlineServerReachable()) {
      final data = await _post('/desktop/avulsas', payload);
      final out = _asStringMap(data);
      await _cacheCreatedAvulsa(payload, out);
      return out;
    }

    final mutationId = await _enqueueMutation(
      method: 'POST',
      path: '/desktop/avulsas',
      body: payload,
    );
    await _upsertDraftAvulsa(mutationId, payload);
    return <String, dynamic>{
      'ok': true,
      'queued_offline': true,
      'codigo_avulsa': _offlineCodeFromMutationId(mutationId),
      'message':
          'Sem conexao. Avulsa registrada na fila para sincronizar depois.',
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
      return _asStringMap(await _post(path, body));
    }
    await _enqueueMutation(method: 'POST', path: path, body: body);
    return <String, dynamic>{
      'ok': true,
      'queued_offline': true,
      'message':
          'Sem conexao. Conciliacao registrada na fila para sincronizar depois.',
    };
  }

  Future<List<Map<String, dynamic>>> listarAvulsas({
    String status = '',
    String dataDe = '',
    String dataAte = '',
  }) async {
    final query = <String, String>{};
    if (status.trim().isNotEmpty) query['status'] = status.trim();
    if (dataDe.trim().isNotEmpty) query['data_de'] = dataDe.trim();
    if (dataAte.trim().isNotEmpty) query['data_ate'] = dataAte.trim();
    query['limit'] = '200';
    final qs = query.entries
        .map((entry) => '${entry.key}=${Uri.encodeQueryComponent(entry.value)}')
        .join('&');

    if (await isOnlineServerReachable()) {
      final data = await _get('/desktop/avulsas?$qs');
      final out = _asMapList(data);
      await _upsertAvulsasCache(
        out,
        replaceAll: status.trim().isEmpty &&
            dataDe.trim().isEmpty &&
            dataAte.trim().isEmpty,
      );
      return _mergeAvulsasWithDrafts(
        serverItems: out,
        status: status,
        dataDe: dataDe,
        dataAte: dataAte,
      );
    }

    final cached = _asMapList(await store.getJson(_kCacheAvulsas));
    return _mergeAvulsasWithDrafts(
      serverItems: cached,
      status: status,
      dataDe: dataDe,
      dataAte: dataAte,
    );
  }

  Future<Map<String, dynamic>> detalheAvulsa(String codigoAvulsa) async {
    final code = codigoAvulsa.trim().toUpperCase();
    if (code.startsWith('OFFLINE-')) {
      final draft = await _findDraftAvulsaByCode(code);
      if (draft != null) {
        return <String, dynamic>{
          'avulsa': _toAvulsaResumo(draft),
          'itens': _asMapList(draft['itens']),
        };
      }
    }

    if (await isOnlineServerReachable()) {
      final data = _asStringMap(await _get('/desktop/avulsas/$code'));
      await store.putJson('$_kCacheAvulsaPrefix$code', data);
      return data;
    }

    final cached = await store.getJson('$_kCacheAvulsaPrefix$code');
    final out = _asStringMap(cached);
    if (out.isNotEmpty) return out;
    throw Exception('Detalhe indisponivel offline para $code.');
  }
}
