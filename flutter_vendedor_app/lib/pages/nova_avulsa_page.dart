import 'package:flutter/material.dart';

import '../models/app_config.dart';
import '../services/vendedor_api_service.dart';
import '../ui/app_visuals.dart';

class NovaAvulsaPage extends StatefulWidget {
  const NovaAvulsaPage({
    super.key,
    required this.config,
    required this.api,
    required this.onCreated,
  });

  final AppConfig config;
  final VendedorApiService api;
  final Future<void> Function() onCreated;

  @override
  State<NovaAvulsaPage> createState() => _NovaAvulsaPageState();
}

class _NovaAvulsaPageState extends State<NovaAvulsaPage> {
  final _formKey = GlobalKey<FormState>();
  final _observacaoCtrl = TextEditingController(text: 'DOMINGO AVULSA');
  final _equipeCtrl = TextEditingController();
  final _clienteBuscaCtrl = TextEditingController();

  DateTime _dataSelecionada = DateTime.now();
  bool _loadingRefs = true;
  bool _buscandoClientes = false;
  bool _saving = false;
  List<Map<String, dynamic>> _motoristas = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _veiculos = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _ajudantes = <Map<String, dynamic>>[];
  List<Map<String, dynamic>> _clientesBusca = <Map<String, dynamic>>[];
  final List<Map<String, dynamic>> _itensSelecionados =
      <Map<String, dynamic>>[];
  final Set<int> _ajudantesSelecionados = <int>{};
  Map<String, dynamic>? _motoristaSelecionado;
  Map<String, dynamic>? _veiculoSelecionado;
  String? _localRotaSelecionada;

  @override
  void initState() {
    super.initState();
    _loadReferences();
  }

  @override
  void dispose() {
    _observacaoCtrl.dispose();
    _equipeCtrl.dispose();
    _clienteBuscaCtrl.dispose();
    super.dispose();
  }

  String _formatDate(DateTime value) {
    final month = value.month.toString().padLeft(2, '0');
    final day = value.day.toString().padLeft(2, '0');
    return '${value.year}-$month-$day';
  }

  Map<String, dynamic>? _reselectItem(
    List<Map<String, dynamic>> items,
    Map<String, dynamic>? current,
    String key,
  ) {
    if (items.isEmpty) return null;
    final currentValue = (current?[key] ?? '').toString().trim().toUpperCase();
    if (currentValue.isEmpty) return items.first;
    for (final item in items) {
      if ((item[key] ?? '').toString().trim().toUpperCase() == currentValue) {
        return item;
      }
    }
    return items.first;
  }

  Future<void> _loadReferences() async {
    setState(() {
      _loadingRefs = true;
    });

    try {
      final motoristas = await widget.api.listarMotoristas(
        preferCacheWhenOffline: true,
      );
      final veiculos = await widget.api.listarVeiculos(
        preferCacheWhenOffline: true,
      );
      final ajudantes = await widget.api.listarAjudantes(
        preferCacheWhenOffline: true,
      );

      if (!mounted) return;
      setState(() {
        _motoristas = motoristas;
        _veiculos = veiculos;
        _ajudantes = ajudantes;
        _motoristaSelecionado =
            _reselectItem(_motoristas, _motoristaSelecionado, 'codigo');
        _veiculoSelecionado =
            _reselectItem(_veiculos, _veiculoSelecionado, 'placa');
      });
      await _buscarClientes();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao carregar base: $error')),
      );
    } finally {
      if (mounted) {
        setState(() {
          _loadingRefs = false;
        });
      }
    }
  }

  Future<void> _buscarClientes() async {
    setState(() {
      _buscandoClientes = true;
    });

    try {
      final filtrosVendedor = <String>[
        widget.config.vendedorPadrao.trim(),
        if (widget.config.vendedorLogin.trim().isNotEmpty &&
            widget.config.vendedorLogin.trim().toUpperCase() !=
                widget.config.vendedorPadrao.trim().toUpperCase())
          widget.config.vendedorLogin.trim(),
        '',
      ];
      final limit = _clienteBuscaCtrl.text.trim().isEmpty ? 80 : 250;

      List<Map<String, dynamic>> result = <Map<String, dynamic>>[];
      for (final filtroVendedor in filtrosVendedor) {
        result = await widget.api.buscarClientes(
          _clienteBuscaCtrl.text,
          vendedor: filtroVendedor,
          cidade: widget.config.cidadePadrao,
          limit: limit,
          preferCacheWhenOffline: true,
        );
        if (result.isNotEmpty || filtroVendedor.isEmpty) {
          break;
        }
      }

      if (!mounted) return;
      setState(() {
        _clientesBusca = result;
      });
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha na busca de clientes: $error')),
      );
    } finally {
      if (mounted) {
        setState(() {
          _buscandoClientes = false;
        });
      }
    }
  }

  Future<void> _pickDate() async {
    final picked = await showDatePicker(
      context: context,
      initialDate: _dataSelecionada,
      firstDate: DateTime.now().subtract(const Duration(days: 30)),
      lastDate: DateTime.now().add(const Duration(days: 365)),
    );
    if (picked == null) return;
    setState(() {
      _dataSelecionada = picked;
    });
  }

  void _syncEquipeField() {
    final nomes = _ajudantes
        .where((item) =>
            _ajudantesSelecionados.contains((item['id'] as num?)?.toInt() ?? 0))
        .map((item) => (item['nome'] ?? '').toString().trim())
        .where((item) => item.isNotEmpty)
        .toList();
    _equipeCtrl.text = nomes.join(' / ').toUpperCase();
  }

  void _toggleAjudante(Map<String, dynamic> ajudante, bool selected) {
    final id = (ajudante['id'] as num?)?.toInt() ?? 0;
    setState(() {
      if (selected) {
        _ajudantesSelecionados.add(id);
      } else {
        _ajudantesSelecionados.remove(id);
      }
      _syncEquipeField();
    });
  }

  void _adicionarCliente(Map<String, dynamic> cliente) {
    final codigo =
        (cliente['cod_cliente'] ?? '').toString().trim().toUpperCase();
    final jaExiste = _itensSelecionados.any(
      (item) =>
          (item['cod_cliente'] ?? '').toString().trim().toUpperCase() == codigo,
    );
    if (jaExiste) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Cliente ja adicionado nesta avulsa.')),
      );
      return;
    }

    setState(() {
      _itensSelecionados.add(
        <String, dynamic>{
          ...cliente,
          'observacao': '',
          'ordem': _itensSelecionados.length + 1,
        },
      );
    });
  }

  void _renumerarItens() {
    for (var i = 0; i < _itensSelecionados.length; i += 1) {
      _itensSelecionados[i]['ordem'] = i + 1;
    }
  }

  void _removerCliente(int index) {
    setState(() {
      _itensSelecionados.removeAt(index);
      _renumerarItens();
    });
  }

  void _moverCliente(int index, int delta) {
    final novoIndice = index + delta;
    if (novoIndice < 0 || novoIndice >= _itensSelecionados.length) return;
    setState(() {
      final item = _itensSelecionados.removeAt(index);
      _itensSelecionados.insert(novoIndice, item);
      _renumerarItens();
    });
  }

  Future<void> _editarObservacaoItem(int index) async {
    final controller = TextEditingController(
      text: (_itensSelecionados[index]['observacao'] ?? '').toString(),
    );
    final result = await showDialog<String>(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: const Text('Observacao do item'),
          content: TextField(
            controller: controller,
            maxLines: 3,
            decoration: const InputDecoration(
              hintText: 'Ex.: entregar cedo, falar com gerente...',
            ),
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: const Text('Cancelar'),
            ),
            FilledButton(
              onPressed: () =>
                  Navigator.of(context).pop(controller.text.trim()),
              child: const Text('Salvar'),
            ),
          ],
        );
      },
    );
    controller.dispose();
    if (result == null) return;
    setState(() {
      _itensSelecionados[index]['observacao'] = result;
    });
  }

  Map<String, dynamic> _buildPayload() {
    final motorista = _motoristaSelecionado ?? <String, dynamic>{};
    final veiculo = _veiculoSelecionado ?? <String, dynamic>{};

    return <String, dynamic>{
      'data_programada': _formatDate(_dataSelecionada),
      'motorista_id': (motorista['id'] as num?)?.toInt(),
      'motorista_codigo': (motorista['codigo'] ?? '').toString().trim(),
      'motorista_nome': (motorista['nome'] ?? '').toString().trim(),
      'veiculo': (veiculo['placa'] ?? '').toString().trim(),
      'equipe': _equipeCtrl.text.trim(),
      'local_rota': (_localRotaSelecionada ?? '').trim(),
      'observacao': _observacaoCtrl.text.trim(),
      'criado_por': widget.config.vendedorPadrao.trim().toUpperCase(),
      'itens': List<Map<String, dynamic>>.generate(
        _itensSelecionados.length,
        (index) {
          final item = _itensSelecionados[index];
          return <String, dynamic>{
            'cod_cliente': (item['cod_cliente'] ?? '').toString().trim(),
            'nome_cliente': (item['nome_cliente'] ?? '').toString().trim(),
            'endereco': (item['endereco'] ?? '').toString().trim(),
            'cidade': (item['cidade'] ?? '').toString().trim(),
            'bairro': (item['bairro'] ?? '').toString().trim(),
            'ordem': index + 1,
            'observacao': (item['observacao'] ?? '').toString().trim(),
          };
        },
      ),
    };
  }

  Future<void> _submit() async {
    if (!_formKey.currentState!.validate()) return;
    if (_motoristaSelecionado == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Selecione um motorista.')),
      );
      return;
    }
    if (_veiculoSelecionado == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Selecione um veiculo.')),
      );
      return;
    }
    if (_equipeCtrl.text.trim().isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Informe a equipe da avulsa.')),
      );
      return;
    }
    if (_itensSelecionados.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Adicione ao menos um cliente.')),
      );
      return;
    }

    setState(() {
      _saving = true;
    });

    try {
      final response = await widget.api.criarAvulsa(_buildPayload());
      final codigo = (response['codigo_avulsa'] ?? '').toString();
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            (response['message'] ?? 'Avulsa registrada. Codigo: $codigo')
                .toString(),
          ),
        ),
      );
      setState(() {
        _observacaoCtrl.text = 'DOMINGO AVULSA';
        _equipeCtrl.clear();
        _clienteBuscaCtrl.clear();
        _clientesBusca = <Map<String, dynamic>>[];
        _itensSelecionados.clear();
        _ajudantesSelecionados.clear();
      });
      await widget.onCreated();
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Falha ao criar avulsa: $error')),
      );
    } finally {
      if (mounted) {
        setState(() {
          _saving = false;
        });
      }
    }
  }

  Widget _buildClienteBuscaCard() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const PanelHeader(
            title: 'Base de clientes',
            subtitle:
                'Use a mesma leitura operacional do app do motorista para montar a avulsa.',
            icon: Icons.people_alt_outlined,
          ),
          const SizedBox(height: 16),
          Row(
            children: [
              Expanded(
                child: TextField(
                  controller: _clienteBuscaCtrl,
                  onSubmitted: (_) => _buscarClientes(),
                  decoration: const InputDecoration(
                    labelText: 'Buscar por codigo, nome ou cidade',
                  ),
                ),
              ),
              const SizedBox(width: 12),
              ElevatedButton.icon(
                onPressed: _buscandoClientes ? null : _buscarClientes,
                icon: _buscandoClientes
                    ? const SizedBox(
                        width: 16,
                        height: 16,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    : const Icon(Icons.search),
                label: const Text('Buscar'),
              ),
            ],
          ),
          const SizedBox(height: 12),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              StatusBadge(
                label: 'Vendedor: ${widget.config.vendedorPadrao}',
                color: VendorUiColors.primary,
              ),
              if (widget.config.cidadePadrao.trim().isNotEmpty)
                StatusBadge(
                  label: 'Cidade: ${widget.config.cidadePadrao}',
                  color: Colors.blueGrey.shade700,
                ),
            ],
          ),
          const SizedBox(height: 12),
          if (_clientesBusca.isEmpty)
            const Text('Nenhum cliente carregado para a busca atual.')
          else
            ..._clientesBusca.take(8).map(
                  (cliente) => Container(
                    margin: const EdgeInsets.only(bottom: 10),
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: VendorUiColors.surfaceAlt,
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(color: VendorUiColors.border),
                    ),
                    child: Row(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                '${cliente['cod_cliente'] ?? '-'} - ${cliente['nome_cliente'] ?? '-'}',
                                style: const TextStyle(
                                  color: VendorUiColors.heading,
                                  fontWeight: FontWeight.w800,
                                ),
                              ),
                              const SizedBox(height: 4),
                              Text(
                                '${cliente['cidade'] ?? ''} ${cliente['bairro'] ?? ''}'
                                    .trim(),
                                style: const TextStyle(
                                  color: VendorUiColors.muted,
                                  fontSize: 12,
                                  fontWeight: FontWeight.w700,
                                ),
                              ),
                            ],
                          ),
                        ),
                        IconButton(
                          onPressed: () => _adicionarCliente(cliente),
                          icon: const Icon(Icons.add_circle_outline),
                        ),
                      ],
                    ),
                  ),
                ),
        ],
      ),
    );
  }

  Widget _buildItensSelecionadosCard() {
    return AppPanel(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          PanelHeader(
            title: 'Clientes na avulsa (${_itensSelecionados.length})',
            subtitle: 'Ordem de entrega, observacao e composicao final da rota.',
            icon: Icons.assignment_outlined,
            trailing: StatusBadge(
              label: '${_itensSelecionados.length} itens',
              color: VendorUiColors.primary,
            ),
          ),
          const SizedBox(height: 12),
          if (_itensSelecionados.isEmpty)
            const Text('Nenhum cliente adicionado ainda.')
          else
            ...List.generate(_itensSelecionados.length, (index) {
              final item = _itensSelecionados[index];
              final observacao = (item['observacao'] ?? '').toString().trim();
              return AppPanel(
                margin: const EdgeInsets.only(bottom: 12),
                padding: const EdgeInsets.all(12),
                backgroundColor: VendorUiColors.surfaceAlt,
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      '${item['ordem']}. ${item['cod_cliente']} - ${item['nome_cliente']}',
                      style: const TextStyle(
                        color: VendorUiColors.heading,
                        fontWeight: FontWeight.w800,
                      ),
                    ),
                    const SizedBox(height: 6),
                    Text(
                      [
                        (item['cidade'] ?? '').toString().trim(),
                        (item['bairro'] ?? '').toString().trim(),
                        if (observacao.isNotEmpty) 'Obs: $observacao',
                      ].where((part) => part.isNotEmpty).join(' | '),
                      style: const TextStyle(
                        color: VendorUiColors.muted,
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                    const SizedBox(height: 8),
                    Wrap(
                      spacing: 4,
                      children: [
                        IconButton(
                          onPressed:
                              index == 0 ? null : () => _moverCliente(index, -1),
                          icon: const Icon(Icons.keyboard_arrow_up),
                        ),
                        IconButton(
                          onPressed: index == _itensSelecionados.length - 1
                              ? null
                              : () => _moverCliente(index, 1),
                          icon: const Icon(Icons.keyboard_arrow_down),
                        ),
                        IconButton(
                          onPressed: () => _editarObservacaoItem(index),
                          icon: const Icon(Icons.edit_outlined),
                        ),
                        IconButton(
                          onPressed: () => _removerCliente(index),
                          icon: const Icon(Icons.delete_outline),
                        ),
                      ],
                    ),
                  ],
                ),
              );
            }),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    if (_loadingRefs) {
      return const Center(child: CircularProgressIndicator());
    }

    return RefreshIndicator(
      onRefresh: _loadReferences,
      child: ListView(
        padding: const EdgeInsets.fromLTRB(16, 8, 16, 24),
        children: [
          AppPanel(
            child: Form(
              key: _formKey,
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  PanelHeader(
                    title: 'Nova programacao avulsa',
                    subtitle:
                        'Monte a rota com motorista, equipe, local e clientes no mesmo padrao visual do motorista.',
                    icon: Icons.add_business,
                    trailing: StatusBadge(
                      label: _formatDate(_dataSelecionada),
                      color: VendorUiColors.primary,
                    ),
                  ),
                  const SizedBox(height: 16),
                  ListTile(
                    contentPadding: EdgeInsets.zero,
                    title: const Text('Data programada'),
                    subtitle: Text(_formatDate(_dataSelecionada)),
                    trailing: OutlinedButton(
                      onPressed: _pickDate,
                      child: const Text('Alterar'),
                    ),
                  ),
                  const SizedBox(height: 8),
                  DropdownButtonFormField<Map<String, dynamic>>(
                    initialValue: _motoristaSelecionado,
                    isExpanded: true,
                    decoration: const InputDecoration(labelText: 'Motorista'),
                    items: _motoristas
                        .map(
                          (item) => DropdownMenuItem<Map<String, dynamic>>(
                            value: item,
                            child: Text('${item['codigo']} - ${item['nome']}'),
                          ),
                        )
                        .toList(),
                    onChanged: (value) {
                      setState(() {
                        _motoristaSelecionado = value;
                      });
                    },
                  ),
                  const SizedBox(height: 16),
                  DropdownButtonFormField<Map<String, dynamic>>(
                    initialValue: _veiculoSelecionado,
                    isExpanded: true,
                    decoration: const InputDecoration(labelText: 'Veiculo'),
                    items: _veiculos
                        .map(
                          (item) => DropdownMenuItem<Map<String, dynamic>>(
                            value: item,
                            child: Text(
                              '${item['placa']} ${item['modelo'] ?? ''}'.trim(),
                            ),
                          ),
                        )
                        .toList(),
                    onChanged: (value) {
                      setState(() {
                        _veiculoSelecionado = value;
                      });
                    },
                  ),
                  const SizedBox(height: 16),
                  Text('Equipe', style: Theme.of(context).textTheme.titleMedium),
                  const SizedBox(height: 8),
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: _ajudantes
                        .map(
                          (ajudante) => FilterChip(
                            label: Text((ajudante['nome'] ?? '-').toString()),
                            selected: _ajudantesSelecionados.contains(
                              (ajudante['id'] as num?)?.toInt() ?? 0,
                            ),
                            onSelected: (selected) =>
                                _toggleAjudante(ajudante, selected),
                          ),
                        )
                        .toList(),
                  ),
                  const SizedBox(height: 12),
                  TextFormField(
                    controller: _equipeCtrl,
                    decoration: const InputDecoration(
                      labelText: 'Equipe consolidada',
                      hintText: 'Ex.: MARCIO / PAULO',
                    ),
                    validator: (value) {
                      if ((value ?? '').trim().isEmpty) {
                        return 'Informe a equipe da avulsa.';
                      }
                      return null;
                    },
                  ),
                  const SizedBox(height: 16),
                  DropdownButtonFormField<String>(
                    initialValue: _localRotaSelecionada,
                    decoration: const InputDecoration(
                      labelText: 'Local da rota',
                    ),
                    items: const [
                      DropdownMenuItem(
                        value: 'SERRA',
                        child: Text('SERRA'),
                      ),
                      DropdownMenuItem(
                        value: 'SERTAO',
                        child: Text('SERTAO'),
                      ),
                    ],
                    onChanged: (value) {
                      setState(() {
                        _localRotaSelecionada = value;
                      });
                    },
                    validator: (value) {
                      if ((value ?? '').trim().isEmpty) {
                        return 'Selecione o local da rota.';
                      }
                      return null;
                    },
                  ),
                  const SizedBox(height: 16),
                  TextFormField(
                    controller: _observacaoCtrl,
                    maxLines: 2,
                    decoration: const InputDecoration(
                      labelText: 'Observacao',
                    ),
                  ),
                  const SizedBox(height: 16),
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      StatusBadge(
                        label:
                            'Motoristas: ${_motoristas.length} | Veiculos: ${_veiculos.length}',
                        color: Colors.blueGrey.shade700,
                      ),
                      StatusBadge(
                        label: 'Clientes selecionados: ${_itensSelecionados.length}',
                        color: VendorUiColors.primary,
                      ),
                    ],
                  ),
                  const SizedBox(height: 20),
                  ElevatedButton.icon(
                    onPressed: _saving ? null : _submit,
                    icon: _saving
                        ? const SizedBox(
                            width: 16,
                            height: 16,
                            child: CircularProgressIndicator(strokeWidth: 2),
                          )
                        : const Icon(Icons.save_alt),
                    label: const Text('Registrar avulsa'),
                  ),
                ],
              ),
            ),
          ),
          const SizedBox(height: 8),
          _buildClienteBuscaCard(),
          const SizedBox(height: 8),
          _buildItensSelecionadosCard(),
        ],
      ),
    );
  }
}
