import 'package:flutter_test/flutter_test.dart';
import 'package:flutter_vendedor_app/models/app_config.dart';

void main() {
  test('AppConfig normaliza base url e vendedor', () {
    const config = AppConfig(
      baseUrl: 'http://localhost:8000///',
      desktopSecret: 'segredo',
      vendedorPadrao: 'joao',
      vendedorLogin: 'Joao',
      cidadePadrao: 'ipu',
    );

    expect(config.normalizedBaseUrl, 'http://localhost:8000');
    expect(config.toJson()['vendedor_padrao'], 'JOAO');
    expect(config.toJson()['vendedor_login'], 'Joao');
    expect(config.toJson()['cidade_padrao'], 'IPU');
  });
}
