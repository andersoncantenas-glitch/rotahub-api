import 'package:flutter/material.dart';

import 'pages/bootstrap_page.dart';
import 'ui/app_visuals.dart';

class AppVendedorApp extends StatelessWidget {
  const AppVendedorApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'RotaHub Vendedor',
      debugShowCheckedModeBanner: false,
      theme: buildVendorTheme(),
      builder: (context, child) {
        final media = MediaQuery.of(context);
        final width = media.size.width;
        final baseScale = (width / 390.0).clamp(0.92, 1.08);
        final finalScale = media.textScaler.scale(1.0) * baseScale;

        return MediaQuery(
          data: media.copyWith(
            textScaler: TextScaler.linear(finalScale.clamp(0.92, 1.12)),
          ),
          child: child ?? const SizedBox.shrink(),
        );
      },
      home: const BootstrapPage(),
    );
  }
}
