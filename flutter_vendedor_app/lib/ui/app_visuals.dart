import 'package:flutter/material.dart';

class VendorUiColors {
  static const Color primary = Color(0xFF0B2A3A);
  static const Color heading = Color(0xFF102A43);
  static const Color muted = Color(0xFF607D8B);
  static const Color border = Color(0xFFE0E6ED);
  static const Color background = Color(0xFFF4F6F8);
  static const Color surfaceAlt = Color(0xFFF7FAFC);
  static const Color success = Color(0xFF2E7D32);
  static const Color warning = Color(0xFFE65100);
  static const Color danger = Color(0xFFD32F2F);
}

ThemeData buildVendorTheme() {
  final base = ThemeData(
    colorScheme: ColorScheme.fromSeed(seedColor: VendorUiColors.primary),
    useMaterial3: true,
    visualDensity: VisualDensity.adaptivePlatformDensity,
  );

  final outline = OutlineInputBorder(
    borderRadius: BorderRadius.circular(12),
    borderSide: const BorderSide(color: VendorUiColors.border),
  );

  return base.copyWith(
    scaffoldBackgroundColor: VendorUiColors.background,
    appBarTheme: const AppBarTheme(
      backgroundColor: VendorUiColors.background,
      foregroundColor: VendorUiColors.heading,
      elevation: 0,
      scrolledUnderElevation: 0,
      surfaceTintColor: Colors.transparent,
      titleTextStyle: TextStyle(
        color: VendorUiColors.heading,
        fontSize: 20,
        fontWeight: FontWeight.w800,
      ),
    ),
    inputDecorationTheme: InputDecorationTheme(
      isDense: true,
      filled: true,
      fillColor: Colors.white,
      contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 12),
      border: outline,
      enabledBorder: outline,
      disabledBorder: outline,
      focusedBorder: outline.copyWith(
        borderSide: const BorderSide(color: VendorUiColors.primary, width: 1.3),
      ),
      errorBorder: outline.copyWith(
        borderSide: const BorderSide(color: VendorUiColors.danger),
      ),
      focusedErrorBorder: outline.copyWith(
        borderSide: const BorderSide(color: VendorUiColors.danger, width: 1.3),
      ),
      labelStyle: const TextStyle(
        color: VendorUiColors.muted,
        fontWeight: FontWeight.w700,
      ),
      hintStyle: const TextStyle(color: VendorUiColors.muted),
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      color: Colors.white,
      surfaceTintColor: Colors.transparent,
      margin: EdgeInsets.zero,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(14),
        side: const BorderSide(color: VendorUiColors.border),
      ),
    ),
    elevatedButtonTheme: ElevatedButtonThemeData(
      style: ElevatedButton.styleFrom(
        backgroundColor: VendorUiColors.primary,
        foregroundColor: Colors.white,
        minimumSize: const Size(0, 44),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
        textStyle: const TextStyle(fontWeight: FontWeight.w800),
      ),
    ),
    filledButtonTheme: FilledButtonThemeData(
      style: FilledButton.styleFrom(
        backgroundColor: VendorUiColors.primary,
        foregroundColor: Colors.white,
        minimumSize: const Size(0, 44),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
        textStyle: const TextStyle(fontWeight: FontWeight.w800),
      ),
    ),
    outlinedButtonTheme: OutlinedButtonThemeData(
      style: OutlinedButton.styleFrom(
        foregroundColor: VendorUiColors.primary,
        minimumSize: const Size(0, 44),
        side: const BorderSide(color: VendorUiColors.primary),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
        textStyle: const TextStyle(fontWeight: FontWeight.w800),
      ),
    ),
    textButtonTheme: TextButtonThemeData(
      style: TextButton.styleFrom(
        foregroundColor: VendorUiColors.primary,
        textStyle: const TextStyle(fontWeight: FontWeight.w800),
      ),
    ),
    chipTheme: base.chipTheme.copyWith(
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(999)),
      side: const BorderSide(color: VendorUiColors.border),
      labelStyle: const TextStyle(
        color: VendorUiColors.heading,
        fontWeight: FontWeight.w700,
      ),
    ),
    navigationBarTheme: NavigationBarThemeData(
      backgroundColor: Colors.white,
      surfaceTintColor: Colors.transparent,
      indicatorColor: VendorUiColors.primary.withValues(alpha: 0.12),
      labelTextStyle: WidgetStateProperty.resolveWith((states) {
        final selected = states.contains(WidgetState.selected);
        return TextStyle(
          color: selected ? VendorUiColors.primary : VendorUiColors.muted,
          fontWeight: selected ? FontWeight.w800 : FontWeight.w700,
        );
      }),
      iconTheme: WidgetStateProperty.resolveWith((states) {
        final selected = states.contains(WidgetState.selected);
        return IconThemeData(
          color: selected ? VendorUiColors.primary : VendorUiColors.muted,
        );
      }),
    ),
    dividerColor: VendorUiColors.border,
    textTheme: base.textTheme.copyWith(
      titleLarge: const TextStyle(
        color: VendorUiColors.heading,
        fontWeight: FontWeight.w800,
      ),
      titleMedium: const TextStyle(
        color: VendorUiColors.heading,
        fontWeight: FontWeight.w800,
      ),
      bodyLarge: const TextStyle(
        color: VendorUiColors.heading,
      ),
      bodyMedium: const TextStyle(
        color: VendorUiColors.heading,
      ),
      bodySmall: const TextStyle(
        color: VendorUiColors.muted,
        fontWeight: FontWeight.w600,
      ),
    ),
  );
}

class AppPanel extends StatelessWidget {
  const AppPanel({
    super.key,
    required this.child,
    this.padding = const EdgeInsets.all(16),
    this.margin,
    this.backgroundColor,
  });

  final Widget child;
  final EdgeInsetsGeometry padding;
  final EdgeInsetsGeometry? margin;
  final Color? backgroundColor;

  @override
  Widget build(BuildContext context) {
    return Container(
      margin: margin,
      padding: padding,
      decoration: BoxDecoration(
        color: backgroundColor ?? Colors.white,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(color: VendorUiColors.border),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withValues(alpha: 0.04),
            blurRadius: 10,
            offset: const Offset(0, 4),
          ),
        ],
      ),
      child: child,
    );
  }
}

class PanelHeader extends StatelessWidget {
  const PanelHeader({
    super.key,
    required this.title,
    this.subtitle,
    this.icon,
    this.trailing,
  });

  final String title;
  final String? subtitle;
  final IconData? icon;
  final Widget? trailing;

  @override
  Widget build(BuildContext context) {
    return Row(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        if (icon != null) ...[
          Container(
            width: 42,
            height: 42,
            decoration: BoxDecoration(
              color: VendorUiColors.primary,
              borderRadius: BorderRadius.circular(12),
            ),
            child: Icon(icon, color: Colors.white),
          ),
          const SizedBox(width: 12),
        ],
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(
                title,
                style: const TextStyle(
                  color: VendorUiColors.heading,
                  fontWeight: FontWeight.w900,
                  fontSize: 16,
                ),
              ),
              if (subtitle != null && subtitle!.trim().isNotEmpty) ...[
                const SizedBox(height: 4),
                Text(
                  subtitle!,
                  style: const TextStyle(
                    color: VendorUiColors.muted,
                    fontWeight: FontWeight.w600,
                    fontSize: 12,
                  ),
                ),
              ],
            ],
          ),
        ),
        if (trailing != null) trailing!,
      ],
    );
  }
}

class StatusBadge extends StatelessWidget {
  const StatusBadge({
    super.key,
    required this.label,
    required this.color,
  });

  final String label;
  final Color color;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
        color: color.withValues(alpha: 0.12),
        borderRadius: BorderRadius.circular(999),
        border: Border.all(color: color.withValues(alpha: 0.28)),
      ),
      child: Text(
        label,
        style: TextStyle(
          color: color,
          fontSize: 12,
          fontWeight: FontWeight.w800,
        ),
      ),
    );
  }
}

class InfoRow extends StatelessWidget {
  const InfoRow({
    super.key,
    required this.label,
    required this.value,
  });

  final String label;
  final String value;

  @override
  Widget build(BuildContext context) {
    if (value.trim().isEmpty) return const SizedBox.shrink();
    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          SizedBox(
            width: 124,
            child: Text(
              '$label:',
              style: const TextStyle(
                color: VendorUiColors.muted,
                fontSize: 12,
                fontWeight: FontWeight.w800,
              ),
            ),
          ),
          Expanded(
            child: Text(
              value,
              style: const TextStyle(
                color: VendorUiColors.heading,
                fontSize: 12,
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
        ],
      ),
    );
  }
}
