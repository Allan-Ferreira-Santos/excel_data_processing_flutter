import 'package:flutter/material.dart';
import 'package:flutter_excel/view/excel_view.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Processor',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: const ExcelView(),
    );
  }
}