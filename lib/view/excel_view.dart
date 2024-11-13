import 'package:flutter/material.dart';
import 'package:flutter_excel/controller/excel_controller.dart';

class ExcelView extends StatefulWidget {
  const ExcelView({super.key});

  @override
  State<ExcelView> createState() => _ExcelViewState();
}

class _ExcelViewState extends State<ExcelView> {
  final ExcelController controller = ExcelController();

  @override
  void initState() {
    super.initState();
    controller.progressNotifier.addListener(() {
      setState(() {}); // Rebuild when progress changes
      if (controller.progressNotifier.value == 1.0) {
        // Show dialog when progress reaches 100%
        showCompletionDialog();
      }
    });
  }

  void showCompletionDialog() {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text("Processamento Concluído"),
        content: const Text("O processamento do arquivo Excel foi concluído."),
        actions: [
          TextButton(
            onPressed: () {
              Navigator.of(context).pop();
              setState(() {
                controller.progressNotifier.value = 0.0;
              });
            },
            child: const Text("OK"),
          ),
        ],
      ),
    );
  }

  @override
  void dispose() {
    controller.progressNotifier.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Processador de Excel')),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            ElevatedButton(
              onPressed: controller.progressNotifier.value == 0.0
                  ? controller.pickFile
                  : null,
              child: const Text('Selecionar Arquivo Excel'),
            ),
            const SizedBox(height: 20),
            if (controller.progressNotifier.value > 0.0)
              Column(
                children: [
                  const Text('Processando...'),
                  const SizedBox(height: 10),
                  CircularProgressIndicator(
                    value: controller.progressNotifier.value,
                  ),
                  const SizedBox(height: 10),
                  Text(
                    '${(controller.progressNotifier.value * 100).toStringAsFixed(0)}% concluído',
                  ),
                ],
              ),
          ],
        ),
      ),
    );
  }
}
