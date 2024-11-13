import 'dart:convert';
import 'dart:developer';
import 'dart:html' as html;
import 'package:intl/intl.dart';
import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:file_picker/file_picker.dart';
import 'package:flutter_excel/helper/helper.dart';
import 'package:flutter_excel/model/new_header_model.dart';
import 'package:flutter_excel/model/types_of_device_model.dart';

class ExcelController {
  final progressNotifier = ValueNotifier<double>(0.0);
  List<List<dynamic>>? data;

  Future<void> pickFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );
    progressNotifier.value = 0.1;

    if (result?.files.single.bytes != null) {
      final excel = Excel.decodeBytes(result!.files.single.bytes!);
      progressNotifier.value = 0.3;
      processExcel(excel);
    } else {
      //TODO: codlocar um Dialog se der merda
      print('Erro ao ler o arquivo: bytes não disponíveis');
    }
  }

  void processExcel(Excel excel) async {
    List<List<dynamic>> data = [];
    List<List<dynamic>> allRows = [];

    for (var table in excel.tables.keys) {
      allRows.addAll(excel.tables[table]!.rows);
    }

    for (var row in allRows) {
      var filteredRow = row.map((cell) => cell?.value).toList();

      if (filteredRow
          .any((cell) => cell != null && cell.toString().isNotEmpty)) {
        data.add(filteredRow);
      }
    }

    progressNotifier.value = 0.5;
    await validateCep(data, excel);
  }

  Future<void> validateCep(List<List<dynamic>> data, Excel excel) async {
    for (var i = 1; i < data.length; i++) {
      String cep = data[i][7]?.toString() ?? "";
      if (cep.isNotEmpty && cep != "null") {
        Map<String, dynamic>? cepInfo = await requestCep(cep);

        String? neighborhood =
            cepInfo?['neighborhood']?.toString().isEmpty ?? true
                ? null
                : cepInfo?['neighborhood']?.toString();
        String? city = cepInfo?['city']?.toString();
        String? state = cepInfo?['state']?.toString();

        log(cepInfo.toString());

        data[i][6] = neighborhood != null ? TextCellValue(neighborhood) : null;
        data[i][8] = city != null ? TextCellValue(city) : null;
        data[i][9] = state != null ? TextCellValue(state) : null;
        progressNotifier.value = 0.5 + (0.3 * (i / data.length));
      }
    }

    await _preprocessarDados(data, excel);
  }

  Future<Map<String, dynamic>?> requestCep(String cep) async {
    final url = 'https://brasilapi.com.br/api/cep/v1/$cep';
    try {
      final response = await http.get(Uri.parse(url));
      if (response.statusCode == 200) {
        return json.decode(response.body);
      } else {
        return null;
      }
    } catch (e) {
      return null;
    }
  }

  Future<void> _preprocessarDados(List<List<dynamic>> data, Excel excel) async {
    final header = NewHeaderModel().newHeader;
    List<List<dynamic>> newData = [];

    newData.add(header);

    for (var i = 1; i < data.length; i++) {
      List<dynamic> row = data[i];

      var deviceTypeName1 = getRowValue(row, 20) ?? '';

      deviceTypeName1 = typesOfDevice(deviceTypeName1);

      String statusAgendamento =
          processAppointmentStatus(row[42]?.toString() ?? '0');

      List<dynamic> newRow = [
        getRowValue(row, 0) ?? '0', // ID_ROBO
        "DE_LB", // ACRONIMO
        (1111111 + i).toString(), // INSTANCIA
        '0', // DOCUMENTO
        cleanText(getRowValue(row, 2) ?? '0'), // NOME_CLIENTE
        cleanText(getRowValue(row, 3) ?? '0'), // ENDERECO
        getRowValue(row, 4) ?? '0', // NUMERO
        cleanText(getRowValue(row, 5) ?? '0'), // COMPLEMENTO
        cleanText(getRowValue(row, 6) ?? '0'), // BAIRRO
        getRowValue(row, 7) ?? '0', // CEP
        cleanText(getRowValue(row, 8) ?? '0'), // CIDADE
        getRowValue(row, 9) ?? '0', // UF
        cleanText(getRowValue(row, 3) ?? '0'), // ENDERECO_COBRANCA
        getRowValue(row, 4) ?? '0', // NUMERO_COBRANCA
        cleanText(getRowValue(row, 5) ?? '0'), // COMPLEMENTO_COBRANCA
        cleanText(getRowValue(row, 6) ?? '0'), // BAIRRO_COBRANCA
        getRowValue(row, 7) ?? '0', // CEP_COBRANCA
        cleanText(getRowValue(row, 8) ?? '0'), // CIDADE_COBRANCA
        getRowValue(row, 9) ?? '0', // UF_COBRANCA
        cleanText(getRowValue(row, 10)?.replaceFirst('+', '') ??
            '0'), // TELEFONE_COMERCIAL
        cleanText(getRowValue(row, 11)?.replaceFirst('+', '') ??
            '0'), // TELEFONE_RESIDENCIA
        cleanText(getRowValue(row, 12)?.replaceFirst('+', '') ??
            '0'), // TELEFONE_CELULAR
        getRowValue(row, 13) ?? '0', // NOME_CONTATO
        (getRowValue(row, 14)?.toUpperCase().trim() ?? '0'), // EMAIL_CONTATO
        getRowValue(row, 15) ?? '0', // DATA_CRIACAO
        cleanText(getRowValue(row, 16) ?? '0'), // MOTIVO_DESCONEXAO
        cleanText(getRowValue(row, 0) ?? '0'), // PON
        "VIVO2", // REDE
        getRowValue(row, 17) ?? '0', // TECNOLOGIA_ACESSO
        DateFormat('dd-MM-yyyy').format(DateTime.now()), // INSERT_DATE
        getRowValue(row, 18) ?? '0', // QTD_EQUIPAMENTOS
        getRowValue(row, 19) ?? '', // SERIAL_NUMBER_1
        deviceTypeName1, // DEVICE_TYPE_NAME_1 (aplicando a substituição aqui)
        '', // MANUFACTURER_1
        '', // DEVICE_TIPO_1
        getRowValue(row, 21) ?? '', // SERIAL_NUMBER_2
        getRowValue(row, 22) ?? '', // DEVICE_TYPE_NAME_2
        '', // MANUFACTURER_2
        '', // DEVICE_TIPO_2
        getRowValue(row, 23) ?? '', // SERIAL_NUMBER_3
        getRowValue(row, 24) ?? '', // DEVICE_TYPE_NAME_3
        '', // MANUFACTURER_3
        '', // DEVICE_TIPO_3
        getRowValue(row, 25) ?? '', // SERIAL_NUMBER_4
        getRowValue(row, 26) ?? '', // DEVICE_TYPE_NAME_4
        '', // MANUFACTURER_4
        '', // DEVICE_TIPO_4
        getRowValue(row, 27) ?? '', // SERIAL_NUMBER_5
        getRowValue(row, 28) ?? '', // DEVICE_TYPE_NAME_5
        '', // MANUFACTURER_5
        '', // DEVICE_TIPO_5
        getRowValue(row, 29) ?? '', // SERIAL_NUMBER_6
        getRowValue(row, 30) ?? '', // DEVICE_TYPE_NAME_6
        '', // MANUFACTURER_6
        '', // DEVICE_TIPO_6
        getRowValue(row, 31) ?? '', // SERIAL_NUMBER_7
        getRowValue(row, 32) ?? '', // DEVICE_TYPE_NAME_7
        '', // MANUFACTURER_7
        '', // DEVICE_TIPO_7
        getRowValue(row, 33) ?? '', // SERIAL_NUMBER_8
        getRowValue(row, 34) ?? '', // DEVICE_TYPE_NAME_8
        '', // MANUFACTURER_8
        '', // DEVICE_TIPO_8
        getRowValue(row, 35) ?? '', // SERIAL_NUMBER_9
        getRowValue(row, 36) ?? '', // DEVICE_TYPE_NAME_9
        '', // MANUFACTURER_9
        '', // DEVICE_TIPO_9
        getRowValue(row, 37) ?? '', // SERIAL_NUMBER_10
        getRowValue(row, 38) ?? '', // DEVICE_TYPE_NAME_10
        '', // MANUFACTURER_10
        '', // DEVICE_TIPO_10
        'PREMIUM', // OBS
        'RL', // EMPRESA
        cleanText(getRowValue(row, 41) ?? ''), // CLUSTER
        '0', // RETORNO ATENTO
        statusAgendamento, // RETORNO AURA
        'NÃO', // ENDEREÇO CORRIGIDO?
        getRowValue(row, 45) ?? '', // NOVO_LOGRADOURO
        getRowValue(row, 46) ?? '', // NOVO_NUMERO
        getRowValue(row, 47) ?? '', // NOVO_COMPLEMENTO
        getRowValue(row, 48) ?? '', // NOVO_BAIRRO
        getRowValue(row, 49) ?? '', // NOVO_CIDADE
        getRowValue(row, 50) ?? '', // NOVO_ESTADO
        getRowValue(row, 51) ?? '', // NOVO_CEP
        getRowValue(row, 52) ?? '' // PONTO_REFERENCIA
      ];

      while (newRow.length < header.length) {
        newRow.add('');
      }
      newData.add(newRow);
    }
    progressNotifier.value = 0.9;
    await saveExcel(newData, excel);
    progressNotifier.value = 1.0;
  }

  Future<void> saveExcel(List<List<dynamic>> data, Excel originalExcel) async {
    var newExcel = Excel.createExcel();
    var sheet = newExcel['Sheet1'];

    for (int rowIndex = 0; rowIndex < data.length; rowIndex++) {
      List<CellValue?> cellValues = data[rowIndex].map((item) {
        if (item is String) return TextCellValue(item);
        if (item is int) return IntCellValue(item);
        if (item is double) return DoubleCellValue(item);
        if (item is bool) return BoolCellValue(item);
        return null;
      }).toList();

      sheet.appendRow(cellValues);

      // Ajusta a largura da coluna com base no comprimento do texto da primeira linha
      if (rowIndex == 0) {
        for (int colIndex = 0; colIndex < cellValues.length; colIndex++) {
          if (cellValues[colIndex] is TextCellValue) {
            int textLength =
                (cellValues[colIndex] as TextCellValue).value.text!.length;
            double columnWidth = (textLength / 1.5).clamp(10, 50).toDouble();
            sheet.setColumnWidth(colIndex, 15 + columnWidth);
          }
        }
      }
    }

    applyHeaderStyle(sheet, data[0].length);

    final bytes = newExcel.encode() ?? [];
    final blob = html.Blob([bytes]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    html.AnchorElement(href: url)
      ..setAttribute('download', 'processada.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  }

  void applyHeaderStyle(Sheet sheet, int columnCount) {
    var headerStyle = CellStyle(
      bold: true,
      fontSize: 12,
      fontColorHex: ExcelColor.black,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    );

    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      var cell = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: 0));
      cell.cellStyle = headerStyle;
    }
  }

  String? getRowValue(List<dynamic> row, int index) {
    if (index < row.length) {
      return row[index]?.toString();
    }
    return null;
  }

  String typesOfDevice(String value) {
    final devices = TypesOfDeviceModel().replacementDict;

    String normalizedValue = value.trim();

    return devices[normalizedValue] ?? value;
  }

  String processAppointmentStatus(String status) {
    if (status.contains('Coleta já agendada para') ||
        status.contains('Coleta agendada para')) {
      return status
          .replaceAll(
              RegExp(
                  r'Coleta já agendada para:|Coleta ja agendada para:|Coleta agendada para|COLETA AGENDADA PARA:|Coleta agendada para:|Coleta agendada para'),
              '')
          .trim();
    }
    if (status.contains('Visita Surpresa') ||
        status.contains(
            'Coleta já agendada para: Expressa 2 DU (Mudança de Endereço)') ||
        status.contains('Expressa 2 DU (Mudança de Endereço)')) {
      return '0';
    }

    return '0';
  }

  String cleanText(String text) {
    return Helper().cleanText(text);
  }
}
