import 'dart:io';
import 'package:excel/excel.dart';

// Função de limpeza (mantida para evitar o erro XML ao salvar)
String limparValorParaExcel(String valor) {
  if (valor.isEmpty) return "";
  return valor.replaceAll(RegExp(r'[\x00-\x08\x0b\x0c\x0e-\x1f]'), '');
}

// Função para carregar o arquivo XLSX (Binário)
Excel carregarExcel(String filePath) {
  final file = File(filePath);
  if (!file.existsSync()) {
    throw FileSystemException("ERRO: Arquivo não encontrado", filePath);
  }
  var bytes = file.readAsBytesSync();
  return Excel.decodeBytes(bytes);
}

void main() {
  // Arquivos de entrada
  const file1 = "input/planilha1.xlsx";
  const file2 = "input/planilha2.xlsx";
  const file3 = "input/planilha3.xlsx";

  // Conjuntos para evitar duplicação
  final Set<String> emailsSet = {};
  final Set<String> telefonesSet = {};

  // Lista final com cabeçalho
  final List<List<String>> resultado = [
    ["NOME", "CNPJ", "CPF", "EMAIL", "TELEFONE"]
  ];

  // CORREÇÃO FINAL: Apenas aceita números que resultam em 11 dígitos.
  String? formatarTelefone(String telefone) {
    // 1. Limpa todos os caracteres não numéricos
    var apenasNumeros = telefone.replaceAll(RegExp(r'[^0-9]'), '');

    // 2. Remove o prefixo 55 se já existir, para contar apenas o DDD + Número
    // Verificamos se tem 12 ou mais dígitos para garantir que é 55 + telefone.
    if (apenasNumeros.startsWith('55') && apenasNumeros.length >= 12) {
      apenasNumeros = apenasNumeros.substring(2);
    }

    // 3. Valida e formata: SOMENTE 11 DÍGITOS SÃO ACEITOS.
    if (apenasNumeros.length == 11) {
      // Formata com o código do país
      return "55$apenasNumeros";
    }

    // Se não tiver 11 dígitos, é descartado.
    return null;
  }

  // Função para extrair registros (sem mudanças na lógica de desduplicação)
  void extrairRegistros(
    Excel excel, {
    required int nomeIndex,
    required int cnpjIndex,
    required int cpfIndex,
    required int emailIndex,
    required int telefoneIndex,
    int skipRows = 0,
  }) {
    for (var table in excel.tables.keys) {
      var sheet = excel.tables[table]!;

      for (var row in sheet.rows.skip(skipRows)) {
        if (row.isEmpty) continue;

        String getStringValue(int index) {
          if (index < 0 || index >= row.length || row[index] == null) {
            return "";
          }
          return (row[index]!.value?.toString().trim() ?? "").trim();
        }

        var nome = getStringValue(nomeIndex);
        var cnpj = getStringValue(cnpjIndex);
        var cpf = getStringValue(cpfIndex);
        var email = getStringValue(emailIndex).toLowerCase();
        var telefoneRaw = getStringValue(telefoneIndex);

        // Usa a função de formatação estrita
        String? telefone =
            telefoneRaw.isNotEmpty ? formatarTelefone(telefoneRaw) : null;

        if (email.isEmpty && telefone == null) continue;

        var jaExiste = false;
        if (email.isNotEmpty && emailsSet.contains(email)) {
          jaExiste = true;
        }
        if (telefone != null && telefonesSet.contains(telefone)) {
          jaExiste = true;
        }

        if (!jaExiste) {
          if (email.isNotEmpty) emailsSet.add(email);
          if (telefone != null) telefonesSet.add(telefone);
          resultado.add([nome, cnpj, cpf, email, telefone ?? ""]);
        }
      }
    }
  }

  // --- CARREGAMENTO E PROCESSAMENTO (Mantendo os índices corrigidos) ---
  try {
    var excel1 = carregarExcel(file1);
    var excel2 = carregarExcel(file2);
    var excel3 = carregarExcel(file3);

    // Planilha 1: Pula 4 linhas
    extrairRegistros(
      excel1,
      skipRows: 4,
      nomeIndex: 1,
      cnpjIndex: 2,
      cpfIndex: 4,
      emailIndex: 7,
      telefoneIndex: 8,
    );

    // Planilha 2: Pula 1 linha
    extrairRegistros(
      excel2,
      skipRows: 1,
      nomeIndex: -1,
      cnpjIndex: 0,
      cpfIndex: -1,
      emailIndex: -1,
      telefoneIndex: 1,
    );

    // Planilha 3: Pula 1 linha
    extrairRegistros(
      excel3,
      skipRows: 1,
      nomeIndex: 0,
      cnpjIndex: -1,
      cpfIndex: -1,
      emailIndex: 1,
      telefoneIndex: -1,
    );
  } catch (e) {
    print("❌ ERRO NO PROCESSAMENTO: $e");
    return;
  }

  // Verificação final dos resultados antes de salvar
  if (resultado.length <= 1) {
    print(
        "\n⚠️ A planilha final está vazia ou todos os telefones válidos foram descartados por não terem 11 dígitos.");
    return;
  }

  // --- CRIAÇÃO E SALVAMENTO DO ARQUIVO FINAL ---
  var excelFinal = Excel.createExcel();
  var sheet = excelFinal['dados'];

  for (var row in resultado) {
    sheet.appendRow(row
        .map((valor) => TextCellValue(limparValorParaExcel(valor)))
        .toList());
  }

  var fileBytes = excelFinal.encode();
  File("output/planilha_final.xlsx")
    ..createSync(recursive: true)
    ..writeAsBytesSync(fileBytes!);

  print(
      "\n✅ Planilha final gerada com sucesso! ${resultado.length - 1} registros únicos salvos.");
}
