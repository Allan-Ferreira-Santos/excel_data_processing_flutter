class Helper {
  String cleanText(String texto) {
    return texto
        .replaceAll(RegExp(r'[ÁÀÂÃÄ]'), 'A')
        .replaceAll(RegExp(r'[áàâãä]'), 'a')
        .replaceAll(RegExp(r'[ÉÈÊË]'), 'E')
        .replaceAll(RegExp(r'[éèêë]'), 'e')
        .replaceAll(RegExp(r'[ÍÌÎÏ]'), 'I')
        .replaceAll(RegExp(r'[íìîï]'), 'i')
        .replaceAll(RegExp(r'[ÓÒÔÕÖ]'), 'O')
        .replaceAll(RegExp(r'[óòôõö]'), 'o')
        .replaceAll(RegExp(r'[ÚÙÛÜ]'), 'U')
        .replaceAll(RegExp(r'[úùûü]'), 'u')
        .replaceAll(RegExp(r'[Ç]'), 'C')
        .replaceAll(RegExp(r'[ç]'), 'c')
        .toUpperCase()
        .trim();
  }

  String formatCep(String cep) {
    cep = cep.replaceAll(
        RegExp(r'[^0-9]'), ''); 
    if (cep.length < 8) {
      cep = cep.padLeft(
          8, '0'); 
    }
    return cep;
  }
}
