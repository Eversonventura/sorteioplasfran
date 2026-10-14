/**
 * Sorteio Grupo Plasfran - recebe o formulário e grava na planilha.
 * Abas: "Plaspet" e "Mirimóveis" (criadas automaticamente se não existirem).
 * Número do sorteio = sequência única entre as duas abas (1, 2, 3...).
 * Mesmo telefone OU mesmo e-mail => devolve o número já existente, sem nova linha.
 */
var LIMITE = 1000;
var ABAS = ['Plaspet', 'Mirimóveis'];
var CABECALHO = ['Nº Sorteio', 'Data/Hora', 'Nome', 'Telefone', 'E-mail', 'Marca', 'Consentimento LGPD'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // fila: um envio por vez => números nunca repetem
    var d = JSON.parse(e.postData.contents);

    var marca = String(d.marca || '');
    if (ABAS.indexOf(marca) < 0) return out({ ok: false, erro: 'marca' });
    var nome = String(d.nome || '').trim();
    var tel = String(d.telefone || '').trim();
    var email = String(d.email || '').trim();
    if (!nome || !tel || !email || d.consentimento !== true) return out({ ok: false, erro: 'dados' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abas = ABAS.map(function (n) { return aba(ss, n); });

    var telKey = normTel(tel), emailKey = email.toLowerCase();
    var maior = 0;
    for (var i = 0; i < abas.length; i++) {
      var n = abas[i].getLastRow();
      if (n < 2) continue;
      var v = abas[i].getRange(2, 1, n - 1, 5).getValues();
      for (var r = 0; r < v.length; r++) {
        var num = Number(v[r][0]) || 0;
        if (num > maior) maior = num;
        if (normTel(v[r][3]) === telKey || String(v[r][4]).trim().toLowerCase() === emailKey) {
          return out({ ok: true, numero: num, repetido: true });
        }
      }
    }

    var numero = maior + 1;
    if (numero > LIMITE) return out({ ok: false, erro: 'limite' });

    var sh = abas[ABAS.indexOf(marca)];
    var linha = sh.getLastRow() + 1;
    sh.getRange(linha, 4).setNumberFormat('@');
    sh.getRange(linha, 1, 1, 7).setValues([[numero, new Date(), nome, tel, email, marca, 'Sim']]);
    SpreadsheetApp.flush();
    return out({ ok: true, numero: numero });
  } catch (err) {
    return out({ ok: false, erro: 'servidor' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() { return out({ ok: true, status: 'online' }); }

function aba(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]).setFontWeight('bold').setBackground('#e8f0e8');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange('D:D').setNumberFormat('@');
    sh.setColumnWidths(1, 7, 140);
  }
  return sh;
}

function normTel(t) {
  var d = String(t).replace(/\D/g, '');
  if (d.length > 11 && d.indexOf('55') === 0) d = d.slice(2);
  return d;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
