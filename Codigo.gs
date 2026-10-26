/** @OnlyCurrentDoc */
/**
 * Rotina AET – Férias (app ligado à planilha DATAS TPB 13º Biênio)
 * Lê a aba CONTROLE_TPB e grava as marcações da rotina na própria planilha.
 */

var ABA_CONTROLE = 'CONTROLE_TPB';
var ABA_REGISTRO = 'REGISTRO_PROGRAMA';   // visível: o que foi feito, por quem e quando
var ABA_ESTADO   = '_ESTADO_PROGRAMA';    // oculta: dados internos do programa

function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.dados = JSON.stringify(carregarDados_()).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Rotina AET – Férias')
    .setFaviconUrl('https://raw.githubusercontent.com/aet1ciaind-cmd/rotina-aet/main/icon-512.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function carregarDados_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_CONTROLE);
  if (!sh) throw new Error('Aba ' + ABA_CONTROLE + ' não encontrada.');
  return {
    controle: sh.getDataRange().getDisplayValues(),
    estado: lerEstado_(ss),
    usuario: Session.getActiveUser().getEmail(),
    lidoEm: Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm')
  };
}

function lerEstado_(ss) {
  var sh = ss.getSheetByName(ABA_ESTADO);
  if (!sh || sh.getLastRow() === 0) return null;
  var partes = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; });
  try { return JSON.parse(partes.join('')); } catch (e) { return null; }
}

/** Chamado pelo programa a cada marcação. */
function salvarEstado(json, linhas) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) estado interno (dividido em pedaços de 40.000 caracteres)
    var est = ss.getSheetByName(ABA_ESTADO) || ss.insertSheet(ABA_ESTADO);
    est.clear();
    var pedacos = [];
    for (var i = 0; i < json.length; i += 40000) pedacos.push([json.substring(i, i + 40000)]);
    if (pedacos.length) est.getRange(1, 1, pedacos.length, 1).setNumberFormat('@').setValues(pedacos);
    if (!est.isSheetHidden()) est.hideSheet();

    // 2) registro legível
    var reg = ss.getSheetByName(ABA_REGISTRO) || ss.insertSheet(ABA_REGISTRO);
    var cab = ['Chave', 'Tipo', 'Data', 'Atividade / militar', 'Feito', 'Quem', 'Quando / etapas', 'Observação / registro'];
    reg.clear();
    var dados = [cab].concat(linhas || []);
    reg.getRange(1, 1, dados.length, cab.length).setNumberFormat('@').setValues(dados);
    reg.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#9b8a5c');
    reg.setFrozenRows(1);
    reg.hideColumns(1);

    return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'HH:mm');
  } finally {
    lock.releaseLock();
  }
}

/**
 * EXECUTAR UMA ÚNICA VEZ (no editor: escolher "limparAbasAntigas" e clicar em Executar).
 * Apaga as abas de rotina que foram substituídas pelo programa.
 * Se precisar desfazer: Arquivo › Histórico de versões.
 */
function limparAbasAntigas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ['HOJE', 'CHECKLIST FÉRIAS', 'PENDÊNCIAS TPB', 'ATIVIDADES AET',
              'SICI PASSO A PASSO', 'MODELOS PA', 'INFORMAÇÕES ÚTEIS'];
  var apagadas = [];
  abas.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) { ss.deleteSheet(sh); apagadas.push(n); }
  });
  var ctrl = ss.getSheetByName(ABA_CONTROLE);
  if (ctrl) ss.setActiveSheet(ctrl);
  Logger.log('Abas apagadas: ' + (apagadas.join(', ') || 'nenhuma'));
}

/**
 * EXECUTAR (no editor: escolher "protegerPlanilha" e clicar em Executar).
 * Pode ser executada de novo sem problema: refaz as proteções.
 * - Todas as abas: editáveis só pela conta dona (AET).
 * - CONTROLE_TPB: liberadas para os editores (conta S.Adm)
 *   a) as colunas RESULTADO TPB e CF SICI, em todas as turmas;
 *   b) NOTA EAD, DATA CF, RESULTADO CF (1º e 2º), LISTAGEM CTP, INTRANET,
 *      CF REAVALIAÇÃO e REAVALIAÇÃO, só nas turmas que começam em novembro ou dezembro de 2026
 *      (pela coluna DATA; ex.: 30/11 a 02/12 conta como novembro).
 * - REGISTRO_PROGRAMA e _ESTADO_PROGRAMA: só aviso, para o programa conseguir gravar.
 */
var COLS_NOV_DEZ = ['NOTA EAD', 'DATA CF', 'RESULTADO CF', 'LISTAGEM CTP',
                    'INTRANET', 'CF REAVALIAÇÃO', 'REAVALIAÇÃO'];

function protegerPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dono = Session.getEffectiveUser();
  var log = [];
  ss.getSheets().forEach(function (sh) {
    var nome = sh.getName();
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    var p = sh.protect().setDescription('Rotina AET – proteção automática');
    if (nome === ABA_REGISTRO || nome === ABA_ESTADO) {
      p.setWarningOnly(true);
      log.push(nome + ': só aviso (o programa grava aqui)');
      return;
    }
    p.addEditor(dono);
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== dono.getEmail(); }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
    if (nome === ABA_CONTROLE) {
      var libera = intervalosControle_(sh);
      p.setUnprotectedRanges(libera.ranges);
      log.push(nome + ': travada, exceto ' + libera.resumo);
    } else {
      log.push(nome + ': travada');
    }
  });
  Logger.log(log.join('\n'));
}

/** Monta os intervalos liberados da CONTROLE_TPB. */
function intervalosControle_(sh) {
  var v = sh.getDataRange().getDisplayValues();
  var hi = -1;
  for (var r = 0; r < v.length && hi < 0; r++) {
    if (v[r].indexOf('NUMERO') >= 0 && v[r].indexOf('NOME') >= 0) hi = r;
  }
  if (hi < 0) throw new Error('Cabeçalho (NUMERO / NOME) não encontrado na ' + ABA_CONTROLE + '. Nada foi liberado nela.');
  var cab = v[hi];
  var colData = cab.indexOf('DATA');
  var colTpb = cab.indexOf('RESULTADO TPB');
  var colSici = cab.indexOf('CF SICI');
  if (colData < 0 || colTpb < 0 || colSici < 0) throw new Error('Colunas DATA, RESULTADO TPB ou CF SICI não encontradas na ' + ABA_CONTROLE + '. Nada foi liberado nela.');

  var primeiraLinha = hi + 2;                       // 1ª linha de dados (base 1)
  var ultimaLinha = sh.getMaxRows();
  var ranges = [
    sh.getRange(primeiraLinha, colTpb + 1, ultimaLinha - primeiraLinha + 1, 1),
    sh.getRange(primeiraLinha, colSici + 1, ultimaLinha - primeiraLinha + 1, 1)
  ];

  // colunas de novembro/dezembro (inclui as duas "RESULTADO CF")
  var colsNov = [];
  cab.forEach(function (h, i) { if (COLS_NOV_DEZ.indexOf(h) >= 0) colsNov.push(i); });

  // linhas de turmas que começam em nov/dez de 2026
  var linhas = [];
  for (var k = hi + 1; k < v.length; k++) {
    var m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(v[k][colData] || '');
    if (m && m[3] === '2026' && (m[2] === '11' || m[2] === '12')) linhas.push(k + 1);
  }
  // agrupa linhas seguidas em blocos
  var blocos = [];
  linhas.forEach(function (lin) {
    var ult = blocos[blocos.length - 1];
    if (ult && ult[1] === lin - 1) ult[1] = lin; else blocos.push([lin, lin]);
  });
  colsNov.forEach(function (c) {
    blocos.forEach(function (b) {
      ranges.push(sh.getRange(b[0], c + 1, b[1] - b[0] + 1, 1));
    });
  });

  var resumo = 'as colunas RESULTADO TPB e CF SICI (todas as turmas)';
  resumo += linhas.length
    ? ' e as colunas de resultado e CF das turmas de novembro e dezembro (' + linhas.length + ' linhas)'
    : ' (nenhuma turma de novembro/dezembro encontrada)';
  return { ranges: ranges, resumo: resumo };
}