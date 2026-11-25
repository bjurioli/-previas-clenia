// Prévias Clênia: recebe comentários e decisões do site e guarda na planilha.
// Cole este código em Extensões > Apps Script da planilha "Prévias Clênia | Aprovações".

const ABA = 'Registros';
const CABECALHO = ['Data', 'Prévia', 'Post', 'Título do post', 'Tipo', 'Onde', 'Nome', 'Texto', 'Decisão'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(ABA);
  if (!s) {
    s = ss.insertSheet(ABA);
    s.appendRow(CABECALHO);
    s.setFrozenRows(1);
  }
  return s;
}

function resposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function limpa_(valor, max) {
  let t = String(valor || '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t; // evita que o texto vire fórmula
  return t;
}

function doGet(e) {
  const previa = String((e && e.parameter && e.parameter.previa) || '').slice(0, 60);
  const valores = aba_().getDataRange().getValues();
  valores.shift();
  const linhas = valores
    .filter(function (r) { return !previa || r[1] === previa; })
    .map(function (r) {
      return {
        data: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        post: String(r[2]), tipo: String(r[4]), onde: String(r[5]),
        nome: String(r[6]).replace(/^'/, ''), texto: String(r[7]).replace(/^'/, ''), decisao: String(r[8])
      };
    });
  return resposta_({ ok: true, linhas: linhas });
}

function doPost(e) {
  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const tipo = String(d.tipo || '');
    if (['comentario', 'decisao'].indexOf(tipo) < 0) return resposta_({ ok: false, erro: 'tipo' });
    const nome = limpa_(d.nome, 80).trim();
    if (!nome) return resposta_({ ok: false, erro: 'nome' });
    const texto = limpa_(d.texto, 2000);
    if (tipo === 'comentario' && !texto.trim()) return resposta_({ ok: false, erro: 'texto' });
    const decisao = tipo === 'decisao' ? (d.decisao === 'aprovado' ? 'aprovado' : 'correcao') : '';
    const onde = ['video', 'legenda', 'post'].indexOf(d.onde) >= 0 ? d.onde : 'post';
    aba_().appendRow([new Date(), limpa_(d.previa, 60), limpa_(d.post, 40), limpa_(d.titulo, 120), tipo, onde, nome, texto, decisao]);
    return resposta_({ ok: true });
  } catch (err) {
    return resposta_({ ok: false, erro: 'falha' });
  } finally {
    trava.releaseLock();
  }
}
