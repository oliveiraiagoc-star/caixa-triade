// Caixa da Tríade · back-end na planilha.
// Publicado como app da Web: Executar como "Eu", Acesso "Qualquer pessoa".
// A senha fica só aqui (não vai pro GitHub). Pra trocar: mude SENHA e publique uma nova versão.
const SENHA = '__SENHA__';
const PLANILHA_ID = '1zB4IGCV6OHi3LzQQ77LAA1PbCPp2n1-z09bp3aYPSHA'; // planilha "Caixa da Tríade" no Drive da Tríade
let _ss = null;
function planilha() { return _ss || (_ss = SpreadsheetApp.openById(PLANILHA_ID)); }

const ABAS = {
  lancamentos: ['id', 'mes', 'data', 'tipo', 'valor (R$)', 'descricao', 'categoria', 'conta', 'pra conta', 'socio', 'lancado por', 'json'],
  fechamentos: ['mes', 'entradas (R$)', 'saidas (R$)', 'lucro (R$)', 'fechado em', 'json'],
  config: ['chave', 'json']
};

function doPost(e) { return responder(JSON.parse(e.postData.contents || '{}')); }
function doGet(e) { return responder(e.parameter || {}); }

function responder(p) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (p.senha !== SENHA) return out({ ok: false, erro: 'senha' });
  const f = ACOES[p.acao];
  if (!f) return out({ ok: false, erro: 'acao desconhecida' });
  try { return out({ ok: true, data: f(p) }); } catch (err) { return out({ ok: false, erro: String(err && err.message || err) }); }
}

function aba(nome) {
  const ss = planilha();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    // id, mês e data como texto, senão a planilha vira "2026-10" em data
    sh.getRange(nome === 'lancamentos' ? 'A:C' : 'A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function linhas(nome) {
  const sh = aba(nome), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, ABAS[nome].length).getValues();
}
function acharLinha(nome, chave) {
  const sh = aba(nome), n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(chave)) return i + 2;
  return -1;
}
function gravar(nome, chave, valores) {
  const sh = aba(nome), r = acharLinha(nome, chave);
  if (r > 0) sh.getRange(r, 1, 1, valores.length).setValues([valores]);
  else sh.appendRow(valores);
}
function apagarLinha(nome, chave) {
  const r = acharLinha(nome, chave);
  if (r > 0) aba(nome).deleteRow(r);
}
function comTrava(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function linhaLanc(d) {
  const sinal = d.tipo === 'saida' ? -1 : 1;
  const tipo = d.tipo === 'entrada' ? 'Entrada' : d.tipo === 'saida' ? 'Saída' : 'Transferência';
  const limpo = Object.assign({}, d);
  ['_conta', '_destino', '_socio', '_por'].forEach(k => delete limpo[k]);
  // o apóstrofo força texto; sem ele a planilha vira "2026-10" em data
  return ["'" + d.id, "'" + d.mes, "'" + d.data, tipo, sinal * d.valor / 100, d.descricao || '', d.categoria || '', d._conta || '', d._destino || '', d._socio || '', d._por || '', JSON.stringify(limpo)];
}

const ACOES = {
  estado(p) {
    const cfg = linhas('config').filter(l => l[0] === 'main')[0];
    const fechamentos = {};
    linhas('fechamentos').forEach(l => { fechamentos[String(l[0])] = JSON.parse(l[5]); });
    const lancamentos = linhas('lancamentos').filter(l => String(l[1]) === p.mes).map(l => JSON.parse(l[11]));
    return { config: cfg ? JSON.parse(cfg[1]) : null, fechamentos, lancamentos };
  },
  chaves(p) {
    const meses = p.meses || [];
    return linhas('lancamentos').filter(l => meses.indexOf(String(l[1])) >= 0).map(l => JSON.parse(l[11]).importKey).filter(Boolean);
  },
  salvar(p) { return comTrava(() => { gravar('lancamentos', p.doc.id, linhaLanc(p.doc)); return true; }); },
  salvarVarios(p) {
    return comTrava(() => {
      const sh = aba('lancamentos'), rows = p.docs.map(linhaLanc);
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      return rows.length;
    });
  },
  apagar(p) { return comTrava(() => { apagarLinha('lancamentos', p.id); return true; }); },
  salvarConfig(p) { return comTrava(() => { gravar('config', 'main', ['main', JSON.stringify(p.config)]); return true; }); },
  fechar(p) {
    const d = p.doc;
    return comTrava(() => { gravar('fechamentos', p.mes, ["'" + p.mes, d.entradas / 100, d.saidas / 100, d.lucro / 100, new Date(d.fechadoEm), JSON.stringify(d)]); return true; });
  },
  reabrir(p) { return comTrava(() => { apagarLinha('fechamentos', p.mes); return true; }); }
};

// Rode uma vez pelo editor pra criar as abas e gravar os sócios e contas.
function preparar() {
  Object.keys(ABAS).forEach(aba);
  if (!linhas('config').length) gravar('config', 'main', ['main', JSON.stringify({
    socios: [{ id: 's1', nome: 'Iago', pct: 33.34 }, { id: 's2', nome: 'Miguel', pct: 33.33 }, { id: 's3', nome: 'Luciano', pct: 33.33 }],
    contas: [{ id: 'mp', nome: 'Mercado Pago', saldoInicial: 0 }, { id: 'shopee', nome: 'Carteira Shopee', saldoInicial: 0 }, { id: 'banco', nome: 'Conta da empresa', saldoInicial: 0 }]
  })]);
  const padrao = planilha().getSheetByName('Página1') || planilha().getSheetByName('Sheet1');
  if (padrao && planilha().getSheets().length > 1) planilha().deleteSheet(padrao);
}
