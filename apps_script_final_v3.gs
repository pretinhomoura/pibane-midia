// ════════════════════════════════════════════════════════════════
//  APPS SCRIPT COMPLETO — PIBANE MÍDIA
//  1. Apague TUDO que está no Apps Script
//  2. Cole este código inteiro
//  3. Salve com Ctrl+S
//  4. Implantar → Gerenciar implantações → lápis → Nova versão → Implantar
// ════════════════════════════════════════════════════════════════

const SHEET_ID    = '1jr_nznZt1xeE9_YkC10GHvl9Uz4khLyf3yPubN4fH0o';
const EMAIL_ADMIN = 'pretinho@corredor5.art';
const NOME_IGREJA = 'Pibane';
const NOME_ADMIN  = 'Pretinho Moura';

// ── ROTEADOR ────────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    let r;
    switch (d.acao) {
      case 'listar_voluntarios':    r = listarVoluntarios();          break;
      case 'listar_eventos':        r = listarEventos();              break;
      case 'listar_escalas':        r = listarEscalas(d.volNome);     break;
      case 'listar_disponibilidade':r = listarDisponibilidade();      break;
      case 'disponibilidade':       r = salvarDisponibilidade(d);     break;
      case 'confirmar':             r = confirmarPresenca(d);         break;
      case 'troca':                 r = registrarTroca(d);            break;
      case 'enviar_whatsapp':       r = montarMensagensWA();          break;
      case 'montar_escala':         r = montarEscalaAutomatica();    break;
      case 'corrigir_datas':        corrigirDatasDisponibilidade(); r = {status:'ok'}; break;
      default: r = { status: 'erro', msg: 'Ação desconhecida' };
    }
    return ok(r);
  } catch (err) {
    return ok({ status: 'erro', msg: err.toString() });
  }
}

function doGet(e) {
  try {
    const acao = (e && e.parameter && e.parameter.acao) ? e.parameter.acao : '';
    let r;
    switch (acao) {
      case 'listar_voluntarios':    r = listarVoluntarios();    break;
      case 'listar_eventos':        r = listarEventos();        break;
      case 'listar_escalas':        r = listarEscalas('');      break;
      case 'listar_disponibilidade':r = listarDisponibilidade();break;
      default: r = { status: 'ok', msg: 'Pibane API ativa' };
    }
    return ok(r);
  } catch (err) {
    return ok({ status: 'erro', msg: err.toString() });
  }
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPER: lê aba e retorna array de objetos ────────────────────
function lerAba(nome) {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const sh   = ss.getSheetByName(nome);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const cols = rows[0];
  const lista = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] && rows[i][0] !== 0) continue;
    const obj = {};
    cols.forEach((c, j) => {
      let val = rows[i][j];
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      obj[String(c).toLowerCase().trim().replace(/\s+/g,'_')] = val;
    });
    lista.push(obj);
  }
  return lista;
}

// ── LEITURA ──────────────────────────────────────────────────────
function listarVoluntarios() {
  const dados = lerAba('voluntarios');
  return { status: 'ok', dados };
}

function listarEventos() {
  const dados = lerAba('eventos');
  return { status: 'ok', dados };
}

function listarEscalas(volNome) {
  const escalas  = lerAba('escalas');
  const eventos  = lerAba('eventos');

  // Monta dicionário de eventos por id
  const evMap = {};
  eventos.forEach(ev => {
    const id = String(ev.id || '').trim();
    evMap[id] = ev;
  });

  const lista = escalas
    .filter(e => {
      if (!volNome) return true;
      const nome = String(e.voluntario_id || e.voluntario || e.nome || '').toLowerCase().trim();
      return nome.includes(volNome.toLowerCase().trim());
    })
    .map(e => {
      const evId = String(e.evento_id || e.evento || '').trim();
      const ev   = evMap[evId] || {};
      return {
        id:         e.id,
        evento_id:  evId,
        voluntario: String(e.voluntario_id || e.voluntario || '').trim(),
        funcao:     String(e.funcao || '').trim(),
        status:     String(e.status || 'pendente').trim(),
        data:       String(ev.data  || '').trim(),
        hora:       String(ev.hora  || '').trim(),
        tipo:       String(ev.tipo  || '').trim(),
      };
    });

  return { status: 'ok', dados: lista };
}

function listarDisponibilidade() {
  const dados = lerAba('disponibilidade');
  return { status: 'ok', dados };
}

// ── SALVAR DISPONIBILIDADE ───────────────────────────────────────
function salvarDisponibilidade(d) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh   = ss.getSheetByName('disponibilidade');
  if (!sh) {
    sh = ss.insertSheet('disponibilidade');
    sh.appendRow(['timestamp','voluntario','funcao','data_id','data','culto','disponivel','funcoes']);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
    sh.setFrozenRows(1);
  }

  // Remove respostas antigas desse voluntário
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === String(d.voluntarioNome).trim()) sh.deleteRow(i + 1);
  }

  // Insere novas
  d.respostas.forEach(r => {
    const disp = r.resposta === 'sim' ? '✅ Sim' : r.resposta === 'nao' ? '❌ Não' : '— Sem resposta';
    sh.appendRow([d.timestamp, d.voluntarioNome, d.funcao, r.dataId, r.data, r.culto, disp, r.funcoes||'']);
    const cor = r.resposta === 'sim' ? '#dcfce7' : r.resposta === 'nao' ? '#fee2e2' : '#fef3c7';
    sh.getRange(sh.getLastRow(), 1, 1, 7).setBackground(cor);
  });

  atualizarPainelAdmin(ss);
  return { status: 'ok', msg: 'Disponibilidade salva!' };
}

// ── CONFIRMAR PRESENÇA ───────────────────────────────────────────
function confirmarPresenca(d) {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  const sh  = ss.getSheetByName('escalas');
  if (!sh) return { status: 'erro', msg: 'Aba escalas não encontrada' };

  const rows = sh.getDataRange().getValues();
  const cols = rows[0];
  const iVol    = cols.findIndex(c => /voluntario/i.test(String(c)));
  const iStatus = cols.findIndex(c => /status/i.test(String(c)));
  const iConf   = cols.findIndex(c => /confirmad/i.test(String(c)));

  for (let i = 1; i < rows.length; i++) {
    const nomeNaPlanilha = String(rows[i][iVol] || '').toLowerCase().trim();
    const nomeBuscado    = String(d.voluntarioNome || '').toLowerCase().trim();
    if (nomeNaPlanilha.includes(nomeBuscado) || nomeBuscado.includes(nomeNaPlanilha)) {
      if (iStatus >= 0) sh.getRange(i+1, iStatus+1).setValue('confirmado');
      if (iConf   >= 0) sh.getRange(i+1, iConf+1).setValue(d.timestamp);
      sh.getRange(i+1, 1, 1, cols.length).setBackground('#dcfce7');
      atualizarPainelAdmin(ss);
      return { status: 'ok', msg: 'Presença confirmada!' };
    }
  }

  // Se não achou, registra em aba separada
  let shC = ss.getSheetByName('confirmacoes');
  if (!shC) {
    shC = ss.insertSheet('confirmacoes');
    shC.appendRow(['timestamp','voluntario','funcao','escala_id']);
    shC.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
  }
  shC.appendRow([d.timestamp, d.voluntarioNome, d.funcao, d.escalaId]);
  shC.getRange(shC.getLastRow(),1,1,4).setBackground('#dcfce7');
  return { status: 'ok', msg: 'Confirmação registrada!' };
}

// ── REGISTRAR TROCA ──────────────────────────────────────────────
function registrarTroca(d) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh   = ss.getSheetByName('trocas');
  if (!sh) {
    sh = ss.insertSheet('trocas');
    sh.appendRow(['timestamp','voluntario','funcao','escala_id','substituto','status','resolvido_em']);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
    sh.setFrozenRows(1);
  }

  sh.appendRow([d.timestamp, d.voluntarioNome, d.funcao, d.escalaId, d.substitutoNome, 'pendente', '']);
  sh.getRange(sh.getLastRow(), 1, 1, 7).setBackground('#fef3c7');

  // Marca na aba escalas
  const shEsc = ss.getSheetByName('escalas');
  if (shEsc) {
    const rows = shEsc.getDataRange().getValues();
    const cols = rows[0];
    const iVol    = cols.findIndex(c => /voluntario/i.test(String(c)));
    const iStatus = cols.findIndex(c => /status/i.test(String(c)));
    for (let i = 1; i < rows.length; i++) {
      const nome = String(rows[i][iVol]||'').toLowerCase().trim();
      if (nome.includes(d.voluntarioNome.toLowerCase().trim())) {
        if (iStatus >= 0) shEsc.getRange(i+1, iStatus+1).setValue('troca_pedida');
        shEsc.getRange(i+1, 1, 1, cols.length).setBackground('#fef3c7');
        break;
      }
    }
  }

  enviarEmailTroca(d, ss);
  atualizarPainelAdmin(ss);
  return { status: 'ok', msg: 'Pedido de troca registrado!' };
}

// ── E-MAIL NA TROCA ──────────────────────────────────────────────
function enviarEmailTroca(d, ss) {
  try {
    const vols = lerAba('voluntarios');
    let emailVol = '', emailSub = '';
    vols.forEach(v => {
      const nome = String(v.nome || '').toLowerCase().trim();
      if (nome.includes(d.voluntarioNome.toLowerCase().trim())) emailVol = v.email || '';
      if (nome.includes((d.substitutoNome||'').toLowerCase().trim())) emailSub = v.email || '';
    });

    const assunto = `🔄 Pedido de troca — ${NOME_IGREJA} Mídia`;
    const html = (titulo, corpo) => `
      <div style="font-family:sans-serif;max-width:480px;margin:0 auto;background:#f8fafc;padding:24px;border-radius:12px">
        <div style="background:#0a0f1e;border-radius:8px;padding:16px 20px;margin-bottom:20px">
          <h2 style="color:#fff;margin:0;font-size:18px">⛪ ${NOME_IGREJA} Mídia</h2>
        </div>
        <h3 style="color:#0a0f1e">${titulo}</h3>
        ${corpo}
        <p style="color:#94a3b8;font-size:11px;margin-top:20px">${NOME_IGREJA} Mídia · Sistema de Voluntários</p>
      </div>`;

    if (emailVol) {
      GmailApp.sendEmail(emailVol, assunto, '', { htmlBody: html(
        `Olá, ${d.voluntarioNome}! Sua troca foi registrada 🔄`,
        `<div style="background:#fef3c7;border:1px solid #fcd34d;border-radius:8px;padding:14px;margin:16px 0">
          <p style="margin:0;color:#78350f">Função: <strong>${d.funcao}</strong></p>
          <p style="margin:6px 0 0;color:#78350f">Substituto sugerido: <strong>${d.substitutoNome||'A definir'}</strong></p>
          <p style="margin:4px 0 0;color:#78350f">Status: <strong>Aguardando confirmação do líder</strong></p>
        </div>
        <p style="color:#334155">O líder ${NOME_ADMIN} foi notificado e confirmará em breve.</p>`
      )});
    }

    if (emailSub && emailSub !== emailVol) {
      GmailApp.sendEmail(emailSub, `📲 Você foi sugerido como substituto — ${NOME_IGREJA}`, '', { htmlBody: html(
        `Você foi sugerido como substituto! 🙌`,
        `<div style="background:#dbeafe;border:1px solid #93c5fd;border-radius:8px;padding:14px;margin:16px 0">
          <p style="margin:0;color:#1e40af"><strong>${d.voluntarioNome}</strong> precisa de troca na função <strong>${d.funcao}</strong>.</p>
          <p style="margin:6px 0 0;color:#1d4ed8">Aguarde o líder <strong>${NOME_ADMIN}</strong> confirmar.</p>
        </div>`
      )});
    }

    GmailApp.sendEmail(EMAIL_ADMIN, `⚠️ Troca pedida — ${d.voluntarioNome} (${d.funcao})`, '', { htmlBody: html(
      `⚠️ Pedido de troca`,
      `<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:14px;margin:16px 0">
        <p style="margin:0;color:#7f1d1d"><strong>${d.voluntarioNome}</strong> pediu troca</p>
        <p style="margin:6px 0 0;color:#991b1b">Função: <strong>${d.funcao}</strong></p>
        <p style="margin:4px 0 0;color:#991b1b">Substituto: <strong>${d.substitutoNome||'Nenhum'}</strong></p>
        <p style="margin:4px 0 0;color:#991b1b">Em: ${d.timestamp}</p>
      </div>
      <a href="https://docs.google.com/spreadsheets/d/${SHEET_ID}" style="display:inline-block;margin-top:12px;padding:10px 20px;background:#1d4ed8;color:#fff;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600">Abrir planilha →</a>`
    )});
  } catch(err) {
    Logger.log('Erro e-mail: ' + err.toString());
  }
}

// ── GERAR MENSAGENS WHATSAPP ─────────────────────────────────────
function montarMensagensWA() {
  const vols    = lerAba('voluntarios');
  const escalas = lerAba('escalas');
  const eventos = lerAba('eventos');

  const evMap = {};
  eventos.forEach(ev => evMap[String(ev.id).trim()] = ev);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let shWA = ss.getSheetByName('mensagens_wa');
  if (!shWA) {
    shWA = ss.insertSheet('mensagens_wa');
    shWA.appendRow(['nome','telefone','funcao','mensagem','gerado_em']);
    shWA.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
    shWA.setFrozenRows(1);
  } else {
    if (shWA.getLastRow() > 1) shWA.getRange(2,1,shWA.getLastRow()-1,5).clear();
  }

  const mensagens = [];
  escalas.forEach(e => {
    const nomeEsc = String(e.voluntario_id || e.voluntario || '').trim();
    const evId    = String(e.evento_id || e.evento || '').trim();
    const ev      = evMap[evId] || {};
    const vol     = vols.find(v => String(v.nome||'').trim().toLowerCase() === nomeEsc.toLowerCase()) || {};
    const funcao  = String(e.funcao || vol.funcoes || '').trim();
    const tel     = String(vol.telefone || '').trim();
    const data    = String(ev.data || '').trim();
    const hora    = String(ev.hora || '').trim();
    const tipo    = String(ev.tipo || 'Culto').trim();

    const msg =
      `Olá, ${nomeEsc}! 👋\n\n` +
      `Você está escalado(a):\n` +
      `📅 ${data} às ${hora}\n` +
      `⛪ ${NOME_IGREJA} — ${tipo}\n` +
      `🎬 Função: *${funcao}*\n\n` +
      `Responda:\n` +
      `✅ *1* — Confirmar presença\n` +
      `🔄 *2* — Preciso de troca`;

    shWA.appendRow([nomeEsc, tel, funcao, msg, new Date().toLocaleString('pt-BR')]);
    shWA.getRange(shWA.getLastRow(),1,1,5).setBackground('#f0fdf4');
    shWA.getRange(shWA.getLastRow(),4).setWrap(true);
    mensagens.push({ nome: nomeEsc, tel, funcao, mensagem: msg });
  });

  shWA.setColumnWidth(4, 320);
  return { status: 'ok', mensagens };
}

// ── PAINEL ADMIN ─────────────────────────────────────────────────
function atualizarPainelAdmin(ss) {
  let sh = ss.getSheetByName('resumo_lider');
  if (!sh) sh = ss.insertSheet('resumo_lider');
  sh.clearContents(); sh.clearFormats();

  sh.getRange(1,1).setValue('PAINEL DO LÍDER — PIBANE MÍDIA');
  sh.getRange(1,1,1,5).merge().setBackground('#0a0f1e').setFontColor('#fff').setFontSize(14).setFontWeight('bold');
  sh.getRange(2,1).setValue('Atualizado: ' + new Date().toLocaleString('pt-BR')).setFontColor('#64748b');

  const disp = lerAba('disponibilidade');
  if (disp.length) {
    sh.getRange(4,1).setValue('DISPONIBILIDADE');
    sh.getRange(4,1,1,5).merge().setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
    sh.getRange(5,1).setValue('Voluntário'); sh.getRange(5,2).setValue('Data');
    sh.getRange(5,3).setValue('Culto'); sh.getRange(5,4).setValue('Disponível');
    sh.getRange(5,1,1,4).setBackground('#334155').setFontColor('#fff').setFontWeight('bold');

    let l = 6;
    disp.forEach(r => {
      sh.getRange(l,1).setValue(r.voluntario); sh.getRange(l,2).setValue(r.data);
      sh.getRange(l,3).setValue(r.culto);      sh.getRange(l,4).setValue(r.disponivel);
      const cor = String(r.disponivel).includes('Sim') ? '#dcfce7' : String(r.disponivel).includes('Não') ? '#fee2e2' : '#fef3c7';
      sh.getRange(l,1,1,4).setBackground(cor);
      l++;
    });

    // Resumo por data
    l += 2;
    sh.getRange(l,1).setValue('RESUMO POR DATA');
    sh.getRange(l,1,1,4).merge().setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold'); l++;
    sh.getRange(l,1).setValue('Data'); sh.getRange(l,2).setValue('Podem');
    sh.getRange(l,3).setValue('Não podem'); sh.getRange(l,4).setValue('Total');
    sh.getRange(l,1,1,4).setBackground('#334155').setFontColor('#fff').setFontWeight('bold'); l++;

    const por = {};
    disp.forEach(r => {
      const dt = r.data || '?';
      if (!por[dt]) por[dt] = { sim: [], nao: [] };
      if (String(r.disponivel).includes('Sim')) por[dt].sim.push(r.voluntario);
      else por[dt].nao.push(r.voluntario);
    });

    Object.keys(por).forEach(dt => {
      const p = por[dt];
      sh.getRange(l,1).setValue(dt);
      sh.getRange(l,2).setValue(p.sim.join(', ')||'—').setBackground('#dcfce7');
      sh.getRange(l,3).setValue(p.nao.join(', ')||'—').setBackground('#fee2e2');
      sh.getRange(l,4).setValue(p.sim.length + ' pessoa(s)');
      sh.getRange(l,4).setBackground(p.sim.length>=4?'#dcfce7':p.sim.length>=2?'#fef3c7':'#fee2e2');
      l++;
    });
  }

  [1,2,3,4,5].forEach(c => { try { sh.autoResizeColumn(c); } catch(e){} });
}

// ── TESTES ───────────────────────────────────────────────────────
function testarLeituraVoluntarios() {
  const r = listarVoluntarios();
  Logger.log('Total: ' + r.dados.length);
  r.dados.forEach(v => Logger.log(v.nome + ' | ' + v.funcoes));
}

function testarLeituraEventos() {
  const r = listarEventos();
  Logger.log('Total: ' + r.dados.length);
  r.dados.forEach(e => Logger.log(e.id + ' | ' + e.data + ' | ' + e.tipo));
}

function testarLeituraEscalas() {
  const r = listarEscalas('');
  Logger.log('Total: ' + r.dados.length);
  r.dados.forEach(e => Logger.log(e.voluntario + ' | ' + e.funcao + ' | ' + e.data + ' | ' + e.status));
}

function testarTudo() {
  Logger.log('=== VOLUNTÁRIOS ===');
  testarLeituraVoluntarios();
  Logger.log('=== EVENTOS ===');
  testarLeituraEventos();
  Logger.log('=== ESCALAS ===');
  testarLeituraEscalas();
}

// ════════════════════════════════════════════════════════════════
//  MONTAR ESCALA AUTOMATICAMENTE
//  Execute esta função manualmente quando quiser gerar a escala
//  ou chame pelo HTML com acao: 'montar_escala'
// ════════════════════════════════════════════════════════════════

function montarEscalaAutomatica() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const disp = lerAba('disponibilidade');
  const evs  = lerAba('eventos');
  const shEsc= ss.getSheetByName('escalas');

  if (!shEsc) { Logger.log('Aba escalas não encontrada'); return; }

  // Limpa escalas pendentes antigas (mantém confirmadas)
  const rowsEsc = shEsc.getDataRange().getValues();
  for (let i = rowsEsc.length - 1; i >= 1; i--) {
    const status = String(rowsEsc[i][4] || '').toLowerCase();
    if (status === 'pendente' || status === '') {
      shEsc.deleteRow(i + 1);
    }
  }

  // Agrupa disponibilidade por data
  const porData = {};
  disp.forEach(d => {
    const dataId = String(d.data_id || d.dataid || '').trim();
    if (!porData[dataId]) porData[dataId] = [];
    if (String(d.disponivel || '').includes('Sim')) {
      porData[dataId].push({
        nome:   String(d.voluntario || '').trim(),
        funcoes: String(d.funcoes || '').trim()
      });
    }
  });

  // Para cada evento, distribui voluntários pelas funções
  const FUNS_ORDEM = ['Câmera Móvel', 'Câmera Fixa', 'Áudio', 'Corte'];
  let idContador = shEsc.getLastRow(); // começa do último id
  let totalEscalados = 0;

  evs.forEach(ev => {
    const evId   = String(ev.id || '').trim();
    const dispEv = porData[evId] || [];
    if (!dispEv.length) return;

    const usados = new Set();

    FUNS_ORDEM.forEach(funcao => {
      // Acha voluntário disponível para essa função
      const candidato = dispEv.find(d => {
        if (usados.has(d.nome)) return false;
        // Se o voluntário marcou funções específicas, verifica
        if (d.funcoes) {
          return d.funcoes.toLowerCase().includes(funcao.toLowerCase().split(' ')[0]);
        }
        return true; // se não marcou função, pode qualquer
      });

      if (candidato) {
        idContador++;
        usados.add(candidato.nome);
        shEsc.appendRow([idContador, evId, candidato.nome, funcao, 'pendente', '']);
        shEsc.getRange(shEsc.getLastRow(), 1, 1, 6).setBackground('#fef3c7');
        totalEscalados++;
      }
    });

    // Voluntários disponíveis sem função atribuída ficam como reserva
    dispEv.forEach(d => {
      if (!usados.has(d.nome)) {
        idContador++;
        shEsc.appendRow([idContador, evId, d.nome, 'Reserva', 'pendente', '']);
        shEsc.getRange(shEsc.getLastRow(), 1, 1, 6).setBackground('#f0f9ff');
      }
    });
  });

  atualizarPainelAdmin(ss);
  Logger.log('✅ Escala montada! ' + totalEscalados + ' voluntários escalados.');
  return { status: 'ok', msg: totalEscalados + ' voluntários escalados com sucesso!' };
}

// Corrige formato de data no painel do líder
function corrigirDatasDisponibilidade() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('disponibilidade');
  if (!sh) return;
  const rows = sh.getDataRange().getValues();
  const cols = rows[0];
  const iData = cols.findIndex(c => /^data$/i.test(String(c)));
  if (iData < 0) return;
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][iData];
    if (val instanceof Date) {
      sh.getRange(i+1, iData+1).setValue(
        Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy")
      );
    }
  }
  Logger.log('Datas corrigidas!');
}
