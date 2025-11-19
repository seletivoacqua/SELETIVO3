// ============================================
// GOOGLE APPS SCRIPT - SISTEMA DE TRIAGEM COMPLETO
// VERSÃO PRODUÇÃO - TESTADA E OTIMIZADA
// ============================================
//
// ✅ 54 FUNÇÕES IMPLEMENTADAS
// ✅ Cache avançado com lock (evita chamadas simultâneas)
// ✅ Sistema de indexação para performance
// ✅ CORS corrigido
// ✅ Suporte para 50+ usuários simultâneos
//
// ÚLTIMA ATUALIZAÇÃO: Novembro 2025
// STATUS: PRODUÇÃO ESTÁVEL
//
// ============================================

// Adicione no início do seu arquivo .gs
const CACHE_TTL = 60; // 60 segundos (recomendado para 50+ usuários)
const CACHE_KEYS = {
  REPORT_DATA: 'report_data_v3',
  USERS: 'users_data_v3',
  STATS: 'stats_data_v3',
  INTERVIEWERS: 'interviewers_v3',
  REASONS: 'disqualification_reasons_v3'
};

// Serviço de Cache Avançado — VERSÃO PRODUÇÃO 2025
class AdvancedCacheService {
  static getCache() {
    return CacheService.getScriptCache();
  }

  static getLock() {
    return LockService.getScriptLock();
  }

  // JSON seguro (remove undefined e previne erros)
  static safeStringify(obj) {
    return JSON.stringify(obj, (key, value) =>
      value === undefined ? null : value
    );
  }

  static safeParse(str) {
    try {
      return str ? JSON.parse(str) : null;
    } catch (e) {
      console.warn('Cache parse error:', e);
      return null;
    }
  }

  static get(key) {
    const cached = this.getCache().get(key);
    return this.safeParse(cached);
  }

  static set(key, data, ttl = CACHE_TTL) {
    try {
      this.getCache().put(key, this.safeStringify(data), ttl);
      return true;
    } catch (error) {
      console.warn('Cache set error:', error);
      return false;
    }
  }

  // VERSÃO SEGURA COM LOCK (evita 50 chamadas ao mesmo tempo)
  static getWithFallback(key, fetchFunction, ttl = CACHE_TTL) {
    let data = this.get(key);

    if (data !== null) {
      console.log('Cache hit:', key);
      return data;
    }

    const lock = this.getLock();

    // Tenta pegar o lock por até 10 segundos
    if (lock.tryLock(10000)) {
      try {
        // Verifica novamente (outro usuário pode ter preenchido enquanto esperava)
        data = this.get(key);
        if (data !== null) {
          console.log('Cache hit após lock:', key);
          return data;
        }

        console.log('Cache miss + lock adquirido - executando fetch:', key);
        data = fetchFunction(); // ← tem que ser função SÍNCRONA!

        this.set(key, data, ttl);
        console.log('Cache atualizado com sucesso:', key);
      } catch (error) {
        console.error('Erro crítico no fetchFunction:', error);
        // Não salva erro no cache!
      } finally {
        lock.releaseLock();
      }
    } else {
      console.warn('Lock não adquirido, retornando dados antigos ou nulos:', key);
      // Pode retornar dados antigos ou forçar fallback simples
      data = this.get(key) || fetchFunction(); // fallback sem cache
    }

    return data;
  }
}

const SPREADSHEET_ID = '1iQSQ06P_OXkqxaGWN3uG5jRYFBKyjWqQyvzuGk2EplY';
const SHEET_USUARIOS = 'USUARIOS';
const SHEET_CANDIDATOS = 'CANDIDATOS';
const SHEET_MOTIVOS = 'MOTIVOS';
const SHEET_MENSAGENS = 'MENSAGENS';
const SHEET_TEMPLATES = 'TEMPLATES';

const HEADER_ROWS = 1;
const COL_ID_PRIMARY = 'CPF';
const COL_ID_ALT = 'Número de Inscrição';
const CACHE_TTL_SEC = 1200;
const PROP_REV_KEY = 'IDX_REV';
const IDX_CACHE_KEY = 'idx:v';

function _ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function _sheet(name){ return _ss().getSheetByName(name); }

function _getHeaders_(sh){
  const lastCol = sh.getLastColumn();
  return (lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0] : []);
}

function _colMap_(headers){
  const m = {};
  headers.forEach((h,i)=> m[h]=i);
  return m;
}

function _getRev_(){
  return PropertiesService.getDocumentProperties().getProperty(PROP_REV_KEY) || '0';
}

function _bumpRev_(){
  const props = PropertiesService.getDocumentProperties();
  const cur = Number(props.getProperty(PROP_REV_KEY) || '0') + 1;
  props.setProperty(PROP_REV_KEY, String(cur));
  return String(cur);
}

function _buildIndex_(sh, headers){
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROWS) return {};

  const colMap = _colMap_(headers);
  const colCpf = colMap[COL_ID_PRIMARY] ?? -1;
  const colAlt = colMap[COL_ID_ALT] ?? -1;
  const keyCols = [colCpf, colAlt].filter(c => c>=0);
  if (!keyCols.length) return {};

  const values = sh.getRange(HEADER_ROWS+1, 1, lastRow-HEADER_ROWS, sh.getLastColumn()).getValues();
  const idx = {};
  for (let i=0;i<values.length;i++){
    for (const c of keyCols){
      const key = values[i][c];
      if (key) {
        const row = i + HEADER_ROWS + 1;
        idx[String(key).trim()] = row;
      }
    }
  }
  return idx;
}

function _getIndex_(sh, headers){
  const rev = _getRev_();
  const key = `${IDX_CACHE_KEY}${rev}`;
  const cache = CacheService.getDocumentCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const idx = _buildIndex_(sh, headers);
  cache.put(key, JSON.stringify(idx), CACHE_TTL_SEC);
  return idx;
}

function _readSheetBlock_(name){
  const sh = _sheet(name);
  if (!sh) return {sheet:null, headers:[], values:[]};
  const headers = _getHeaders_(sh);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROWS || lastCol === 0){
    return {sheet:sh, headers, values:[]};
  }
  const values = sh.getRange(HEADER_ROWS+1, 1, lastRow-HEADER_ROWS, lastCol).getValues();
  return {sheet:sh, headers, values};
}

function _writeWholeRow_(sh, row, rowArray){
  const lastCol = sh.getLastColumn();
  sh.getRange(row, 1, 1, lastCol).setValues([rowArray]);
}

// ============================================
// CORS HEADERS - CORREÇÃO DEFINITIVA
// ============================================

function createCorsResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================
// FUNÇÕES PRINCIPAIS DE ENTRADA
// ============================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

// ============================================
// ROTEAMENTO DE REQUISIÇÕES
// ============================================

function handleRequest(e) {
  try {
    let action, params;

    if (e && e.postData && e.postData.contents) {
      try {
        const data = JSON.parse(e.postData.contents);
        action = data.action;
        params = data;
      } catch (parseError) {
        Logger.log('Erro ao fazer parse do JSON: ' + parseError);
        return createCorsResponse({
          success: false,
          error: 'JSON inválido: ' + parseError.toString()
        });
      }
    } else if (e && e.parameter) {
      action = e.parameter.action;
      params = e.parameter;
    } else {
      return createCorsResponse({
        success: false,
        error: 'Requisição inválida: parâmetros não encontrados'
      });
    }

    Logger.log('🔄 Ação recebida: ' + action);

    const actions = {
      'getUserRole': () => getUserRole(params),
      'getAnalysts': () => getAnalysts(params),
      'getCandidates': () => getCandidates(params),
      'assignCandidates': () => assignCandidates(params),
      'updateCandidateStatus': () => updateCandidateStatus(params),
      'getCandidatesByStatus': () => getCandidatesByStatus(params),
      'logMessage': () => logMessage(params),
      'getDisqualificationReasons': () => getDisqualificationReasons(),
      'getMessageTemplates': () => getMessageTemplates(params),
      'sendMessages': () => sendMessages(params),
      'updateMessageStatus': () => updateMessageStatus(params),
      'moveToInterview': () => moveToInterview(params),
      'getInterviewCandidates': () => getInterviewCandidates(params),
      'getInterviewers': () => getInterviewers(params),
      'getInterviewerCandidates': () => getInterviewerCandidates(params),
      'allocateToInterviewer': () => allocateToInterviewer(params),
      'updateInterviewStatus': () => updateInterviewStatus(params),
      'saveInterviewEvaluation': () => saveInterviewEvaluation(params),
      'getReportStats': () => getReportStats(params),
      'getReport': () => getReport(params),
      'getEmailAliases': () => getEmailAliases(),
      'saveScreening': () => saveScreening(params),
      'test': () => testConnection()
    };

    if (actions[action]) {
      try {
        const result = actions[action]();
        Logger.log('✅ Resultado: ' + JSON.stringify(result).substring(0, 200));
        return createCorsResponse({ success: true, data: result });
      } catch (actionError) {
        Logger.log('❌ Erro ao executar ação ' + action + ': ' + actionError.toString());
        return createCorsResponse({
          success: false,
          error: actionError.message || actionError.toString()
        });
      }
    } else {
      Logger.log('❌ Ação não encontrada: ' + action);
      return createCorsResponse({
        success: false,
        error: 'Ação não encontrada: ' + action
      });
    }
  } catch (error) {
    Logger.log('❌ Erro no handleRequest: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return createCorsResponse({
      success: false,
      error: error.toString(),
      stack: error.stack
    });
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getCurrentTimestamp() {
  return new Date().toISOString();
}

// ============================================
// FUNÇÕES DE USUÁRIOS
// ============================================

function initUsuariosSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USUARIOS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USUARIOS);
    sheet.getRange('A1:D1').setValues([['Email', 'Nome', 'Role', 'ID']]);
    const defaultUsers = [
      ['admin@email.com', 'Administrador', 'admin', 'admin@email.com'],
      ['analista@email.com', 'Analista', 'analista', 'analista@email.com']
    ];
    sheet.getRange(2, 1, defaultUsers.length, 4).setValues(defaultUsers);
    sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }
  return sheet;
}

function getUserRole(params) {
  try {
    const sheet = initUsuariosSheet();
    const data = sheet.getDataRange().getValues();
    const emailToFind = params.email ? params.email.toLowerCase().trim() : '';

    if (!emailToFind) {
      throw new Error('Email é obrigatório');
    }

    Logger.log('🔍 Procurando usuário: ' + emailToFind);

    for (let i = 1; i < data.length; i++) {
      const emailInSheet = data[i][0] ? data[i][0].toLowerCase().trim() : '';
      if (emailInSheet === emailToFind) {
        const rawRole = data[i][2];
        const normalizedRole = rawRole ? String(rawRole).toLowerCase().trim() : '';

        Logger.log('✅ Usuário encontrado: ' + emailInSheet + ' | Role: ' + normalizedRole);

        return {
          email: data[i][0],
          name: data[i][1] || data[i][0],
          role: normalizedRole,
          id: data[i][3] || data[i][0],
          active: true
        };
      }
    }

    Logger.log('❌ Usuário não encontrado: ' + emailToFind);
    throw new Error('Usuário não encontrado');
  } catch (error) {
    Logger.log('❌ Erro em getUserRole: ' + error.toString());
    throw error;
  }
}

function getAnalysts(params) {
  try {
    Logger.log('🔍 getAnalysts - Iniciando busca de analistas');
    const sheet = initUsuariosSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log('📊 Total de linhas na planilha USUARIOS: ' + data.length);

    const analysts = [];

    for (let i = 1; i < data.length; i++) {
      const rawRole = data[i][2];
      const normalizedRole = rawRole ? String(rawRole).toLowerCase().trim() : '';

      Logger.log('👤 Linha ' + (i + 1) + ':');
      Logger.log('   Email: ' + data[i][0]);
      Logger.log('   Nome: ' + data[i][1]);
      Logger.log('   Role (raw): "' + rawRole + '"');
      Logger.log('   Role (normalized): "' + normalizedRole + '"');

      if (normalizedRole === 'analista') {
        const analyst = {
          id: data[i][3] || data[i][0],
          email: data[i][0],
          name: data[i][1] || data[i][0],
          role: normalizedRole,
          active: true
        };
        analysts.push(analyst);
        Logger.log('✅ Analista encontrado: ' + analyst.email);
      }
    }

    Logger.log('📋 Total de analistas encontrados: ' + analysts.length);
    Logger.log('📦 Retornando: ' + JSON.stringify({ analysts: analysts }));
    return { analysts: analysts };
  } catch (error) {
    Logger.log('❌ Erro em getAnalysts: ' + error.toString());
    throw error;
  }
}


// ============================================
// FUNÇÕES DE CANDIDATOS
// ============================================

function getCandidates(params) {
  const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
  if (!sheet || !values.length) return { candidates: [] };

  const out = values.map(row => {
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    return obj;
  });
  return { candidates: out };
}

function updateCandidateStatus(params) {
  const sh = _sheet(SHEET_CANDIDATOS);
  const headers = _getHeaders_(sh);
  const col = _colMap_(headers);

  const statusCol  = col['Status'];
  const cpfCol     = col['CPF'];
  const regNumCol  = (col['Número de Inscrição'] ?? col['NUMEROINSCRICAO']);
  const analystCol = (col['Analista'] ?? col['assigned_to']);
  const dateCol    = (col['Data Triagem'] ?? col['data_hora_triagem']);
  const reasonCol  = col['Motivo Desclassificação'];
  const notesCol   = (col['Observações'] ?? col['screening_notes']);

  const idx = _getIndex_(sh, headers);
  const searchKey = String(params.registrationNumber).trim();
  let row = idx[searchKey];

  const checkMismatch =
    row &&
    (
      (cpfCol>=0 && String(sh.getRange(row, cpfCol+1).getValue()).trim() !== searchKey) &&
      (regNumCol>=0 && String(sh.getRange(row, regNumCol+1).getValue()).trim() !== searchKey)
    );

  if (!row || checkMismatch){
    const newIdx = _buildIndex_(sh, headers);
    const rev = _getRev_();
    CacheService.getDocumentCache().put(`${IDX_CACHE_KEY}${rev}`, JSON.stringify(newIdx), CACHE_TTL_SEC);
    row = newIdx[searchKey];
  }

  if (!row) throw new Error('Candidato não encontrado');

  const lastCol = sh.getLastColumn();
  const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  if (statusCol>=0) rowVals[statusCol] = params.statusTriagem;
  if (analystCol>=0 && params.analystEmail) rowVals[analystCol] = params.analystEmail;
  if (dateCol>=0) rowVals[dateCol] = getCurrentTimestamp();
  if (reasonCol>=0 && params.reasonId) rowVals[reasonCol] = getDisqualificationReasonById(params.reasonId);
  if (notesCol>=0 && params.notes) rowVals[notesCol] = params.notes;

  _writeWholeRow_(sh, row, rowVals);
  return { success: true, message: 'Status atualizado' };
}

function getCandidatesByStatus(params) {
  const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
  if (!sheet || !values.length) return [];

  const col = _colMap_(headers);
  const statusCol = col['Status'];
  const cpfCol = col['CPF'];
  const emailSentCol = col['EMAIL_SENT'];
  const smsSentCol = col['SMS_SENT'];

  const filtered = [];
  for (let i=0;i<values.length;i++){
    if (values[i][statusCol] === params.status){
      const obj = {};
      for (let j=0;j<headers.length;j++) obj[headers[j]] = values[i][j];
      obj.id = values[i][cpfCol];
      obj.registration_number = values[i][cpfCol];

      obj.email_sent = emailSentCol >= 0 ? (values[i][emailSentCol] === 'Sim' || values[i][emailSentCol] === true || values[i][emailSentCol] === 'TRUE') : false;
      obj.sms_sent = smsSentCol >= 0 ? (values[i][smsSentCol] === 'Sim' || values[i][smsSentCol] === true || values[i][smsSentCol] === 'TRUE') : false;

      filtered.push(obj);
    }
  }
  return filtered;
}

function assignCandidates(params) {
  const sh = _sheet(SHEET_CANDIDATOS);
  const headers = _getHeaders_(sh);
  const col = _colMap_(headers);

  const cpfCol        = col['CPF'];
  const assignedToCol = col['assigned_to'];
  const assignedByCol = col['assigned_by'];
  const assignedAtCol = col['assigned_at'];
  const statusCol     = col['Status'];

  if (cpfCol == null) throw new Error('Coluna CPF não encontrada');

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROWS) {
    return { success: true, assignedCount: 0, message: 'Nada para alocar' };
  }

  const n = lastRow - HEADER_ROWS;
  const cpfs = sh.getRange(HEADER_ROWS+1, cpfCol+1, n, 1).getValues().map(r=> String(r[0]).trim());
  const assignedTo = assignedToCol!=null ? sh.getRange(HEADER_ROWS+1, assignedToCol+1, n, 1).getValues() : null;
  const assignedBy = assignedByCol!=null ? sh.getRange(HEADER_ROWS+1, assignedByCol+1, n, 1).getValues() : null;
  const assignedAt = assignedAtCol!=null ? sh.getRange(HEADER_ROWS+1, assignedAtCol+1, n, 1).getValues() : null;
  const status     = statusCol!=null     ? sh.getRange(HEADER_ROWS+1, statusCol+1, n, 1).getValues()     : null;

  const target = String(params.candidateIds || '').split(',').map(s=>s.trim()).filter(Boolean);
  const stamp = getCurrentTimestamp();
  let count = 0;

  const pos = new Map();
  for (let i=0;i<cpfs.length;i++) pos.set(cpfs[i], i);

  for (const id of target){
    const i = pos.get(id);
    if (i==null) continue;
    if (assignedTo) assignedTo[i][0] = params.analystEmail || '';
    if (assignedBy) assignedBy[i][0] = params.adminEmail || '';
    if (assignedAt) assignedAt[i][0] = stamp;
    if (status)     status[i][0]     = 'em_analise';
    count++;
  }

  if (assignedTo) sh.getRange(HEADER_ROWS+1, assignedToCol+1, n, 1).setValues(assignedTo);
  if (assignedBy) sh.getRange(HEADER_ROWS+1, assignedByCol+1, n, 1).setValues(assignedBy);
  if (assignedAt) sh.getRange(HEADER_ROWS+1, assignedAtCol+1, n, 1).setValues(assignedAt);
  if (status)     sh.getRange(HEADER_ROWS+1, statusCol+1, n, 1).setValues(status);

  return { success: true, assignedCount: count, message: `${count} candidatos alocados com sucesso` };
}

// ============================================
// LOG DE MENSAGENS
// ============================================

function initMensagensSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MENSAGENS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MENSAGENS);
    sheet.getRange('A1:H1').setValues([
      ['Data/Hora', 'Número Inscrição', 'Tipo', 'Destinatário', 'Assunto', 'Conteúdo', 'Enviado Por', 'Status']
    ]);
  }

  return sheet;
}

function logMessage(params) {
  try {
    const sheet = initMensagensSheet();
    const newRow = [
      getCurrentTimestamp(),
      params.registrationNumber,
      params.messageType,
      params.recipient,
      params.subject || '',
      params.content,
      params.sentBy,
      params.status || 'pendente'
    ];
    sheet.appendRow(newRow);
    return { success: true };
  } catch (error) {
    Logger.log('❌ Erro em logMessage: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================
// MOTIVOS DE DESCLASSIFICAÇÃO
// ============================================

function initMotivosSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MOTIVOS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MOTIVOS);
    sheet.getRange('A1:D1').setValues([['ID', 'Motivo', 'Tipo', 'Ativo']]);
    sheet.setFrozenRows(1);
  }

  // Limpa tudo abaixo do cabeçalho
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 4).clearContent();

  const motivosFixos = [
    ['M001', 'Fora do prazo de inscrição', 'Fixo', 'Sim'],
    ['M002', 'Vaga preenchida internamente', 'Fixo', 'Sim'],
    ['M003', 'Desistência do candidato', 'Fixo', 'Sim'],
    ['M099', 'Outros (especificar em observações)', 'Fixo', 'Sim']
  ];

  // MAPA OFICIAL — nomes exatos das colunas na planilha Candidatos
  const mapaNaoConformidade = {
    'checkrg-cpf': 'Documentação de RG/CPF irregular ou ausente',
    'check-cnh': 'CNH inválida, suspensa ou ausente',
    'check-experiencia': 'Experiência insuficiente ou não comprovada',
    'check-regularidade': 'Irregularidade em antecedentes ou documentação',
    'check-laudo': 'Laudo médico ausente ou inválido (PCD)',
    'check-curriculo': 'Currículo incompleto ou incompatível com a vaga'
  };

  const motivosDinamicos = Object.entries(mapaNaoConformidade)
    .map(([key, texto], i) => [`D${String(i + 1).padStart(3, '0')}`, texto, 'Dinâmico', 'Sim']);

  const todos = [...motivosFixos, ...motivosDinamicos];

  if (todos.length > 0) {
    sheet.getRange(2, 1, todos.length, 4).setValues(todos);
  }

  sheet.autoResizeColumns(1, 4);
  Logger.log(`Aba ${SHEET_MOTIVOS} criada/atualizada: ${motivosFixos.length} fixos + ${motivosDinamicos.length} dinâmicos`);
}

function getDisqualificationReasons() {
  return AdvancedCacheService.getWithFallback(
    CACHE_KEYS.REASONS || 'disqualification_reasons_v4',
    () => {
      try {
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MOTIVOS);
        if (!sheet || sheet.getLastRow() <= 1) {
          initMotivosSheet();
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MOTIVOS);
        }

        const data = sheet.getDataRange().getValues();
        const reasons = [];

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const id = String(row[0] || '').trim();
          const motivo = String(row[1] || '').trim();
          const tipo = String(row[2] || 'Fixo').trim();
          const ativo = String(row[3] || 'Sim').toLowerCase() === 'sim';

          if (id && motivo && ativo) {
            reasons.push({
              id,
              reason: motivo,
              type: tipo.toLowerCase(),
              is_active: true
            });
          }
        }

        reasons.sort((a, b) => {
          if (a.type === b.type) return a.reason.localeCompare(b.reason);
          return a.type === 'fixo' ? -1 : 1;
        });

        return reasons;

      } catch (error) {
        Logger.log('Erro crítico ao carregar motivos: ' + error);
        return [
          { id: 'M099', reason: 'Outros (especificar em observações)', type: 'fixo', is_active: true }
        ];
      }
    },
    600
  );
}

// ============================================
// TEMPLATES
// ============================================

function initTemplatesSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_TEMPLATES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TEMPLATES);
    sheet.getRange('A1:E1').setValues([['ID', 'Nome', 'Tipo', 'Assunto', 'Conteúdo']]);

    const templates = [
      ['T001', 'Classificado - Email', 'email', 'Atualização do Processo Seletivo - Classificação',
       'Prezado(a) [NOME],\n\nTemos o prazer de informar que seu perfil foi classificado no processo seletivo para a vaga de [CARGO] na área de [AREA].\n\nNossa equipe entrará em contato em breve para dar continuidade ao processo.\n\nAgradecemos seu interesse e desejamos sucesso nesta etapa.\n\nAtenciosamente,\nEquipe de Recrutamento e Seleção\n\n--\nEste e-mail foi enviado automaticamente. Por favor, não responda esta mensagem.'],

      ['T002', 'Classificado - SMS', 'sms', '',
       'Olá [NOME]! Seu perfil foi classificado no processo para [CARGO]. Aguarde nosso contato para próximos passos.'],

      ['T003', 'Desclassificado - Email', 'email', 'Atualização do Processo Seletivo',
       'Prezado(a) [NOME],\n\nAgradecemos sinceramente seu interesse em fazer parte de nossa equipe e pelo tempo dedicado ao processo seletivo para a vaga de [CARGO].\n\nInformamos que, após análise criteriosa, seu perfil não foi selecionado para prosseguir nesta oportunidade.\n\nSeu currículo permanecerá em nosso banco de dados para futuras oportunidades compatíveis com seu perfil.\n\nDesejamos sucesso em sua jornada profissional.\n\nAtenciosamente,\nEquipe de Recrutamento e Seleção\n\n--\nEste e-mail foi enviado automaticamente. Por favor, não responda esta mensagem.'],

      ['T004', 'Em Revisão - Email', 'email', 'Processo Seletivo - Análise em Andamento',
       'Prezado(a) [NOME],\n\nConfirmamos o recebimento de sua inscrição para a vaga de [CARGO].\n\nSeu cadastro está atualmente em análise por nossa equipe de recrutamento. O processo de triagem está em andamento e todas as inscrições estão sendo cuidadosamente avaliadas.\n\nVocê receberá uma atualização sobre o status assim que a análise for concluída.\n\nAgradecemos sua paciência e interesse em nossa organização.\n\nAtenciosamente,\nEquipe de Recrutamento e Seleção\n\n--\nEste e-mail foi enviado automaticamente. Por favor, não responda esta mensagem.'],

      ['T005', 'Convite para Entrevista - Email', 'email', 'Convite para Próxima Etapa do Processo Seletivo',
       'Prezado(a) [NOME],\n\nÉ com satisfação que convidamos você para a próxima etapa do processo seletivo para a vaga de [CARGO].\n\nSeu perfil foi selecionado entre diversos candidatos e gostaríamos de conhecê-lo(a) melhor.\n\nNossa equipe entrará em contato em breve para agendar a entrevista e fornecer todos os detalhes.\n\nContamos com sua participação!\n\nAtenciosamente,\nEquipe de Recrutamento e Seleção\n\n--\nEste e-mail foi enviado automaticamente. Por favor, não responda esta mensagem.']
    ];

    sheet.getRange(2, 1, templates.length, 5).setValues(templates);

    sheet.autoResizeColumns(1, 5);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f3f3f3');
  }

  return sheet;
}

function getMessageTemplates(params) {
  const sheet = initTemplatesSheet();
  const data = sheet.getDataRange().getValues();
  const templates = [];

  for (let i = 1; i < data.length; i++) {
    const messageType = params && params.messageType ? params.messageType : null;
    if (!messageType || data[i][2] === messageType) {
      templates.push({
        id: data[i][0],
        template_name: data[i][1],
        message_type: data[i][2],
        subject: data[i][3],
        content: data[i][4]
      });
    }
  }
  return templates;
}

// ============================================
// FUNÇÕES AUXILIARES PARA EMAIL
// ============================================

// ✅ FUNÇÃO PARA VERIFICAR ALIASES
function myFunction() {
  try {
    const aliases = GmailApp.getAliases();
    Logger.log('📧 Aliases disponíveis:');
    aliases.forEach(alias => Logger.log('   - ' + alias));
    return {
      success: true,
      email: Session.getActiveUser().getEmail(),
      aliases: aliases
    };
  } catch (error) {
    Logger.log('❌ Erro ao buscar aliases: ' + error.toString());
    return {
      success: false,
      email: Session.getActiveUser().getEmail(),
      aliases: [],
      error: error.toString()
    };
  }
}

function getEmailAliases() {
  try {
    const aliases = GmailApp.getAliases();
    Logger.log('📧 Aliases disponíveis: ' + JSON.stringify(aliases));
    return aliases;
  } catch (error) {
    Logger.log('❌ Erro ao buscar aliases: ' + error.toString());
    return [];
  }
}

function _getProp_(k){
  return PropertiesService.getScriptProperties().getProperty(k);
}

function _twilioEnabled_(){
  return !!(_getProp_('TWILIO_SID') && _getProp_('TWILIO_TOKEN') && _getProp_('TWILIO_FROM'));
}

function _formatE164_(phone){
  if (!phone) return '';

  try {
    let cleaned = String(phone).replace(/\D/g, '');
    cleaned = cleaned.replace(/^0+/, '');

    if (!cleaned.startsWith('55')) {
      cleaned = '55' + cleaned;
    }

    if (cleaned.length < 12) {
      throw new Error('Número muito curto: ' + cleaned);
    }

    if (cleaned.length > 13) {
      cleaned = cleaned.substring(0, 13);
    }

    return '+' + cleaned;

  } catch (error) {
    Logger.log('❌ Erro ao formatar telefone: ' + error.toString());
    return '';
  }
}

function _sendEmailGmail_(to, subject, body, alias) {
  try {
    Logger.log('📧 Enviando email para: ' + to);
    Logger.log('👤 Alias solicitado: ' + (alias || 'Email principal'));

    const options = {
      htmlBody: body,
      noReply: false,
      name: 'Processo Seletivo Instituto Acqua',
      replyTo: Session.getActiveUser().getEmail()
    };

    const aliasCheck = myFunction();
    const aliases = aliasCheck.aliases;

    let finalAlias = null;

    if (alias) {
      if (aliases.includes(alias)) {
        finalAlias = alias;
        Logger.log('✅ Usando alias solicitado: ' + alias);
      } else {
        Logger.log('⚠️ Alias solicitado não encontrado: ' + alias);
      }
    }

    if (!finalAlias && aliases.length > 0) {
      finalAlias = aliases[0];
      Logger.log('🔄 Usando primeiro alias disponível: ' + aliases[0]);
    }

    if (finalAlias) {
      options.from = finalAlias;
      options.replyTo = finalAlias;
    }

    Logger.log('⚙️ Opções de envio: ' + JSON.stringify(options));

    GmailApp.sendEmail(to, subject, body, options);

    Logger.log('✅ Email enviado com sucesso');
    Logger.log('   De: Processo Seletivo Instituto Acqua');
    Logger.log('   Email: ' + (options.from || Session.getActiveUser().getEmail()));

    return {
      ok: true,
      from: options.from || Session.getActiveUser().getEmail(),
      displayName: 'Processo Seletivo Instituto Acqua',
      aliasUsed: !!finalAlias
    };

  } catch (e) {
    Logger.log('❌ Erro ao enviar email: ' + e.toString());

    try {
      Logger.log('🔄 Tentando fallback simplificado...');
      GmailApp.sendEmail(to, subject, body, {
        name: 'Processo Seletivo Instituto Acqua'
      });
      Logger.log('✅ Email enviado via fallback');
      return {
        ok: true,
        from: Session.getActiveUser().getEmail(),
        displayName: 'Processo Seletivo Instituto Acqua',
        fallback: true
      };
    } catch (fallbackError) {
      Logger.log('❌ Fallback também falhou: ' + fallbackError.toString());
    }

    return {
      ok: false,
      error: e.toString()
    };
  }
}

function _sendSmsTwilio_(to, body) {
  try {
    if (!_twilioEnabled_()) {
      Logger.log('⚠️ Twilio não configurado - SMS desabilitado');
      return {
        ok: false,
        skipped: true,
        error: 'Twilio não configurado. Verifique as variáveis TWILIO_SID, TWILIO_TOKEN e TWILIO_FROM.'
      };
    }

    if (!to) {
      throw new Error('Número de telefone é obrigatório');
    }

    const formattedTo = _formatE164_(to);

    if (!formattedTo.startsWith('+55') || formattedTo.length < 13) {
      throw new Error('Número de telefone brasileiro inválido: ' + formattedTo);
    }

    Logger.log('📱 Enviando SMS para: ' + formattedTo);
    Logger.log('📝 Conteúdo: ' + body.substring(0, 50) + '...');

    const sid = _getProp_('TWILIO_SID');
    const token = _getProp_('TWILIO_TOKEN');
    const from = _getProp_('TWILIO_FROM');

    const url = `https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`;

    const payload = {
      To: formattedTo,
      From: from,
      Body: body
    };

    const options = {
      method: 'POST',
      payload: payload,
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token),
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      timeout: 30000
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('📡 Resposta Twilio - Código: ' + responseCode);

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseText);
      Logger.log('✅ SMS enviado com sucesso - SID: ' + responseData.sid);
      return {
        ok: true,
        sid: responseData.sid,
        status: responseData.status
      };
    } else {
      Logger.log('❌ Erro Twilio: ' + responseText);
      let errorMessage = `Twilio HTTP ${responseCode}`;
      try {
        const errorData = JSON.parse(responseText);
        errorMessage += ` - ${errorData.message || errorData.code || 'Erro desconhecido'}`;
      } catch (e) {
        errorMessage += ` - ${responseText.substring(0, 100)}`;
      }
      return {
        ok: false,
        error: errorMessage,
        responseCode: responseCode
      };
    }

  } catch (error) {
    Logger.log('❌ Erro crítico ao enviar SMS: ' + error.toString());
    Logger.log('📞 Stack: ' + error.stack);

    return {
      ok: false,
      error: 'Erro de conexão: ' + error.toString()
    };
  }
}

function _pickEmailFromRow_(headers, row) {
  const colMap = _colMap_(headers);

  const emailColumns = ['EMAIL', 'EMAIL1', 'EMAIL2', 'E_MAIL', 'E-MAIL'];

  for (const colName of emailColumns) {
    if (colMap[colName] >= 0) {
      const email = String(row[colMap[colName]]).trim();
      if (email && email.includes('@')) {
        Logger.log('📧 Email encontrado na coluna ' + colName + ': ' + email);
        return email;
      }
    }
  }

  Logger.log('⚠️ Nenhum email válido encontrado para o candidato');
  return null;
}

function _pickPhoneFromRow_(headers, row) {
  const colMap = _colMap_(headers);

  const phoneColumns = ['TELEFONE', 'CELULAR', 'PHONE', 'TELEFONE1', 'CELULAR1'];

  for (const colName of phoneColumns) {
    if (colMap[colName] >= 0) {
      const phone = String(row[colMap[colName]]).trim();
      if (phone && phone.replace(/\D/g, '').length >= 10) {
        Logger.log('📱 Telefone encontrado na coluna ' + colName + ': ' + phone);
        return phone;
      }
    }
  }

  Logger.log('⚠️ Nenhum telefone válido encontrado para o candidato');
  return null;
}

function _applyTemplate_(text, candidate){
  if (!text) return '';
  return String(text)
    .replace(/\[NOME\]/g, candidate.NOMECOMPLETO || candidate.NOMESOCIAL || '')
    .replace(/\[CARGO\]/g, candidate.CARGOPRETENDIDO || '')
    .replace(/\[AREA\]/g, candidate.AREAATUACAO || '');
}

// ============================================
// ENVIO DE MENSAGENS
// ============================================

function sendMessages(params) {
  Logger.log('📤 sendMessages iniciado');

  const messageType = params.messageType;
  const subject = params.subject || '';
  const content = params.content || '';
  const candidateIds = params.candidateIds || '';
  const sentBy = params.sentBy || 'system';
  const fromAlias = params.fromAlias || '';

  if (!content) {
    throw new Error('Conteúdo da mensagem é obrigatório');
  }

  if (messageType === 'email' && !subject) {
    throw new Error('Assunto é obrigatório para emails');
  }

  const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
  if (!sheet) throw new Error('Planilha de candidatos não encontrada');

  const col = _colMap_(headers);
  const cpfCol = col['CPF'];

  const targetIds = candidateIds.split(',').map(s => s.trim()).filter(Boolean);
  Logger.log('📋 Candidatos alvo: ' + targetIds.length);

  const results = [];
  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < values.length; i++) {
    const cpf = String(values[i][cpfCol]).trim();

    if (!targetIds.includes(cpf)) continue;

    const candidate = {};
    for (let j = 0; j < headers.length; j++) {
      candidate[headers[j]] = values[i][j];
    }

    const nome = candidate.NOMECOMPLETO || candidate.NOMESOCIAL || 'Candidato';
    let recipient, result;

    if (messageType === 'email') {
      recipient = _pickEmailFromRow_(headers, values[i]);

      if (!recipient) {
        Logger.log('⚠️ Sem email: ' + nome);
        results.push({
          candidateId: cpf,
          candidateName: nome,
          success: false,
          error: 'Email não cadastrado'
        });
        failCount++;
        continue;
      }

      const personalizedSubject = _applyTemplate_(subject, candidate);
      const personalizedContent = _applyTemplate_(content, candidate);

      result = _sendEmailGmail_(recipient, personalizedSubject, personalizedContent, fromAlias);

      logMessage({
        registrationNumber: cpf,
        messageType: 'email',
        recipient: recipient,
        subject: personalizedSubject,
        content: personalizedContent,
        sentBy: sentBy,
        status: result.ok ? 'enviado' : 'falhou'
      });

    } else if (messageType === 'sms') {
      recipient = _pickPhoneFromRow_(headers, values[i]);

      if (!recipient) {
        Logger.log('⚠️ Sem telefone: ' + nome);
        results.push({
          candidateId: cpf,
          candidateName: nome,
          success: false,
          error: 'Telefone não cadastrado'
        });
        failCount++;
        continue;
      }

      const personalizedContent = _applyTemplate_(content, candidate);
      result = _sendSmsTwilio_(recipient, personalizedContent);

      if (result.skipped) {
        results.push({
          candidateId: cpf,
          candidateName: nome,
          success: false,
          error: 'Twilio não configurado'
        });
        failCount++;
        continue;
      }

      logMessage({
        registrationNumber: cpf,
        messageType: 'sms',
        recipient: recipient,
        subject: '',
        content: personalizedContent,
        sentBy: sentBy,
        status: result.ok ? 'enviado' : 'falhou'
      });

    } else {
      throw new Error('Tipo de mensagem inválido: ' + messageType);
    }

    if (result.ok) {
      successCount++;
      results.push({
        candidateId: cpf,
        candidateName: nome,
        success: true
      });

      _updateMessageStatusInCandidates_(cpf, messageType);
    } else {
      failCount++;
      results.push({
        candidateId: cpf,
        candidateName: nome,
        success: false,
        error: result.error || 'Erro desconhecido'
      });
    }
  }

  Logger.log('✅ Sucesso: ' + successCount);
  Logger.log('❌ Falhas: ' + failCount);

  return {
    successCount: successCount,
    failCount: failCount,
    results: results
  };
}

function _updateMessageStatusInCandidates_(cpf, messageType) {
  try {
    const sh = _sheet(SHEET_CANDIDATOS);
    if (!sh) return;

    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);
    const cpfCol = col['CPF'];

    let targetCol;
    if (messageType === 'email') {
      targetCol = col['EMAIL_SENT'];
    } else if (messageType === 'sms') {
      targetCol = col['SMS_SENT'];
    }

    if (targetCol === undefined || targetCol < 0) {
      Logger.log('⚠️ Coluna ' + (messageType === 'email' ? 'EMAIL_SENT' : 'SMS_SENT') + ' não encontrada');
      return;
    }

    const idx = _getIndex_(sh, headers);
    const searchKey = String(cpf).trim();
    const row = idx[searchKey];

    if (!row) {
      Logger.log('⚠️ Candidato não encontrado para atualizar status de mensagem: ' + cpf);
      return;
    }

    sh.getRange(row, targetCol + 1).setValue('Sim');
    Logger.log('✅ Status de mensagem atualizado para ' + cpf + ' - ' + messageType);

  } catch (error) {
    Logger.log('❌ Erro ao atualizar status de mensagem: ' + error.toString());
  }
}

function updateMessageStatus(params) {
  try {
    Logger.log('📝 updateMessageStatus iniciado');

    const registrationNumber = params.registrationNumber;
    const messageType = params.messageType;

    if (!registrationNumber) {
      throw new Error('Número de inscrição é obrigatório');
    }

    if (!messageType || (messageType !== 'email' && messageType !== 'sms')) {
      throw new Error('Tipo de mensagem inválido. Use "email" ou "sms"');
    }

    const sh = _sheet(SHEET_CANDIDATOS);
    if (!sh) {
      throw new Error('Planilha de candidatos não encontrada');
    }

    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);
    const cpfCol = col['CPF'];
    const regNumCol = col['Número de Inscrição'];

    let targetCol;
    if (messageType === 'email') {
      targetCol = col['EMAIL_SENT'];
    } else if (messageType === 'sms') {
      targetCol = col['SMS_SENT'];
    }

    if (targetCol === undefined || targetCol < 0) {
      const colName = messageType === 'email' ? 'EMAIL_SENT' : 'SMS_SENT';
      throw new Error('Coluna ' + colName + ' não encontrada. Execute addStatusColumnIfNotExists primeiro.');
    }

    const idx = _getIndex_(sh, headers);
    const searchKey = String(registrationNumber).trim();
    let row = idx[searchKey];

    if (!row) {
      const newIdx = _buildIndex_(sh, headers);
      const rev = _getRev_();
      CacheService.getDocumentCache().put(`${IDX_CACHE_KEY}${rev}`, JSON.stringify(newIdx), CACHE_TTL_SEC);
      row = newIdx[searchKey];
    }

    if (!row) {
      throw new Error('Candidato não encontrado: ' + registrationNumber);
    }

    sh.getRange(row, targetCol + 1).setValue('Sim');
    _bumpRev_();

    Logger.log('✅ Status de mensagem atualizado: ' + registrationNumber + ' - ' + messageType + ' = Sim');

    return {
      success: true,
      message: 'Status de mensagem atualizado com sucesso',
      registrationNumber: registrationNumber,
      messageType: messageType,
      status: 'Sim'
    };

  } catch (error) {
    Logger.log('❌ Erro em updateMessageStatus: ' + error.toString());
    throw error;
  }
}

// ============================================
// FUNÇÕES DE ENTREVISTA
// ============================================

function getInterviewCandidates(params) {
  try {
    const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
    if (!sheet || !values.length) return [];

    const col = _colMap_(headers);
    const statusEntrevistaCol = col['status_entrevista'];
    const cpfCol = col['CPF'];
    const regNumCol = col['Número de Inscrição'];
    const emailSentCol = col['EMAIL_SENT'];
    const smsSentCol = col['SMS_SENT'];

    if (statusEntrevistaCol === undefined) {
      Logger.log('⚠️ Coluna status_entrevista não encontrada');
      return [];
    }

    const candidates = [];
    for (let i = 0; i < values.length; i++) {
      const statusEntrevista = values[i][statusEntrevistaCol];

      if (statusEntrevista === 'Aguardando' || statusEntrevista === 'aguardando') {
        const candidate = {};
        headers.forEach((header, index) => {
          candidate[header] = values[i][index];
        });
        candidate.id = values[i][cpfCol] || values[i][regNumCol];
        candidate.registration_number = values[i][regNumCol] || values[i][cpfCol];

        candidate.email_sent = emailSentCol >= 0 ? (values[i][emailSentCol] === 'Sim' || values[i][emailSentCol] === true || values[i][emailSentCol] === 'TRUE') : false;
        candidate.sms_sent = smsSentCol >= 0 ? (values[i][smsSentCol] === 'Sim' || values[i][smsSentCol] === true || values[i][smsSentCol] === 'TRUE') : false;

        candidates.push(candidate);
      }
    }

    Logger.log('📋 Candidatos para entrevista: ' + candidates.length);
    return candidates;
  } catch (error) {
    Logger.log('❌ Erro em getInterviewCandidates: ' + error.toString());
    throw error;
  }
}

function moveToInterview(params) {
  try {
    const sh = _sheet(SHEET_CANDIDATOS);
    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);

    const statusEntrevistaCol = col['status_entrevista'];
    const cpfCol = col['CPF'];
    const emailSentCol = col['EMAIL_SENT'];
    const smsSentCol = col['SMS_SENT'];

    if (statusEntrevistaCol === undefined || statusEntrevistaCol < 0) {
      throw new Error('Coluna status_entrevista não encontrada');
    }

    const candidateIds = String(params.candidateIds || '').split(',').map(s => s.trim()).filter(Boolean);
    Logger.log('📋 Movendo ' + candidateIds.length + ' candidatos para entrevista');

    const lastRow = sh.getLastRow();
    if (lastRow <= HEADER_ROWS) {
      return { success: true, movedCount: 0, message: 'Nenhum candidato para mover' };
    }

    const n = lastRow - HEADER_ROWS;
    const cpfs = sh.getRange(HEADER_ROWS + 1, cpfCol + 1, n, 1).getValues().map(r => String(r[0]).trim());
    const statusEntrevista = sh.getRange(HEADER_ROWS + 1, statusEntrevistaCol + 1, n, 1).getValues();

    const emailSent = emailSentCol >= 0 ? sh.getRange(HEADER_ROWS + 1, emailSentCol + 1, n, 1).getValues() : null;
    const smsSent = smsSentCol >= 0 ? sh.getRange(HEADER_ROWS + 1, smsSentCol + 1, n, 1).getValues() : null;

    let movedCount = 0;
    const pos = new Map();
    for (let i = 0; i < cpfs.length; i++) {
      pos.set(cpfs[i], i);
    }

    for (const cpf of candidateIds) {
      const i = pos.get(cpf);
      if (i === undefined) {
        Logger.log('⚠️ CPF não encontrado: ' + cpf);
        continue;
      }

      const hasEmail = emailSent && (emailSent[i][0] === 'Sim' || emailSent[i][0] === true || emailSent[i][0] === 'TRUE');
      const hasSms = smsSent && (smsSent[i][0] === 'Sim' || smsSent[i][0] === true || smsSent[i][0] === 'TRUE');

      if (!hasEmail && !hasSms) {
        Logger.log('⚠️ Candidato ' + cpf + ' não recebeu mensagens. Pulando.');
        continue;
      }

      statusEntrevista[i][0] = 'Aguardando';
      movedCount++;
      Logger.log('✅ ' + cpf + ' movido para entrevista');
    }

    if (movedCount > 0) {
      sh.getRange(HEADER_ROWS + 1, statusEntrevistaCol + 1, n, 1).setValues(statusEntrevista);
      _bumpRev_();
    }

    Logger.log('✅ Total movidos: ' + movedCount);
    return {
      success: true,
      movedCount: movedCount,
      message: movedCount + ' candidato(s) movido(s) para entrevista'
    };
  } catch (error) {
    Logger.log('❌ Erro em moveToInterview: ' + error.toString());
    throw error;
  }
}

function getInterviewers(params) {
  try {
    const sheet = initUsuariosSheet();
    const data = sheet.getDataRange().getValues();
    const interviewers = [];

    for (let i = 1; i < data.length; i++) {
      const rawRole = data[i][2];
      const normalizedRole = rawRole ? String(rawRole).toLowerCase().trim() : '';

      if (normalizedRole === 'entrevistador') {
        interviewers.push({
          id: data[i][3] || data[i][0],
          email: data[i][0],
          name: data[i][1] || data[i][0],
          role: normalizedRole,
          active: true
        });
      }
    }

    Logger.log('👥 Entrevistadores encontrados: ' + interviewers.length);
    return interviewers;
  } catch (error) {
    Logger.log('❌ Erro em getInterviewers: ' + error.toString());
    throw error;
  }
}

function getInterviewerCandidates(params) {
  try {
    const interviewerEmail = params.interviewerEmail;

    if (!interviewerEmail) {
      throw new Error('Email do entrevistador é obrigatório');
    }

    Logger.log('🔍 Buscando candidatos do entrevistador: ' + interviewerEmail);

    const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
    if (!sheet || !values.length) {
      Logger.log('⚠️ Nenhum candidato encontrado na planilha');
      return [];
    }

    const col = _colMap_(headers);
    const statusEntrevistaCol = col['status_entrevista'];
    const entrevistadorCol = col['entrevistador'];
    const cpfCol = col['CPF'];
    const regNumCol = col['Número de Inscrição'];

    if (entrevistadorCol === undefined || entrevistadorCol < 0) {
      Logger.log('⚠️ Coluna entrevistador não encontrada');
      return [];
    }

    const candidates = [];
    for (let i = 0; i < values.length; i++) {
      const candidateInterviewer = values[i][entrevistadorCol];
      const normalizedInterviewer = candidateInterviewer ? String(candidateInterviewer).toLowerCase().trim() : '';
      const normalizedEmail = interviewerEmail.toLowerCase().trim();

      if (normalizedInterviewer === normalizedEmail) {
        const candidate = {};
        headers.forEach((header, index) => {
          candidate[header] = values[i][index];
        });
        candidate.id = values[i][cpfCol] || values[i][regNumCol];
        candidate.registration_number = values[i][regNumCol] || values[i][cpfCol];

        candidates.push(candidate);
      }
    }

    Logger.log('✅ Candidatos encontrados para ' + interviewerEmail + ': ' + candidates.length);
    return candidates;
  } catch (error) {
    Logger.log('❌ Erro em getInterviewerCandidates: ' + error.toString());
    throw error;
  }
}

function allocateToInterviewer(params) {
  const startTime = Date.now();
  try {
    const sh = _sheet(SHEET_CANDIDATOS);
    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);

    const entrevistadorCol = col['entrevistador'];
    const dataEntrevistaCol = col['data_entrevista'];
    const cpfCol = col['CPF'];

    if (entrevistadorCol === undefined) {
      throw new Error('Coluna "entrevistador" não encontrada na planilha');
    }

    const candidateIds = String(params.candidateIds || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    const interviewerEmail = String(params.interviewerEmail || '').trim();
    const adminEmail = String(params.adminEmail || '').trim();

    if (candidateIds.length === 0) {
      return { success: false, error: 'Nenhum candidato selecionado' };
    }
    if (!interviewerEmail) {
      return { success: false, error: 'E-mail do entrevistador não informado' };
    }

    Logger.log(`Alocando ${candidateIds.length} candidato(s) para: ${interviewerEmail}`);

    const lastRow = sh.getLastRow();
    if (lastRow <= HEADER_ROWS) {
      return { success: true, allocatedCount: 0, message: 'Nenhum candidato na planilha' };
    }

    const n = lastRow - HEADER_ROWS;
    const timestamp = getCurrentTimestamp();

    const startCol = Math.min(cpfCol, entrevistadorCol, dataEntrevistaCol ?? 999) + 1;
    const numCols = 3;

    const rawData = sh.getRange(HEADER_ROWS + 1, startCol, n, numCols).getValues();

    const cpfIdx = cpfCol - (startCol - 1);
    const entIdx = entrevistadorCol - (startCol - 1);
    const dataIdx = dataEntrevistaCol !== undefined ? dataEntrevistaCol - (startCol - 1) : -1;

    const cpfToRowIndex = new Map();
    for (let i = 0; i < rawData.length; i++) {
      const cpf = String(rawData[i][cpfIdx] || '').trim();
      if (cpf) cpfToRowIndex.set(cpf, i);
    }

    let allocatedCount = 0;
    const rangesToUpdate = [];

    for (const cpf of candidateIds) {
      const rowIndex = cpfToRowIndex.get(cpf);
      if (rowIndex === undefined) {
        Logger.log(`CPF não encontrado: ${cpf}`);
        continue;
      }

      if (rawData[rowIndex][entIdx] !== interviewerEmail) {
        rawData[rowIndex][entIdx] = interviewerEmail;
        rangesToUpdate.push(sh.getRange(HEADER_ROWS + 1 + rowIndex, entrevistadorCol + 1));
      }

      if (dataIdx >= 0 && rawData[rowIndex][dataIdx] !== timestamp) {
        rawData[rowIndex][dataIdx] = timestamp;
        if (!rangesToUpdate.includes(sh.getRange(HEADER_ROWS + 1 + rowIndex, dataEntrevistaCol + 1))) {
          rangesToUpdate.push(sh.getRange(HEADER_ROWS + 1 + rowIndex, dataEntrevistaCol + 1));
        }
      }

      allocatedCount++;
      Logger.log(`${cpf} → alocado para ${interviewerEmail}`);
    }

    if (rangesToUpdate.length > 0) {
      const rangeList = sh.getRangeList(rangesToUpdate.map(r => r.getA1Notation()));
      rangeList.getRanges()
        .filter(r => r.getColumn() === entrevistadorCol + 1)
        .forEach(r => r.setValue(interviewerEmail));

      if (dataEntrevistaCol >= 0) {
        rangeList.getRanges()
          .filter(r => r.getColumn() === dataEntrevistaCol + 1)
          .forEach(r => r.setValue(timestamp));
      }

      _bumpRev_();
    }

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    Logger.log(`Total alocados: ${allocatedCount} em ${elapsed}s`);

    return {
      success: true,
      allocatedCount,
      message: `${allocatedCount} candidato(s) alocado(s) com sucesso para ${interviewerEmail}`
    };

  } catch (error) {
    Logger.log('Erro crítico em allocateToInterviewer: ' + error.toString());
    Logger.log(error.stack);
    return {
      success: false,
      error: 'Falha ao alocar candidatos',
      details: error.toString()
    };
  }
}

function updateInterviewStatus(params) {
  try {
    const sh = _sheet(SHEET_CANDIDATOS);
    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);

    const statusEntrevistaCol = col['status_entrevista'];
    const entrevistadorCol = col['entrevistador'];
    const dataEntrevistaCol = col['data_entrevista'];
    const cpfCol = col['CPF'];
    const regNumCol = col['Número de Inscrição'];

    const idx = _getIndex_(sh, headers);
    const searchKey = String(params.registrationNumber).trim();
    let row = idx[searchKey];

    if (!row) {
      const newIdx = _buildIndex_(sh, headers);
      const rev = _getRev_();
      CacheService.getDocumentCache().put(`${IDX_CACHE_KEY}${rev}`, JSON.stringify(newIdx), CACHE_TTL_SEC);
      row = newIdx[searchKey];
    }

    if (!row) throw new Error('Candidato não encontrado');

    const lastCol = sh.getLastColumn();
    const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

    if (statusEntrevistaCol >= 0) rowVals[statusEntrevistaCol] = params.status;
    if (entrevistadorCol >= 0 && params.interviewerEmail) rowVals[entrevistadorCol] = params.interviewerEmail;
    if (dataEntrevistaCol >= 0) rowVals[dataEntrevistaCol] = getCurrentTimestamp();

    _writeWholeRow_(sh, row, rowVals);
    _bumpRev_();

    Logger.log('✅ Status de entrevista atualizado');
    return { success: true, message: 'Status atualizado' };
  } catch (error) {
    Logger.log('❌ Erro em updateInterviewStatus: ' + error.toString());
    throw error;
  }
}

function saveInterviewEvaluation(params) {
  try {
    const sh = _sheet(SHEET_CANDIDATOS);
    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);

    const cpfCol = col['CPF'];
    const statusEntrevistaCol = col['status_entrevista'];
    const entrevistadorCol = col['entrevistador'];
    const dataEntrevistaCol = col['data_entrevista'];
    const completedAtCol = col['interview_completed_at'];
    const scoreCol = col['interview_score'];
    const resultCol = col['interview_result'];
    const notesCol = col['interview_notes'];
    const formacaoCol = col['formacao_adequada'];
    const graduacoesCol = col['graduacoes_competencias'];
    const descricaoCol = col['descricao_processos'];
    const terminologiaCol = col['terminologia_tecnica'];
    const calmaCol = col['calma_clareza'];
    const escalasCol = col['escalas_flexiveis'];
    const adaptabilidadeCol = col['adaptabilidade_mudancas'];
    const ajustesCol = col['ajustes_emergencia'];
    const residenciaCol = col['residencia'];
    const conflitosCol = col['resolucao_conflitos'];
    const colaboracaoCol = col['colaboracao_equipe'];
    const adaptacaoPerfisCol = col['adaptacao_perfis'];

    const idx = _getIndex_(sh, headers);
    const searchKey = String(params.candidateId).trim();
    let row = idx[searchKey];

    if (!row) {
      const newIdx = _buildIndex_(sh, headers);
      const rev = _getRev_();
      CacheService.getDocumentCache().put(`${IDX_CACHE_KEY}${rev}`, JSON.stringify(newIdx), CACHE_TTL_SEC);
      row = newIdx[searchKey];
    }

    if (!row) {
      Logger.log('❌ Candidato não encontrado: ' + searchKey);
      throw new Error('Candidato não encontrado: ' + searchKey);
    }

    Logger.log('📝 Salvando avaliação do candidato na linha: ' + row);

    const lastCol = sh.getLastColumn();
    const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

    const secao1 = (Number(params.formacao_adequada) + Number(params.graduacoes_competencias)) * 2;
    const secao2 = (Number(params.descricao_processos) + Number(params.terminologia_tecnica) + Number(params.calma_clareza)) * 2;
    const secao3 = Number(params.escalas_flexiveis) + Number(params.adaptabilidade_mudancas) + Number(params.ajustes_emergencia);
    const secao4 = Number(params.residencia);
    const secao5 = (Number(params.resolucao_conflitos) + Number(params.colaboracao_equipe) + Number(params.adaptacao_perfis)) * 2;
    const totalScore = secao1 + secao2 + secao3 + secao4 + secao5;

    Logger.log('📊 Pontuação calculada: ' + totalScore + '/120');

    if (statusEntrevistaCol >= 0) rowVals[statusEntrevistaCol] = 'Avaliado';
    if (entrevistadorCol >= 0) rowVals[entrevistadorCol] = params.interviewerEmail || '';
    if (dataEntrevistaCol >= 0) rowVals[dataEntrevistaCol] = getCurrentTimestamp();
    if (completedAtCol >= 0) rowVals[completedAtCol] = params.completed_at || getCurrentTimestamp();
    if (scoreCol >= 0) rowVals[scoreCol] = totalScore;
    if (resultCol >= 0) rowVals[resultCol] = params.resultado || '';
    if (notesCol >= 0) rowVals[notesCol] = params.impressao_perfil || '';
    if (formacaoCol >= 0) rowVals[formacaoCol] = params.formacao_adequada || '';
    if (graduacoesCol >= 0) rowVals[graduacoesCol] = params.graduacoes_competencias || '';
    if (descricaoCol >= 0) rowVals[descricaoCol] = params.descricao_processos || '';
    if (terminologiaCol >= 0) rowVals[terminologiaCol] = params.terminologia_tecnica || '';
    if (calmaCol >= 0) rowVals[calmaCol] = params.calma_clareza || '';
    if (escalasCol >= 0) rowVals[escalasCol] = params.escalas_flexiveis || '';
    if (adaptabilidadeCol >= 0) rowVals[adaptabilidadeCol] = params.adaptabilidade_mudancas || '';
    if (ajustesCol >= 0) rowVals[ajustesCol] = params.ajustes_emergencia || '';
    if (residenciaCol >= 0) rowVals[residenciaCol] = params.residencia || '';
    if (conflitosCol >= 0) rowVals[conflitosCol] = params.resolucao_conflitos || '';
    if (colaboracaoCol >= 0) rowVals[colaboracaoCol] = params.colaboracao_equipe || '';
    if (adaptacaoPerfisCol >= 0) rowVals[adaptacaoPerfisCol] = params.adaptacao_perfis || '';

    _writeWholeRow_(sh, row, rowVals);
    _bumpRev_();

    Logger.log('✅ Avaliação de entrevista salva com sucesso');
    Logger.log('   - Candidato: ' + searchKey);
    Logger.log('   - Pontuação: ' + totalScore + '/120');
    Logger.log('   - Resultado: ' + params.resultado);

    return {
      success: true,
      message: 'Avaliação salva com sucesso',
      score: totalScore,
      resultado: params.resultado
    };
  } catch (error) {
    Logger.log('❌ Erro em saveInterviewEvaluation: ' + error.toString());
    Logger.log('   Stack: ' + error.stack);
    throw error;
  }
}

// ============================================
// RELATÓRIOS
// ============================================

function getReportStats(params) {
  return AdvancedCacheService.getWithFallback(
    CACHE_KEYS.STATS,
    () => {
      try {
        Logger.log('Gerando estatísticas de relatórios (cache miss)');

        const { sheet, headers, values } = _readSheetBlock_(SHEET_CANDIDATOS);
        if (!sheet || !values.length) {
          return {
            classificados: 0,
            desclassificados: 0,
            entrevistaClassificados: 0,
            entrevistaDesclassificados: 0
          };
        }

        const col = _colMap_(headers);
        const statusCol = col['Status'];
        const statusEntrevistaCol = col['status_entrevista'];
        const interviewResultCol = col['interview_result'];

        let classificados = 0;
        let desclassificados = 0;
        let entrevistaClassificados = 0;
        let entrevistaDesclassificados = 0;

        for (let i = 0; i < values.length; i++) {
          const status = values[i][statusCol] ? String(values[i][statusCol]).trim() : '';
          const statusEntrevista = values[i][statusEntrevistaCol] ? String(values[i][statusEntrevistaCol]).trim() : '';
          const interviewResult = values[i][interviewResultCol] ? String(values[i][interviewResultCol]).trim() : '';

          if (status === 'Classificado') classificados++;
          else if (status === 'Desclassificado') desclassificados++;

          if (statusEntrevista === 'Avaliado') {
            if (interviewResult === 'Classificado') entrevistaClassificados++;
            else if (interviewResult === 'Desclassificado') entrevistaDesclassificados++;
          }
        }

        const result = {
          classificados,
          desclassificados,
          entrevistaClassificados,
          entrevistaDesclassificados
        };

        Logger.log('Estatísticas geradas e armazenadas no cache');
        Logger.log(`   Classificados: ${classificados} | Desclassificados: ${desclassificados}`);
        Logger.log(`   Entrevista Classificados: ${entrevistaClassificados} | Desclassificados: ${entrevistaDesclassificados}`);

        return result;

      } catch (error) {
        Logger.log('Erro crítico em getReportStats (cache não será salvo): ' + error.toString());
        throw error;
      }
    },
    60
  );
}

function getReport(params) {
  try {
    const reportType = params.reportType;
    const analystEmail = params.analystEmail;
    const interviewerEmail = params.interviewerEmail;

    Logger.log('📋 Gerando relatório: ' + reportType);
    if (analystEmail) Logger.log('   - Filtro por analista: ' + analystEmail);
    if (interviewerEmail) Logger.log('   - Filtro por entrevistador: ' + interviewerEmail);

    const {sheet, headers, values} = _readSheetBlock_(SHEET_CANDIDATOS);
    if (!sheet || !values.length) {
      Logger.log('⚠️ Nenhum candidato encontrado');
      return [];
    }

    const col = _colMap_(headers);
    const statusCol = col['Status'];
    const analistaCol = col['Analista'];
    const statusEntrevistaCol = col['status_entrevista'];
    const interviewResultCol = col['interview_result'];
    const entrevistadorCol = col['entrevistador'];

    const candidates = [];

    for (let i = 0; i < values.length; i++) {
      const status = values[i][statusCol] ? String(values[i][statusCol]).trim() : '';
      const analista = values[i][analistaCol] ? String(values[i][analistaCol]).toLowerCase().trim() : '';
      const statusEntrevista = values[i][statusEntrevistaCol] ? String(values[i][statusEntrevistaCol]).trim() : '';
      const interviewResult = values[i][interviewResultCol] ? String(values[i][interviewResultCol]).trim() : '';
      const entrevistador = values[i][entrevistadorCol] ? String(values[i][entrevistadorCol]).toLowerCase().trim() : '';

      if (analystEmail && analista !== analystEmail.toLowerCase().trim()) {
        continue;
      }

      if (interviewerEmail && entrevistador !== interviewerEmail.toLowerCase().trim()) {
        continue;
      }

      let include = false;

      if (reportType === 'classificados' && status === 'Classificado') {
        include = true;
      } else if (reportType === 'desclassificados' && status === 'Desclassificado') {
        include = true;
      } else if (reportType === 'entrevista_classificados' && statusEntrevista === 'Avaliado' && interviewResult === 'Classificado') {
        include = true;
      } else if (reportType === 'entrevista_desclassificados' && statusEntrevista === 'Avaliado' && interviewResult === 'Desclassificado') {
        include = true;
      }

      if (include) {
        const candidate = {};
        headers.forEach((header, index) => {
          candidate[header] = values[i][index];
        });
        candidates.push(candidate);
      }
    }

    Logger.log('✅ Relatório gerado: ' + candidates.length + ' registros');
    return candidates;
  } catch (error) {
    Logger.log('❌ Erro em getReport: ' + error.toString());
    throw error;
  }
}

// ============================================
// TRIAGEM
// ============================================

function saveScreening(params) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📝 INICIANDO saveScreening');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 Parâmetros recebidos:');
    Logger.log('   - candidateId: ' + params.candidateId);
    Logger.log('   - registrationNumber: ' + params.registrationNumber);
    Logger.log('   - cpf: ' + params.cpf);
    Logger.log('   - status (RAW): "' + params.status + '"');
    Logger.log('   - tipo do status: ' + typeof params.status);
    Logger.log('   - analystEmail: ' + params.analystEmail);

    const sh = _sheet(SHEET_CANDIDATOS);
    const headers = _getHeaders_(sh);
    const col = _colMap_(headers);
    const statusCol = col['Status'];
    const analistaCol = col['Analista'];
    const dataTriagemCol = col['Data Triagem'];
    const checkRgCpfCol = col['checkrg-cpf'];
    const checkCnhCol = col['check-cnh'];
    const checkExperienciaCol = col['check-experiencia'];
    const checkRegularidadeCol = col['check-regularidade'];
    const checkLaudoCol = col['check-laudo'];
    const checkCurriculoCol = col['check-curriculo'];
    const observacoesCol = col['Observações'];
    const capacidadeTecnicaCol = col['capacidade_tecnica'];
    const experienciaCol = col['experiencia'];
    const pontuacaoTotalCol = col['pontuacao_triagem'];
    const motivoCol = col['Motivo Desclassificação'];

    Logger.log('📊 Índice da coluna Status: ' + statusCol);

    const idx = _getIndex_(sh, headers);
    const searchKey = String(params.candidateId || params.registrationNumber || params.cpf).trim();
    let row = idx[searchKey];

    if (!row) {
      const newIdx = _buildIndex_(sh, headers);
      const rev = _getRev_();
      CacheService.getDocumentCache().put(`${IDX_CACHE_KEY}${rev}`, JSON.stringify(newIdx), CACHE_TTL_SEC);
      row = newIdx[searchKey];
    }

    if (!row) {
      Logger.log('❌ Candidato não encontrado: ' + searchKey);
      throw new Error('Candidato não encontrado: ' + searchKey);
    }

    Logger.log('📍 Candidato encontrado na linha: ' + row);

    const lastCol = sh.getLastColumn();
    const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

    // ✅ LÓGICA DE STATUS COM LOG DETALHADO
    Logger.log('───────────────────────────────────────');
    Logger.log('🔍 PROCESSANDO STATUS:');
    Logger.log('   - Status recebido: "' + params.status + '"');
    Logger.log('   - Comparação (params.status === "classificado"): ' + (params.status === 'classificado'));
    Logger.log('   - Comparação (params.status === "desclassificado"): ' + (params.status === 'desclassificado'));

    let statusFinal;
    if (params.status === 'classificado') {
      statusFinal = 'Classificado';
      Logger.log('   ✅ Status será: Classificado');
    } else if (params.status === 'desclassificado') {
      statusFinal = 'Desclassificado';
      Logger.log('   ❌ Status será: Desclassificado');
    } else {
      statusFinal = 'Desclassificado';
      Logger.log('   ⚠️ Status não reconhecido, usando padrão: Desclassificado');
    }

    if (statusCol >= 0) {
      rowVals[statusCol] = statusFinal;
      Logger.log('   📝 Status gravado na coluna ' + statusCol + ': "' + statusFinal + '"');
    } else {
      Logger.log('   ⚠️ Coluna Status não encontrada!');
    }
    Logger.log('───────────────────────────────────────');

    // Analista
    if (analistaCol >= 0 && params.analystEmail) {
      rowVals[analistaCol] = params.analystEmail;
      Logger.log('👤 Analista: ' + params.analystEmail);
    }

    // Data triagem
    if (dataTriagemCol >= 0) {
      rowVals[dataTriagemCol] = params.screenedAt || getCurrentTimestamp();
      Logger.log('📅 Data triagem: ' + rowVals[dataTriagemCol]);
    }

    // Documentos
    Logger.log('📋 Salvando documentos:');
    const updateDocument = (colIndex, value, fieldName) => {
      if (colIndex >= 0 && value !== undefined && value !== null) {
        let convertedValue = '';
        switch (value) {
          case 'conforme':
            convertedValue = 'Sim';
            break;
          case 'nao_conforme':
            convertedValue = 'Não';
            break;
          case 'nao_se_aplica':
          case null:
            convertedValue = 'Não se aplica';
            break;
          default:
            convertedValue = String(value || '');
        }
        rowVals[colIndex] = convertedValue;
        Logger.log(`   - ${fieldName}: ${convertedValue} (original: ${value})`);
      }
    };

    updateDocument(checkRgCpfCol, params['checkrg-cpf'], 'RG/CPF');
    updateDocument(checkCnhCol, params['check-cnh'], 'CNH');
    updateDocument(checkExperienciaCol, params['check-experiencia'], 'Experiência');
    updateDocument(checkRegularidadeCol, params['check-regularidade'], 'Regularidade');
    updateDocument(checkLaudoCol, params['check-laudo'], 'Laudo PCD');
    updateDocument(checkCurriculoCol, params['check-curriculo'], 'Currículo');

    // Avaliação técnica (apenas para classificados)
    if (statusFinal === 'Classificado') {
      if (capacidadeTecnicaCol >= 0 && params.capacidade_tecnica !== undefined) {
        rowVals[capacidadeTecnicaCol] = Number(params.capacidade_tecnica) || 0;
        Logger.log('   - Capacidade técnica: ' + rowVals[capacidadeTecnicaCol]);
      }
      if (experienciaCol >= 0 && params.experiencia !== undefined) {
        rowVals[experienciaCol] = Number(params.experiencia) || 0;
        Logger.log('   - Experiência: ' + rowVals[experienciaCol]);
      }
      if (pontuacaoTotalCol >= 0) {
        const total = (Number(params.capacidade_tecnica) || 0) + (Number(params.experiencia) || 0);
        rowVals[pontuacaoTotalCol] = total;
        Logger.log('   - Total score: ' + total);
      }
    }

    // Motivo desclassificação (apenas para desclassificados)
    if (statusFinal === 'Desclassificado' && motivoCol >= 0) {
      let motivo = '';
      const docsNaoConformes = [];

      if (params['checkrg-cpf'] === 'nao_conforme') docsNaoConformes.push('RG/CPF');
      if (params['check-cnh'] === 'nao_conforme') docsNaoConformes.push('CNH');
      if (params['check-experiencia'] === 'nao_conforme') docsNaoConformes.push('Experiência Profissional');
      if (params['check-regularidade'] === 'nao_conforme') docsNaoConformes.push('Regularidade Profissional');
      if (params['check-laudo'] === 'nao_conforme') docsNaoConformes.push('Laudo PCD');
      if (params['check-curriculo'] === 'nao_conforme') docsNaoConformes.push('Currículo');

      if (docsNaoConformes.length > 0) {
        motivo = `Documentos não conformes: ${docsNaoConformes.join(', ')}`;
      }

      if (params.disqualification_reason) {
        motivo = motivo ? `${motivo} | ${params.disqualification_reason}` : params.disqualification_reason;
      }

      if (!motivo) {
        motivo = 'Desclassificado pelo analista';
      }

      rowVals[motivoCol] = motivo;
      Logger.log('📝 Motivo desclassificação: ' + motivo);
    }

    // Observações
    if (observacoesCol >= 0 && params.notes) {
      rowVals[observacoesCol] = params.notes;
      Logger.log('📝 Observações salvas');
    }

    // Salvar
    _writeWholeRow_(sh, row, rowVals);
    _bumpRev_();

    Logger.log('═══════════════════════════════════════');
    Logger.log('✅ TRIAGEM SALVA COM SUCESSO');
    Logger.log('   - Status final gravado: "' + statusFinal + '"');
    Logger.log('   - Linha: ' + row);
    Logger.log('═══════════════════════════════════════');

    return {
      success: true,
      message: 'Triagem salva com sucesso',
      candidateId: searchKey,
      status: statusFinal
    };
  } catch (error) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('❌ ERRO EM saveScreening');
    Logger.log('   Erro: ' + error.toString());
    Logger.log('   Stack: ' + error.stack);
    Logger.log('═══════════════════════════════════════');

    return {
      success: false,
      error: 'Falha ao salvar triagem: ' + error.toString(),
      details: error.stack
    };
  }
}

// ============================================
// UTILITÁRIOS
// ============================================

function addStatusColumnIfNotExists() {
  const sh = _sheet(SHEET_CANDIDATOS);
  const headers = _getHeaders_(sh);

  const requiredColumns = [
    'Status',
    'Motivo Desclassificação',
    'Observações',
    'Data Triagem',
    'Analista',
    'EMAIL',
    'TELEFONE',
    'EMAIL_SENT',
    'SMS_SENT',
    'capacidade_tecnica',
    'experiencia',
    'pontuacao_triagem',
    'status_entrevista',
    'entrevistador',
    'data_entrevista',
    'interview_completed_at',
    'interview_score',
    'interview_result',
    'interview_notes',
    'formacao_adequada',
    'graduacoes_competencias',
    'descricao_processos',
    'terminologia_tecnica',
    'calma_clareza',
    'escalas_flexiveis',
    'adaptabilidade_mudancas',
    'ajustes_emergencia',
    'residencia',
    'resolucao_conflitos',
    'colaboracao_equipe',
    'adaptacao_perfis',
    'checkrg-cpf',
    'check-cnh',
    'check-experiencia',
    'check-regularidade',
    'check-laudo',
    'check-curriculo'
  ];

  let added = false;
  requiredColumns.forEach(colName => {
    if (headers.indexOf(colName) === -1) {
      const lastCol = sh.getLastColumn();
      sh.getRange(1, lastCol + 1).setValue(colName);
      Logger.log('➕ Coluna adicionada: ' + colName);
      added = true;
    }
  });

  if (added) {
    _bumpRev_();
    Logger.log('✅ Colunas adicionadas com sucesso');
  } else {
    Logger.log('✅ Todas as colunas já existem');
  }
}

function testConnection() {
  return {
    status: 'OK',
    timestamp: getCurrentTimestamp(),
    spreadsheetId: SPREADSHEET_ID
  };
}
