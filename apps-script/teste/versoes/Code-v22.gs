// Servidor do sistema de agendamento CEPRESS (Apps Script ligado à planilha).
// Regras de horário/conflito ficam em Logica.gs; aqui só leitura, gravação e telas.

const ABAS = {
  'Config': ['Chave', 'Valor'],
  'Usuários': ['Nome', 'Código', 'Perfil'], // Perfil: Direção (vê Indicadores e muda a meta) ou Recepção (em branco = Recepção)
  'Pacientes': ['Prontuário', 'Nome', 'Responsável', 'Telefone', 'Status', 'Data Inativação', 'Motivo Inativação'],
  'Profissionais': ['ID', 'Nome', 'Especialidade', 'Tipo de Agenda', 'Duração (min)', 'Máx. por Dia', 'Ativo'],
  'Expediente': ['Profissional ID', 'Dia', 'Início', 'Fim', 'Intervalo Início', 'Intervalo Fim', 'Sala'],
  'Vagas Fixas': ['ID', 'Prontuário', 'Paciente', 'Profissional ID', 'Profissional', 'Primeira Data', 'Horário', 'Frequência', 'Data Fim', 'Status', 'Criado Por', 'Criado Em'],
  'Atendimentos': ['ID', 'Data', 'Horário', 'Prontuário', 'Paciente', 'Profissional ID', 'Profissional', 'Vaga ID', 'Status', 'Motivo', 'Registrado Por', 'Registrado Em'],
  'Histórico': ['ID', 'Data', 'Horário', 'Prontuário', 'Paciente', 'Profissional ID', 'Profissional', 'Vaga ID', 'Status', 'Motivo', 'Registrado Por', 'Registrado Em'],
  'Fila Interna': ['ID', 'Prontuário', 'Paciente', 'Especialidade', 'Entrada', 'Prioridade', 'Status', 'Obs'],
  'Bloqueios': ['Data Início', 'Data Fim', 'Profissional ID', 'Motivo'], // Profissional ID vazio = clínica toda (feriado)
};
const LISTAS = {
  'Usuários!C': ['Direção', 'Recepção'],
  'Pacientes!E': ['Ativo', 'Inativo'],
  'Profissionais!D': ['Horário', 'Ordem de chegada', 'Livre'],
  'Profissionais!G': ['Sim', 'Não'],
  'Expediente!B': ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'],
  'Vagas Fixas!H': ['Semanal', 'Quinzenal', 'Mensal'],
  'Vagas Fixas!J': ['Ativa', 'Encerrada'],
  'Atendimentos!I': ['Agendado', 'Realizado', 'Faltou', 'Cancelado', 'Remarcado'],
  'Fila Interna!F': ['Alta', 'Normal'],
  'Fila Interna!G': ['Aguardando', 'Agendado', 'Cancelado'],
};
const STATUS = LISTAS['Atendimentos!I'];
const LIBERA = ['Cancelado', 'Remarcado']; // status que liberam o horário para outra pessoa
const DIAS_ATIVOS = 90; // a aba Atendimentos guarda só os últimos 90 dias e o futuro; o resto vai para Histórico
const CFG_PRONT = ['Prontuários: planilha (link ou ID)', 'Prontuários: aba', 'Prontuários: coluna nº', 'Prontuários: coluna nome',
  'Prontuários: coluna responsável', 'Prontuários: coluna telefone', 'Prontuários: coluna status'];

// ---------- Planilha ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CEPRESS')
    .addItem('Configurar planilha', 'configurar')
    .addItem('Novo usuário', 'novoUsuario')
    .addItem('Sincronizar pacientes agora', 'sincronizarAgora')
    .addItem('Sincronizar pacientes a cada hora', 'ativarSincronizacao')
    .addToUi();
}

function configurar() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(ABAS).forEach(nome => {
    const sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    sh.getRange('A:Z').setNumberFormat('@'); // tudo texto: prontuário com zero à esquerda e datas sem fuso
    if (!sh.getRange('A1').getValue()) sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  Object.keys(LISTAS).forEach(ref => {
    const [aba, col] = ref.split('!');
    const regra = SpreadsheetApp.newDataValidation().requireValueInList(LISTAS[ref]).setAllowInvalid(true).build();
    ss.getSheetByName(aba).getRange(col + '2:' + col).setDataValidation(regra);
  });
  // Arquivamento diário, de madrugada, para a tela continuar rápida com o passar dos meses.
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'arquivarAntigos'))
    ScriptApp.newTrigger('arquivarAntigos').timeBased().everyDays(1).atHour(3).create();
  const vazia = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (vazia && vazia.getLastRow() === 0) ss.deleteSheet(vazia);

  const add = (aba, linhas) => { const sh = ss.getSheetByName(aba); if (sh.getLastRow() < 2) sh.getRange(2, 1, linhas.length, linhas[0].length).setValues(linhas); };
  add('Config', [['Capacidade', '400'], [CFG_PRONT[0], ''], [CFG_PRONT[1], ''], [CFG_PRONT[2], 'Nº Prontuário'], [CFG_PRONT[3], 'Nome'],
    [CFG_PRONT[4], 'Responsável'], [CFG_PRONT[5], 'Telefone'], [CFG_PRONT[6], 'Status']]);
  add('Pacientes', [['0001', 'Ana Exemplo', 'Marta', '(00) 90000-0001', 'Ativo', '', ''], ['0002', 'João Exemplo', 'Carlos', '(00) 90000-0002', 'Ativo', '', ''],
    ['0003', 'Bia Exemplo', 'Rosa', '(00) 90000-0003', 'Inativo', '2026-09-01', 'Alta']]);
  add('Profissionais', [['P01', 'Carla (exemplo)', 'Fonoaudiologia', 'Horário', '30', '', 'Sim'], ['P02', 'Paulo (exemplo)', 'Terapia Ocupacional', 'Horário', '30', '', 'Sim'],
    ['P03', 'Dra. Rita (exemplo)', 'Psiquiatria', 'Ordem de chegada', '', '30', 'Sim'], ['P04', 'Serviço Social (exemplo)', 'Serviço Social', 'Livre', '', '', 'Sim']]);
  // Feriados nacionais (Carnaval e Corpus Christi são ponto facultativo: apague se a clínica abrir). Inclua os municipais.
  add('Bloqueios', [['2026-10-12', '', '', 'Feriado da Nossa Senhora Aparecida'], ['2026-11-02', '', '', 'Feriado de Finados'], ['2026-11-20', '', '', 'Feriado da Consciência Negra'],
    ['2026-12-25', '', '', 'Feriado de Natal'], ['2027-01-01', '', '', 'Feriado de Ano-Novo'], ['2027-02-08', '2027-02-09', '', 'Feriado de Carnaval'],
    ['2027-03-26', '', '', 'Feriado da Sexta-feira Santa'], ['2027-04-21', '', '', 'Feriado de Tiradentes'], ['2027-05-27', '', '', 'Feriado de Corpus Christi'],
    ['2027-09-07', '', '', 'Feriado da Independência'], ['2027-10-12', '', '', 'Feriado da Nossa Senhora Aparecida'], ['2027-11-02', '', '', 'Feriado de Finados'],
    ['2027-11-15', '', '', 'Feriado da Proclamação da República'], ['2027-11-20', '', '', 'Feriado da Consciência Negra'], ['2027-12-25', '', '', 'Feriado de Natal'],
    ['2028-01-01', '', '', 'Feriado de Ano-Novo'], ['2028-02-28', '2028-02-29', '', 'Feriado de Carnaval'], ['2028-04-14', '', '', 'Feriado da Sexta-feira Santa'],
    ['2028-04-21', '', '', 'Feriado de Tiradentes'], ['2028-05-01', '', '', 'Feriado do Dia do Trabalho'], ['2028-06-15', '', '', 'Feriado de Corpus Christi'],
    ['2028-09-07', '', '', 'Feriado da Independência'], ['2028-10-12', '', '', 'Feriado da Nossa Senhora Aparecida'], ['2028-11-02', '', '', 'Feriado de Finados'],
    ['2028-11-15', '', '', 'Feriado da Proclamação da República'], ['2028-11-20', '', '', 'Feriado da Consciência Negra'], ['2028-12-25', '', '', 'Feriado de Natal']]);
  add('Expediente', [['P01', 'Seg', '08:00', '12:00', '10:00', '10:30', 'Sala 3'], ['P01', 'Qua', '13:00', '17:00', '', '', 'Sala 3'],
    ['P02', 'Qua', '13:00', '17:00', '15:00', '15:30', 'Sala 1'], ['P03', 'Ter', '08:00', '12:00', '', '', 'Consultório 1']]);
  lidas = null;
}

// Leitura rápida: as abas do dia a dia vêm todas numa só chamada da API do Sheets (serviço avançado "Sheets").
// Medido com 1 ano de dados: 1,4 s pela API contra ~12 s lendo aba por aba com SpreadsheetApp.
// O resultado vale até a próxima gravação (gravar/atualizar descartam), e cada pedido da tela começa do zero (usuario()).
const ABAS_DIA = ['Usuários', 'Config', 'Pacientes', 'Profissionais', 'Expediente', 'Vagas Fixas', 'Atendimentos', 'Bloqueios', 'Fila Interna'];
let lidas = null;

function ler(aba) {
  if (!lidas || !(aba in lidas)) carregar(ABAS_DIA.indexOf(aba) >= 0 ? ABAS_DIA : [aba]);
  return lidas[aba];
}

function carregar(abas) {
  const ss = SpreadsheetApp.getActive();
  let tabelas = null;
  if (typeof Sheets !== 'undefined') {
    try { tabelas = Sheets.Spreadsheets.Values.batchGet(ss.getId(), { ranges: abas.map(a => "'" + a + "'") }).valueRanges.map(v => v.values || []); }
    catch (e) { tabelas = null; } // ex.: aba que ainda não existe; cai para o jeito antigo
  }
  if (!tabelas) tabelas = abas.map(a => { const sh = ss.getSheetByName(a); return sh ? sh.getDataRange().getDisplayValues() : []; });
  lidas = lidas || {};
  abas.forEach((a, k) => { lidas[a] = objetos(tabelas[k]); });
}

// Linhas da planilha → objetos {Cabeçalho: valor}. A API omite células vazias no fim da linha; aqui viram ''.
function objetos(v) {
  const cab = v[0] || [];
  return v.slice(1).map((r, i) => {
    const o = { _linha: i + 2 };
    cab.forEach((c, j) => { o[c] = r[j] === undefined ? '' : String(r[j]).trim(); });
    return o;
  }).filter(o => o[cab[0]]);
}

function cabecalho(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }

function gravar(aba, obj) {
  lidas = null;
  const sh = SpreadsheetApp.getActive().getSheetByName(aba), cab = cabecalho(sh);
  // Não usar appendRow: ele ignora o formato texto e grava o prontuário '0001' como o número 1.
  sh.getRange(sh.getLastRow() + 1, 1, 1, cab.length).setNumberFormat('@').setValues([cab.map(c => obj[c] === undefined ? '' : String(obj[c]))]);
}

function atualizar(aba, linha, campos) {
  lidas = null;
  const sh = SpreadsheetApp.getActive().getSheetByName(aba);
  const cab = cabecalho(sh);
  Object.keys(campos).forEach(c => sh.getRange(linha, cab.indexOf(c) + 1).setNumberFormat('@').setValue(String(campos[c])));
}

function config() {
  const c = {};
  ler('Config').forEach(r => { c[r.Chave] = r.Valor; });
  return c;
}

// Dia anterior ao período guardado em Atendimentos: os registros dele estão na aba Histórico.
function antigo(data) { return isoData(data) < somaDias(hoje(), -DIAS_ATIVOS); }
function hoje() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function agora() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function novoId(prefixo) { return prefixo + Utilities.getUuid().slice(0, 8); }

// Duas recepcionistas clicando juntas não podem pegar o mesmo horário.
function comTrava(fn) {
  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try { const r = fn(); SpreadsheetApp.flush(); return r; } finally { trava.releaseLock(); }
}

// Com Gmail comum o Google não informa quem acessa, então cada pessoa tem um código próprio (aba Usuários).
// Sem usuários cadastrados ninguém entra. Códigos com menos de 6 caracteres são ignorados.
// acesso = { cod, disp }: disp identifica o navegador, para que errar o código bloqueie só aquele aparelho.
// Quem ataca pode trocar o disp a cada tentativa, por isso existe também um limite geral bem mais alto.
// comHistorico: a tela também vai precisar da aba Histórico; já vem na mesma chamada à API.
function usuario(acesso, comHistorico) {
  lidas = null; // cada pedido da tela lê a planilha de novo
  if (comHistorico) carregar(ABAS_DIA.concat(['Histórico']));
  direcao = false;
  const a = acesso || {}, cache = CacheService.getScriptCache();
  const chave = 'erros-' + String(a.disp || 'sem-id').slice(0, 64);
  const meus = +(cache.get(chave) || 0), todos = +(cache.get('erros') || 0);
  const aparelho = String(a.disp || 'sem-id').slice(0, 8);
  if (meus >= 5) { console.warn('Acesso recusado: aparelho ' + aparelho + ' bloqueado (5 erros em 10 min)'); throw new Error('ACESSO: Muitas tentativas erradas neste aparelho. Aguarde 10 minutos.'); }
  if (todos >= 100) { console.warn('Acesso recusado: limite geral de erros atingido'); throw new Error('ACESSO: Muitas tentativas erradas no sistema. Aguarde 10 minutos.'); }
  const c = String(a.cod || '').trim();
  const u = ler('Usuários').find(x => x['Código'].length >= 6 && x['Código'] === c);
  if (!u) {
    cache.put(chave, String(meus + 1), 600);
    cache.put('erros', String(todos + 1), 600);
    console.warn('Acesso recusado: código inválido (aparelho ' + aparelho + ', ' + (meus + 1) + 'º erro)');
    throw new Error('ACESSO: Código de acesso inválido.');
  }
  direcao = u.Perfil === 'Direção';
  return u.Nome;
}
let direcao = false; // perfil de quem fez o pedido atual (preenchido por usuario())
function exigeDirecao(acao) { if (!direcao) throw new Error('Só a Direção pode ' + (acao || 'ver os indicadores e mudar a meta') + '.'); }

// Menu da planilha: cria a pessoa com um código aleatório (10 letras/números, impossível de adivinhar na prática).
function novoUsuario() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Novo usuário', 'Nome da pessoa:', ui.ButtonSet.OK_CANCEL);
  const nome = r.getResponseText().trim();
  if (r.getSelectedButton() !== ui.Button.OK || !nome) return;
  const cod = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  gravar('Usuários', { Nome: nome, 'Código': cod });
  ui.alert('Código de acesso de ' + nome + ':\n\n' + cod + '\n\nPasse este código só para essa pessoa. Para cancelar o acesso, apague a linha dela na aba Usuários.');
}

// Nome digitado direto na aba Usuários com o Código vazio: gera o código na hora.
function onEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== 'Usuários' || r.getColumn() !== 1) return;
  for (let i = 0; i < r.getNumRows(); i++) {
    const lin = r.getRow() + i;
    if (lin < 2) continue;
    const [nome, cod] = sh.getRange(lin, 1, 1, 2).getDisplayValues()[0];
    if (nome.trim() && !cod) sh.getRange(lin, 2).setNumberFormat('@').setValue(Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  }
}

// ---------- Sincronização com a planilha de prontuários ----------

function sincronizarPacientes() {
  const cfg = config();
  const ref = cfg[CFG_PRONT[0]];
  if (!ref) throw new Error('Preencha "' + CFG_PRONT[0] + '" na aba Config.');
  const origem = /^https?:/.test(ref) ? SpreadsheetApp.openByUrl(ref) : SpreadsheetApp.openById(ref);
  const sh = cfg[CFG_PRONT[1]] ? origem.getSheetByName(cfg[CFG_PRONT[1]]) : origem.getSheets()[0];
  if (!sh) throw new Error('Aba "' + cfg[CFG_PRONT[1]] + '" não encontrada na planilha de prontuários.');
  const v = sh.getDataRange().getDisplayValues();
  const cab = v.shift().map(c => c.trim().toLowerCase());
  const col = (k, obrigatoria) => {
    const i = cab.indexOf(String(cfg[k] || '').trim().toLowerCase());
    if (i < 0 && obrigatoria) throw new Error('Coluna "' + cfg[k] + '" não encontrada na planilha de prontuários.');
    return i;
  };
  const [cN, cNome, cResp, cTel, cSt] = [col(CFG_PRONT[2], true), col(CFG_PRONT[3], true), col(CFG_PRONT[4]), col(CFG_PRONT[5]), col(CFG_PRONT[6], true)];
  const status = s => /inativ/i.test(s) ? 'Inativo' : /ativ/i.test(s) ? 'Ativo' : s;
  const linhas = v.filter(r => String(r[cN]).trim()).map(r => [String(r[cN]).trim(), r[cNome], cResp < 0 ? '' : r[cResp], cTel < 0 ? '' : r[cTel], status(r[cSt])]);
  // Confere antes de apagar: uma coluna errada na Config deixaria todo mundo inativo e esvaziaria a agenda.
  const estranhos = linhas.map(l => l[4]).filter((x, i, a) => x !== 'Ativo' && x !== 'Inativo' && a.indexOf(x) === i);
  if (!linhas.length) throw new Error('A planilha de prontuários veio vazia. Nada foi alterado.');
  if (estranhos.length) throw new Error('A coluna de status tem valores que não são Ativo/Inativo: ' + estranhos.slice(0, 5).map(x => '"' + x + '"').join(', ') +
    '. Confira a coluna configurada na aba Config. Nada foi alterado.');
  const ativosNovos = linhas.filter(l => l[4] === 'Ativo').length;
  const ativosAntes = ler('Pacientes').filter(p => p.Status === 'Ativo').length;
  if (ativosAntes >= 20 && ativosNovos < ativosAntes / 2)
    throw new Error('A sincronização deixaria só ' + ativosNovos + ' pacientes ativos (hoje são ' + ativosAntes + '). Confira a planilha de prontuários. Nada foi alterado.');
  // A aba Pacientes vira espelho do prontuário: o prontuário continua sendo a fonte da verdade.
  return comTrava(() => {
    const dest = SpreadsheetApp.getActive().getSheetByName('Pacientes');
    if (dest.getLastRow() > 1) dest.getRange(2, 1, dest.getLastRow() - 1, ABAS.Pacientes.length).clearContent();
    dest.getRange(2, 1, linhas.length, 5).setNumberFormat('@').setValues(linhas);
    lidas = null;
    return linhas.length;
  });
}

// Move para a aba Histórico os atendimentos com mais de DIAS_ATIVOS dias. Roda todo dia de madrugada.
function arquivarAntigos() {
  return comTrava(() => {
    const ss = SpreadsheetApp.getActive(), at = ss.getSheetByName('Atendimentos'), hist = ss.getSheetByName('Histórico');
    const v = at.getDataRange().getDisplayValues(), cab = v.shift(), iData = cab.indexOf('Data');
    const corte = somaDias(hoje(), -DIAS_ATIVOS), velho = r => { const d = isoData(r[iData]); return d && d < corte; };
    const linhas = v.filter(r => r[0]), velhos = linhas.filter(velho);
    if (!velhos.length) return 0;
    const ficam = linhas.filter(r => !velho(r));
    // Grava no Histórico antes de apagar: se der erro no meio, sobra cópia, nunca perda.
    hist.getRange(hist.getLastRow() + 1, 1, velhos.length, cab.length).setNumberFormat('@').setValues(velhos);
    at.getRange(2, 1, v.length, cab.length).clearContent();
    if (ficam.length) at.getRange(2, 1, ficam.length, cab.length).setNumberFormat('@').setValues(ficam);
    lidas = null;
    return velhos.length;
  });
}

function sincronizarAgora() {
  SpreadsheetApp.getUi().alert(sincronizarPacientes() + ' pacientes sincronizados.');
}

function ativarSincronizacao() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sincronizarPacientes').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarPacientes').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Pronto: os pacientes serão sincronizados a cada hora.');
}

// ---------- Contexto (tudo que as regras precisam, lido de uma vez) ----------

// historico = true também lê a aba Histórico (só para datas antigas e para a ficha do paciente).
function contexto(historico) {
  const pacientes = {};
  ler('Pacientes').forEach(p => { pacientes[p['Prontuário']] = p; });
  const profs = {};
  ler('Profissionais').filter(p => p.Ativo !== 'Não').forEach(p => {
    profs[p.ID] = { id: p.ID, nome: p.Nome, esp: p.Especialidade, tipo: p['Tipo de Agenda'] || 'Horário', dur: +p['Duração (min)'] || 30, max: +p['Máx. por Dia'] || 0 };
  });
  const exp = {};
  ler('Expediente').forEach(e => {
    const id = e['Profissional ID'];
    const d = diaIdx(e.Dia);
    exp[id] = exp[id] || {};
    (exp[id][d] = exp[id][d] || []).push({ inicio: isoHora(e['Início']), fim: isoHora(e.Fim), intIni: isoHora(e['Intervalo Início']), intFim: isoHora(e['Intervalo Fim']), sala: e.Sala });
  });
  const ativo = pr => !!pacientes[pr] && pacientes[pr].Status === 'Ativo';
  const nome = pr => pacientes[pr] ? pacientes[pr].Nome : pr;
  const turnos = (prof, iso) => (exp[prof] || {})[diaSemana(iso)] || [];
  const bloqueios = ler('Bloqueios').map(b => ({ ini: isoData(b['Data Início']), fim: isoData(b['Data Fim']) || isoData(b['Data Início']),
    prof: b['Profissional ID'], motivo: b.Motivo || 'Dia bloqueado' })).filter(b => b.ini);
  const bloqueio = (prof, iso) => { const b = bloqueios.find(x => (!x.prof || x.prof === prof) && iso >= x.ini && iso <= x.fim); return b ? b.motivo : ''; };

  if (historico && !(lidas && 'Histórico' in lidas)) carregar(ABAS_DIA.concat(['Histórico'])); // tudo numa chamada só
  const arquivo = historico ? ler('Histórico').map(r => Object.assign(r, { _arquivo: true })) : [];
  const atend = arquivo.concat(ler('Atendimentos')).map(a => ({ id: a.ID, data: isoData(a.Data), hora: isoHora(a['Horário']), pront: a['Prontuário'], prof: a['Profissional ID'],
    vaga: a['Vaga ID'], status: a.Status || 'Agendado', motivo: a.Motivo, linha: a._linha, arquivado: !!a._arquivo }));
  const excecoes = {};
  atend.forEach(a => { if (a.vaga && LIBERA.indexOf(a.status) >= 0) (excecoes[a.vaga] = excecoes[a.vaga] || {})[a.data] = true; });

  const evento = o => {
    const p = profs[o.prof], dur = p ? p.dur : 30, ts = turnos(o.prof, o.inicio), t = turnoDe({ hora: o.hora, dur }, ts) || ts[0];
    return Object.assign(o, { dur, sala: t ? t.sala : '', nome: nome(o.pront), profNome: p ? p.nome : o.prof });
  };
  const vagas = ler('Vagas Fixas').map(v => evento({ id: v.ID, pront: v['Prontuário'], prof: v['Profissional ID'], inicio: isoData(v['Primeira Data']),
    hora: isoHora(v['Horário']), freq: v['Frequência'], fim: isoData(v['Data Fim']), status: v.Status, linha: v._linha, excecoes: excecoes[v.ID] || {} }));
  // Vaga encerrada continua valendo até a Data Fim. Paciente inativo continua ocupando o horário até a vaga ser
  // encerrada: se ele voltar a ficar ativo, não pode encontrar outra pessoa no lugar dele.
  const vigentes = vagas.filter(v => FREQ[v.freq] && v.inicio && (v.status === 'Ativa' || v.fim));
  // Sessões em dia bloqueado não ocupam o horário (como se tivessem sido canceladas).
  bloqueios.forEach(b => {
    for (let d = b.ini, n = 0; d <= b.fim && n < 400; d = somaDias(d, 1), n++)
      vigentes.forEach(v => { if (!b.prof || v.prof === b.prof) v.excecoes[d] = true; });
  });
  const eventos = vigentes.concat(
    atend.filter(a => !a.vaga && a.hora && LIBERA.indexOf(a.status) < 0)
      .map(a => evento({ id: a.id, pront: a.pront, prof: a.prof, inicio: a.data, hora: a.hora, freq: 'Única' })));
  return { pacientes, profs, exp, atend, vagas, vigentes, eventos, ativo, nome, turnos, bloqueio, bloqueios };
}

// Devolve a lista de problemas do horário candidato (vazia = pode agendar).
function validar(ctx, c) {
  const p = ctx.profs[c.prof];
  if (!p) throw new Error('Profissional não encontrado ou inativo.');
  if (c.pront && !ctx.ativo(c.pront)) throw new Error('Paciente ' + c.pront + ' não encontrado ou inativo.');
  if (!c.inicio || !c.hora) throw new Error('Informe data e horário.');
  const ts = ctx.turnos(c.prof, c.inicio);
  c.dur = p.dur;
  const t = turnoDe(c, ts);
  c.sala = t ? t.sala : '';
  const fora = foraDoExpediente(c, ts) || (ctx.bloqueio(c.prof, c.inicio) && 'Sem atendimento neste dia: ' + ctx.bloqueio(c.prof, c.inicio));
  return fora ? [fora] : conflitos(c, ctx.eventos);
}

// ---------- Funções chamadas pela tela (Index.html) ----------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Espaço TEA Vila Entre Rios').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function inicio(cod) {
  const u = usuario(cod), ctx = contexto();
  // dias = dias da semana em que atende (0 = domingo), para a tela sugerir a data certa.
  // horarios = turnos de cada dia da semana (tela Profissionais e cabeçalho da Agenda).
  const profs = Object.keys(ctx.profs).map(k => Object.assign({ dias: Object.keys(ctx.exp[k] || {}).map(Number).filter(d => d >= 0).sort(),
    horarios: Object.keys(ctx.exp[k] || {}).map(Number).filter(d => d >= 0).sort().map(d => ({ dia: d, turnos: ctx.exp[k][d] })) }, ctx.profs[k]));
  const esps = profs.map(p => p.esp).concat(ler('Fila Interna').map(f => f.Especialidade));
  return {
    usuario: u, direcao, hoje: hoje(), profs,
    especialidades: esps.filter((e, i) => e && esps.indexOf(e) === i).sort(),
    pacientes: Object.keys(ctx.pacientes).filter(ctx.ativo).map(k => ({ pront: k, nome: ctx.pacientes[k].Nome })),
  };
}

// Tela Pacientes: todos, inclusive inativos (a agenda só oferece os ativos).
function listaPacientes(cod) {
  usuario(cod);
  return ler('Pacientes').map(p => ({ pront: p['Prontuário'], nome: p.Nome, responsavel: p['Responsável'], telefone: p.Telefone, status: p.Status || 'Ativo' }))
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function paciente(cod, pront) {
  usuario(cod, true);
  const ctx = contexto(true), p = ctx.pacientes[pront];
  if (!p) throw new Error('Prontuário ' + pront + ' não encontrado.');
  const nomeProf = id => (ctx.profs[id] || {}).nome || id;
  return {
    pront, nome: p.Nome, responsavel: p['Responsável'], telefone: p.Telefone, status: p.Status,
    vagas: ctx.vagas.filter(v => v.pront === pront).map(v => ({ id: v.id, prof: v.profNome, dia: v.inicio ? DIAS[diaSemana(v.inicio)] : '', hora: v.hora,
      freq: v.freq, inicio: v.inicio, fim: v.fim, status: v.status })),
    fila: ler('Fila Interna').filter(f => f['Prontuário'] === pront).map(f => ({ esp: f.Especialidade, entrada: f.Entrada, status: f.Status })),
    historico: ctx.atend.filter(a => a.pront === pront).sort((x, y) => (x.data + x.hora < y.data + y.hora ? 1 : -1)).slice(0, 20)
      .map(a => ({ data: a.data, hora: a.hora, prof: nomeProf(a.prof), status: a.status, motivo: a.motivo })),
  };
}

function horariosLivres(cod, q) {
  usuario(cod);
  const ctx = contexto(), p = ctx.profs[q.prof], data = isoData(q.data), ts = p && data ? ctx.turnos(q.prof, data) : [];
  if (!ts.length) return [];
  return gradeHorarios(ts, p.dur).map(h => {
    const m = validar(ctx, { prof: q.prof, pront: ctx.ativo(q.pront) ? q.pront : '', inicio: data, hora: h, freq: q.freq, fim: isoData(q.fim) });
    return { hora: h, livre: !m.length, motivo: m.join('; ') };
  });
}

function agendar(cod, q) {
  const u = usuario(cod);
  return comTrava(() => {
    const ctx = contexto(), p = ctx.profs[q.prof], data = isoData(q.data);
    if (!p) throw new Error('Profissional não encontrado ou inativo.');
    if (!ctx.ativo(q.pront)) throw new Error('Paciente ' + q.pront + ' não encontrado ou inativo.');
    if (!data) throw new Error('Informe a data.');
    const base = { 'Prontuário': q.pront, Paciente: ctx.nome(q.pront), 'Profissional ID': p.id, Profissional: p.nome, Status: 'Agendado' };
    const reg = { 'Registrado Por': u, 'Registrado Em': agora() };
    let id;
    if (p.tipo === 'Livre') throw new Error(p.esp + ' é atendimento livre, sem agendamento.');
    if (p.tipo === 'Ordem de chegada') {
      checarOC(ctx, p, data, q.pront);
      id = novoId('A');
      gravar('Atendimentos', Object.assign({ ID: id, Data: data }, base, reg));
    } else {
      if (!(q.freq in FREQ)) throw new Error('Frequência inválida.');
      const m = validar(ctx, { prof: p.id, pront: q.pront, inicio: data, hora: isoHora(q.hora), freq: q.freq, fim: isoData(q.fim) });
      if (m.length) throw new Error(m.join('\n'));
      if (q.freq === 'Única') {
        id = novoId('A');
        gravar('Atendimentos', Object.assign({ ID: id, Data: data, 'Horário': isoHora(q.hora), Motivo: q.motivo || '' }, base, reg));
      } else {
        id = novoId('V');
        gravar('Vagas Fixas', Object.assign({ ID: id, 'Primeira Data': data, 'Horário': isoHora(q.hora), 'Frequência': q.freq, 'Data Fim': isoData(q.fim),
          'Criado Por': u, 'Criado Em': agora() }, base, { Status: 'Ativa' }));
      }
    }
    ler('Fila Interna').filter(f => f['Prontuário'] === q.pront && f.Especialidade === p.esp && f.Status === 'Aguardando')
      .forEach(f => atualizar('Fila Interna', f._linha, { Status: 'Agendado' }));
    return id;
  });
}

// Regras da lista de ordem de chegada (usadas ao agendar e ao passar para outro dia).
function checarOC(ctx, p, data, pront) {
  if (!ctx.turnos(p.id, data).length) throw new Error(p.nome + ' não atende ' + DIAS[diaSemana(data)] + '.');
  if (ctx.bloqueio(p.id, data)) throw new Error('Sem atendimento neste dia: ' + ctx.bloqueio(p.id, data) + '.');
  const dia = ctx.atend.filter(a => a.prof === p.id && a.data === data && !a.hora && LIBERA.indexOf(a.status) < 0);
  if (dia.some(a => a.pront === pront)) throw new Error('O paciente já está na lista deste dia.');
  if (p.max && dia.length >= p.max) throw new Error('LOTADO: ' + p.nome + ' já tem ' + dia.length + '/' + p.max + ' neste dia. Coloque o paciente na fila interna.');
}

function agendaDe(ctx, data) {
  const reg = {};
  ctx.atend.forEach(a => { if (a.vaga) reg[a.vaga + '|' + a.data] = a; });
  return Object.keys(ctx.profs).map(k => ctx.profs[k]).filter(p => p.tipo !== 'Livre' && ctx.turnos(p.id, data).length).map(p => {
    let itens;
    if (p.tipo === 'Ordem de chegada') {
      itens = ctx.atend.filter(a => a.prof === p.id && a.data === data && !a.hora)
        .map(a => ({ atend: a.id, pront: a.pront, nome: ctx.nome(a.pront), status: a.status, motivo: a.motivo }));
    } else {
      itens = ctx.vigentes.filter(v => v.prof === p.id && ocorre(Object.assign({}, v, { excecoes: null }), data))
        .map(v => {
          const a = reg[v.id + '|' + data];
          return { vaga: v.id, atend: a ? a.id : '', hora: v.hora, pront: v.pront, nome: v.nome, status: a ? a.status : 'Agendado', motivo: a ? a.motivo : '', freq: v.freq };
        })
        .concat(ctx.atend.filter(a => a.prof === p.id && a.data === data && a.hora && !a.vaga)
          .map(a => ({ atend: a.id, hora: a.hora, pront: a.pront, nome: ctx.nome(a.pront), status: a.status, motivo: a.motivo, freq: 'Avulsa' })))
        .sort((x, y) => x.hora < y.hora ? -1 : x.hora > y.hora ? 1 : 0);
    }
    // Em dia bloqueado as sessões continuam visíveis, para a recepção avisar ou remarcar os pacientes.
    const bloqueio = ctx.bloqueio(p.id, data);
    if (bloqueio) itens.forEach(i => { if (i.status === 'Agendado') i.status = 'Sem atendimento'; });
    // Vaga de paciente inativo segura o horário até ser encerrada (tela Paciente → Encerrar).
    else if (data >= hoje()) itens.forEach(i => { if (i.vaga && i.status === 'Agendado' && !ctx.ativo(i.pront)) i.status = 'Paciente inativo'; });
    itens.forEach(i => { i.tel = (ctx.pacientes[i.pront] || {}).Telefone || ''; });
    const salas = ctx.turnos(p.id, data).map(t => t.sala).filter((x, i, a) => x && a.indexOf(x) === i);
    return { id: p.id, nome: p.nome, esp: p.esp, tipo: p.tipo, max: p.max, sala: salas.join(' / '), bloqueio, itens };
  });
}

function agenda(cod, data) {
  usuario(cod, antigo(data));
  const d = isoData(data);
  if (!d) throw new Error('Data inválida.');
  const ctx = contexto(antigo(d));
  return { feriado: ctx.bloqueio('', d), profs: agendaDe(ctx, d) }; // feriado = bloqueio da clínica toda
}

// Semana de um profissional (segunda a sábado), a partir de qualquer dia da semana pedida.
function agendaSemana(cod, prof, data) {
  const d = isoData(data);
  if (!d) throw new Error('Data inválida.');
  const seg = somaDias(d, -((diaSemana(d) + 6) % 7));
  usuario(cod, antigo(seg));
  const ctx = contexto(antigo(seg));
  if (!ctx.profs[prof]) throw new Error('Profissional não encontrado ou inativo.');
  const dias = [];
  for (let i = 0; i < 6; i++) {
    const dia = somaDias(seg, i), col = agendaDe(ctx, dia).find(p => p.id === prof);
    dias.push({ data: dia, feriado: ctx.bloqueio('', dia), atende: !!col, bloqueio: col ? col.bloqueio : '', sala: col ? col.sala : '', itens: col ? col.itens : [] });
  }
  return { inicio: seg, prof: ctx.profs[prof], dias };
}

function marcar(ctx, q, u) {
  const campos = { Status: q.status, Motivo: q.motivo || '', 'Registrado Por': u, 'Registrado Em': agora() };
  const a = q.atend ? ctx.atend.find(x => x.id === q.atend) : ctx.atend.find(x => x.vaga === q.vaga && x.data === q.data);
  if (a && a.arquivado) throw new Error('Este atendimento já foi para o Histórico e não pode mais ser alterado.');
  if (a) return atualizar('Atendimentos', a.linha, campos);
  const v = ctx.vagas.find(x => x.id === q.vaga);
  if (!v) throw new Error('Agendamento não encontrado.');
  gravar('Atendimentos', Object.assign({ ID: novoId('A'), Data: q.data, 'Horário': v.hora, 'Prontuário': v.pront, Paciente: v.nome,
    'Profissional ID': v.prof, Profissional: v.profNome, 'Vaga ID': v.id }, campos));
}

function atualizarStatus(cod, q) {
  const u = usuario(cod, antigo(q.data));
  if (STATUS.indexOf(q.status) < 0) throw new Error('Status inválido.');
  return comTrava(() => {
    const ctx = contexto(antigo(q.data));
    // Voltar uma sessão cancelada (ou remarcada) para ativa ocupa o horário de novo: confere antes, como num agendamento.
    const a = q.atend ? ctx.atend.find(x => x.id === q.atend) : ctx.atend.find(x => x.vaga === q.vaga && x.data === q.data);
    const atual = a ? a.status : 'Agendado';
    if (LIBERA.indexOf(atual) >= 0 && LIBERA.indexOf(q.status) < 0) {
      if (atual === 'Remarcado') throw new Error('Esta sessão foi remarcada. Para voltar atrás, cancele a sessão nova e agende de novo neste horário.');
      const v = q.vaga ? ctx.vagas.find(x => x.id === q.vaga) : null, orig = v || a, p = ctx.profs[orig.prof];
      if (!p) throw new Error('Profissional não encontrado ou inativo.');
      if (!orig.hora) checarOC(ctx, p, q.data, orig.pront);
      else {
        const m = validar(ctx, { id: v ? '' : a.id, prof: orig.prof, pront: '', inicio: q.data, hora: orig.hora, freq: 'Única' });
        if (m.length) throw new Error('O horário já foi ocupado: ' + m.join('; '));
      }
    }
    marcar(ctx, q, u);
  });
}

function remarcar(cod, q) {
  const u = usuario(cod, antigo(q.data));
  return comTrava(() => {
    const ctx = contexto(antigo(q.data));
    const orig = q.vaga ? ctx.vagas.find(v => v.id === q.vaga) : ctx.atend.find(a => a.id === q.atend);
    if (!orig) throw new Error('Agendamento não encontrado.');
    if (!q.vaga && !orig.hora) { // ordem de chegada: muda só o dia e entra no fim da lista
      const po = ctx.profs[orig.prof], nova = isoData(q.novaData);
      if (!po) throw new Error('Profissional não encontrado ou inativo.');
      if (!nova) throw new Error('Informe a nova data.');
      checarOC(ctx, po, nova, orig.pront);
      marcar(ctx, { atend: q.atend, data: q.data, status: 'Remarcado', motivo: q.motivo }, u);
      gravar('Atendimentos', { ID: novoId('A'), Data: nova, 'Prontuário': orig.pront, Paciente: ctx.nome(orig.pront), 'Profissional ID': po.id,
        Profissional: po.nome, Status: 'Agendado', Motivo: 'Remarcado de ' + q.data, 'Registrado Por': u, 'Registrado Em': agora() });
      return;
    }
    if (q.vaga) ctx.eventos.filter(e => e.id === q.vaga).forEach(e => { e.excecoes[q.data] = true; }); // a sessão original sai do caminho
    const c = { id: q.vaga ? '' : q.atend, prof: orig.prof, pront: orig.pront, inicio: isoData(q.novaData), hora: isoHora(q.novaHora), freq: 'Única' };
    const m = validar(ctx, c);
    if (m.length) throw new Error(m.join('\n'));
    marcar(ctx, { atend: q.atend, vaga: q.vaga, data: q.data, status: 'Remarcado', motivo: q.motivo }, u);
    const p = ctx.profs[orig.prof];
    gravar('Atendimentos', { ID: novoId('A'), Data: c.inicio, 'Horário': c.hora, 'Prontuário': orig.pront, Paciente: ctx.nome(orig.pront), 'Profissional ID': p.id,
      Profissional: p.nome, Status: 'Agendado', Motivo: 'Remarcado de ' + q.data, 'Registrado Por': u, 'Registrado Em': agora() });
  });
}

function encerrarVaga(cod, id, fim) {
  usuario(cod);
  return comTrava(() => {
    const v = contexto().vagas.find(x => x.id === id);
    if (!v) throw new Error('Vaga não encontrada.');
    const f = isoData(fim);
    if (!f) throw new Error('Data de fim inválida.');
    atualizar('Vagas Fixas', v.linha, { Status: 'Encerrada', 'Data Fim': f });
  });
}

function fila(cod) {
  usuario(cod);
  const pr = { Alta: 0, Normal: 1 };
  return ler('Fila Interna').filter(f => f.Status === 'Aguardando')
    .map(f => ({ id: f.ID, pront: f['Prontuário'], nome: f.Paciente, esp: f.Especialidade, entrada: isoData(f.Entrada), prioridade: f.Prioridade || 'Normal', obs: f.Obs }))
    .sort((a, b) => (pr[a.prioridade] - pr[b.prioridade]) || (a.entrada < b.entrada ? -1 : 1));
}

function addFila(cod, q) {
  usuario(cod);
  return comTrava(() => {
    const ctx = contexto();
    if (!ctx.ativo(q.pront)) throw new Error('Paciente ' + q.pront + ' não encontrado ou inativo.');
    if (!q.esp) throw new Error('Informe a especialidade.');
    const daEsp = Object.keys(ctx.profs).map(k => ctx.profs[k]).filter(p => p.esp === q.esp);
    if (daEsp.length && daEsp.every(p => p.tipo === 'Livre')) throw new Error(q.esp + ' é atendimento livre: não precisa de fila.');
    if (ler('Fila Interna').some(f => f['Prontuário'] === q.pront && f.Especialidade === q.esp && f.Status === 'Aguardando'))
      throw new Error('O paciente já está na fila de ' + q.esp + '.');
    gravar('Fila Interna', { ID: novoId('F'), 'Prontuário': q.pront, Paciente: ctx.nome(q.pront), Especialidade: q.esp, Entrada: hoje(),
      Prioridade: q.prioridade || 'Normal', Status: 'Aguardando', Obs: q.obs || '' });
  });
}

function statusFila(cod, id, status) {
  usuario(cod);
  return comTrava(() => {
    const f = ler('Fila Interna').find(x => x.ID === id);
    if (!f) throw new Error('Registro não encontrado.');
    atualizar('Fila Interna', f._linha, { Status: status });
  });
}

function painel(cod) {
  usuario(cod);
  const ctx = contexto(), h = hoje(), mes = h.slice(0, 7), cfg = config();
  const ag = agendaDe(ctx, h);
  const itens = [].concat.apply([], ag.map(p => p.itens));
  const conta = (lista, st) => lista.filter(x => x.status === st).length;
  const vale = x => LIBERA.indexOf(x.status) < 0 && x.status !== 'Sem atendimento' && x.status !== 'Paciente inativo';
  const doMes = ctx.atend.filter(a => a.data.slice(0, 7) === mes);
  const porEsp = {};
  const lin = esp => (porEsp[esp] = porEsp[esp] || { esp, profs: 0, vagas: 0, fila: 0 });
  Object.keys(ctx.profs).forEach(k => lin(ctx.profs[k].esp).profs++);
  ctx.eventos.filter(e => FREQ[e.freq] && (!e.fim || e.fim >= h)).forEach(e => lin((ctx.profs[e.prof] || {}).esp || 'Sem profissional ativo').vagas++);
  ler('Fila Interna').filter(f => f.Status === 'Aguardando').forEach(f => lin(f.Especialidade).fila++);
  const pac = Object.keys(ctx.pacientes).map(k => ctx.pacientes[k]);
  // Pacientes com 2+ faltas nos últimos 60 dias (a aba Atendimentos guarda 90 dias, então os dados estão todos aqui).
  const desde = somaDias(h, -60), faltas = {};
  ctx.atend.filter(a => a.status === 'Faltou' && a.data >= desde && a.data <= h).forEach(a => {
    const f = faltas[a.pront] = faltas[a.pront] || { nome: ctx.nome(a.pront), faltas: 0, ultima: '' };
    f.faltas++; if (a.data > f.ultima) f.ultima = a.data;
  });
  const ultimoFeriado = ctx.bloqueios.filter(b => !b.prof).map(b => b.fim).sort().pop() || '';
  return {
    ativos: pac.filter(p => p.Status === 'Ativo').length, capacidade: +cfg.Capacidade || 0,
    faltosos: Object.keys(faltas).map(k => faltas[k]).filter(f => f.faltas >= 2).sort((a, b) => b.faltas - a.faltas).slice(0, 15),
    inativosComVaga: ctx.vigentes.filter(v => !ctx.ativo(v.pront) && (!v.fim || v.fim >= h))
      .map(v => ({ pront: v.pront, nome: v.nome, prof: v.profNome, quando: DIAS[diaSemana(v.inicio)] + ' ' + v.hora + ' · ' + v.freq.toLowerCase() })),
    feriadosAte: ultimoFeriado < somaDias(h, 60) ? ultimoFeriado || 'nenhum' : '',
    hoje: { total: itens.filter(vale).length, realizados: conta(itens, 'Realizado'), faltas: conta(itens, 'Faltou') },
    mes: { faltas: conta(doMes, 'Faltou'), cancelamentos: conta(doMes, 'Cancelado'), remarcacoes: conta(doMes, 'Remarcado') },
    ordemChegada: ag.filter(p => p.tipo === 'Ordem de chegada').map(p => ({ nome: p.nome, ocupadas: p.itens.filter(vale).length, max: p.max })),
    especialidades: Object.keys(porEsp).sort().map(k => porEsp[k]),
  };
}

// ---------- Indicadores (só Direção) ----------
// Cada sessão conta uma vez: vaga fixa em cada dia em que ocorre, consulta avulsa e lista de ordem de chegada.
// Não contam: feriado/bloqueio sem registro, sessão remarcada (conta no dia novo) e vaga futura de paciente inativo.
// O período vai do primeiro registro até 2 meses à frente (vagas fixas não têm fim, então o futuro precisa de limite).
function sessoes(ctx, ate) {
  const h = hoje(), out = [], usados = {};
  const reg = {};
  ctx.atend.forEach(a => { if (a.vaga) reg[a.vaga + '|' + a.data] = a; });
  // ~70 mil sessões num ano: guarda as contas repetidas (data + n dias, bloqueio do profissional no dia).
  const somas = {}, bloq = {};
  const mais = (d, n) => somas[d + n] || (somas[d + n] = somaDias(d, n));
  const bloqueado = (prof, data) => { const k = prof + data; return k in bloq ? bloq[k] : (bloq[k] = !!ctx.bloqueio(prof, data)); };
  const poe = (data, prof, pront, status) => {
    if (status === 'Remarcado') return;
    if (status === 'Agendado' && bloqueado(prof, data)) return;
    out.push({ data, prof, pront, status: STATUS.indexOf(status) >= 0 ? status : 'Agendado' });
  };
  ctx.vigentes.forEach(v => {
    const passo = FREQ[v.freq] * 7, fim = v.fim && v.fim < ate ? v.fim : ate;
    for (let d = v.inicio, n = 0; d <= fim && n < 2000; d = mais(d, passo), n++) {
      const a = reg[v.id + '|' + d];
      if (a) usados[a.id] = true;
      if (!a && d >= h && !ctx.ativo(v.pront)) continue;
      poe(d, v.prof, v.pront, a ? a.status : 'Agendado');
    }
  });
  // Avulsas, ordem de chegada e registros de vagas que já não existem.
  ctx.atend.forEach(a => { if (!usados[a.id] && a.data && a.data <= ate) poe(a.data, a.prof, a.pront, a.status); });
  return out;
}

function indicadores(cod, q) {
  usuario(cod, true);
  exigeDirecao();
  q = q || {};
  const ctx = contexto(true), h = hoje(), cfg = config();
  const limite = somaDias(somaDias(h.slice(0, 7) + '-01', 92).slice(0, 7) + '-01', -1); // fim do 2º mês à frente
  const mes = /^\d{4}-\d{2}$/.test(q.mes || '') ? q.mes : h.slice(0, 7);
  const ate = mes + '-31' > limite ? somaDias(somaDias(mes + '-01', 32).slice(0, 7) + '-01', -1) : limite;
  const profs = {};
  ler('Profissionais').forEach(p => { profs[p.ID] = { nome: p.Nome, esp: p.Especialidade }; }); // inclui inativos: o histórico deles conta
  const todas = sessoes(ctx, ate);
  const filtradas = todas.filter(x => (!q.prof || x.prof === q.prof) && (!q.esp || (profs[x.prof] || {}).esp === q.esp));
  const soma = lista => {
    const c = { agendamentos: lista.length, realizados: 0, faltas: 0, cancelados: 0, pendentes: 0 }, pacs = {};
    lista.forEach(x => {
      if (x.status === 'Realizado') { c.realizados++; pacs[x.pront] = true; }
      else if (x.status === 'Faltou') c.faltas++;
      else if (x.status === 'Cancelado') c.cancelados++;
      else c.pendentes++;
    });
    c.pacientes = Object.keys(pacs).length;
    c.pctFaltas = c.realizados + c.faltas ? Math.round(100 * c.faltas / (c.realizados + c.faltas)) : null;
    return c;
  };
  const doMes = filtradas.filter(x => x.data.slice(0, 7) === mes);
  const agrupa = chave => {
    const g = {};
    doMes.forEach(x => { const k = chave(x); (g[k] = g[k] || []).push(x); });
    return Object.keys(g).sort().map(k => Object.assign({ nome: k }, soma(g[k])));
  };
  const meses = {};
  filtradas.forEach(x => { const m = x.data.slice(0, 7); (meses[m] = meses[m] || []).push(x); });
  const primeiro = Object.keys(meses).sort()[0] || mes;
  const lista = [];
  for (let m = limite.slice(0, 7); m >= primeiro && lista.length < 120; m = somaDias(m + '-01', -1).slice(0, 7)) lista.push(Object.assign({ mes: m }, soma(meses[m] || [])));
  const fila = {};
  ler('Fila Interna').filter(f => f.Status === 'Aguardando' && (!q.esp || f.Especialidade === q.esp)).forEach(f => { fila[f.Especialidade] = (fila[f.Especialidade] || 0) + 1; });
  const metaV = +cfg['Meta ' + mes] || 0;
  return {
    mes, hoje: h, limite,
    meses: lista.map(m => m.mes),
    profs: Object.keys(profs).map(k => ({ id: k, nome: profs[k].nome, esp: profs[k].esp })),
    esps: Object.keys(profs).map(k => profs[k].esp).filter((e, i, a) => e && a.indexOf(e) === i).sort(),
    total: soma(doMes),
    acumulado: soma(filtradas),
    historico: lista,
    porProf: agrupa(x => x.prof).map(o => Object.assign(o, { nome: (profs[o.nome] || { nome: o.nome }).nome, esp: (profs[o.nome] || {}).esp || '' })),
    porEsp: agrupa(x => (profs[x.prof] || {}).esp || 'Sem especialidade'),
    fila: Object.keys(fila).sort().map(k => ({ esp: k, n: fila[k] })),
    meta: { valor: metaV, realizados: soma(todas.filter(x => x.data.slice(0, 7) === mes)).realizados }, // a meta é da unidade inteira, sem filtros
  };
}

function salvarMeta(cod, mes, valor) {
  usuario(cod);
  exigeDirecao();
  if (!/^\d{4}-\d{2}$/.test(mes || '')) throw new Error('Mês inválido.');
  const v = String(valor).trim();
  if (v && !/^\d{1,6}$/.test(v)) throw new Error('A meta precisa ser um número inteiro.');
  return comTrava(() => {
    const chave = 'Meta ' + mes, r = ler('Config').find(x => x.Chave === chave);
    if (r) atualizar('Config', r._linha, { Valor: v });
    else gravar('Config', { Chave: chave, Valor: v });
    return true;
  });
}

// ---------- Cadastro pela tela ----------
// Paciente: qualquer pessoa da equipe. Prontuário não muda depois de criado (é a chave dos agendamentos).
function salvarPaciente(cod, p) {
  usuario(cod);
  p = p || {};
  const pront = String(p.pront || '').trim(), nome = String(p.nome || '').trim();
  if (!/^[0-9A-Za-z.\-]{1,20}$/.test(pront)) throw new Error('Informe o número do prontuário (só números, letras, ponto ou traço).');
  if (!nome) throw new Error('Informe o nome do paciente.');
  const status = p.status === 'Inativo' ? 'Inativo' : 'Ativo';
  return comTrava(() => {
    const atual = ler('Pacientes').find(x => x['Prontuário'] === pront);
    if (p.novo && atual) throw new Error('Já existe o prontuário ' + pront + ' (' + atual.Nome + ').');
    if (!p.novo && !atual) throw new Error('Prontuário ' + pront + ' não encontrado.');
    const campos = { Nome: nome, 'Responsável': String(p.responsavel || '').trim(), Telefone: String(p.telefone || '').trim(), Status: status,
      'Data Inativação': status === 'Inativo' ? (atual && atual.Status === 'Inativo' && atual['Data Inativação'] || hoje()) : '',
      'Motivo Inativação': status === 'Inativo' ? String(p.motivo || '').trim() : '' };
    if (atual) atualizar('Pacientes', atual._linha, campos);
    else gravar('Pacientes', Object.assign({ 'Prontuário': pront }, campos));
    return pront;
  });
}

// Tela Profissionais: todos (inclusive inativos, para poder reativar), com os turnos de cada dia.
function listaProfissionais(cod) {
  usuario(cod);
  const exp = ler('Expediente');
  return ler('Profissionais').map(p => ({ id: p.ID, nome: p.Nome, esp: p.Especialidade, tipo: p['Tipo de Agenda'] || 'Horário', dur: p['Duração (min)'], max: p['Máx. por Dia'],
    ativo: p.Ativo !== 'Não', turnos: exp.filter(e => e['Profissional ID'] === p.ID).map(e => ({ dia: diaIdx(e.Dia), inicio: isoHora(e['Início']), fim: isoHora(e.Fim),
      intIni: isoHora(e['Intervalo Início']), intFim: isoHora(e['Intervalo Fim']), sala: e.Sala })).sort((a, b) => a.dia - b.dia || (a.inicio < b.inicio ? -1 : 1)) }))
    .sort((a, b) => String(a.nome).localeCompare(String(b.nome), 'pt-BR'));
}

// Profissional: só Direção. O expediente enviado substitui todo o expediente dele.
// Devolve as vagas fixas ativas que ficaram fora do novo horário (não são apagadas: a recepção decide).
function salvarProfissional(cod, q) {
  usuario(cod);
  exigeDirecao('cadastrar ou mudar profissionais');
  q = q || {};
  const nome = String(q.nome || '').trim(), esp = String(q.esp || '').trim(), tipo = q.tipo;
  if (!nome) throw new Error('Informe o nome do profissional.');
  if (!esp) throw new Error('Informe a especialidade.');
  if (LISTAS['Profissionais!D'].indexOf(tipo) < 0) throw new Error('Escolha o tipo de agenda.');
  const dur = +q.dur, max = +q.max;
  if (tipo === 'Horário' && !(Number.isInteger(dur) && dur >= 5 && dur <= 240)) throw new Error('Duração da sessão: de 5 a 240 minutos.');
  if (tipo === 'Ordem de chegada' && !(Number.isInteger(max) && max >= 1 && max <= 200)) throw new Error('Máximo por dia: de 1 a 200.');
  const turnos = (q.turnos || []).map(t => ({ dia: +t.dia, inicio: isoHora(t.inicio), fim: isoHora(t.fim), intIni: isoHora(t.intIni), intFim: isoHora(t.intFim), sala: String(t.sala || '').trim() }));
  turnos.forEach(t => {
    const qual = (DIAS[t.dia] || '?') + ' ' + (t.inicio || '');
    if (!(t.dia >= 1 && t.dia <= 6)) throw new Error('Dia inválido no expediente (use segunda a sábado).');
    if (!t.inicio || !t.fim || t.inicio >= t.fim) throw new Error('Expediente de ' + qual + ': o fim precisa ser depois do início.');
    if (!t.intIni !== !t.intFim || (t.intIni && !(t.inicio < t.intIni && t.intIni < t.intFim && t.intFim < t.fim))) throw new Error('Expediente de ' + qual + ': o intervalo precisa ter início e fim, dentro do horário.');
  });
  turnos.slice().sort((a, b) => a.dia - b.dia || (a.inicio < b.inicio ? -1 : 1)).forEach((t, i, l) => {
    if (i && l[i - 1].dia === t.dia && t.inicio < l[i - 1].fim) throw new Error('Expediente de ' + DIAS[t.dia] + ': dois turnos se sobrepõem.');
  });
  return comTrava(() => {
    const todos = ler('Profissionais');
    let id = q.id;
    const atual = id ? todos.find(x => x.ID === id) : null;
    if (id && !atual) throw new Error('Profissional não encontrado.');
    if (!id) id = 'P' + String(Math.max(0, ...todos.map(x => +((/^P(\d+)$/.exec(x.ID) || [])[1] || 0))) + 1).padStart(2, '0');
    const campos = { Nome: nome, Especialidade: esp, 'Tipo de Agenda': tipo, 'Duração (min)': tipo === 'Horário' ? dur : '',
      'Máx. por Dia': tipo === 'Ordem de chegada' ? max : '', Ativo: q.ativo === false ? 'Não' : 'Sim' };
    if (atual) atualizar('Profissionais', atual._linha, campos);
    else gravar('Profissionais', Object.assign({ ID: id }, campos));
    const sh = SpreadsheetApp.getActive().getSheetByName('Expediente');
    ler('Expediente').filter(e => e['Profissional ID'] === id).map(e => e._linha).sort((a, b) => b - a).forEach(l => sh.deleteRow(l)); // de baixo para cima
    lidas = null;
    turnos.forEach(t => gravar('Expediente', { 'Profissional ID': id, Dia: DIAS[t.dia], 'Início': t.inicio, Fim: t.fim, 'Intervalo Início': t.intIni, 'Intervalo Fim': t.intFim, Sala: t.sala }));
    const ctx = contexto(), h = hoje();
    const fora = ctx.vigentes.filter(v => v.prof === id && (!v.fim || v.fim >= h) && foraDoExpediente(v, ctx.turnos(id, v.inicio)))
      .map(v => v.nome + ' · ' + DIAS[diaSemana(v.inicio)] + ' ' + v.hora);
    return { id, fora };
  });
}
