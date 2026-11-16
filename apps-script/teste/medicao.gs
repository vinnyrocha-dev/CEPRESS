// Medição de velocidade com 1 ano de dados. Usar SÓ numa CÓPIA da planilha (gerarAno apaga e recria os dados).
// No editor do Apps Script da cópia: rodar gerarAno() uma vez, depois medir(). Os tempos saem no registro de execução.

function gerarAno() {
  const ss = SpreadsheetApp.getActive();
  const limpa = aba => { const sh = ss.getSheetByName(aba); if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent(); return sh; };
  const poe = (aba, linhas) => { const sh = limpa(aba); if (linhas.length) sh.getRange(2, 1, linhas.length, linhas[0].length).setNumberFormat('@').setValues(linhas); };
  const h = hoje(), esp = ['Fonoaudiologia', 'Psicologia', 'Terapia Ocupacional', 'Fisioterapia', 'Psicopedagogia', 'Musicoterapia', 'Psicomotricidade', 'Nutrição'];
  const n4 = i => ('000' + i).slice(-4);

  poe('Usuários', [['Medição', 'medicao123', 'Direção']]);
  poe('Pacientes', Array.from({ length: 520 }, (_, i) => [n4(i + 1), 'Paciente ' + (i + 1), 'Responsável ' + (i + 1), '(00) 90000-' + n4(i), i < 400 ? 'Ativo' : 'Inativo', '', '']));
  const profs = [];
  for (let i = 1; i <= 25; i++) profs.push(['T' + i, 'Terapeuta ' + i, esp[i % esp.length], 'Horário', i % 5 === 0 ? '45' : '30', '', 'Sim']);
  ['Psiquiatria', 'Neuropediatria', 'Neuropsicologia'].forEach((e, i) => profs.push(['M' + (i + 1), 'Médico ' + (i + 1), e, 'Ordem de chegada', '', i === 2 ? '16' : '30', 'Sim']));
  profs.push(['L1', 'Serviço Social', 'Serviço Social', 'Livre', '', '', 'Sim'], ['L2', 'Enfermagem', 'Enfermagem', 'Livre', '', '', 'Sim']);
  poe('Profissionais', profs);
  const exp = [], dias = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex'];
  for (let i = 1; i <= 25; i++) dias.forEach((d, j) => {
    if ((i + j) % 5 === 4) return; // cada terapeuta folga um dia
    exp.push(['T' + i, d, '08:00', '12:00', '', '', 'Sala ' + i], ['T' + i, d, '13:00', '17:00', '15:00', '15:30', 'Sala ' + i]);
  });
  for (let i = 1; i <= 3; i++) exp.push(['M' + i, dias[i], '08:00', '12:00', '', '', 'Consultório ' + i], ['M' + i, dias[i + 1], '08:00', '12:00', '', '', 'Consultório ' + i]);
  poe('Expediente', exp);

  // 800 vagas quinzenais (2 por paciente ativo), começando há até 1 ano
  const vagas = [];
  for (let k = 0; k < 800; k++) {
    const pac = n4((k % 400) + 1), t = 'T' + ((k % 25) + 1), ini = somaDias(h, -364 + (k % 7) - 7 * (k % 50));
    vagas.push(['V' + k, pac, 'Paciente ' + ((k % 400) + 1), t, 'Terapeuta ' + ((k % 25) + 1), ini, ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '13:00'][k % 8],
      'Quinzenal', '', 'Ativa', 'Medição', h + ' 08:00']);
  }
  poe('Vagas Fixas', vagas);

  // 40 mil atendimentos no último ano (presenças/faltas das vagas, avulsas e listas dos médicos)
  const st = ['Realizado', 'Realizado', 'Realizado', 'Realizado', 'Faltou', 'Cancelado'], at = [];
  for (let k = 0; k < 40000; k++) {
    const d = somaDias(h, -365 + Math.floor(k * 395 / 40000)), pac = n4((k % 400) + 1);
    if (k % 3 === 0) at.push(['A' + k, d, '', pac, 'Paciente ' + pac, 'M' + ((k % 3) + 1), 'Médico', '', d > h ? 'Agendado' : st[k % 6], '', 'Medição', d + ' 08:00']);
    else if (k % 3 === 1) at.push(['A' + k, d, '14:00', pac, 'Paciente ' + pac, 'T' + ((k % 25) + 1), 'Terapeuta', '', d > h ? 'Agendado' : st[k % 6], '', 'Medição', d + ' 08:00']);
    else at.push(['A' + k, d, '08:00', pac, 'Paciente ' + pac, 'T' + ((k % 25) + 1), 'Terapeuta', 'V' + (k % 800), d > h ? 'Agendado' : st[k % 6], '', 'Medição', d + ' 08:00']);
  }
  poe('Atendimentos', at);
  limpa('Histórico'); limpa('Fila Interna');
  console.log('Pronto: ' + at.length + ' atendimentos, ' + vagas.length + ' vagas, ' + exp.length + ' linhas de expediente.');
}

function medir() {
  const C = { cod: 'medicao123', disp: 'medicao' }, h = hoje();
  const qua = (() => { let d = h; while (diaSemana(d) !== 3) d = somaDias(d, 1); return d; })();
  const t = (nome, fn) => { const ini = Date.now(); fn(); const ms = Date.now() - ini; console.log(nome + ': ' + ms + ' ms'); return ms; };
  console.log('Linhas em Atendimentos: ' + (SpreadsheetApp.getActive().getSheetByName('Atendimentos').getLastRow() - 1) +
    ' | Histórico: ' + (SpreadsheetApp.getActive().getSheetByName('Histórico').getLastRow() - 1));
  ['Usuários', 'Pacientes', 'Profissionais', 'Expediente', 'Vagas Fixas', 'Atendimentos', 'Bloqueios', 'Fila Interna'].forEach(a => t('  ler ' + a, () => ler(a)));
  t('contexto()', () => contexto());
  t('inicio (abrir o sistema)', () => inicio(C));
  t('agenda do dia', () => agenda(C, qua));
  t('horarios livres', () => horariosLivres(C, { prof: 'T3', pront: '0007', data: qua, freq: 'Quinzenal' }));
  t('painel', () => painel(C));
  t('ficha do paciente (com histórico)', () => paciente(C, '0007'));
  t('indicadores (mês atual, com histórico)', () => indicadores(C, {}));
  t('indicadores (filtro de profissional)', () => indicadores(C, { prof: 'T3' }));
  t('agendar (grava)', () => { try { agendar(C, { prof: 'M1', pront: '0399', data: somaDias(qua, -1), }); } catch (e) { console.log('    ' + e.message); } });
}

// Compara formas de ler a planilha (getDisplayValues é o que o sistema usa hoje).
function compararLeitura() {
  const ss = SpreadsheetApp.getActive();
  const t = (nome, fn) => { const ini = Date.now(); const r = fn(); console.log(nome + ': ' + (Date.now() - ini) + ' ms'); return r; };
  ['Atendimentos', 'Pacientes', 'Expediente'].forEach(aba => {
    const sh = ss.getSheetByName(aba);
    t(aba + ' getDisplayValues', () => sh.getDataRange().getDisplayValues());
    const v = t(aba + ' getValues', () => sh.getDataRange().getValues());
    console.log('   tipos na 2ª linha: ' + v[1].map(x => typeof x).join(','));
  });
  t('CacheService get (500 KB)', () => { const c = CacheService.getScriptCache(); c.put('x', 'a'.repeat(90000), 600); return c.get('x'); });
}

// Arquiva e mede de novo (para comparar antes/depois do arquivamento).
function arquivarEMedir() { console.log('Arquivados: ' + arquivarAntigos()); medir(); }

// Lê várias abas numa só chamada pela API do Sheets (serviço avançado "Sheets").
function compararAPI() {
  const id = SpreadsheetApp.getActive().getId();
  const t = (nome, fn) => { const ini = Date.now(); const r = fn(); console.log(nome + ': ' + (Date.now() - ini) + ' ms'); return r; };
  const pequenas = ['Usuários', 'Config', 'Pacientes', 'Profissionais', 'Expediente', 'Vagas Fixas', 'Bloqueios', 'Fila Interna'];
  t('jeito atual: 8 abas pequenas, uma por vez', () => pequenas.forEach(a => ler(a)));
  const r = t('API: as mesmas 8 abas numa chamada', () => Sheets.Spreadsheets.Values.batchGet(id, { ranges: pequenas.map(a => "'" + a + "'") }));
  console.log('   linhas recebidas: ' + r.valueRanges.map(v => (v.values || []).length).join(','));
  t('API: Atendimentos (40 mil linhas)', () => Sheets.Spreadsheets.Values.get(id, "'Atendimentos'"));
  t('API: tudo junto numa chamada', () => Sheets.Spreadsheets.Values.batchGet(id, { ranges: pequenas.concat(['Atendimentos']).map(a => "'" + a + "'") }));
}
