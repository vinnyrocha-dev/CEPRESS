// Testes das correções da auditoria (cada bloco começa com uma planilha nova): node teste/correcoes.js
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const J = o => JSON.stringify(o);
function nova() {
  const c = vm.createContext({ console, crypto, Intl });
  for (const f of ['teste/simulador.js', 'Logica.gs', 'Code.gs']) vm.runInContext(fs.readFileSync(__dirname + '/../' + f, 'utf8'), c, { filename: f });
  vm.runInContext("configurar(); gravar('Usuários', { Nome: 'Maria', 'Código': 'abc123' }); var C = { cod: 'abc123', disp: 'x' };", c);
  const run = code => vm.runInContext(code, c);
  const erro = (code, re) => assert.throws(() => run(code), e => re.test(e.message), code);
  const prox = dia => run(`(() => { let d = hoje(); do d = somaDias(d, 1); while (diaSemana(d) !== ${dia}); return d; })()`);
  const itens = (data, prof) => run(`agenda(C, '${data}')`).profs.find(p => p.id === prof).itens.map(i => i.hora + ' ' + i.pront + ' ' + i.status);
  return { run, erro, prox, itens };
}

{ // 1. Desfazer não pode colocar duas pessoas no mesmo horário
  const { run, erro, prox, itens } = nova(), qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Quinzenal' })})`);
  const vaga = run(`ler('Vagas Fixas')[0].ID`);
  run(`atualizarStatus(C, ${J({ vaga, data: qua, status: 'Cancelado' })})`);
  run(`atualizarStatus(C, ${J({ vaga, data: qua, status: 'Agendado' })})`); // horário ainda livre: pode desfazer
  run(`atualizarStatus(C, ${J({ vaga, data: qua, status: 'Cancelado' })})`);
  // Atendimento em dupla: o 2º paciente no mesmo horário entra com aviso; o 3º não
  run(`agendar(C, ${J({ prof: 'P01', pront: '0002', data: qua, hora: '13:00', freq: 'Única' })})`);
  assert.strictEqual(J(itens(qua, 'P01')), J(['13:00 0001 Cancelado', '13:00 0002 Agendado']));
  run(`gravar('Pacientes', { 'Prontuário': '0009', Nome: 'Terceiro', Status: 'Ativo' })`);
  run(`atualizarStatus(C, ${J({ vaga, data: qua, status: 'Agendado' })})`); // desfazer vira dupla
  erro(`agendar(C, ${J({ prof: 'P01', pront: '0009', data: qua, hora: '13:00', freq: 'Única' })})`, /já tem 2 pacientes/);
  erro(`atualizarStatus(C, ${J({ vaga, data: qua, status: 'Cancelado' })}); agendar(C, ${J({ prof: 'P01', pront: '0009', data: qua, hora: '13:00', freq: 'Única' })}); atualizarStatus(C, ${J({ vaga, data: qua, status: 'Agendado' })})`, /já foi ocupado/);
  run(`agendar(C, ${J({ prof: 'P02', pront: '0002', data: qua, hora: '16:00', freq: 'Única' })})`);
  const r2 = run(`agendar(C, ${J({ prof: 'P02', pront: '0009', data: qua, hora: '16:00', freq: 'Única', comAvisos: true })})`);
  assert.ok(/em dupla com João/.test(r2.avisos[0]), J(r2));
  // Remarcado não volta: a sessão nova continua valendo
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '14:00', freq: 'Única' })})`);
  const avulsa = run(`ler('Atendimentos').find(a => a['Horário'] === '14:00').ID`);
  run(`remarcar(C, ${J({ atend: avulsa, data: qua, novaData: qua, novaHora: '15:00' })})`);
  erro(`atualizarStatus(C, ${J({ atend: avulsa, data: qua, status: 'Agendado' })})`, /foi remarcada/);
}
{ // 1b. Ordem de chegada: desfazer não passa do limite do dia
  const { run, erro, prox } = nova(), ter = prox(2);
  run(`atualizar('Profissionais', 4, { 'Máx. por Dia': '1' })`);
  run(`agendar(C, ${J({ prof: 'P03', pront: '0001', data: ter })})`);
  const a1 = run(`ler('Atendimentos')[0].ID`);
  run(`atualizarStatus(C, ${J({ atend: a1, data: ter, status: 'Cancelado' })})`);
  run(`agendar(C, ${J({ prof: 'P03', pront: '0002', data: ter })})`);
  erro(`atualizarStatus(C, ${J({ atend: a1, data: ter, status: 'Agendado' })})`, /LOTADO/);
}
{ // 2. Sincronização: não esvazia a agenda com coluna errada, prontuário vazio ou queda grande de ativos
  const { run, erro } = nova();
  const externa = linhas => run(`(() => { const e = new Planilha(); e.insertSheet('P').d = ${J([['Nº Prontuário', 'Nome', 'Status']].concat(linhas))}; externas['X'] = e; atualizar('Config', 3, { Valor: 'X' }); })()`);
  externa([['0001', 'Ana', 'Sim'], ['0002', 'João', 'Não']]);
  erro('sincronizarPacientes()', /"Sim", "Não".*Nada foi alterado/);
  externa([]);
  erro('sincronizarPacientes()', /vazia/);
  assert.strictEqual(run(`ler('Pacientes').length`), 3); // nada mudou
  // queda grande: 30 ativos hoje, a planilha nova traria só 5
  run(`(() => { const s = SpreadsheetApp.getActive().getSheetByName('Pacientes'); for (let i = 10; i < 40; i++) gravar('Pacientes', { 'Prontuário': '00' + i, Nome: 'P' + i, Status: 'Ativo' }); })()`);
  externa(Array.from({ length: 30 }, (_, i) => ['00' + (10 + i), 'P' + i, i < 5 ? 'Ativo' : 'Inativo']));
  erro('sincronizarPacientes()', /só 5 pacientes ativos \(hoje são 32\)/);
  externa([['0001', 'Ana', 'Ativo'], ['0002', 'João', 'ATIVO']].concat(Array.from({ length: 30 }, (_, i) => ['00' + (10 + i), 'P' + i, 'Ativo'])));
  assert.strictEqual(run('sincronizarPacientes()'), 32);
}
{ // 3. Paciente inativo continua segurando o horário (e aparece como "Paciente inativo" na agenda)
  const { run, erro, prox, itens } = nova(), qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  run(`atualizar('Pacientes', 2, { Status: 'Inativo' })`);
  run(`gravar('Pacientes', { 'Prontuário': '0009', Nome: 'Terceiro', Status: 'Ativo' })`);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0002', data: qua, hora: '13:00', freq: 'Semanal' })})`); // dupla com a vaga da Ana (inativa)
  erro(`agendar(C, ${J({ prof: 'P01', pront: '0009', data: qua, hora: '13:00', freq: 'Semanal' })})`, /já tem 2 pacientes/);
  assert.strictEqual(J(itens(qua, 'P01')), J(['13:00 0001 Paciente inativo', '13:00 0002 Agendado']));
  const p = run('painel(C)');
  assert.strictEqual(p.inativosComVaga.length, 1);
  assert.strictEqual(p.inativosComVaga[0].nome, 'Ana Exemplo');
  // encerrada a vaga, o horário libera
  run(`encerrarVaga(C, ${J(run(`ler('Vagas Fixas')[0].ID`))}, hoje())`);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0009', data: qua, hora: '13:00', freq: 'Semanal' })})`);
}
{ // 4. Dois turnos no mesmo dia: os dois aparecem, cada um com a sua sala
  const { run, prox } = nova(), qua = prox(3);
  run(`gravar('Expediente', { 'Profissional ID': 'P02', Dia: 'Qua', 'Início': '08:00', Fim: '09:00', Sala: 'Sala 2' })`);
  assert.strictEqual(J(run(`horariosLivres(C, ${J({ prof: 'P02', data: qua, freq: 'Única' })})`).map(h => h.hora).slice(0, 4)), J(['08:00', '08:30', '13:00', '13:30']));
  run(`agendar(C, ${J({ prof: 'P02', pront: '0001', data: qua, hora: '08:00', freq: 'Única' })})`);
  // às 08:00 o Paulo está na Sala 2; a Carla (Sala 3 à tarde) não é afetada, e a sala do cartão mostra as duas
  assert.strictEqual(run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P02').sala, 'Sala 1 / Sala 2');
  assert.strictEqual(run(`contexto().eventos.find(e => e.prof === 'P02').sala`), 'Sala 2');
}
{ // 5. Encerrar vaga só com data válida
  const { run, erro, prox } = nova(), qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  const id = run(`ler('Vagas Fixas')[0].ID`);
  erro(`encerrarVaga(C, '${id}', 'amanhã')`, /inválida/);
  erro(`encerrarVaga(C, '${id}', '15/13/2026')`, /inválida/);
  assert.strictEqual(run(`ler('Vagas Fixas')[0].Status`), 'Ativa');
  run(`encerrarVaga(C, '${id}', '31/12/2026')`);
  assert.strictEqual(run(`ler('Vagas Fixas')[0]['Data Fim']`), '2026-12-31');
}
{ // 7. Arquivamento: atendimentos antigos vão para Histórico; agenda antiga e ficha do paciente continuam vendo tudo
  const { run, erro } = nova();
  assert.ok(run(`gatilhos.some(t => t.getHandlerFunction() === 'arquivarAntigos')`)); // rotina diária criada pelo configurar
  run(`configurar()`);
  assert.strictEqual(run(`gatilhos.filter(t => t.getHandlerFunction() === 'arquivarAntigos').length`), 1); // sem duplicar
  const velho = run(`somaDias(hoje(), -200)`), recente = run(`somaDias(hoje(), -10)`);
  run(`gravar('Atendimentos', { ID: 'A-velho', Data: '${velho}', 'Prontuário': '0001', 'Profissional ID': 'P03', Status: 'Faltou' })`);
  run(`gravar('Atendimentos', { ID: 'A-novo', Data: '${recente}', 'Prontuário': '0001', 'Profissional ID': 'P03', Status: 'Faltou' })`);
  assert.strictEqual(run('arquivarAntigos()'), 1);
  assert.strictEqual(J(run(`ler('Atendimentos').map(a => a.ID)`)), J(['A-novo']));
  assert.strictEqual(J(run(`ler('Histórico').map(a => a.ID)`)), J(['A-velho']));
  assert.strictEqual(run('arquivarAntigos()'), 0); // rodar de novo não move nada
  assert.strictEqual(J(run(`paciente(C, '0001')`).historico.map(h => h.data)), J([recente, velho]));
  // a agenda de um dia antigo lê o Histórico, mas o registro arquivado não pode ser alterado
  if (run(`diaSemana('${velho}')`) === 2) assert.strictEqual(run(`agenda(C, '${velho}')`).profs[0].itens.length, 1);
  erro(`atualizarStatus(C, ${J({ atend: 'A-velho', data: velho, status: 'Realizado' })})`, /Histórico/);
  erro(`agenda(C, '2026-02-31')`, /Data inválida/);
}
{ // 9 e 10. Painel avisa quando os feriados estão acabando e lista quem mais falta
  const { run } = nova();
  assert.strictEqual(run('painel(C)').feriadosAte, ''); // feriados cadastrados até 2028
  run("SpreadsheetApp.getActive().getSheetByName('Bloqueios').d.length = 1");
  assert.strictEqual(run('painel(C)').feriadosAte, 'nenhum');
  for (const d of [-5, -20, -40, -100]) run(`gravar('Atendimentos', { ID: 'F${d}', Data: somaDias(hoje(), ${d}), 'Prontuário': '0002', 'Profissional ID': 'P03', Status: 'Faltou' })`);
  run(`gravar('Atendimentos', { ID: 'F1', Data: somaDias(hoje(), -3), 'Prontuário': '0001', 'Profissional ID': 'P03', Status: 'Faltou' })`);
  const f = run('painel(C)').faltosos;
  assert.strictEqual(f.length, 1); // a Ana tem só 1 falta; o João tem 3 nos últimos 60 dias (a de 100 dias não conta)
  assert.strictEqual(f[0].nome, 'João Exemplo');
  assert.strictEqual(f[0].faltas, 3);
}
{ // 7b. Leitura rápida: uma tela inteira faz uma só chamada à API; sem a API (ou faltando aba) cai para o jeito antigo
  const { run, prox } = nova(), qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  run('chamadasAPI = 0'); run(`agenda(C, '${qua}')`);
  assert.strictEqual(run('chamadasAPI'), 1);
  run('chamadasAPI = 0'); run(`horariosLivres(C, ${J({ prof: 'P01', data: qua, freq: 'Única' })})`);
  assert.strictEqual(run('chamadasAPI'), 1);
  run('chamadasAPI = 0'); run(`paciente(C, '0001')`); // ficha do paciente: Histórico vem na mesma chamada
  assert.strictEqual(run('chamadasAPI'), 1);
  // célula vazia no fim da linha (a API não manda): continua lida como ''
  assert.strictEqual(run(`ler('Pacientes').find(p => p['Prontuário'] === '0001')['Motivo Inativação']`), '');
  const comAPI = J(run(`agenda(C, '${qua}')`));
  run("var SheetsGuardado = Sheets; Sheets = undefined");
  assert.strictEqual(J(run(`agenda(C, '${qua}')`)), comAPI); // mesmo resultado pelo jeito antigo
  run("Sheets = SheetsGuardado; delete SpreadsheetApp.getActive().abas['Fila Interna']"); // aba faltando: a API recusa o pedido todo
  assert.strictEqual(J(run(`agenda(C, '${qua}')`)), comAPI);
}
{ // Nome digitado na aba Usuários ganha código; quem já tem código não muda
  const { run } = nova();
  run(`(() => { const sh = SpreadsheetApp.getActive().getSheetByName('Usuários');
    sh.getRange(3, 1, 2, 1).setValues([['Ana'], ['Bia']]);
    onEdit({ range: { getSheet: () => sh, getColumn: () => 1, getRow: () => 2, getNumRows: () => 3 } }); })()`);
  const u = run("ler('Usuários')");
  assert.strictEqual(u[0]['Código'], 'abc123');
  assert.ok(/^[0-9a-f]{10}$/.test(u[1]['Código']) && /^[0-9a-f]{10}$/.test(u[2]['Código']) && u[1]['Código'] !== u[2]['Código']);
}
{ // Indicadores: só Direção; cada sessão conta uma vez; feriado, remarcado e vaga futura de inativo não contam
  const { run, erro } = nova();
  erro(`indicadores(C, {})`, /Só a Direção/);
  erro(`salvarMeta(C, '2026-10', '10')`, /Só a Direção/);
  run(`gravar('Usuários', { Nome: 'Dir', 'Código': 'dir123', Perfil: 'Direção' }); var DIR = { cod: 'dir123', disp: 'y' };`);
  // Agenda vazia: nada quebra e não aparece porcentagem sem base.
  const vazio = run(`indicadores(DIR, {})`);
  assert.strictEqual(vazio.total.agendamentos, 0);
  assert.strictEqual(vazio.total.pctFaltas, null);
  assert.ok(vazio.historico.length >= 1);
  assert.strictEqual(run(`inicio(DIR)`).direcao, true);
  assert.strictEqual(run(`inicio(C)`).direcao, false);

  const h = run('hoje()');
  const w0 = run(`(() => { let d = somaDias(hoje(), -21); while (diaSemana(d) !== 3) d = somaDias(d, -1); return d; })()`); // quarta passada
  const ter = run(`(() => { let d = somaDias(hoje(), -1); while (diaSemana(d) !== 2) d = somaDias(d, -1); return d; })()`);
  const d7 = run(`somaDias('${w0}', 7)`), d14 = run(`somaDias('${w0}', 14)`);
  const A = (id, data, hora, pront, prof, vaga, st) => run(`gravar('Atendimentos', ${J({ ID: id, Data: data, 'Horário': hora, 'Prontuário': pront, 'Profissional ID': prof, 'Vaga ID': vaga, Status: st })})`);
  run(`gravar('Vagas Fixas', ${J({ ID: 'V1', 'Prontuário': '0001', 'Profissional ID': 'P01', 'Primeira Data': w0, 'Horário': '13:00', 'Frequência': 'Semanal', Status: 'Ativa' })})`);
  A('a1', w0, '13:00', '0001', 'P01', 'V1', 'Realizado');
  A('a2', d7, '13:00', '0001', 'P01', 'V1', 'Faltou');
  A('a3', d14, '13:00', '0001', 'P01', 'V1', 'Cancelado');
  A('a4', w0, '14:00', '0002', 'P01', '', 'Realizado');   // avulsa
  A('a5', d7, '14:00', '0002', 'P01', '', 'Remarcado');   // não conta
  A('a6', ter, '', '0002', 'P03', '', 'Faltou');          // ordem de chegada
  // Feriado numa quarta futura da vaga: essa sessão não conta.
  const fer = run(`(() => { let d = somaDias(hoje(), 1); while (diaSemana(d) !== 3) d = somaDias(d, 1); return d; })()`);
  run(`gravar('Bloqueios', { 'Data Início': '${fer}', Motivo: 'Feriado X' })`);

  const r = run(`indicadores(DIR, {})`);
  // Sessões futuras esperadas da vaga: quartas de w0+21 até o limite, menos o feriado (e as que já passaram sem registro também são "sem registro").
  let pend = 0;
  for (let d = run(`somaDias('${w0}', 21)`); d <= r.limite; d = run(`somaDias('${d}', 7)`)) if (d !== fer) pend++;
  assert.strictEqual(J([r.acumulado.realizados, r.acumulado.faltas, r.acumulado.cancelados, r.acumulado.pendentes]), J([2, 2, 1, pend]));
  assert.strictEqual(r.acumulado.agendamentos, 5 + pend);
  assert.strictEqual(r.acumulado.pacientes, 2);
  assert.strictEqual(r.acumulado.pctFaltas, 50);
  // A soma dos meses do histórico bate com o acumulado.
  assert.strictEqual(r.historico.reduce((s, m) => s + m.agendamentos, 0), r.acumulado.agendamentos);
  // Filtros por profissional e por especialidade.
  const rita = run(`indicadores(DIR, { prof: 'P03' })`).acumulado;
  assert.strictEqual(J([rita.agendamentos, rita.faltas]), J([1, 1]));
  assert.strictEqual(run(`indicadores(DIR, { esp: 'Psiquiatria' })`).acumulado.agendamentos, 1);
  // Mês do primeiro registro: tabelas por profissional e por especialidade somam o total do mês.
  const m0 = run(`indicadores(DIR, { mes: '${w0.slice(0, 7)}' })`);
  assert.strictEqual(m0.porProf.reduce((s, x) => s + x.agendamentos, 0), m0.total.agendamentos);
  assert.strictEqual(m0.porEsp.reduce((s, x) => s + x.agendamentos, 0), m0.total.agendamentos);
  assert.ok(m0.porProf.every(x => x.nome && x.esp));
  // Paciente inativado: as sessões futuras da vaga dele deixam de contar; o passado continua.
  run(`(() => { const p = ler('Pacientes').find(x => x['Prontuário'] === '0001'); atualizar('Pacientes', p._linha, { Status: 'Inativo' }); })()`);
  const ina = run(`indicadores(DIR, {})`).acumulado;
  assert.strictEqual(J([ina.realizados, ina.faltas, ina.cancelados]), J([2, 2, 1]));
  assert.ok(ina.pendentes < pend);
  // Meta: só Direção grava, só número, e regravar não duplica a linha.
  run(`salvarMeta(DIR, '2026-10', '300')`); run(`salvarMeta(DIR, '2026-10', '350')`);
  assert.strictEqual(run(`indicadores(DIR, { mes: '2026-10' })`).meta.valor, 350);
  assert.strictEqual(run(`ler('Config').filter(x => x.Chave === 'Meta 2026-10').length`), 1);
  erro(`salvarMeta(DIR, '2026-10', 'muito')`, /número inteiro/);
  erro(`salvarMeta(DIR, 'outubro', '10')`, /Mês inválido/);
  // Mês inválido vira o mês atual.
  assert.strictEqual(run(`indicadores(DIR, { mes: 'abc' })`).mes, h.slice(0, 7));
}
{ // Tela Pacientes lista todos (com inativos) em ordem de nome; inicio traz os turnos de cada profissional
  const { run, erro } = nova();
  const l = run(`listaPacientes(C)`);
  assert.strictEqual(J(l.map(p => p.nome + ':' + p.status)), J(['Ana Exemplo:Ativo', 'Bia Exemplo:Inativo', 'João Exemplo:Ativo']));
  erro(`listaPacientes({ cod: 'errado1', disp: 'q' })`, /inválido/);
  const carla = run(`inicio(C)`).profs.find(p => p.id === 'P01');
  assert.strictEqual(J(carla.horarios.map(h => h.dia + ' ' + h.turnos.map(t => t.inicio + '-' + t.fim + ' ' + t.sala).join())), J(['1 08:00-12:00 Sala 3', '3 13:00-17:00 Sala 3']));
}
{ // Agenda da semana: segunda a sábado do profissional, com feriado e telefone do responsável
  const { run, prox } = nova();
  const qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  const w = run(`agendaSemana(C, 'P01', '${qua}')`);
  assert.strictEqual(w.dias.length, 6);
  assert.strictEqual(run(`diaSemana('${w.inicio}')`), 1); // começa na segunda
  const q = w.dias.find(d => d.data === qua);
  assert.ok(q.atende && q.itens.length === 1 && q.itens[0].nome === 'Ana Exemplo' && q.itens[0].tel === '(00) 90000-0001');
  assert.ok(!w.dias.find(d => d.data === run(`somaDias('${w.inicio}', 1)`)).atende); // P01 não atende terça
  assert.strictEqual(run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01').itens[0].tel, '(00) 90000-0001');
}

{ // Cadastro de paciente pela tela
  const { run, erro } = nova();
  run(`salvarPaciente(C, ${J({ novo: true, pront: '0099', nome: 'Novo Paciente', responsavel: 'Mãe', telefone: '(21) 99999-0000' })})`);
  let p = run(`ler('Pacientes').find(x => x['Prontuário'] === '0099')`);
  assert.strictEqual(J([p.Nome, p.Status, p.Telefone]), J(['Novo Paciente', 'Ativo', '(21) 99999-0000']));
  erro(`salvarPaciente(C, ${J({ novo: true, pront: '0099', nome: 'Outro' })})`, /Já existe o prontuário 0099/);
  erro(`salvarPaciente(C, ${J({ novo: true, pront: '', nome: 'X' })})`, /prontuário/);
  erro(`salvarPaciente(C, ${J({ novo: true, pront: '0100', nome: '  ' })})`, /nome/);
  erro(`salvarPaciente(C, ${J({ pront: '7777', nome: 'X' })})`, /não encontrado/);
  run(`salvarPaciente(C, ${J({ pront: '0099', nome: 'Novo Paciente', status: 'Inativo', motivo: 'Alta' })})`);
  p = run(`ler('Pacientes').find(x => x['Prontuário'] === '0099')`);
  assert.strictEqual(J([p.Status, p['Motivo Inativação'], p['Data Inativação'] === run('hoje()')]), J(['Inativo', 'Alta', true]));
  run(`salvarPaciente(C, ${J({ pront: '0099', nome: 'Novo Paciente', status: 'Ativo' })})`);
  p = run(`ler('Pacientes').find(x => x['Prontuário'] === '0099')`);
  assert.strictEqual(J([p.Status, p['Data Inativação'], p['Motivo Inativação']]), J(['Ativo', '', '']));
  assert.strictEqual(run(`ler('Pacientes').filter(x => x['Prontuário'] === '0099').length`), 1);
}

{ // Cadastro de profissional pela tela: só Direção, validações, expediente substituído e aviso de vagas fora do horário
  const { run, erro, prox } = nova();
  const nv = { nome: 'Lia', esp: 'Musicoterapia', tipo: 'Horário', dur: 40, turnos: [{ dia: 2, inicio: '08:00', fim: '12:00', intIni: '10:00', intFim: '10:20', sala: 'Sala 9' }] };
  erro(`salvarProfissional(C, ${J(nv)})`, /Recepção não pode cadastrar ou mudar profissionais/);
  run(`gravar('Usuários', { Nome: 'Dir', 'Código': 'dir123', Perfil: 'Direção' }); var DIR = { cod: 'dir123', disp: 'y' };`);
  const r = run(`salvarProfissional(DIR, ${J(nv)})`);
  assert.strictEqual(r.id, 'P05'); // exemplos vão até P04
  const lia = run(`listaProfissionais(DIR)`).find(p => p.id === 'P05');
  assert.strictEqual(J([lia.esp, lia.tipo, lia.dur, lia.ativo, lia.turnos.length, lia.turnos[0].sala]), J(['Musicoterapia', 'Horário', '40', true, 1, 'Sala 9']));
  assert.ok(run(`inicio(DIR)`).especialidades.indexOf('Musicoterapia') >= 0); // especialidade nova aparece nas listas
  const bad = (mud, re) => erro(`salvarProfissional(DIR, ${J(Object.assign({}, nv, mud))})`, re);
  bad({ nome: '' }, /nome/); bad({ esp: '' }, /especialidade/); bad({ tipo: 'X' }, /tipo/); bad({ dur: 0 }, /Duração/);
  bad({ tipo: 'Ordem de chegada', max: 0 }, /Máximo/);
  bad({ turnos: [{ dia: 0, inicio: '08:00', fim: '12:00' }] }, /segunda a sábado/);
  bad({ turnos: [{ dia: 2, inicio: '12:00', fim: '08:00' }] }, /fim precisa ser depois/);
  bad({ turnos: [{ dia: 2, inicio: '08:00', fim: '12:00', intIni: '11:00', intFim: '' }] }, /intervalo/);
  bad({ turnos: [{ dia: 2, inicio: '08:00', fim: '12:00' }, { dia: 2, inicio: '11:00', fim: '15:00' }] }, /sobrepõem/);
  // Vaga numa terça às 08:00; depois o expediente passa a ser só quarta: a vaga é avisada, não apagada.
  const ter = prox(2);
  run(`agendar(DIR, ${J({ prof: 'P05', pront: '0001', data: ter, hora: '08:00', freq: 'Semanal' })})`);
  const r2 = run(`salvarProfissional(DIR, ${J(Object.assign({}, nv, { id: 'P05', turnos: [{ dia: 3, inicio: '13:00', fim: '17:00' }] }))})`);
  assert.strictEqual(J(r2.fora), J(['Ana Exemplo · Ter 08:00']));
  assert.strictEqual(J(run(`ler('Expediente').filter(e => e['Profissional ID'] === 'P05').map(e => e.Dia)`)), J(['Qua'])); // o expediente antigo saiu
  assert.ok(run(`ler('Expediente').some(e => e['Profissional ID'] === 'P01')`)); // dos outros não mexeu
  assert.strictEqual(run(`ler('Vagas Fixas').length`), 1);
  // Desativar e reativar
  run(`salvarProfissional(DIR, ${J(Object.assign({}, nv, { id: 'P05', ativo: false }))})`);
  assert.ok(!run(`inicio(DIR)`).profs.some(p => p.id === 'P05') && run(`listaProfissionais(DIR)`).find(p => p.id === 'P05').ativo === false);
}

{ // Recusas de acesso ficam registradas (sem o código digitado)
  const { run, erro } = nova();
  run(`var avisos = []; console.warn = m => avisos.push(m);`);
  erro(`inicio({ cod: 'segredo1', disp: 'aparelhoX' })`, /inválido/);
  const a = run('avisos');
  assert.ok(a.length === 1 && /código inválido \(aparelho aparelho, 1º erro\)/.test(a[0]) && !/segredo1/.test(a[0]), a[0]);
}
{ // Ao entregar a página, o Google corta a linha onde acha // fora de aspas duplas (quebrou o site inteiro na versão 21).
  // Só são seguros: linha que começa com // (comentário) e comentário no fim da linha depois de código fechado.
  const html = fs.readFileSync(__dirname + '/../Index.html', 'utf8'), js = html.slice(html.indexOf('<script>'), html.lastIndexOf('</script>'));
  const ruins = js.split('\n').map((l, i) => [i + 1, l]).filter(([, l]) => {
    if (/^\s*\/\//.test(l)) return false;
    const k = l.replace(/"[^"]*"/g, '""').indexOf('//');
    if (k < 0) return false;
    const antes = l.replace(/"[^"]*"/g, '""').slice(0, k);
    return (antes.split("'").length - 1) % 2 === 1 || (antes.split('`').length - 1) % 2 === 1; // // dentro de aspas simples ou crases
  });
  assert.strictEqual(J(ruins.map(([i, l]) => i + ': ' + l.trim().slice(0, 80))), '[]');
}
{ // Etapa 1a: configurar acrescenta colunas novas em aba antiga sem perder dados; perfis e permissões
  const { run, erro, prox } = nova();
  run(`(() => { const sh = SpreadsheetApp.getActive().getSheetByName('Pacientes'); sh.d[0] = sh.d[0].slice(0, 7); sh.d.forEach((l, i) => { if (i) sh.d[i] = l.slice(0, 7); }); })()`);
  run('configurar()');
  const cab = run(`SpreadsheetApp.getActive().getSheetByName('Pacientes').d[0]`);
  assert.strictEqual(J(cab.slice(6, 9)), J(['Motivo Inativação', 'CPF', 'CNS']));
  assert.strictEqual(run(`ler('Pacientes')[0].Nome`), 'Ana Exemplo');
  const U = (nome, cod, perfil, prof, ativo) => run(`gravar('Usuários', ${J({ Nome: nome, 'Código': cod, Perfil: perfil, 'Profissional ID': prof || '', Ativo: ativo || '' })})`);
  U('Geral', 'geral123', 'Geral'); U('Cep', 'cepress123', 'CEPRESS'); U('Adm', 'admin123', 'Administração'); U('Prof Carla', 'carla123', 'Profissional', 'P01');
  U('Sem vínculo', 'semv123', 'Profissional'); U('Saiu', 'saiu123', 'Recepção', '', 'Não');
  const A = c => ({ cod: c, disp: 'z' + c });
  const ac = c => `${J(A(c))}`;
  // agenda de exemplo: Ana com Carla (P01) e João com Paulo (P02)
  const qua = prox(3);
  run(`agendar(${ac('geral123')}, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  run(`agendar(${ac('geral123')}, ${J({ prof: 'P02', pront: '0002', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  const ini = run(`inicio(${ac('carla123')})`);
  assert.strictEqual(J([ini.perfil, ini.perfilNome, ini.prof, ini.pode.recepcao, ini.pode.gestao]), J(['Profissional', 'Profissional', 'P01', false, false]));
  assert.strictEqual(J(ini.pacientes.map(p => p.pront)), J(['0001'])); // só quem ela atende
  assert.strictEqual(J(run(`agenda(${ac('carla123')}, '${qua}')`).profs.map(p => p.id)), J(['P01']));
  assert.strictEqual(J(run(`listaPacientes(${ac('carla123')})`).map(p => p.pront)), J(['0001']));
  erro(`paciente(${ac('carla123')}, '0002')`, /não tem atendimentos com você/);
  erro(`agendar(${ac('carla123')}, ${J({ prof: 'P02', pront: '0001', data: qua, hora: '14:00', freq: 'Única' })})`, /sua própria agenda/);
  erro(`painel(${ac('carla123')})`, /Profissional não pode ver o painel/);
  erro(`fila(${ac('carla123')})`, /não pode ver a fila/);
  erro(`salvarPaciente(${ac('carla123')}, ${J({ novo: true, pront: '0009', nome: 'X' })})`, /não pode cadastrar/);
  erro(`agendaSemana(${ac('semv123')}, 'P01', '${qua}')`, /não está ligado a um profissional/);
  erro(`inicio(${ac('saiu123')})`, /acesso foi desativado/);
  assert.ok(run(`indicadores(${ac('cepress123')}, {})`).total);           // CEPRESS vê indicadores
  erro(`indicadores(${ac('admin123')}, {})`, /Só a Direção, a CEPRESS ou a Administradora geral/);
  assert.ok(run(`salvarProfissional(${ac('admin123')}, ${J({ nome: 'Nova', esp: 'Nutrição', tipo: 'Livre', turnos: [] })})`).id); // Administração cadastra profissional
  assert.strictEqual(run(`inicio(${ac('geral123')})`).pode.equipe, true);
}

{ // Etapa 1b: paciente com documentos, endereço, perfil clínico e suporte; ficha com totais e terapias
  const { run, erro, prox } = nova();
  const q = { novo: true, pront: '0060', nome: 'Caio', cpf: '529.982.247-25', cns: '798000000000006', nascimento: '2018-03-10', cep: '26100000', logradouro: 'Rua A',
    numero: '10', bairro: 'Centro', cidade: 'Belford Roxo', uf: 'rj', perfis: ['TEA', 'TDAH', 'OUTROS'], perfilOutros: 'seletividade', suporte: 'Outros', suporteOutros: 'em avaliação' };
  run(`salvarPaciente(C, ${J(q)})`);
  const f = run(`paciente(C, '0060')`);
  assert.strictEqual(J([f.cpf, f.cns, f.nascimento, f.cep, f.uf, f.perfis, f.perfilOutros, f.suporte, f.suporteOutros]),
    J(['529.982.247-25', '798000000000006', '2018-03-10', '26100-000', 'RJ', ['TEA', 'TDAH', 'OUTROS'], 'seletividade', 'Outros', 'em avaliação']));
  assert.strictEqual(f.idade, run(`idade('2018-03-10', hoje())`));
  const bad = (m, re) => erro(`salvarPaciente(C, ${J(Object.assign({}, q, { pront: '0061', cpf: '' }, m))})`, re);
  bad({ cpf: '111.111.111-11' }, /CPF inválido/); bad({ cns: '123' }, /CNS/); bad({ nascimento: '2999-01-01' }, /nascimento/); bad({ cep: '123' }, /CEP/);
  bad({ uf: 'Rio' }, /UF/); bad({ suporte: 'Nível 9' }, /suporte/);
  erro(`salvarPaciente(C, ${J(Object.assign({}, q, { pront: '0061' }))})`, /CPF já está no prontuário 0060/);
  assert.ok(run(`listaPacientes(C)`).some(p => p.cns === '798000000000006')); // busca por CNS/CPF
  // Totais e terapias: 1 presença e 1 falta com Carla, e vaga ativa com Paulo
  let w = run(`(() => { let d = somaDias(hoje(), -21); while (diaSemana(d) !== 3) d = somaDias(d, -1); return d; })()`);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0060', data: w, hora: '13:00', freq: 'Semanal' })})`);
  const v1 = run(`ler('Vagas Fixas')[0].ID`);
  run(`atualizarStatus(C, ${J({ vaga: v1, data: w, status: 'Realizado' })})`);
  run(`atualizarStatus(C, ${J({ vaga: v1, data: run(`somaDias('${w}', 7)`), status: 'Faltou', motivo: 'Justificada: Doença' })})`);
  run(`agendar(C, ${J({ prof: 'P02', pront: '0060', data: prox(3), hora: '14:00', freq: 'Quinzenal' })})`);
  const f2 = run(`paciente(C, '0060')`);
  assert.strictEqual(J([f2.resumo.realizados, f2.resumo.faltas]), J([1, 1]));
  assert.ok(f2.resumo.pendentes >= 1); // a quarta mais recente, sem registro
  assert.strictEqual(J(f2.terapias.map(t => t.esp + ':' + t.realizados + '/' + t.faltas)), J(['Fonoaudiologia:1/1', 'Terapia Ocupacional:0/0']));
  assert.ok(f2.historico.every(h => h.esp));
}

{ // Etapa 1c: mesmo paciente 2x no dia com o mesmo profissional; capacidade diária; Confirmado; 3 faltas; fechamento; registro
  const { run, erro, prox } = nova();
  const qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Única' })})`);
  erro(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '15:00', freq: 'Única' })})`, /Ana Exemplo já tem horário com Carla \(exemplo\) neste dia \(13:00\)/);
  run(`agendar(C, ${J({ prof: 'P02', pront: '0001', data: qua, hora: '16:00', freq: 'Única' })})`); // com outro profissional pode
  run(`(() => { const p = ler('Profissionais').find(x => x.ID === 'P01'); atualizar('Profissionais', p._linha, { 'Máx. por Dia': '1' }); })()`);
  erro(`agendar(C, ${J({ prof: 'P01', pront: '0002', data: qua, hora: '14:00', freq: 'Única' })})`, /Capacidade diária de Carla \(exemplo\) atingida/);
  const hs = run(`horariosLivres(C, ${J({ prof: 'P01', pront: '0002', data: qua, freq: 'Única' })})`);
  assert.ok(hs.every(h => !h.livre)); // a grade mostra tudo ocupado
  // Confirmado
  const a1 = run(`ler('Atendimentos').find(a => a['Horário'] === '13:00').ID`);
  run(`atualizarStatus(C, ${J({ atend: a1, data: qua, status: 'Confirmado' })})`);
  assert.strictEqual(run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01').itens[0].status, 'Confirmado');
  run(`gravar('Usuários', { Nome: 'Dir', 'Código': 'dir123', Perfil: 'Direção' }); gravar('Usuários', { Nome: 'Geral', 'Código': 'geral123', Perfil: 'Geral' });`);
  const D = `{ cod: 'dir123', disp: 'd' }`, G = `{ cod: 'geral123', disp: 'g' }`;
  assert.strictEqual(run(`indicadores(${D}, { mes: '${qua.slice(0, 7)}' })`).total.pendentes, 2);
  // 3 faltas seguidas sem justificativa (João com Paulo, quartas passadas)
  const w = n => run(`(() => { let d = somaDias(hoje(), -7 * ${n} - 1); while (diaSemana(d) !== 3) d = somaDias(d, -1); return d; })()`);
  run(`agendar(C, ${J({ prof: 'P02', pront: '0002', data: w(4), hora: '13:00', freq: 'Semanal' })})`);
  const vj = run(`ler('Vagas Fixas').find(v => v['Prontuário'] === '0002').ID`);
  [4, 3, 2, 1].forEach((n, i) => run(`atualizarStatus(C, ${J({ vaga: vj, data: w(n), status: i ? 'Faltou' : 'Realizado' })})`));
  let f3 = run(`painel(C)`).faltas3;
  assert.strictEqual(J(f3.map(x => x.nome + '·' + x.esp + '·' + x.datas.length)), J(['João Exemplo·Terapia Ocupacional·3']));
  run(`atualizarStatus(C, ${J({ vaga: vj, data: w(1), status: 'Faltou', motivo: 'Justificada: Doença' })})`);
  assert.strictEqual(run(`painel(C)`).faltas3.length, 0); // falta justificada quebra a regra
  run(`atualizarStatus(C, ${J({ vaga: vj, data: w(1), status: 'Faltou', motivo: '' })})`);
  run(`faltasAcao(C, ${J({ pront: '0002', prof: 'P02', acao: 'fila' })})`);
  assert.strictEqual(run(`ler('Vagas Fixas').find(v => v.ID === '${vj}').Status`), 'Encerrada');
  assert.ok(run(`fila(C)`).some(f => f.pront === '0002' && f.esp === 'Terapia Ocupacional' && /3 faltas/.test(f.obs)));
  run(`faltasAcao(C, ${J({ pront: '0002', prof: 'P02', acao: 'inativar' })})`);
  assert.strictEqual(run(`ler('Pacientes').find(p => p['Prontuário'] === '0002').Status`), 'Inativo');
  // Fechamento do mês
  const mesPassado = run(`somaDias(hoje().slice(0, 7) + '-01', -1).slice(0, 7)`);
  erro(`fecharMes(C, '${mesPassado}', true)`, /Só a Direção/);
  erro(`fecharMes(${D}, hoje().slice(0, 7), true)`, /já terminou/);
  run(`fecharMes(${D}, '${mesPassado}', true)`);
  assert.ok(/por Dir/.test(run(`indicadores(${D}, { mes: '${mesPassado}' })`).fechado));
  const diaPassado = mesPassado + '-15';
  erro(`agendar(C, ${J({ prof: 'P02', pront: '0001', data: diaPassado, hora: '13:00', freq: 'Única' })})`, /está fechado/);
  erro(`atualizarStatus(C, ${J({ vaga: vj, data: diaPassado, status: 'Realizado' })})`, /está fechado/);
  erro(`fecharMes(${D}, '${mesPassado}', false)`, /Só a Administradora geral reabre/);
  run(`fecharMes(${G}, '${mesPassado}', false)`);
  assert.strictEqual(run(`indicadores(${D}, { mes: '${mesPassado}' })`).fechado, '');
  // Histórico de alterações
  const reg = run(`historicoAlteracoes(${D})`);
  ['Agendou', 'Marcou Confirmado', 'Marcou Faltou', '3 faltas: vaga encerrada e paciente na fila', '3 faltas: paciente inativado', 'Fechou o mês', 'Reabriu o mês']
    .forEach(a => assert.ok(reg.some(r => r.acao === a), 'falta no registro: ' + a));
  assert.ok(reg[0].acao === 'Reabriu o mês' && reg[0].quem === 'Geral' && reg[0].perfil === 'Administradora geral');
  erro(`historicoAlteracoes(C)`, /Só a Direção/);
}
{ // 2a. Equipe e acessos (só Administradora geral)
  const { run, erro } = nova();
  run(`gravar('Usuários', { Nome: 'Geral', 'Código': 'geral123', Perfil: 'Geral' })`);
  const G = `{ cod: 'geral123', disp: 'g' }`;
  erro(`listaEquipe(C)`, /não pode gerenciar a equipe/);
  assert.ok(run(`listaEquipe(${G})`).some(u => u.nome === 'Geral' && u.eu));
  erro(`salvarUsuario(${G}, ${J({ nome: 'Carla', perfil: 'Profissional' })})`, /escolha a qual profissional/);
  erro(`salvarUsuario(${G}, ${J({ nome: 'Carla', perfil: 'Profissional', prof: 'P01', email: 'x@' })})`, /E-mail inválido/);
  const r = run(`salvarUsuario(${G}, ${J({ nome: 'Carla', perfil: 'Profissional', prof: 'P01', funcao: 'Fonoaudióloga' })})`);
  assert.ok(/^[0-9a-f]{10}$/.test(r.codigo));
  assert.strictEqual(run(`usuario({ cod: '${r.codigo}', disp: 'c' })`), 'Carla');
  erro(`salvarUsuario(${G}, ${J({ nome: 'carla', perfil: 'Recepção' })})`, /Já existe alguém/);
  erro(`salvarUsuario(${G}, ${J({ nome: 'Outra', perfil: 'Profissional', prof: 'P01' })})`, /já está ligado/);
  const carla = run(`listaEquipe(${G})`).find(u => u.nome === 'Carla');
  assert.strictEqual(carla.profNome, 'Carla (exemplo) · Fonoaudiologia');
  assert.ok(!('codigo' in carla) && !JSON.stringify(run(`listaEquipe(${G})`)).includes(r.codigo)); // o código nunca volta na lista
  assert.strictEqual(run(`salvarUsuario(${G}, ${J({ linha: carla.linha, nomeAntigo: 'Carla', nome: 'Carla', perfil: 'Profissional', prof: 'P01', ativo: false })})`).codigo, '');
  erro(`usuario({ cod: '${r.codigo}', disp: 'c' })`, /desativado/);
  const nov = run(`salvarUsuario(${G}, ${J({ linha: carla.linha, nomeAntigo: 'Carla', nome: 'Carla', perfil: 'Profissional', prof: 'P01', novoCodigo: true })})`).codigo;
  erro(`usuario({ cod: '${r.codigo}', disp: 'c2' })`, /inválido/);
  assert.strictEqual(run(`usuario({ cod: '${nov}', disp: 'c3' })`), 'Carla');
  const eu = run(`listaEquipe(${G})`).find(u => u.eu);
  erro(`salvarUsuario(${G}, ${J({ linha: eu.linha, nomeAntigo: 'Geral', nome: 'Geral', perfil: 'Direção' })})`, /pelo menos uma Administradora geral/);
  erro(`salvarUsuario(${G}, ${J({ linha: eu.linha, nomeAntigo: 'Outro', nome: 'Geral', perfil: 'Geral' })})`, /mudou na planilha/);
  assert.ok(run(`historicoAlteracoes(${G})`).some(x => x.acao === 'Criou acesso') );
  assert.ok(!JSON.stringify(run(`historicoAlteracoes(${G})`)).includes(nov)); // código fora do registro
}
{ // 2b. Enfermagem, evoluções e compartilhamento
  const { run, erro, prox } = nova();
  const ontem = run(`(() => { let d = somaDias(hoje(), -1); while (diaSemana(d) !== 3) d = somaDias(d, -1); return d; })()`), h = run(`hoje()`); // última quarta
  run(`gravar('Profissionais', { ID: 'P05', Nome: 'Enf. Lia', Especialidade: 'Enfermagem', 'Tipo de Agenda': 'Livre', Ativo: 'Sim' })`);
  run(`gravar('Usuários', { Nome: 'Lia', 'Código': 'lia1234', Perfil: 'Profissional', 'Profissional ID': 'P05' })`);
  run(`gravar('Usuários', { Nome: 'Carla', 'Código': 'carla123', Perfil: 'Profissional', 'Profissional ID': 'P01' })`);
  run(`gravar('Usuários', { Nome: 'Paulo', 'Código': 'paulo123', Perfil: 'Profissional', 'Profissional ID': 'P02' })`);
  run(`gravar('Usuários', { Nome: 'Geral', 'Código': 'geral123', Perfil: 'Geral' })`);
  const L = `{ cod: 'lia1234', disp: 'l' }`, CA = `{ cod: 'carla123', disp: 'ca' }`, PA = `{ cod: 'paulo123', disp: 'pa' }`, G = `{ cod: 'geral123', disp: 'g' }`;
  assert.strictEqual(run(`inicio(${L})`).pode.enfermagem, true);
  assert.strictEqual(run(`inicio(${CA})`).pode.enfermagem, false);
  assert.strictEqual(run(`inicio(C)`).pode.evolucoes, false); // Recepção não evolui
  // Ana tem fono (P01) e TO (P02) no mesmo dia
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: ontem, hora: '13:00', freq: 'Única' })})`);
  run(`agendar(C, ${J({ prof: 'P02', pront: '0001', data: ontem, hora: '14:00', freq: 'Única' })})`);
  const d = run(`enfermagemDia(C, '${ontem}')`);
  assert.strictEqual(d.lista.length, 1); // um paciente, uma vez só
  assert.strictEqual(d.lista[0].terapias.length, 2);
  assert.strictEqual(d.podeRegistrar, false); // a recepção só vê
  erro(`registrarEnfermagem(C, ${J({ data: ontem, pront: '0001' })})`, /Só a conta ligada a um profissional de atendimento livre/);
  erro(`enfermagemDia(${CA}, '${ontem}')`, /atendimento livre é para/);
  erro(`registrarEnfermagem(${L}, ${J({ data: run('somaDias(hoje(), 1)'), pront: '0001' })})`, /dia futuro/);
  const n1 = run(`registrarEnfermagem(${L}, ${J({ data: ontem, pront: '0001' })})`);
  erro(`registrarEnfermagem(${L}, ${J({ data: ontem, pront: '0001' })})`, /já foi registrada/);
  run(`registrarEnfermagem(${L}, ${J({ data: ontem, pront: '0002', livre: true })})`); // atendimento livre
  const d2 = run(`enfermagemDia(${L}, '${ontem}')`);
  assert.ok(d2.podeRegistrar && d2.lista.length === 2 && d2.lista.every(x => x.passou));
  assert.strictEqual(d2.lista.find(x => x.pront === '0002').passou.tipo, 'Atendimento livre');
  // Evolução: exige presença e o próprio profissional
  const a1 = run(`ler('Atendimentos').find(a => a['Profissional ID'] === 'P01').ID`);
  erro(`abrirEvolucao(${CA}, '${a1}')`, /Registre a presença/);
  run(`atualizarStatus(C, ${J({ atend: a1, data: ontem, status: 'Realizado' })})`);
  erro(`abrirEvolucao(${PA}, '${a1}')`, /de outro profissional/);
  erro(`abrirEvolucao(C, '${a1}')`, /ligada a um profissional/);
  const m = run(`abrirEvolucao(${CA}, '${a1}')`);
  assert.ok(m.modelo.condutas.includes('Comunicação funcional') && m.paciente === 'Ana Exemplo');
  assert.strictEqual(J(run(`pendentesEvolucao(${CA})`).map(x => x.atend)), J([a1]));
  erro(`salvarEvolucao(${CA}, ${J({ atend: a1, conteudo: {} })})`, /Preencha pelo menos/);
  erro(`salvarEvolucao(${CA}, ${J({ atend: a1, conteudo: { objetivo: 'x', duracao: '0' } })})`, /Duração/);
  const e1 = run(`salvarEvolucao(${CA}, ${J({ atend: a1, conteudo: { objetivo: 'Ampliar vocabulário', duracao: '40', condutas: ['Comunicação funcional', 'Inventado'], chegada: ['Calmo(a)'] } })})`);
  assert.strictEqual(run(`pendentesEvolucao(${CA})`).length, 0);
  erro(`salvarEvolucao(${CA}, ${J({ atend: a1, conteudo: { objetivo: 'de novo' } })})`, /já tem evolução/);
  erro(`atualizarStatus(C, ${J({ atend: a1, data: ontem, status: 'Cancelado' })})`, /já tem evolução/);
  erro(`remarcar(C, ${J({ atend: a1, data: ontem, novaData: h, novaHora: '15:00' })})`, /já tem evolução/);
  const ev = run(`evolucoesPaciente(${CA}, '0001')`);
  assert.strictEqual(J(ev[0].conteudo.condutas), J(['Comunicação funcional'])); // item fora do modelo é descartado
  // Complemento: só o autor; a original não muda
  erro(`salvarEvolucao(${PA}, ${J({ ref: e1, texto: 'x' })})`, /Só o autor/);
  run(`salvarEvolucao(${CA}, ${J({ ref: e1, texto: 'Mãe relatou melhora em casa.' })})`);
  const ev2 = run(`evolucoesPaciente(${CA}, '0001')`);
  assert.ok(ev2.length === 1 && ev2[0].complementos.length === 1 && ev2[0].conteudo.objetivo === 'Ampliar vocabulário');
  // Sigilo: Paulo (outro profissional) não vê; Geral vê com registro
  assert.strictEqual(run(`evolucoesPaciente(${PA}, '0001')`).length, 0);
  assert.strictEqual(run(`evolucoesPaciente(${G}, '0001')`)[0].porque, 'geral');
  assert.ok(run(`historicoAlteracoes(${G})`).some(x => x.acao === 'Consultou evoluções' && /acesso geral/.test(x.detalhe)));
  // Compartilhar com Paulo (só leitura) e revogar
  erro(`compartilhar(${PA}, ${J({ alcance: 'Evolução', ref: e1, destinatario: 'Carla' })})`, /só pode compartilhar as suas/);
  erro(`compartilhar(${CA}, ${J({ alcance: 'Evolução', ref: e1, destinatario: 'Maria' })})`, /Escolha para quem/); // recepção não recebe
  run(`compartilhar(${CA}, ${J({ alcance: 'Evolução', ref: e1, destinatario: 'Paulo' })})`);
  erro(`compartilhar(${CA}, ${J({ alcance: 'Evolução', ref: e1, destinatario: 'Paulo' })})`, /já existe/);
  const vis = run(`evolucoesPaciente(${PA}, '0001')`);
  assert.ok(vis.length === 1 && vis[0].porque === 'compartilhada' && vis[0].complementos.length === 1);
  erro(`salvarEvolucao(${PA}, ${J({ ref: e1, texto: 'x' })})`, /Só o autor/); // só leitura
  assert.strictEqual(J(run(`compartilhamentos(${PA})`).recebi), J([{ pront: '0001', paciente: 'Ana Exemplo' }]));
  const c1 = run(`compartilhamentos(${CA})`).dei[0];
  assert.ok(/Evolução de Ana Exemplo/.test(c1.descricao));
  erro(`revogar(${PA}, '${c1.id}')`, /não encontrada/);
  run(`revogar(${CA}, '${c1.id}')`);
  assert.strictEqual(run(`evolucoesPaciente(${PA}, '0001')`).length, 0);
  run(`compartilhar(${CA}, ${J({ alcance: 'Geral', destinatario: 'Paulo' })})`);
  assert.strictEqual(run(`evolucoesPaciente(${PA}, '0001')`).length, 1);
  // Rita não atende Ana: lê a evolução compartilhada, mas a ficha continua fechada
  run(`gravar('Usuários', { Nome: 'Rita', 'Código': 'rita1234', Perfil: 'Profissional', 'Profissional ID': 'P03' })`);
  run(`compartilhar(${CA}, ${J({ alcance: 'Paciente', ref: '0001', destinatario: 'Rita' })})`);
  assert.strictEqual(run(`evolucoesPaciente({ cod: 'rita1234', disp: 'r' }, '0001')`).length, 1);
  erro(`paciente({ cod: 'rita1234', disp: 'r' }, '0001')`, /não tem atendimentos com você/);
  // Enfermagem evolui a passagem; aí a passagem não pode ser desfeita
  assert.strictEqual(run(`pendentesEvolucao(${L})`).length, 2);
  assert.ok(run(`abrirEvolucao(${L}, '${n1}')`).modelo.condutas.includes('Sinais vitais aferidos'));
  run(`salvarEvolucao(${L}, ${J({ atend: n1, conteudo: { obs: 'Sem queixas.' } })})`);
  assert.ok(run(`enfermagemDia(${L}, '${ontem}')`).lista.find(x => x.pront === '0001').passou.evolucao);
  erro(`registrarEnfermagem(${L}, ${J({ data: ontem, pront: '0001', remover: true })})`, /já tem evolução/);
  run(`registrarEnfermagem(${L}, ${J({ data: ontem, pront: '0002', remover: true })})`);
  assert.strictEqual(run(`enfermagemDia(${L}, '${ontem}')`).lista.length, 1);
  // Modelo genérico para especialidade sem lista própria
  assert.ok(run(`modeloEvolucao('Arteterapia')`).condutas.includes('Intervenção'));
  assert.ok(run(`modeloEvolucao('serviço social')`).condutas.includes('Benefícios'));
  // Cópia de segurança semanal
  assert.ok(run(`gatilhos.some(t => t.getHandlerFunction() === 'copiaDeSeguranca')`));
  assert.ok(/cópia de segurança/.test(run(`copiaDeSeguranca()`)));
}
{ // 3a. Conflito de sala e bloqueios pela tela
  const { run, erro, prox } = nova();
  run(`gravar('Usuários', { Nome: 'Dir', 'Código': 'dir123', Perfil: 'Direção' })`);
  const D = `{ cod: 'dir123', disp: 'd' }`;
  // P01 usa a Sala 3 na quarta 13–17: outro profissional não pode pegar a mesma sala no mesmo horário
  erro(`salvarProfissional(${D}, ${J({ nome: 'Novo', esp: 'Psicologia', tipo: 'Horário', dur: 30, turnos: [{ dia: 3, inicio: '16:00', fim: '18:00', sala: 'sala 3' }] })})`, /Sala 3 já está com Carla/);
  run(`salvarProfissional(${D}, ${J({ nome: 'Novo', esp: 'Psicologia', tipo: 'Horário', dur: 30, turnos: [{ dia: 3, inicio: '17:00', fim: '18:00', sala: 'Sala 3' }] })})`); // depois das 17h pode
  run(`salvarProfissional(${D}, ${J({ nome: 'Outro', esp: 'Psicologia', tipo: 'Horário', dur: 30, ativo: false, turnos: [{ dia: 3, inicio: '13:00', fim: '14:00', sala: 'Sala 3' }] })})`); // inativo não ocupa
  erro(`listaBloqueios(C)`, /não pode cadastrar bloqueios/);
  const qua = prox(3);
  run(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`);
  erro(`salvarBloqueio(${D}, ${J({ ini: qua, fim: '2000-01-01', motivo: 'x' })})`, /fim não pode ser antes/);
  erro(`salvarBloqueio(${D}, ${J({ ini: qua, prof: 'P01' })})`, /Informe o motivo/);
  run(`salvarBloqueio(${D}, ${J({ ini: qua, prof: 'P01', motivo: 'Férias' })})`);
  assert.strictEqual(run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01').itens[0].status, 'Sem atendimento');
  erro(`agendar(C, ${J({ prof: 'P01', pront: '0002', data: qua, hora: '14:00', freq: 'Única' })})`, /Férias/);
  const b = run(`listaBloqueios(${D})`).find(x => x.motivo === 'Férias');
  assert.ok(b && b.profNome === 'Carla (exemplo)');
  erro(`removerBloqueio(${D}, ${J({ linha: b.linha, ini: b.ini, motivo: 'Outro' })})`, /mudou na planilha/);
  run(`removerBloqueio(${D}, ${J(b)})`);
  assert.strictEqual(run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01').itens[0].status, 'Agendado');
}
{ // 4. Ajustes de 28/10: cadastro, faltas, reposição, regras de marcação, painel e relatórios
  const { run, erro, prox, itens } = nova();
  const h = run('hoje()'), qua = prox(3);
  const w = n => run(`(() => { let d = somaDias(hoje(), -7 * ${n} - 1); while (diaSemana(d) !== 3) d = somaDias(d, -1); return d; })()`);
  // Cadastro: prontuário 001/2025, sexo e abertura
  run(`salvarPaciente(C, ${J({ novo: true, pront: '001/2025', nome: 'Liam', sexo: 'Masculino', nascimento: '2019-05-01', abertura: '2025-02-10' })})`);
  const liam = run(`paciente(C, '001/2025')`);
  assert.ok(liam.sexo === 'Masculino' && liam.abertura === '2025-02-10');
  erro(`salvarPaciente(C, ${J({ novo: true, pront: '002/2025', nome: 'X', sexo: 'Outro' })})`, /Sexo inválido/);
  run(`salvarPaciente(C, ${J({ novo: true, pront: '002/2025', nome: 'Maya', sexo: 'Feminino', perfis: ['TAG', 'NÃO SE APLICA'], suporte: 'N/A — Não se aplica' })})`);
  const maya = run(`paciente(C, '002/2025')`);
  assert.ok(maya.abertura === h && maya.perfis.join() === 'TAG,NÃO SE APLICA' && maya.suporte === 'N/A — Não se aplica');
  // Mesma terapia com outro terapeuta (recorrente) não pode; fila bloqueada para quem já tem vaga
  run(`gravar('Profissionais', { ID: 'P09', Nome: 'Outra Fono', Especialidade: 'Fonoaudiologia', 'Tipo de Agenda': 'Horário', 'Duração (min)': '30', Ativo: 'Sim' })`);
  run(`gravar('Expediente', { 'Profissional ID': 'P09', Dia: 'Qua', 'Início': '08:00', Fim: '12:00' })`);
  run(`agendar(C, ${J({ prof: 'P01', pront: '001/2025', data: qua, hora: '13:00', freq: 'Quinzenal' })})`);
  erro(`agendar(C, ${J({ prof: 'P09', pront: '001/2025', data: qua, hora: '09:00', freq: 'Semanal' })})`, /já faz Fonoaudiologia com Carla/);
  run(`agendar(C, ${J({ prof: 'P09', pront: '001/2025', data: qua, hora: '09:00', freq: 'Única' })})`); // avulsa (reposição/substituição) pode
  erro(`addFila(C, ${J({ pront: '001/2025', esp: 'Fonoaudiologia' })})`, /já está agendado/);
  // Quinzenal: a cada 14 dias, mesmo dia e horário
  const q14 = run(`somaDias('${qua}', 14)`), q7 = run(`somaDias('${qua}', 7)`);
  assert.ok(itens(q14, 'P01').includes('13:00 001/2025 Agendado') && !itens(q7, 'P01').some(x => /001\/2025/.test(x)));
  // Faltas: justificar depois e alerta na 3ª não justificada
  run(`agendar(C, ${J({ prof: 'P02', pront: '0002', data: w(4), hora: '13:00', freq: 'Semanal' })})`);
  const vj = run(`ler('Vagas Fixas').find(v => v['Prontuário'] === '0002').ID`);
  [3, 2, 1].forEach(n => run(`atualizarStatus(C, ${J({ vaga: vj, data: w(n), status: 'Faltou' })})`));
  let fl = run(`faltasLista(C)`);
  assert.ok(fl.alertas.length === 1 && fl.alertas[0].n === 3 && fl.lista.filter(x => x.pront === '0002').every(x => !x.justificada));
  assert.ok(/Terapia Ocupacional/.test(J(run(`paciente(C, '0002')`).alertas)));
  run(`justificarFalta(C, ${J({ atend: fl.lista.find(x => x.pront === '0002').atend, justificada: true, motivo: 'Doença', obs: 'mãe trouxe declaração' })})`);
  fl = run(`faltasLista(C)`);
  assert.ok(fl.alertas.length === 0 && fl.lista.filter(x => x.justificada).length === 1 && /Justificada: Doença — mãe/.test(fl.lista.find(x => x.justificada).motivo));
  assert.strictEqual(run(`painel(C)`).faltas3.length, 0);
  // Serviço Social (atendimento livre) justifica e registra passagem; profissional comum não
  run(`gravar('Usuários', { Nome: 'Assistente', 'Código': 'ass1234', Perfil: 'Profissional', 'Profissional ID': 'P04' })`);
  run(`gravar('Usuários', { Nome: 'Carla', 'Código': 'carla123', Perfil: 'Profissional', 'Profissional ID': 'P01' })`);
  const AS = `{ cod: 'ass1234', disp: 'as' }`, CA = `{ cod: 'carla123', disp: 'ca' }`;
  assert.ok(run(`inicio(${AS})`).pode.faltas && run(`inicio(${AS})`).pode.enfermagem && !run(`inicio(${CA})`).pode.faltas);
  run(`justificarFalta(${AS}, ${J({ atend: fl.lista.find(x => !x.justificada).atend, justificada: true, motivo: 'Problema de transporte' })})`);
  erro(`justificarFalta(${CA}, ${J({ atend: fl.lista[0].atend, justificada: false })})`, /Só a recepção/);
  const sd = run(`enfermagemDia(${AS}, '${w(1)}')`);
  assert.ok(sd.esp === 'Serviço Social' && sd.podeRegistrar);
  run(`registrarEnfermagem(${AS}, ${J({ data: w(1), pront: '0001', livre: true })})`);
  assert.strictEqual(run(`enfermagemDia(C, '${w(1)}', 'P04')`).lista.filter(x => x.passou).length, 1); // recepção consulta o Serviço Social
  // Reposição: sessão que cai em feriado vai para a lista e é remarcada como reposição
  run(`salvarBloqueio(C, ${J({ ini: q14, motivo: 'Ponto facultativo' })})`.replace('C,', `{ cod: 'dir123', disp: 'd' },`).replace("salvarBloqueio({", "(gravar('Usuários', { Nome: 'Dir', 'Código': 'dir123', Perfil: 'Direção' }), salvarBloqueio)({"));
  const rp = run(`reposicoes(C)`).find(x => x.pront === '001/2025');
  assert.ok(rp && rp.data === q14 && rp.motivo === 'Ponto facultativo');
  erro(`remarcar(C, ${J({ vaga: rp.vaga, data: q14, novaData: q14, novaHora: '15:00', reposicao: true })})`, /outro dia/);
  run(`remarcar(C, ${J({ vaga: rp.vaga, data: q14, novaData: run(`somaDias('${q14}', 7)`), novaHora: '14:00', reposicao: true })})`);
  assert.ok(!run(`reposicoes(C)`).some(x => x.pront === '001/2025'));
  assert.ok(run(`ler('Atendimentos')`).some(a => a['Prontuário'] === '001/2025' && a.Motivo === 'Reposição de ' + q14));
  // Paciente inativo com agendamentos: só alerta (não cancela)
  run(`salvarPaciente(C, ${J({ pront: '001/2025', nome: 'Liam', sexo: 'Masculino', status: 'Inativo', motivo: 'Alta' })})`);
  const pn = run(`painel(C)`);
  assert.ok(pn.inativosComVaga.some(x => x.pront === '001/2025' && x.vaga) && pn.inativosComVaga.some(x => x.pront === '001/2025' && x.atend));
  assert.ok(pn.listas.inativos.some(x => x.pront === '001/2025') && pn.listas.ativos.some(x => x.pront === '002/2025'));
  // Relatórios: período livre e perfil dos atendidos
  run(`atualizarStatus(C, ${J({ vaga: vj, data: w(4), status: 'Realizado' })})`);
  run(`(() => { const p = ler('Pacientes').find(x => x['Prontuário'] === '0002'); atualizar('Pacientes', p._linha, { Sexo: 'Masculino', Nascimento: '2018-01-01' }); })()`);
  const ind = run(`indicadores({ cod: 'dir123', disp: 'd' }, ${J({ de: w(4), ate: w(1) })})`);
  assert.ok(ind.periodo && ind.total.realizados === 1 && ind.total.faltas === 3 && ind.perfilPac.sexo.Masculino === 1 && ind.perfilPac.idade['7 a 12 anos'] === 1, J(ind.total));
}
{ // 5. Psiquiatria/neuropediatria: retorno a cada 4 semanas, mesmo dia da semana, na ordem de chegada
  const { run, erro, prox } = nova(), ter = prox(2);
  const lista = d => run(`agenda(C, '${d}')`).profs.find(p => p.id === 'P03').itens.map(i => i.pront + ' ' + i.status);
  const v = run(`agendar(C, ${J({ prof: 'P03', pront: '0001', data: ter, freq: 'Mensal' })})`);
  assert.ok(/^V/.test(v));
  const t4 = run(`somaDias('${ter}', 28)`), t2 = run(`somaDias('${ter}', 14)`);
  assert.strictEqual(J(lista(ter)), J(['0001 Agendado'])); assert.strictEqual(J(lista(t4)), J(['0001 Agendado'])); assert.strictEqual(J(lista(t2)), J([]));
  erro(`agendar(C, ${J({ prof: 'P03', pront: '0001', data: t4 })})`, /já está na lista/);
  // limite do dia conta os retornos
  run(`(() => { const p = ler('Profissionais').find(x => x.ID === 'P03'); atualizar('Profissionais', p._linha, { 'Máx. por Dia': '1' }); })()`);
  erro(`agendar(C, ${J({ prof: 'P03', pront: '0002', data: t4 })})`, /LOTADO/);
  erro(`agendar(C, ${J({ prof: 'P03', pront: '0002', data: t2, freq: 'Quinzenal' })})`, /LOTADO em/);
  // presença, passar para outro dia e encerrar a recorrência
  run(`atualizarStatus(C, ${J({ vaga: v, data: ter, status: 'Realizado' })})`);
  assert.strictEqual(J(lista(ter)), J(['0001 Realizado']));
  run(`remarcar(C, ${J({ vaga: v, data: t4, novaData: run(`somaDias('${t4}', 7)`) })})`);
  assert.strictEqual(J(lista(t4)), J(['0001 Remarcado']));
  assert.strictEqual(J(lista(run(`somaDias('${t4}', 7)`))), J(['0001 Agendado']));
  run(`encerrarVaga(C, '${v}', '${t4}')`);
  assert.strictEqual(J(lista(run(`somaDias('${ter}', 56)`))), J([]));
}
console.log('ok');
