// Fluxo completo do Code.gs com a planilha simulada: node teste/servidor.js
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const ctx = vm.createContext({ console, crypto, Intl });
for (const f of ['teste/simulador.js', 'Logica.gs', 'Code.gs']) vm.runInContext(fs.readFileSync(__dirname + '/../' + f, 'utf8'), ctx, { filename: f });
const run = code => vm.runInContext(code, ctx);
const erro = (code, re) => assert.throws(() => run(code), e => re.test(e.message), code);
const J = o => JSON.stringify(o);
const igual = (a, b) => assert.strictEqual(J(a), J(b));

run('configurar()');
run('configurar()'); // rodar de novo não duplica nada
assert.strictEqual(run("ler('Profissionais').length"), 4);
assert.strictEqual(run("ler('Bloqueios').length"), 27); // feriados nacionais até 2028
run("SpreadsheetApp.getActive().getSheetByName('Bloqueios').d.length = 1"); // tira os feriados para as datas do teste não caírem neles

// Código de acesso: sem usuários ninguém entra; código curto não vale; depois de 20 erros, bloqueia.
run(`var C = { cod: 'abc123', disp: 'aparelho-1' }`);
erro('inicio(C)', /^ACESSO/);
run(`gravar('Usuários', { Nome: 'Maria', 'Código': 'abc123' }); gravar('Usuários', { Nome: 'Curto', 'Código': '12' })`);
erro(`inicio({ cod: '12', disp: 'aparelho-1' })`, /^ACESSO/);
assert.strictEqual(run('inicio(C)').usuario, 'Maria');

// Próxima quarta (Carla atende 13–17h, sala 3) e próxima terça (psiquiatra, ordem de chegada).
const hoje = run('hoje()');
const prox = dia => { let d = run(`somaDiasT = (i,n) => new Date(Date.parse(i+'T12:00:00Z')+n*864e5).toISOString().slice(0,10)`) && hoje;
  do d = run(`somaDiasT('${d}', 1)`); while (run(`diaSemana('${d}')`) !== dia); return d; };
const qua = prox(3), qua2 = run(`somaDiasT('${qua}', 7)`), ter = prox(2);

const ag = q => run(`agendar(C, ${J(q)})`);
assert.match(ag({ prof: 'P01', pront: '0001', data: qua, hora: '13:00', freq: 'Quinzenal' }), /^V/);
assert.match(run(`horariosLivres(C, ${J({ prof: 'P01', pront: '0002', data: qua, freq: 'Quinzenal' })})`).find(h => h.hora === '13:00').aviso, /em dupla com Ana/); // dupla: pode, com aviso
assert.match(ag({ prof: 'P01', pront: '0002', data: qua2, hora: '13:00', freq: 'Quinzenal' }), /^V/); // semana alternada
erro(`agendar(C, ${J({ prof: 'P02', pront: '0001', data: qua, hora: '13:00', freq: 'Semanal' })})`, /Ana Exemplo já está/);
erro(`agendar(C, ${J({ prof: 'P01', pront: '0003', data: qua, hora: '14:00', freq: 'Única' })})`, /inativo/);
erro(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: ter, hora: '13:00', freq: 'Única' })})`, /não atende Ter/);
erro(`agendar(C, ${J({ prof: 'P04', pront: '0001', data: qua, hora: '13:00', freq: 'Única' })})`, /livre/);

let carla = run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01');
igual(carla.itens.map(i => i.hora + ' ' + i.pront), ['13:00 0001']);
const livres = run(`horariosLivres(C, ${J({ prof: 'P01', pront: '0002', data: qua, freq: 'Única' })})`);
assert.strictEqual(livres.find(h => h.hora === '13:00').livre, true); // em dupla
assert.match(livres.find(h => h.hora === '13:00').aviso, /dupla/);
assert.strictEqual(livres.find(h => h.hora === '13:30').livre, true);

// Remarcar a sessão de quarta libera 13:00 só naquele dia.
const vaga = carla.itens[0].vaga;
run(`remarcar(C, ${J({ vaga, data: qua, novaData: qua, novaHora: '14:00', motivo: 'pedido da mãe' })})`);
carla = run(`agenda(C, '${qua}')`).profs.find(p => p.id === 'P01');
igual(carla.itens.map(i => i.hora + ' ' + i.status), ['13:00 Remarcado', '14:00 Agendado']);
assert.match(ag({ prof: 'P01', pront: '0002', data: qua, hora: '13:00', freq: 'Única' }), /^A/);
assert.match(ag({ prof: 'P01', pront: '0002', data: run(`somaDiasT('${qua}', 14)`), hora: '13:00', freq: 'Única' }), /^A/); // em dupla com a série da Ana

// Status de uma sessão da série (cria o registro) e depois desfazer (atualiza o mesmo registro).
const q4 = run(`somaDiasT('${qua}', 14)`);
run(`atualizarStatus(C, ${J({ vaga, data: q4, status: 'Faltou', motivo: 'doente' })})`);
run(`atualizarStatus(C, ${J({ vaga, data: q4, status: 'Agendado' })})`);
assert.strictEqual(run(`ler('Atendimentos')`).filter(a => a['Vaga ID'] === vaga && a.Data === q4).length, 1);

// Ordem de chegada: lista do dia, sem repetir paciente, com limite.
ag({ prof: 'P03', pront: '0001', data: ter });
erro(`agendar(C, ${J({ prof: 'P03', pront: '0001', data: ter })})`, /já está na lista/);
run(`atualizar('Profissionais', 4, { 'Máx. por Dia': '1' })`);
erro(`agendar(C, ${J({ prof: 'P03', pront: '0002', data: ter })})`, /LOTADO: .* 1\/1/);
// Passar para outro dia (ordem de chegada): sai da lista de hoje e entra no fim da lista do novo dia, com as mesmas regras.
const ter2 = run(`somaDiasT('${ter}', 7)`);
const oc1 = run(`agenda(C, '${ter}')`).profs.find(x => x.id === 'P03').itens[0];
erro(`remarcar(C, ${J({ atend: oc1.atend, data: ter, novaData: qua })})`, /não atende Qua/);
erro(`remarcar(C, ${J({ atend: oc1.atend, data: ter, novaData: ter })})`, /já está na lista/);
run(`remarcar(C, ${J({ atend: oc1.atend, data: ter, novaData: ter2, motivo: 'viagem' })})`);
assert.strictEqual(run(`agenda(C, '${ter}')`).profs.find(x => x.id === 'P03').itens[0].status, 'Remarcado');
igual(run(`agenda(C, '${ter2}')`).profs.find(x => x.id === 'P03').itens.map(i => i.pront + ' ' + i.status), ['0001 Agendado']);
ag({ prof: 'P03', pront: '0002', data: ter }); // a vaga de hoje liberou
const oc2 = run(`agenda(C, '${ter}')`).profs.find(x => x.id === 'P03').itens.find(i => i.pront === '0002');
erro(`remarcar(C, ${J({ atend: oc2.atend, data: ter, novaData: ter2 })})`, /LOTADO: .* 1\/1/); // limite 1 por dia
erro(`addFila(C, ${J({ pront: '0002', esp: 'Psiquiatria' })})`, /já está agendado\(a\) em Psiquiatria/); // já tem a consulta de terça
run(`atualizarStatus(C, ${J({ atend: oc2.atend, data: ter, status: 'Cancelado' })})`);
run(`addFila(C, ${J({ pront: '0002', esp: 'Psiquiatria', prioridade: 'Alta' })})`);
erro(`addFila(C, ${J({ pront: '0002', esp: 'Psiquiatria' })})`, /já está na fila/);
erro(`addFila(C, ${J({ pront: '0002', esp: 'Serviço Social' })})`, /atendimento livre/);
assert.strictEqual(run('fila(C)')[0].pront, '0002');

// Encerrar a vaga: some da agenda depois da data fim.
run(`encerrarVaga(C, '${vaga}', '${qua}')`);
assert.strictEqual(run(`agenda(C, '${q4}')`).profs.find(p => p.id === 'P01').itens.filter(i => i.vaga === vaga).length, 0);

// Férias da Carla em qua2: a sessão do João aparece como "Sem atendimento", ninguém agenda com ela nesse dia,
// e o horário do João fica livre para outro profissional.
run(`gravar('Bloqueios', { 'Data Início': '${qua2}', 'Data Fim': '${qua2}', 'Profissional ID': 'P01', Motivo: 'Férias' })`);
const bl = run(`agenda(C, '${qua2}')`).profs.find(x => x.id === 'P01');
assert.strictEqual(bl.bloqueio, 'Férias');
igual(bl.itens.map(i => i.pront + ' ' + i.status), ['0002 Sem atendimento']);
erro(`agendar(C, ${J({ prof: 'P01', pront: '0001', data: qua2, hora: '14:00', freq: 'Única' })})`, /Sem atendimento neste dia: Férias/);
assert.ok(run(`horariosLivres(C, ${J({ prof: 'P01', data: qua2, freq: 'Única' })})`).every(h => !h.livre));
assert.match(ag({ prof: 'P02', pront: '0002', data: qua2, hora: '13:00', freq: 'Única' }), /^A/);
// Feriado (sem profissional) vale para todos, inclusive ordem de chegada.
ag({ prof: 'P03', pront: '0001', data: ter });
run(`gravar('Bloqueios', { 'Data Início': '${ter}', Motivo: 'Feriado' })`);
erro(`agendar(C, ${J({ prof: 'P03', pront: '0002', data: ter })})`, /Feriado/);
assert.strictEqual(run(`agenda(C, '${ter}')`).profs.find(x => x.id === 'P03').itens.find(i => i.pront === '0001' && i.status !== 'Remarcado').status, 'Sem atendimento');
assert.strictEqual(run(`agenda(C, '${ter}')`).feriado, 'Feriado');
assert.strictEqual(run(`agenda(C, '${qua2}')`).feriado, ''); // férias de um profissional não é feriado da clínica
// Depois das férias a série segue normal.
assert.strictEqual(run(`agenda(C, '${run(`somaDiasT('${qua2}', 14)`)}')`).profs.find(x => x.id === 'P01').itens[0].status, 'Agendado');

const p = run('painel(C)');
assert.strictEqual(p.ativos, 2);
assert.strictEqual(p.capacidade, 400);
assert.ok(p.especialidades.find(e => e.esp === 'Psiquiatria').fila === 1);
assert.strictEqual(run(`paciente(C, '0001')`).vagas.length, 1);

assert.strictEqual(run(`ler('Atendimentos')`).at(-1)['Registrado Por'], 'Maria');

// 5 erros bloqueiam só aquele aparelho; os outros continuam entrando.
for (let i = 0; i < 5; i++) erro(`inicio({ cod: 'errado', disp: 'aparelho-2' })`, /inválido/);
erro(`inicio({ cod: 'abc123', disp: 'aparelho-2' })`, /neste aparelho/); // até com o código certo, espera
assert.strictEqual(run('inicio(C)').usuario, 'Maria');
// Quem troca de "aparelho" a cada tentativa esbarra no limite geral.
for (let i = 0; i < 100; i++) try { run(`inicio({ cod: 'x', disp: 'ataque-${i}' })`); } catch (e) {}
erro('inicio(C)', /no sistema/);

// Menu "Novo usuário": gera código de 10 caracteres e grava na aba.
run(`novoUsuario()`);
const novo = run(`ler('Usuários')`).at(-1);
assert.strictEqual(novo.Nome, 'Nova Pessoa');
assert.match(novo['Código'], /^[0-9a-f]{10}$/);

// Sincronização com a planilha de prontuários (colunas com nomes diferentes).
run(`(() => { const e = new Planilha(); e.insertSheet('Cadastro').d = [['Prontuario', 'Paciente', 'Situação', 'Mãe'],
  ['0001', 'Ana Exemplo', 'ATIVO', 'Marta'], ['0002', 'João Exemplo', 'Inativo', ''], ['0099', 'Novo Paciente', 'Ativo', '']]; externas['ID-X'] = e; })()`);
run(`atualizar('Config', 3, { Valor: 'ID-X' }); atualizar('Config', 4, { Valor: 'Cadastro' }); atualizar('Config', 5, { Valor: 'prontuario' });
  atualizar('Config', 6, { Valor: 'Paciente' }); atualizar('Config', 7, { Valor: 'Mãe' }); atualizar('Config', 9, { Valor: 'Situação' })`);
assert.strictEqual(run('sincronizarPacientes()'), 3);
igual(run(`ler('Pacientes')`).map(x => x['Prontuário'] + ':' + x.Status + ':' + x['Responsável']), ['0001:Ativo:Marta', '0002:Inativo:', '0099:Ativo:']);

console.log('ok');
