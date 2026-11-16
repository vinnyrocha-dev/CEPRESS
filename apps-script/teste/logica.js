// Testa Logica.gs fora do Google: node teste/logica.js
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const L = vm.createContext({});
vm.runInContext(fs.readFileSync(__dirname + '/../Logica.gs', 'utf8'), L);

for (const d of ['2026-10-14', '2026-10-18', '2027-02-28', '2024-02-29'])
  assert.strictEqual(L.diaSemana(d), new Date(d + 'T00:00Z').getUTCDay(), d);
assert.strictEqual(L.isoData('14/10/2026'), '2026-10-14');
assert.strictEqual(L.isoData('2026-1-5'), '2026-01-05');
assert.strictEqual(L.isoHora('9:00:00'), '09:00');
assert.strictEqual(L.isoData('15/13/2026'), ''); // mês 13
assert.strictEqual(L.isoData('2026-02-31'), ''); // 31 de fevereiro
assert.strictEqual(L.isoData('amanhã'), '');
assert.strictEqual(L.isoData('2028-02-29'), '2028-02-29'); // ano bissexto
assert.strictEqual(L.diaIdx('Sáb'), 6);

const ev = (o) => Object.assign({ id: Math.random(), prof: 'P1', hora: '09:00', dur: 30, fim: '' }, o);
const q14 = ev({ inicio: '2026-10-14', freq: 'Quinzenal' }); // quarta

// Quinzenais em semanas alternadas não colidem; na mesma semana, sim.
assert.ok(!L.colidem(q14, ev({ inicio: '2026-10-21', freq: 'Quinzenal' })));
assert.ok(L.colidem(q14, ev({ inicio: '2026-10-28', freq: 'Quinzenal' })));
assert.ok(L.colidem(q14, ev({ inicio: '2026-10-21', freq: 'Semanal' })));
assert.ok(!L.colidem(ev({ inicio: '2026-10-14', freq: 'Mensal' }), ev({ inicio: '2026-10-28', freq: 'Mensal' })));
assert.ok(!L.colidem(ev({ inicio: '2026-10-14', freq: 'Mensal' }), ev({ inicio: '2026-10-21', freq: 'Quinzenal' })));
// Outro dia da semana, outro horário, ou série já encerrada: sem conflito.
assert.ok(!L.colidem(q14, ev({ inicio: '2026-10-15', freq: 'Semanal' })));
assert.ok(!L.colidem(q14, ev({ inicio: '2026-10-28', hora: '09:30', freq: 'Quinzenal' })));
assert.ok(L.colidem(ev({ inicio: '2026-10-14', dur: 60, freq: 'Semanal' }), ev({ inicio: '2026-10-28', hora: '09:30', freq: 'Única' })));
assert.ok(!L.colidem(ev({ inicio: '2026-10-14', fim: '2026-10-20', freq: 'Semanal' }), ev({ inicio: '2026-10-28', freq: 'Semanal' })));
// Avulsa contra série: só na data em que a série acontece, e não se aquela sessão foi cancelada.
assert.ok(L.colidem(ev({ inicio: '2026-10-28', freq: 'Única' }), q14));
assert.ok(!L.colidem(ev({ inicio: '2026-10-21', freq: 'Única' }), q14));
assert.ok(!L.colidem(ev({ inicio: '2026-10-28', freq: 'Única' }), Object.assign({}, q14, { excecoes: { '2026-10-28': true } })));

// Conflito de sala e de paciente com outros profissionais.
const outro = ev({ prof: 'P2', profNome: 'Paulo', nome: 'Ana', sala: 'Sala 3', pront: '0412', inicio: '2026-10-14', freq: 'Semanal' });
assert.match(L.conflitos(ev({ sala: 'Sala 3', inicio: '2026-10-28', freq: 'Única' }), [outro])[0], /Sala 3/);
assert.match(L.conflitos(ev({ sala: '1', pront: '0412', inicio: '2026-10-28', freq: 'Única' }), [outro])[0], /Ana já está/);
assert.strictEqual(L.conflitos(ev({ sala: '1', pront: '9', inicio: '2026-10-28', freq: 'Única' }), [outro]).length, 0);

const t = { inicio: '08:00', fim: '10:00', intIni: '09:00', intFim: '09:30', sala: '3' };
assert.deepStrictEqual(Array.from(L.gradeHorarios([t], 30)), ['08:00', '08:30', '09:30']);
assert.strictEqual(L.foraDoExpediente(ev({ inicio: '2026-10-14', hora: '09:45' }), [t]), 'Fora do horário do profissional');
assert.strictEqual(L.foraDoExpediente(ev({ inicio: '2026-10-14', hora: '09:00' }), [t]), 'Horário cai no intervalo');
assert.match(L.foraDoExpediente(ev({ inicio: '2026-10-14' }), []), /não atende Qua/);
// Dois turnos no mesmo dia: manhã na sala 2 e tarde na sala 1; os dois valem e cada horário pega a sua sala.
const manha = { inicio: '08:00', fim: '09:00', sala: 'Sala 2' }, tarde = { inicio: '14:00', fim: '15:00', sala: 'Sala 1' };
assert.deepStrictEqual(Array.from(L.gradeHorarios([tarde, manha], 30)), ['08:00', '08:30', '14:00', '14:30']);
assert.strictEqual(L.turnoDe(ev({ hora: '14:30' }), [manha, tarde]).sala, 'Sala 1');
assert.strictEqual(L.foraDoExpediente(ev({ inicio: '2026-10-14', hora: '08:30' }), [manha, tarde]), '');
assert.strictEqual(L.foraDoExpediente(ev({ inicio: '2026-10-14', hora: '11:00' }), [manha, tarde]), 'Fora do horário do profissional');

// CPF, CNS e idade
assert.ok(L.cpfValido('529.982.247-25') && L.cpfValido('52998224725'));
assert.ok(!L.cpfValido('529.982.247-24') && !L.cpfValido('111.111.111-11') && !L.cpfValido('123'));
assert.ok(L.cnsValido('798 0000 0000 0006') && L.cnsValido('100100000000006'));
assert.ok(!L.cnsValido('798000000000007') && !L.cnsValido('300000000000000') && !L.cnsValido('123'));
assert.strictEqual(L.idade('2018-10-22', '2026-10-22'), 8);
assert.strictEqual(L.idade('2018-10-23', '2026-10-22'), 7);
assert.strictEqual(L.idade('', '2026-10-22'), '');
console.log('ok');
