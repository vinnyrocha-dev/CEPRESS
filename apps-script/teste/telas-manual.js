// Gera a página das capturas de tela dos manuais (dados fictícios): node teste/telas-manual.js
// Cada endereço #tela abre o sistema já logado na tela pedida; o Chrome sem janela tira a foto.
const fs = require('fs'), B = __dirname + '/../';
const base = fs.readFileSync(B + '../prototipo/harness/tela-e2e.html', 'utf8');

const demo = `
(async () => {
  const w = ms => new Promise(r => setTimeout(r, ms));
  const tela = location.hash.slice(1);
  if (!tela) return;
  const C = { cod: 'dir123', disp: 'demo' }, h = hoje();
  let qua = somaDias(h, -1); while (diaSemana(qua) !== 3) qua = somaDias(qua, -1);
  let prox = somaDias(h, 1); while (diaSemana(prox) !== 3) prox = somaDias(prox, 1);
  const ini = somaDias(qua, -21);
  [['0004', 'Miguel Santos', 'Juliana Santos', '(21) 98811-2233'], ['0005', 'Helena Costa', 'Patrícia Costa', '(21) 97722-3344'],
   ['0006', 'Davi Oliveira', 'Marcos Oliveira', '(21) 96633-4455'], ['0007', 'Laura Souza', 'Fernanda Souza', '(21) 95544-5566']]
    .forEach(p => gravar('Pacientes', { 'Prontuário': p[0], Nome: p[1], 'Responsável': p[2], Telefone: p[3], Status: 'Ativo' }));
  gravar('Usuários', { Nome: 'Lia Enfermeira', 'Código': 'lia1234', Perfil: 'Profissional', 'Profissional ID': 'P05' });
  const ag = (prof, pront, hora, freq, data) => agendar(C, { prof, pront, data: data || ini, hora, freq: freq || 'Semanal' });
  ag('P01', '0001', '13:00'); ag('P01', '0004', '13:30'); ag('P01', '0005', '14:00'); ag('P01', '0006', '15:00', 'Quinzenal');
  ag('P02', '0001', '14:00'); ag('P02', '0007', '13:00'); ag('P02', '0004', '15:30');
  agendar(C, { prof: 'P03', pront: '0002', data: somaDias(prox, -1) }); agendar(C, { prof: 'P03', pront: '0007', data: somaDias(prox, -1) });
  const vaga = (prof, pront) => ler('Vagas Fixas').find(v => v['Profissional ID'] === prof && v['Prontuário'] === pront).ID;
  for (let d = ini; d <= qua; d = somaDias(d, 7)) {
    atualizarStatus(C, { vaga: vaga('P01', '0001'), data: d, status: 'Realizado' });
    atualizarStatus(C, { vaga: vaga('P02', '0001'), data: d, status: 'Realizado' });
    atualizarStatus(C, { vaga: vaga('P02', '0007'), data: d, status: d === qua ? 'Faltou' : 'Realizado', motivo: d === qua ? 'Justificada: Doença' : '' });
  }
  atualizarStatus(C, { vaga: vaga('P01', '0004'), data: qua, status: 'Realizado' });
  atualizarStatus(C, { vaga: vaga('P01', '0005'), data: qua, status: 'Faltou' });
  atualizarStatus(C, { vaga: vaga('P02', '0004'), data: qua, status: 'Confirmado' });
  addFila(C, { pront: '0005', esp: 'Terapia Ocupacional', prioridade: 'Alta', obs: 'Indicação da fono' });
  addFila(C, { pront: '0006', esp: 'Psiquiatria' }); addFila(C, { pront: '0002', esp: 'Fonoaudiologia' });
  salvarMeta(C, h.slice(0, 7), '120');
  registrarEnfermagem({ cod: 'lia1234', disp: 'l' }, { data: qua, pront: '0001' });
  registrarEnfermagem({ cod: 'lia1234', disp: 'l' }, { data: qua, pront: '0004' });
  const a1 = ler('Atendimentos').find(a => a['Profissional ID'] === 'P01' && a['Prontuário'] === '0001' && a.Data === somaDias(qua, -7)).ID;
  salvarEvolucao({ cod: 'carla123', disp: 'c' }, { atend: a1, conteudo: { objetivo: 'Ampliar vocabulário funcional', duracao: '30',
    chegada: ['Calmo(a)', 'Receptivo(a)'], condutas: ['Comunicação funcional', 'Linguagem expressiva'], resposta: ['Participou ativamente', 'Contato visual'], saida: ['Regulado(a)'],
    obs: 'Manter pranchas de comunicação em casa.' } });
  salvarBloqueio(C, { ini: somaDias(prox, 14), fim: somaDias(prox, 28), prof: 'P02', motivo: 'Férias' });
  const quem = { site: '', login: '', equipe: 'geral123', 'equipe-novo': 'geral123', enfermagem: 'lia1234', evolucao: 'carla123', ficha: 'dir123',
    'evolucoes': 'carla123', recepcao: 'rec123' }[tela] ?? 'dir123';
  if (tela === 'site' || tela === 'login') { await w(1500); codigo = ''; D = {}; try { sessionStorage.clear(); } catch (e) {} mostrarEquipe(false); $('#login').classList.remove('on'); scrollTo(0, 0); if (tela === 'login') entrar(); return; }
  codigo = quem; await iniciar(); await w(300);
  const vai = async (v, arg) => { ir(v, arg); await w(500); };
  if (tela === 'agenda' || tela === 'recepcao') { await vai('Agenda', qua); }
  else if (tela === 'semana') { diaAg = qua; await vai('Agenda', qua); $('#versem').click(); }
  else if (tela === 'agendar') { await vai('Agendar'); const s = (id, v) => { $(id).value = v; $(id).dispatchEvent(new Event('input')); $(id).dispatchEvent(new Event('change')); };
    s('#pr', 'Laura Souza'); await w(100); s('#es', 'Fonoaudiologia'); await w(100); s('#fq', 'Semanal'); s('#dta', prox); await w(100); $('#ver').click(); }
  else if (tela === 'falta') { await vai('Agenda', qua); await w(300); const it = [...document.querySelectorAll('#ag .it')].find(e => /Davi/.test(e.textContent)) || document.querySelector('#ag .it'); it.querySelector('[data-falta]').click(); }
  else if (tela === 'ficha') { await vai('Pacientes', '0001'); }
  else if (tela === 'paciente-novo') { await vai('Pacientes'); $('#novopac').click(); }
  else if (tela === 'profissional') { await vai('Profissionais'); await w(300); main.querySelector('[data-ed]').click(); }
  else if (tela === 'equipe-novo') { await vai('Equipe'); await w(300); $('#novousu').click(); }
  else if (tela === 'evolucao') { await vai('Evoluções'); await w(400); const b = document.querySelector('#evw [data-ev]'); if (b) b.click(); }
  else if (tela === 'evolucoes') { await vai('Pacientes', '0001'); await w(800); document.querySelector('#evpac') && document.querySelector('#evpac').scrollIntoView(); }
  else if (tela === 'lembretes') { await vai('Lembretes'); vLembretes(prox); }
  else if (tela === 'enfermagem') { await vai('Enfermagem', qua); }
  else { await vai({ fila: 'Fila interna', faltas: 'Faltas', reposicoes: 'Reposições', painel: 'Painel', indicadores: 'Indicadores', alteracoes: 'Alterações', equipe: 'Equipe', bloqueios: 'Bloqueios', profissionais: 'Profissionais', pacientes: 'Pacientes' }[tela]); }
})().catch(e => { document.title = 'ERRO ' + e.message; console.error(e); });
`;
const html = base.replace(/<script>\s*window\.e2e[\s\S]*?<\/script>\s*<\/body>/, '<script>' + demo + '</script>\n</body>');
if (html === base) throw new Error('não achei o roteiro do e2e');
fs.writeFileSync(B + '../prototipo/harness/telas-manual.html', html);
console.log('ok');
