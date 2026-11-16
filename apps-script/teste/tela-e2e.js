// Roteiro de ponta a ponta da TELA (Index.html) com o servidor de verdade (Code.gs) e a planilha simulada.
// Monta: node teste/tela-e2e.js > ../prototipo/harness/tela-e2e.html ; abra no navegador e rode `await e2e()` no console.
const fs = require('fs'), B = __dirname + '/../';
const idx = fs.readFileSync(B + 'Index.html', 'utf8');
const srv = ['teste/simulador.js', 'Logica.gs', 'Code.gs'].map(f => fs.readFileSync(B + f, 'utf8')).join('\n;\n');

const preparo = `
if (!crypto.randomUUID) crypto.randomUUID = () => 'id-' + Math.random().toString(36).slice(2) + Date.now();
configurar();
gravar('Usuários', { Nome: 'Diretora', 'Código': 'dir123', Perfil: 'Direção' });
gravar('Usuários', { Nome: 'Recepção', 'Código': 'rec123' });
gravar('Usuários', { Nome: 'Carla Prof', 'Código': 'carla123', Perfil: 'Profissional', 'Profissional ID': 'P01' });
gravar('Usuários', { Nome: 'Admin Geral', 'Código': 'geral123', Perfil: 'Geral' });
gravar('Profissionais', { ID: 'P05', Nome: 'Enf. Lia', Especialidade: 'Enfermagem', 'Tipo de Agenda': 'Livre', Ativo: 'Sim' });
// google.script.run de mentira: chama a função do servidor e devolve uma cópia (como o Google faz).
window.google = { script: { run: (() => {
  const novo = () => { let ok = () => {}, falha = () => {};
    const chama = n => (...a) => setTimeout(() => { let r; try { r = JSON.parse(JSON.stringify(window[n](...a) ?? null)); } catch (e) { return falha(e); } ok(r); }, 20);
    const p = new Proxy({}, { get: (_, n) => n === 'withSuccessHandler' ? f => { ok = f; return p; } : n === 'withFailureHandler' ? f => { falha = f; return p; } : chama(n) });
    return p; };
  return new Proxy({}, { get: (_, n) => novo()[n] });
})() } };
try { sessionStorage.setItem('cepress-cod', 'dir123'); } catch (e) {}
`;

const roteiro = `
window.e2e = async () => {
  const log = [], erros = [];
  const espera = ms => new Promise(r => setTimeout(r, ms));
  const pronto = async () => { for (let i = 0; i < 200; i++) { await espera(25); if (!document.body.classList.contains('ocupado')) return; } throw new Error('tela travada em Aguarde…'); };
  const toast = () => { const t = $('#toast'); return t.style.display === 'block' ? t.textContent : ''; };
  const ok = (cond, msg) => { (cond ? log : erros).push((cond ? '✔ ' : '✘ ') + msg); };
  const clica = async el => { if (!el) throw new Error('botão não encontrado'); el.click(); await pronto(); };
  const muda = async (el, v) => { el.value = v; el.dispatchEvent(new Event('input')); el.dispatchEvent(new Event('change')); await pronto(); };
  const nav = async nome => { await clica($('#menu-eq [data-v="' + nome + '"]')); };
  const dia = (n, wd) => { let d = maisDias(D.hoje, n); while (new Date(d + 'T12:00:00Z').getUTCDay() !== wd) d = maisDias(d, 1); return d; };
  const item = re => [...document.querySelectorAll('#ag .it')].find(e => re.test(e.textContent.replace(/\\s+/g, ' ')));
  const vis = sel => { const e = $(sel); return e && getComputedStyle(e).display !== 'none'; };
  let quaPassada, quaPassada2, proxTer;
  try {
    await espera(300); await pronto();
    // 0. Site e login
    let podeGuardar = true; try { sessionStorage.setItem('t', '1'); } catch (e) { podeGuardar = false; }
    if (podeGuardar) ok($('#equipe').classList.contains('on') && !vis('#institutional'), 'Código guardado na sessão: abre direto na área da equipe');
    else { ok(vis('#institutional'), 'Sem sessionStorage (só nesta página de teste): começa no site'); codigo = 'dir123'; await iniciar(); await pronto(); }
    await clica($('#sair')); ok(vis('#institutional') && !$('#equipe').classList.contains('on'), 'Sair da conta volta para o site');
    await clica($('#entrar')); ok($('#login').classList.contains('on'), 'Botão "Área da equipe" abre a janela do código');
    $('#cod').value = 'errado1'; $('#login form').requestSubmit(); $('#login form').requestSubmit(); $('#login form').requestSubmit();
    ok(/Verificando/.test($('#login .btn:not(.ghost)').textContent) && $('#login .btn:not(.ghost)').disabled, 'Login: botão trava e mostra "Verificando…"');
    await pronto(); await espera(80);
    ok(CacheService.getScriptCache().get('erros-' + disp) === '1', 'Login: 3 cliques seguidos contam só 1 tentativa errada');
    ok(!$('#login .btn:not(.ghost)').disabled && $('#login .btn:not(.ghost)').textContent === 'Entrar', 'Login: botão volta ao normal depois do erro');
    ok($('#login').classList.contains('on') && /inválido/.test($('#lerro').textContent), 'Código errado: continua no site com aviso (' + $('#lerro').textContent + ')');
    $('#cod').value = 'dir123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    ok($('#equipe').classList.contains('on') && /Diretora · Direção/.test($('#quem').textContent), 'Código certo: entra como Diretora · Direção');
    const botoes = [...document.querySelectorAll('#menu-eq button')].map(b => b.dataset.v);
    ok(botoes.join() === 'Agenda,Agendar,Pacientes,Fila interna,Painel,Profissionais,Bloqueios,Faltas,Reposições,Lembretes,Enfermagem e Social,Evoluções,Indicadores,Alterações', 'Menu da Direção com 14 itens (sem Equipe): ' + botoes.join(', '));
    await clica($('#voltar')); ok(vis('#institutional') && /Área da equipe/.test($('#entrar').textContent), 'Voltar ao site mantém logado');
    await clica($('#entrar')); ok($('#equipe').classList.contains('on'), 'Área da equipe abre de novo sem pedir código');
    quaPassada = dia(-20, 3); quaPassada2 = maisDias(quaPassada, 7); proxTer = dia(1, 2);

    // 1. Agendar
    await nav('Agendar');
    await muda($('#pr'), 'Ana Exemplo'); ok(/Ativo/.test($('#pinfo').textContent), 'Agendar: paciente reconhecido pelo nome');
    await muda($('#es'), 'Fonoaudiologia'); ok($('#pf').value === 'P01', 'Agendar: único profissional já vem escolhido');
    await muda($('#fq'), 'Semanal'); await muda($('#dta'), quaPassada); await clica($('#ver'));
    const b13 = [...$('#slots').querySelectorAll('button')].find(b => b.textContent === '13:00');
    ok(b13 && !b13.disabled, 'Agendar: horários livres aparecem');
    await clica(b13); await clica($('#ok')); ok(/Agendado/.test($('#conf').textContent), 'Agendar: vaga semanal criada desde ' + br(quaPassada));
    await muda($('#pr'), 'João Exemplo'); await muda($('#dta'), quaPassada); await clica($('#ver'));
    const b13j = [...$('#slots').querySelectorAll('button')].find(b => b.textContent === '13:00');
    ok(b13j && !b13j.disabled && b13j.classList.contains('dupla') && /em dupla com Ana/.test(b13j.title), 'Agendar: horário com 1 paciente fica disponível como dupla (com aviso)');
    await muda($('#es'), 'Psiquiatria'); await muda($('#dta'), proxTer); await clica($('#ok'));
    ok(/lista/.test(toast()), 'Agendar: ordem de chegada (' + toast() + ')');

    // 2. Agenda em colunas
    await nav('Agenda'); await muda($('#dt'), quaPassada); await espera(800); await pronto();
    ok(/Fonoaudiologia/.test($('#ag').textContent) && /13:00–17:00/.test($('#ag').textContent), 'Agenda: coluna da Fonoaudiologia com o horário do turno');
    let it = item(/13:00 Ana/); ok(!!it, 'Agenda: Ana às 13:00 em ' + br(quaPassada));
    await clica(it.querySelector('[data-s="Realizado"]')); it = item(/Ana/);
    ok(/Realizado/.test(it.textContent) && it.querySelector('[data-s="Agendado"]'), 'Agenda: Presença registrada e botão Desfazer aparece');
    await clica(it.querySelector('[data-s="Agendado"]')); ok(!/Realizado/.test(item(/Ana/).textContent), 'Agenda: Desfazer volta para Agendado');
    await clica(item(/Ana/).querySelector('[data-s="Realizado"]'));
    await clica(item(/Ana/).querySelector('[data-mais]')); ok(!!$('#ag .extra'), 'Agenda: Remarcar ▾ abre as opções');
    await clica(item(/Ana/).querySelector('[data-mais]')); ok(!$('#ag .extra'), 'Agenda: segundo clique fecha');
    for (let i = 0; i < 7; i++) await clica($('#prox'));
    ok(diaAg === quaPassada2 && !!item(/13:00 Ana/), 'Agenda: setas › andam os dias');
    await clica(item(/13:00 Ana/).querySelector('[data-s="Confirmado"]'));
    ok(/Confirmado/.test(item(/13:00 Ana/).querySelector('.st').textContent) && item(/13:00 Ana/).querySelector('[data-s="Agendado"]'), 'Agenda: Confirmar (e Desfazer aparece)');
    await clica(item(/13:00 Ana/).querySelector('[data-falta]'));
    ok(!!$('#ag .fj'), 'Agenda: Falta abre a justificativa');
    $('#ag .fj .fj-mot').value = 'Justificada: Viagem'; $('#ag .fj .fj-obs').value = 'viagem'; await clica($('#ag .fj .fj-ok'));
    ok(/Faltou: Justificada: Viagem — viagem/.test(item(/Ana/).textContent), 'Agenda: falta justificada com observação');
    $('#bq').value = 'ana'; $('#bq').dispatchEvent(new Event('input'));
    ok([...document.querySelectorAll('#ag .it')].every(e => /Ana/.test(e.textContent)) && document.querySelectorAll('#ag .it').length >= 1, 'Agenda: busca por paciente');
    $('#bq').value = ''; $('#bq').dispatchEvent(new Event('input'));
    await clica(item(/Ana/).querySelector('[data-fic]')); await pronto();
    ok($('#titulo').textContent === 'Pacientes' && /Ana Exemplo/.test($('#pac').textContent) && document.querySelectorAll('#pac .kpi').length === 3, 'Agenda: nome abre a ficha (com os totais)');
    await nav('Agenda'); await muda($('#dt'), quaPassada2); await espera(800); await pronto();
    for (let i = 0; i < 7; i++) await clica($('#prox'));
    const diaRem = diaAg;
    await clica(item(/13:00 Ana/).querySelector('[data-mais]')); $('#ag .extra .nd').value = diaRem; $('#ag .extra .nh').value = '14:00';
    await clica($('#ag .extra .rem'));
    ok(!!item(/13:00 Ana.*Remarcado/) && !!item(/14:00 Ana.*avulsa/), 'Agenda: remarcação (13:00 Remarcado, 14:00 avulsa)');
    ok(!item(/14:00 Ana/).querySelector('.tl a[href^="tel:"]') && item(/14:00 Ana/).querySelector('.tl a[href="https://web.whatsapp.com/send?phone=5500900000001"]'), 'Agenda: no computador, telefone em texto e WhatsApp pelo WhatsApp Web (sem "Abrir xdg-open?")');
    await clica(item(/14:00 Ana/).querySelector('[data-s="Cancelado"]'));
    ok(!/Cancelado/.test(item(/14:00 Ana/).querySelector('.st').textContent) && /Confirmar cancelamento/.test(item(/14:00 Ana/).textContent), 'Agenda: 1º clique em Cancelar só pede confirmação');
    await clica(item(/14:00 Ana/).querySelector('[data-s="Cancelado"]')); ok(/Cancelado/.test(item(/14:00 Ana/).textContent), 'Agenda: 2º clique cancela');
    let imprimiu = 0; const pr0 = window.print; window.print = () => imprimiu++; $('#imp').click(); window.print = pr0;
    ok(imprimiu === 1 && /, \\d\\d\\/\\d\\d\\/\\d{4}/.test($('.soimp').textContent), 'Agenda: botão Imprimir (com a data na folha)');
    // Semana do profissional
    await clica($('#versem')); await pronto();
    const cols = [...document.querySelectorAll('#semana .col .hd b')].map(b => b.textContent);
    ok(/semana de/.test($('#titulo').textContent) && cols.length === 5 && /^segunda/.test(cols[0]), 'Semana: 5 dias a partir de segunda (' + $('#titulo').textContent + ')');
    ok([...document.querySelectorAll('#semana .it')].some(e => /Ana Exemplo/.test(e.textContent)), 'Semana: mostra os pacientes do profissional');
    const d0 = $('#titulo').textContent; await clica($('#sprox')); await pronto(); ok($('#titulo').textContent !== d0, 'Semana: › vai para a semana seguinte');
    await clica($('#sant')); await pronto();
    await clica($('#semana [data-dia]')); await pronto(); ok($('#titulo').textContent === 'Agenda de atendimentos' && !!$('#ag'), 'Semana: "Abrir o dia" volta para a Agenda daquele dia');
    await muda($('#fe'), 'Psiquiatria'); ok(!/Fonoaudiologia/.test([...document.querySelectorAll('#ag .col .hd b')].map(b => b.textContent).join()), 'Agenda: filtro por especialidade');
    await muda($('#fe'), '');
    await muda($('#dt'), proxTer); await espera(800); await pronto();
    ok(/1º/.test($('#ag').textContent) && /João/.test($('#ag').textContent) && /1\\/30 no dia/.test($('#ag').textContent), 'Agenda: lista de ordem de chegada (1º João, 1/30)');
    await clica(item(/João/).querySelector('[data-mais]')); $('#ag .extra .nd').value = maisDias(proxTer, 7);
    await clica($('#ag .extra .rem')); ok(/Passado para a lista/.test(toast()), 'Agenda: passar para outro dia (ordem de chegada)');

    // 3. Pacientes
    await nav('Pacientes');
    const linhas = () => [...document.querySelectorAll('#tp tr')].filter(r => r.style.display !== 'none').map(r => r.textContent.replace(/\\s+/g, ' '));
    ok(linhas().length === 3 && linhas().some(t => /Bia.*Inativo/.test(t)), 'Pacientes: lista todos, inclusive inativos');
    await muda($('#bp'), 'joão'); ok(linhas().length === 1, 'Pacientes: busca');
    await clica($('#tp tr[data-pr="0001"]')); const ficha = $('#pac').textContent;
    ok(/Ana Exemplo/.test(ficha) && /Semanal/.test(ficha) && /Realizado/.test(ficha) && /Faltou/.test(ficha), 'Pacientes: ficha com vaga e histórico');

    // 4. Fila interna
    await nav('Fila interna'); ok([...$('#fpr').options].map(o => o.text).join() === 'Escolha…,Ana Exemplo,João Exemplo', 'Fila interna: lista de escolha com os pacientes ativos em ordem'); await muda($('#fpr'), '0002'); await muda($('#fes'), 'Terapia Ocupacional'); await clica($('#fadd'));
    ok(/João/.test($('#fl').textContent), 'Fila interna: adicionado');
    await clica($('#fl [data-ag]')); ok($('#titulo').textContent === 'Agendar' && /Fila interna de Terapia Ocupacional/.test($('#sug').textContent), 'Fila interna: Agendar leva ao Agendar com a sugestão');

    // 5. Painel, Profissionais
    await nav('Painel'); ok(/Pacientes ativos/.test(main.textContent) && !/undefined|NaN/.test(main.textContent), 'Painel: sem erro');
    await nav('Profissionais'); ok(/Carla \\(exemplo\\)/.test(main.textContent) && /Seg 08:00–12:00/.test(main.textContent) && /Sala 3/.test(main.textContent), 'Profissionais: dias, horários e sala');

    // 6. Indicadores
    await nav('Indicadores'); const ac = main.textContent.replace(/\\s+/g, ' ');
    ok(/1 atendimento realizado · 1 falta · 1 cancelado/.test(ac), 'Indicadores: acumulado no singular certo (1 atendimento realizado · 1 falta · 1 cancelado)');
    ok(!/undefined|NaN/.test(ac), 'Indicadores: sem "undefined"/"NaN"');
    await muda($('#io'), quaPassada.slice(0, 7)); ok($('#im').value === quaPassada.slice(0, 7), 'Indicadores: outro mês');
    await muda($('#ip'), 'P03'); ok(/Dra\\. Rita/.test($('#ip').selectedOptions[0].text), 'Indicadores: filtro profissional'); await muda($('#ip'), '');
    $('#iv').value = '120'; await clica($('#is')); ok(/meta de 120/.test($('.meta').textContent), 'Indicadores: meta salva');
    ok(/faltam 119/.test($('.meta').textContent), 'Indicadores: mostra quanto falta para a meta');
    $('#iv').value = 'abc'; await clica($('#is')); ok(/número inteiro/.test(toast()), 'Indicadores: meta inválida recusada');
    const mesAnt = maisDias(D.hoje.slice(0, 7) + '-01', -1).slice(0, 7);
    await muda($('#io'), mesAnt); await clica($('#ifech')); await clica($('#ifech'));
    ok(/Mês fechado/.test(main.textContent) && !$('#ireab'), 'Indicadores: Direção fecha o mês anterior (e não reabre)');
    await nav('Alterações');
    ok(/Agendou/.test(main.textContent) && /Fechou o mês/.test(main.textContent) && /Diretora/.test(main.textContent), 'Histórico de alterações: lista quem fez o quê');

    // 7. Encerrar vaga
    await nav('Pacientes'); await clica($('#tp tr[data-pr="0001"]'));
    await clica($('#pac [data-v]')); await clica($('#pac td .pri')); ok(/Encerrada/.test($('#pac').textContent), 'Pacientes: encerrar vaga');

    // 7b. Cadastro de paciente e de profissional
    await nav('Pacientes'); await clica($('#novopac'));
    $('#fp-pront').value = '0050'; $('#fp-nome').value = 'Paciente Novo'; $('#fp-resp').value = 'Mãe'; $('#fp-tel').value = '(21) 98888-7777';
    $('#fp-cpf').value = '52998224725'; $('#fp-cns').value = '798000000000006'; $('#fp-nasc').value = '2018-03-10'; $('#fp-nasc').dispatchEvent(new Event('input'));
    ok(/anos/.test($('#fp-idade').textContent), 'Pacientes: idade calculada ao digitar o nascimento');
    $('#fp-cid').value = 'Belford Roxo'; $('#fp-uf').value = 'rj';
    document.querySelector('.fp-perfil[value="TEA"]').checked = true; document.querySelector('.fp-perfil[value="OUTROS"]').checked = true; $('#fp-pout').value = 'seletividade';
    $('#fp-sup').value = 'Nível 2';
    await clica($('#fp-ok')); await pronto();
    ok([...document.querySelectorAll('#tp tr')].some(r => /0050.*Paciente Novo.*Ativo/.test(r.textContent.replace(/\\s+/g, ' '))), 'Pacientes: novo paciente cadastrado pela tela');
    ok(D.pacientes.some(p => p.pront === '0050'), 'Pacientes: novo paciente já aparece nas listas (Agendar, Fila)');
    await clica($('#novopac')); $('#fp-pront').value = '0050'; $('#fp-nome').value = 'Repetido'; await clica($('#fp-ok'));
    ok(/Já existe o prontuário 0050/.test(toast()), 'Pacientes: prontuário repetido é recusado');
    await clica($('#tp tr[data-pr="0050"]')); ok($('#pac').querySelector('a[href="https://web.whatsapp.com/send?phone=5521988887777"]'), 'Pacientes: ficha com WhatsApp');
    const fc = $('#pac').textContent.replace(/\\s+/g, ' ');
    ok(/CPF 529\\.982\\.247-25/.test(fc) && /CNS 798000000000006/.test(fc) && /anos\\)/.test(fc) && /Belford Roxo\\/RJ/.test(fc) && /Perfil: TEA, Outros: seletividade · Suporte: Nível 2/.test(fc), 'Pacientes: ficha mostra documentos, idade, endereço, perfil e suporte');
    await muda($('#bp'), '529.982'); ok([...document.querySelectorAll('#tp tr')].filter(r => r.style.display !== 'none').length === 1, 'Pacientes: busca por CPF');
    await muda($('#bp'), '798000000000006'); ok([...document.querySelectorAll('#tp tr')].filter(r => r.style.display !== 'none').length === 1, 'Pacientes: busca por CNS');
    await muda($('#bp'), '');
    await clica($('#novopac')); $('#fp-pront').value = '0051'; $('#fp-nome').value = 'X'; $('#fp-cpf').value = '11111111111'; await clica($('#fp-ok'));
    ok(/CPF inválido/.test(toast()), 'Pacientes: CPF inválido é recusado');
    await clica($('#fp-x'));
    await clica($('#tp tr[data-pr="0050"]'));
    await clica($('#editpac')); $('#fp-st').value = 'Inativo'; $('#fp-mot').value = 'Alta'; await clica($('#fp-ok')); await pronto();
    ok([...document.querySelectorAll('#tp tr')].some(r => /0050.*Inativo/.test(r.textContent.replace(/\\s+/g, ' '))) && !D.pacientes.some(p => p.pront === '0050'), 'Pacientes: inativar pela tela');
    await nav('Profissionais'); await clica($('#novoprof'));
    $('#pf-nome').value = 'Lia Música'; $('#pf-dur').value = '40';
    ok([...$('#pf-esp-sel').options].some(o => o.text === 'Fonoaudiologia') && $('#pf-esp-sel').options[$('#pf-esp-sel').options.length - 1].text === '+ Nova especialidade…', 'Profissionais: lista de especialidades com "+ Nova especialidade…"');
    await muda($('#pf-esp-sel'), '__nova'); ok(getComputedStyle($('#pf-lnova')).display !== 'none', 'Profissionais: "+ Nova especialidade…" abre o campo para digitar');
    $('#pf-esp-nova').value = 'Musicoterapia';
    $('#pf-add').click(); const tr = $('#pf-turnos .tr'); tr.querySelector('.t-dia').value = '2'; tr.querySelector('.t-sala').value = 'Sala 9';
    await clica($('#pf-ok')); await pronto(); await espera(50); await pronto();
    ok(/Lia Música/.test(main.textContent) && /Musicoterapia/.test(main.textContent) && /Ter 08:00–12:00/.test(main.textContent), 'Profissionais: cadastrado pela tela com o expediente');
    ok(D.especialidades.indexOf('Musicoterapia') >= 0 && D.profs.some(p => p.nome === 'Lia Música'), 'Profissionais: especialidade nova já aparece nas listas');
    await clica([...main.querySelectorAll('[data-ed]')].find(b => /Lia Música/.test(b.closest('tr').textContent)));
    $('#pf-turnos .t-fim').value = '07:00'; await clica($('#pf-ok')); ok(/fim precisa ser depois/.test(toast()), 'Profissionais: horário errado é recusado');
    // Bloqueios (férias de um profissional)
    await nav('Bloqueios'); $('#bl-ini').value = proxTer; $('#bl-prof').value = 'P03'; $('#bl-mot').value = 'Férias'; await clica($('#bl-ok'));
    ok(/Dra\. Rita/.test($('#bl-l').textContent) && /Férias/.test($('#bl-l').textContent), 'Bloqueios: férias registradas pela tela');
    await nav('Agenda'); await muda($('#dt'), proxTer); await espera(800); await pronto();
    ok(/Sem atendimento: Férias/.test($('#ag').textContent), 'Bloqueios: agenda do dia mostra "Sem atendimento: Férias"');
    await nav('Bloqueios'); const rm = [...document.querySelectorAll('#bl-l [data-rm]')].find(b => /Férias/.test(b.closest('tr').textContent));
    await clica(rm); await clica(rm); ok(!/Férias/.test($('#bl-l').textContent), 'Bloqueios: tirar (com confirmação)');
    // Saída automática após 30 min parado
    ultimoUso = Date.now() - 31 * 60e3; verificaInatividade();
    ok(!$('#equipe').classList.contains('on') && $('#login').classList.contains('on') && /30 minutos/.test($('#lerro').textContent), 'Sai sozinho depois de 30 min sem uso');
    $('#cod').value = 'dir123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();

    // 7c. Perfil Profissional: só a própria agenda e os próprios pacientes
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = 'carla123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    const mp = [...document.querySelectorAll('#menu-eq button')].map(b => b.dataset.v);
    ok(mp.join() === 'Agenda,Agendar,Pacientes,Profissionais,Evoluções' && /Carla Prof · Profissional/.test($('#quem').textContent), 'Profissional: menu reduzido (' + mp.join(', ') + ')');
    await muda($('#dt'), quaPassada); await espera(800); await pronto();
    ok([...document.querySelectorAll('#ag .col .hd b')].map(b => b.textContent).join() === 'Fonoaudiologia', 'Profissional: agenda só com a própria coluna');
    await nav('Pacientes'); ok(!$('#novopac') && [...document.querySelectorAll('#tp tr')].every(r => /Ana Exemplo/.test(r.textContent)), 'Profissional: só os próprios pacientes, sem cadastro');

    // 8. Recepção
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = 'rec123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    ok(![...document.querySelectorAll('#menu-eq button')].some(b => b.dataset.v === 'Indicadores') && /Recepção/.test($('#quem').textContent), 'Recepção: sem Indicadores no menu');
    await nav('Profissionais'); ok(!$('#novoprof') && !main.querySelector('[data-ed]'), 'Recepção: não vê cadastro de profissionais');
    await nav('Pacientes'); ok(!!$('#novopac'), 'Recepção: pode cadastrar paciente');
    ok(![...document.querySelectorAll('#menu-eq button')].some(b => b.dataset.v === 'Evoluções' || b.dataset.v === 'Equipe'), 'Recepção: sem Evoluções e sem Equipe no menu');
    // 8b. Faltas: alterar a justificativa depois
    await nav('Faltas');
    const fAna = [...document.querySelectorAll('#ftb tr')].find(tr => /Ana Exemplo/.test(tr.textContent) && /Justificada: Viagem/.test(tr.textContent));
    ok(!!fAna, 'Faltas: lista a falta justificada da Ana');
    await clica(fAna.querySelector('[data-fj]')); fAna.querySelector('.fj-s').value = ''; await clica(fAna.querySelector('.fj-ok')); await pronto();
    ok([...document.querySelectorAll('#ftb tr')].some(tr => /Ana Exemplo/.test(tr.textContent) && /Não justificada/.test(tr.textContent)), 'Faltas: justificativa alterada para não justificada');
    await clica([...document.querySelectorAll('#ftb tr')].find(tr => /Ana Exemplo/.test(tr.textContent)).querySelector('[data-fj]'));
    $('#ftb .fj-s').value = 'Viagem'; $('#ftb .fj-o').value = 'viagem'; await clica($('#ftb .fj-ok')); await pronto();
    // 8c. Painel clicável e reposições
    await nav('Painel'); await clica(main.querySelector('[data-lista="ativos"]'));
    ok(/Pacientes ativos/.test($('#plista').textContent) && /João Exemplo/.test($('#plista').textContent), 'Painel: clicar em Pacientes ativos abre a lista');
    await clica(main.querySelector('[data-lista="inativos"]')); ok(/Bia Exemplo/.test($('#plista').textContent), 'Painel: clicar em Pacientes inativos abre a lista');
    await nav('Reposições'); ok(/Nenhuma sessão para repor|Repor/.test(main.textContent) && !/undefined/.test(main.textContent), 'Reposições: tela abre');
    // 8d. Cadastro: máscaras
    await nav('Pacientes'); await clica($('#novopac'));
    $('#fp-cpf').value = '52998224725'; $('#fp-cpf').dispatchEvent(new Event('input')); $('#fp-tel').value = '21988887777'; $('#fp-tel').dispatchEvent(new Event('input'));
    $('#fp-cep').value = '2612'; $('#fp-cep').dispatchEvent(new Event('input'));
    ok($('#fp-cpf').value === '529.982.247-25' && $('#fp-tel').value === '(21) 98888-7777' && $('#fp-cep').value === '2612', 'Cadastro: CPF e telefone formatados ao digitar');
    ok($('#fp-pront').placeholder === 'Ex.: 001/2025' && !!$('#fp-sexo') && !!$('#fp-abert') && [...document.querySelectorAll('.fp-perfil')].some(c => c.value === 'TAG'), 'Cadastro: campos novos (abertura, sexo, TAG)');
    await clica($('#fp-x'));
    // 9. Lembretes por WhatsApp
    await nav('Lembretes'); await muda($('#ldt'), maisDias(proxTer, 7));
    const lk = $('#lem a[data-env]');
    ok(lk && lk.href.indexOf('https://web.whatsapp.com/send?phone=5500900000002&text=') === 0 && /Jo%C3%A3o%20Exemplo/.test(lk.href) && /ordem%20de%20chegada/.test(lk.href), 'Lembretes: WhatsApp com a mensagem pronta para o João');
    await clica($('#lem [data-cf]')); ok(/Confirmado/.test($('#lem').textContent) && !$('#lem [data-cf]'), 'Lembretes: "Confirmou" marca Confirmado');
    await nav('Enfermagem e Social'); await muda($('#edt'), quaPassada);
    ok(/Ana Exemplo/.test($('#enf').textContent) && !$('#elok') && /Aqui é só consulta/.test($('#enf').textContent), 'Enfermagem: recepção só consulta');

    // 10. Administradora geral: Equipe e acessos
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = 'geral123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    ok([...document.querySelectorAll('#menu-eq button')].some(b => b.dataset.v === 'Equipe'), 'Geral: menu com Equipe');
    await nav('Profissionais'); await clica([...main.querySelectorAll('[data-ed]')].find(b => /Enf\. Lia/.test(b.closest('tr').textContent)));
    await clica($('#pf-acesso')); await pronto();
    ok($('#titulo').textContent === 'Equipe e acessos' && $('#fu-prof').value === 'P05' && $('#fu-perfil').value === 'Profissional', 'Profissionais: "Configurar acesso" abre a Equipe já com o profissional');
    await clica($('#fu-ok')); await pronto();
    const codLia = ($('#fusu .aviso span') || {}).textContent || '';
    ok(/^\\S{10}$/.test(codLia) && /Enf\. Lia/.test(main.textContent), 'Equipe: acesso criado e código mostrado uma vez');
    await nav('Equipe'); ok(!main.textContent.includes(codLia), 'Equipe: o código não aparece de novo na lista');
    await clica([...main.querySelectorAll('[data-ed]')].find(b => /Admin Geral/.test(b.closest('tr').textContent)));
    $('#fu-perfil').value = 'Direção'; await clica($('#fu-ok')); ok(/pelo menos uma Administradora geral/.test(toast()), 'Equipe: não deixa ficar sem Administradora geral');

    // 11. Enfermagem: passagem e evolução
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = codLia; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    await nav('Enfermagem e Social'); await muda($('#edt'), quaPassada);
    await clica($('#enf [data-p="0001"]')); ok(/Passou por Enfermagem/.test($('#enf').textContent), 'Enfermagem: registrou a passagem da Ana');
    await clica($('#enf [data-ev]')); ok([...document.querySelectorAll('.ev-condutas')].some(c => c.value === 'Sinais vitais aferidos'), 'Enfermagem: evolução com o modelo da Enfermagem');
    $('#ev-obs').value = 'Sem queixas.'; await clica($('#ev-ok')); await pronto();
    ok(/evoluído/.test($('#enf').textContent) && !$('#enf [data-d]'), 'Enfermagem: evoluído e a passagem não pode mais ser desfeita');
    await muda($('#elivre'), 'João Exemplo'); await clica($('#elok')); ok(/livre/.test($('#enf').textContent) && /João/.test($('#enf').textContent), 'Enfermagem: atendimento livre');

    // 12. Profissional: evolução, complemento, gráfico e compartilhamento
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = 'carla123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    await nav('Evoluções'); const pend = $('#evw [data-ev]');
    ok(!!pend && /Ana Exemplo/.test(pend.closest('tr').textContent), 'Evoluções: atendimento realizado aparece para evoluir');
    await clica(pend); ok(/Comunicação funcional/.test($('#fev').textContent), 'Evoluções: modelo da Fonoaudiologia');
    await clica($('#ev-ok')); ok(/Preencha pelo menos/.test(toast()), 'Evoluções: não salva vazia');
    $('#ev-obj').value = 'Ampliar vocabulário'; $('#ev-dur').value = '40'; document.querySelector('.ev-condutas[value="Comunicação funcional"]').checked = true;
    await clica($('#ev-ok')); await pronto(); ok(/Evolução registrada/.test(toast()) && !$('#evw [data-ev]'), 'Evoluções: salva e sai da lista de pendentes');
    await nav('Pacientes'); await clica($('#tp tr[data-pr="0001"]')); await pronto();
    ok(/Ampliar vocabulário/.test($('#evpac').textContent) && /100% \\(1 de 1 ficha\\)/.test($('#evpac .g-graf').textContent), 'Ficha: evolução e gráfico mensal (100%, 1 de 1 ficha)');
    await clica($('#evpac [data-cp]')); $('#evpac .ev-acao textarea').value = 'Mãe relatou melhora.'; await clica($('#evpac .ev-acao button')); await pronto();
    ok(/Complemento de Carla Prof/.test($('#evpac').textContent) && /Ampliar vocabulário/.test($('#evpac').textContent), 'Ficha: complemento sem mudar a original');
    await clica($('#evpac [data-sh]')); await pronto(); $('#evpac .sh-d').value = 'Diretora'; await clica($('#evpac .ev-acao .pri'));
    ok(/Leitura autorizada/.test(toast()), 'Ficha: compartilhou a evolução com a Diretora');
    await nav('Agenda'); await muda($('#dt'), quaPassada); await espera(800); await pronto();
    await clica(item(/Ana/).querySelector('[data-s="Agendado"]')); ok(/já tem evolução/.test(toast()), 'Agenda: atendimento evoluído não volta para Agendado');
    await muda($('#dt'), quaPassada2); await espera(800); await pronto();
    await clica(item(/13:00 Ana/).querySelector('[data-s="Realizado"]')); await clica(item(/13:00 Ana/).querySelector('[data-evo]')); await pronto();
    ok($('#titulo').textContent === 'Evoluções' && /Evolução · Fonoaudiologia/.test($('#fev').textContent) && new RegExp(br(quaPassada2)).test($('#fev').textContent), 'Agenda: botão Evoluir abre a ficha de evolução daquele atendimento');

    // 13. Direção lê só o que foi compartilhado
    await clica($('#sair')); await clica($('#entrar')); $('#cod').value = 'dir123'; $('#login form').requestSubmit(); await pronto(); await espera(50); await pronto();
    await nav('Pacientes'); await clica($('#tp tr[data-pr="0001"]')); await pronto();
    const ev = $('#evpac').textContent;
    ok(/compartilhada com você/.test(ev) && /Ampliar vocabulário/.test(ev) && !/Sem queixas/.test(ev) && !$('#evpac [data-cp]'), 'Direção: lê só a evolução compartilhada (a da enfermagem não), sem editar');
    await nav('Evoluções'); ok(/Ana Exemplo/.test($('#evw').textContent), 'Evoluções: "compartilhadas comigo" lista a Ana');
  } catch (e) { erros.push('✘ PAROU: ' + e.message + ' ' + (e.stack || '').split('\\n')[1].slice(-30)); }
  return { passou: log.length, falhou: erros.length, erros, log };
};
`;

const html = idx.replace('<script>\nconst LOGO', '<script>' + srv + '\n' + preparo + '</script>\n<script>\nconst LOGO')
  .replace('</body>', '<script>' + roteiro + '</script>\n</body>');
if (html === idx) throw new Error('não achei onde inserir');
process.stdout.write(html);
