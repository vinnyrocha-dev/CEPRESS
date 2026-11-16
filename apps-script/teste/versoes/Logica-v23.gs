// Regras de agenda sem acesso à planilha — testadas por teste/logica.js com Node.
// Datas são texto 'aaaa-mm-dd' e horários 'hh:mm' (evita os problemas de fuso do Sheets).

var FREQ = { 'Única': 0, 'Semanal': 1, 'Quinzenal': 2, 'Mensal': 4 }; // em semanas
var DIAS = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];

function pad(n) { return ('0' + n).slice(-2); }

// Aceita '2026-10-14', '14/10/2026' ou '14/10/2026 08:00:00'. Data impossível (31/02, mês 13) vira ''.
function isoData(v) {
  var s = String(v || '').trim(), m, a, me, d;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) { a = +m[1]; me = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) { a = +m[3]; me = +m[2]; d = +m[1]; }
  else return '';
  var dt = new Date(Date.UTC(a, me - 1, d));
  if (dt.getUTCFullYear() !== a || dt.getUTCMonth() !== me - 1 || dt.getUTCDate() !== d) return '';
  return a + '-' + pad(me) + '-' + pad(d);
}

function isoHora(v) {
  var m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? pad(m[1]) + ':' + m[2] : '';
}

function diaIdx(nome) {
  var s = String(nome || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').slice(0, 3).toLowerCase();
  return ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab'].indexOf(s);
}

function paraMin(h) { var p = h.split(':'); return +p[0] * 60 + +p[1]; }
function paraHora(m) { return pad(Math.floor(m / 60)) + ':' + pad(m % 60); }
function diaNum(iso) { var p = iso.split('-'); return Date.UTC(+p[0], p[1] - 1, +p[2]) / 864e5; }
function diaSemana(iso) { return (diaNum(iso) + 4) % 7; } // 0 = domingo (01/01/1970 foi quinta)
function somaDias(iso, n) { return new Date((diaNum(iso) + n) * 864e5).toISOString().slice(0, 10); }
function semana(iso) { return Math.floor((diaNum(iso) + 3) / 7); } // semanas começando na segunda

// Evento = vaga fixa ou consulta avulsa: {id, prof, sala, pront, inicio, hora, dur, freq, fim, excecoes}
function ocorre(e, iso) {
  if (iso < e.inicio || (e.fim && iso > e.fim) || (e.excecoes && e.excecoes[iso])) return false;
  var p = FREQ[e.freq];
  if (!p) return iso === e.inicio;
  return diaSemana(iso) === diaSemana(e.inicio) && (semana(iso) - semana(e.inicio)) % p === 0;
}

function sobrepoe(a, b) {
  return paraMin(a.hora) < paraMin(b.hora) + b.dur && paraMin(b.hora) < paraMin(a.hora) + a.dur;
}

function colidem(a, b) {
  if (!sobrepoe(a, b)) return false;
  if (!FREQ[a.freq]) return ocorre(b, a.inicio);
  if (!FREQ[b.freq]) return ocorre(a, b.inicio);
  if (diaSemana(a.inicio) !== diaSemana(b.inicio)) return false;
  if ((a.fim && a.fim < b.inicio) || (b.fim && b.fim < a.inicio)) return false;
  // ponytail: se os períodos se sobrepõem, considera conflito mesmo que a sobreposição seja curta demais para as séries se encontrarem.
  var g = Math.min(FREQ[a.freq], FREQ[b.freq]); // mdc de 1, 2 e 4 é o menor
  return ((semana(a.inicio) - semana(b.inicio)) % g + g) % g === 0;
}

function conflitos(c, eventos) {
  var msgs = [];
  eventos.forEach(function (e) {
    // Comparar nomes é barato: só calcula datas (colidem) para quem divide profissional, sala ou paciente.
    var prof = e.prof === c.prof, sala = e.sala && e.sala === c.sala, pac = c.pront && e.pront === c.pront;
    if (e.id === c.id || !(prof || sala || pac) || !colidem(c, e)) return;
    if (prof) msgs.push(e.profNome + ' já atende ' + e.nome + ' às ' + e.hora);
    else if (sala) msgs.push(e.sala + ' ocupada por ' + e.profNome + ' às ' + e.hora);
    else msgs.push(e.nome + ' já está com ' + e.profNome + ' às ' + e.hora);
  });
  return msgs;
}

// Turno = um expediente do profissional num dia da semana: {inicio, fim, intIni, intFim, sala}.
// Pode haver mais de um turno no mesmo dia (ex.: manhã na sala 2 e tarde na sala 1).
function foraDoTurno(c, t) {
  var i = paraMin(c.hora), f = i + c.dur;
  if (i < paraMin(t.inicio) || f > paraMin(t.fim)) return 'Fora do horário do profissional';
  if (t.intIni && t.intFim && i < paraMin(t.intFim) && paraMin(t.intIni) < f) return 'Horário cai no intervalo';
  return '';
}

function turnoDe(c, turnos) {
  return (turnos || []).filter(function (t) { return !foraDoTurno(c, t); })[0];
}

function foraDoExpediente(c, turnos) {
  if (!turnos || !turnos.length) return 'O profissional não atende ' + DIAS[diaSemana(c.inicio)];
  if (turnoDe(c, turnos)) return '';
  var motivos = turnos.map(function (t) { return foraDoTurno(c, t); });
  return motivos.indexOf('Horário cai no intervalo') >= 0 ? 'Horário cai no intervalo' : 'Fora do horário do profissional';
}

function gradeHorarios(turnos, dur) {
  var out = [];
  turnos.forEach(function (t) {
    for (var m = paraMin(t.inicio); m + dur <= paraMin(t.fim); m += dur) {
      var h = paraHora(m);
      if (!foraDoTurno({ hora: h, dur: dur }, t) && out.indexOf(h) < 0) out.push(h);
    }
  });
  return out.sort();
}

// CPF: 11 dígitos com os dois dígitos verificadores certos (aceita com ou sem pontuação). Vazio = não informado.
function cpfValido(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  for (var t = 9; t < 11; t++) {
    var s = 0;
    for (var i = 0; i < t; i++) s += +d[i] * (t + 1 - i);
    if ((s * 10) % 11 % 10 !== +d[t]) return false;
  }
  return true;
}

// Cartão do SUS (CNS): 15 dígitos; começa com 1/2 (soma ponderada múltipla de 11) ou 7/8/9 (mesma regra).
function cnsValido(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (!/^[12789]\d{14}$/.test(d)) return false;
  var s = 0;
  for (var i = 0; i < 15; i++) s += +d[i] * (15 - i);
  return s % 11 === 0;
}

// Idade em anos completos na data de referência (as duas em 'aaaa-mm-dd').
function idade(nasc, ref) {
  if (!nasc || !ref || nasc > ref) return '';
  var a = +ref.slice(0, 4) - +nasc.slice(0, 4);
  return ref.slice(5) < nasc.slice(5) ? a - 1 : a;
}
