# Monta o Index.html novo (site do Espaço TEA + área da equipe) a partir de:
#   - a página do site (prototipo/site-tea.html: o HTML do site sem as imagens grandes)
#   - as imagens reduzidas (prototipo/img/)
#   - as telas já testadas do Index.html antigo (Agendar, ficha do paciente, Fila, Painel, Indicadores)
#   - teste/equipe.html (casco da área da equipe, Agenda em colunas, Pacientes e Profissionais)
# Uso único: depois de montado, o Index.html passa a ser editado direto.
import re, base64, sys
B = '/home/ubunto/CEPRESS-recuperado/'
site = open(B + 'prototipo/site-tea.html', encoding='utf-8').read()
antigo = open(sys.argv[1], encoding='utf-8').read()          # Index.html antigo
equipe = open(B + 'apps-script/teste/equipe.html', encoding='utf-8').read()
img = lambda f, t: 'data:image/' + t + ';base64,' + base64.b64encode(open(B + 'prototipo/img/' + f, 'rb').read()).decode()

# Trechos do script antigo que continuam iguais (entre os marcadores).
js = antigo[antigo.index('<script>') + 8:antigo.rindex('</script>')]
def trecho(ini, fim):
    a = js.index(ini); b = js.index(fim, a)
    return js[a:b]
reaproveitado = '\n'.join([
    trecho('function aviso(', '// Lista do médico lotada'),
    trecho('// Lista do médico lotada', 'const VIEWS'),
    trecho('// ---------- Agendar ----------', '// ---------- Paciente ----------'),
    trecho('// ---------- Fila interna ----------', 'function entrar('),
])
# Ficha do paciente: o corpo de vPaciente vira fichaPaciente(pr, alvo).
ficha = trecho('function vPaciente(pr) {', '// ---------- Fila interna ----------')
ficha = ficha.replace("""function vPaciente(pr) {
  main.innerHTML = `<div class="row"><label style="flex:1;max-width:420px">Paciente<input id="pr" list="dl-pac" value="${esc(rotulo(pr))}"></label><button class="pri" id="bus">Abrir</button></div><div id="pac"></div>`;
  const abrir = async () => {
    if (!$('#pr').value.trim()) return aviso('Digite o número do prontuário ou o nome.', true);
    const p = await api('paciente', pront($('#pr').value));""", """function fichaPaciente(pr) {
  const abrir = async () => {
    const p = await api('paciente', pr);""")
ficha = ficha.replace("""  $('#bus').onclick = abrir;
  $('#pr').onkeydown = e => e.key === 'Enter' && abrir();
  if (pr) abrir();
}""", """  return abrir();
}""")
assert 'fichaPaciente' in ficha and 'return abrir();' in ficha, 'ficha não adaptada'

css, html, novo = re.search(r'<style>(.*?)</style>\s*<!--HTML-->(.*?)<!--/HTML-->\s*<script>(.*?)</script>', equipe, re.S).groups()
novo = novo.replace('/*REAPROVEITADO*/', reaproveitado + '\n' + ficha)

out = site.replace('</style>', css + '</style>', 1)
out = out.replace('</body>', html + '<script>\nconst LOGO = "' + img('logo.webp', 'webp') + '";\n' + novo + '</script>\n</body>', 1)
out = out.replace('src="IMG_PREFEITURA"', 'src="' + img('p1.jpg', 'jpeg') + '"').replace('src="IMG_CEPRESS"', 'src="' + img('p2.jpg', 'jpeg') + '"')
assert 'IMG_' not in out and out.count('<script>') >= 2
open(B + 'apps-script/Index.html', 'w', encoding='utf-8').write(out)
print('Index.html', len(out) // 1024, 'KB')
