# Gera o JavaScript que aplica só as mudanças (versão publicada -> versão nova) no arquivo aberto no editor
# do Apps Script (Monaco) e confere o SHA-1 final. Uso:
#   python3 teste/gera-delta.py <arquivo publicado> <arquivo novo> <modelo no editor, ex. file_3.html> <saída.js>
# Cole o conteúdo da saída no console da página do editor (Claude in Chrome: javascript_tool, com "await " na frente).
import difflib, json, sys, hashlib
velho, novo, modelo, saida = sys.argv[1:]
a = open(velho, encoding='utf-8').read(); b = open(novo, encoding='utf-8').read()
A = a.splitlines(keepends=True); Bl = b.splitlines(keepends=True)
ops = []
for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, A, Bl, autojunk=False).get_opcodes():
    if tag == 'equal': continue
    c = 1  # linhas de contexto até o trecho ficar único no arquivo
    while True:
        o = ''.join(A[max(0, i1 - c):min(len(A), i2 + c)])
        if a.count(o) == 1 or c > 20: break
        c += 1
    assert a.count(o) == 1, 'contexto não único'
    ops.append([o, ''.join(A[max(0, i1 - c):i1]) + ''.join(Bl[j1:j2]) + ''.join(A[i2:min(len(A), i2 + c)])])
t = a
for o, n in ops:  # confere aqui que aplicar em sequência dá exatamente o arquivo novo
    assert t.count(o) == 1, 'trecho some depois de outra troca'; t = t.replace(o, n)
assert t == b, 'aplicação local não bate'
h = hashlib.sha1(b.encode('utf-8')).hexdigest()[:10]
js = ('(async () => { const ed = monaco.editor.getEditors()[0], m = ed.getModel(); if (!m.uri.path.endsWith(%s)) return "ARQUIVO ERRADO " + m.uri.path;'
      ' let t = m.getValue(); const ops = %s; for (const [o, n] of ops) { if (t.split(o).length !== 2) return "TRECHO NAO ENCONTRADO: " + o.slice(0, 60); t = t.replace(o, () => n); }'
      ' ed.executeEdits("claude", [{ range: m.getFullModelRange(), text: t }]);'
      ' const h = [...new Uint8Array(await crypto.subtle.digest("SHA-1", new TextEncoder().encode(m.getValue())))].map(x => x.toString(16).padStart(2, "0")).join("").slice(0, 10);'
      ' return m.uri.path + " " + h + (h === %s ? " CONFERE" : " DIFERENTE"); })()') % (json.dumps(modelo), json.dumps(ops, ensure_ascii=False), json.dumps(h))
open(saida, 'w', encoding='utf-8').write(js)
print(len(ops), 'trechos,', len(js), 'caracteres, sha1 esperado', h)
