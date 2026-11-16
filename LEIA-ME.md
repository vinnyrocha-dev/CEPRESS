# CEPRESS · Sistema de agendamento

Uma planilha Google guarda os dados, e uma tela web feita em Apps Script fica por cima dela. A recepção usa só a tela; a coordenação pode abrir a planilha para relatórios.

## Arquivos

| Arquivo | O que é | Vai para o Apps Script? |
|---|---|---|
| `apps-script/Code.gs` | Servidor: lê/grava a planilha, menu, sincronização | Sim |
| `apps-script/Logica.gs` | Regras de horário e conflito | Sim |
| `apps-script/Index.html` | A tela: o site do Espaço TEA Vila Entre Rios com o botão **Entrar na equipe**, que abre a área da equipe (Agenda, Agendar, Pacientes, Fila interna, Painel, Profissionais e Indicadores) | Sim |
| `apps-script/appsscript.json` | Fuso horário (São Paulo) e o serviço **Sheets**, que lê todas as abas de uma vez | Sim |
| `apps-script/teste/` | Testes rodados no computador (`node apps-script/teste/logica.js`, `servidor.js` e `correcoes.js`); `medicao.gs` mede a velocidade numa **cópia** da planilha | **Não** |

## Como a equipe usa

- O link do sistema abre o **site** da unidade. O botão **Área da equipe** (canto superior direito) pede o código pessoal e abre a área da equipe no lugar do site.
- **Voltar ao site** mostra o site de novo sem sair; **Sair da conta** apaga o código daquele navegador.
- O código fica guardado só até fechar o navegador. Enquanto o código é conferido, o botão mostra "Verificando…" e fica travado (cliques repetidos não contam como tentativas erradas).
- **Sai sozinho depois de 30 minutos sem uso** (computador compartilhado da recepção).
- **Agenda:** "Ver semana" mostra a semana de um profissional; "Imprimir" gera a folha do dia ou da semana; o telefone do responsável aparece em cada paciente, com atalho para o WhatsApp (no computador abre o WhatsApp Web; no celular, o número também liga com um toque); "Cancelar" pede um segundo clique para confirmar.
- **Pacientes:** "+ Novo paciente" cadastra; na ficha, "Editar dados" muda nome, responsável, telefone ou inativa (o nº do prontuário não muda).
- **Profissionais (só Direção):** "+ Novo profissional" e "Editar" cadastram nome, especialidade (escolha na lista ou "+ Nova especialidade…"), tipo de agenda e os dias/horários/salas. Se uma vaga fixa ficar fora do novo horário, a tela avisa (a vaga não é apagada). Para tirar da agenda sem perder o histórico, desmarque "Ativo".
- As abas **Usuários** e **Config** da planilha são protegidas: só o dono da planilha edita.
- Tentativas de acesso recusadas ficam registradas em Apps Script → Execuções (sem o código digitado).
- `teste/tela-e2e.js` monta uma cópia da tela com a planilha simulada e um roteiro que clica em tudo (`await e2e()` no console). Rode depois de mudar a tela.

## Instalação (uma vez, uns 15 minutos)

1. Na conta Google **da clínica**, crie uma planilha nova chamada `CEPRESS - Agenda`.
2. Na planilha: **Extensões → Apps Script**.
3. No editor:
   - Apague o conteúdo do `Código.gs` e cole o `Code.gs`.
   - Clique em **+ → Script**, nome `Logica`, e cole o `Logica.gs`.
   - Clique em **+ → HTML**, nome `Index`, e cole o `Index.html`.
   - Em **Configurações do projeto** (engrenagem), marque *Mostrar o arquivo de manifesto "appsscript.json"*, volte ao editor e cole o `appsscript.json`.
   - Salve. O serviço **Sheets** aparece sozinho em *Serviços* por causa do `appsscript.json`.
4. Volte à planilha e recarregue a página. Vai aparecer o menu **CEPRESS**. Clique em **CEPRESS → Configurar planilha** e autorize. Isso cria as abas (incluindo **Histórico**) e a rotina que roda toda madrugada às 3h. O Google vai avisar que o app "não foi verificado": isso é normal para scripts próprios. Clique em *Avançado → Acessar*. Em seguida, use **CEPRESS → Novo usuário** para criar o seu código de acesso.
5. As abas são criadas já com exemplos. Preencha os dados reais (veja abaixo) e **apague as linhas de exemplo**.
6. Para publicar a tela: **Implantar → Nova implantação → Tipo: App da Web**
   - *Executar como*: **Eu**
   - *Quem pode acessar*: **qualquer pessoa com Conta do Google**. Assim, além do código de acesso, a pessoa precisa estar logada numa conta Google.
   - Copie o link gerado e envie só para a equipe.
7. Sempre que mudar o código: **Implantar → Gerenciar implantações → editar (lápis) → Versão: Nova versão**. O link continua o mesmo.

## O que preencher em cada aba

- **Config**
  - `Capacidade`: 400.
  - Os itens `Prontuários:` **ficam em branco** enquanto não existir planilha de prontuários separada. Se um dia existir: preencha o link dela, o nome da aba e o nome exato das colunas de nº, nome, responsável, telefone e status.
  - Depois, use **CEPRESS → Sincronizar pacientes agora** e, se der certo, **Sincronizar pacientes a cada hora**. A conta que instalou o script precisa ter acesso à planilha de prontuários.
- **Pacientes**: a CEPRESS não tem planilha de prontuários separada, então os pacientes são cadastrados **direto nesta aba**, uma linha por paciente. O `Status` é escolhido numa lista (Ativo/Inativo). Para inativar, troque o Status e preencha `Data Inativação` e `Motivo Inativação`. Nunca apague a linha de um paciente que já foi atendido: o histórico dele depende dela.
  - *Se um dia* os prontuários passarem a ficar numa planilha separada, dá para ligar a sincronização (itens `Prontuários:` da aba Config). A partir daí, esta aba vira um espelho e não deve ser editada à mão.
- **Profissionais**
  - `Tipo de Agenda`: **Horário** (terapias), **Ordem de chegada** (médicos) ou **Livre** (Serviço Social e Enfermagem, que não entram na agenda).
  - `Duração (min)`: só para Horário.
  - `Máx. por Dia`: só para Ordem de chegada.
  - `Ativo = Não` tira o profissional da agenda.
- **Expediente**: uma linha por profissional por dia da semana, com início, fim, intervalo (opcional) e sala. A sala pode mudar conforme o dia. Quem atende manhã e tarde com um buraco no meio pode ter **duas linhas no mesmo dia** (uma por turno, cada uma com sua sala).
- **Vagas Fixas**: os horários fixos (semanal, quinzenal ou mensal). A `Primeira Data` define o dia da semana e qual semana da quinzena o paciente vem. Para cadastrar os pacientes que já estão em atendimento, o mais seguro é usar a tela **Agendar**, porque ela confere os conflitos.
- **Atendimentos**: preenchida pela tela, com chegadas, faltas, cancelamentos, remarcações, consultas avulsas e listas de ordem de chegada.
- **Histórico**: criada pelo sistema. Toda madrugada, os atendimentos com mais de 90 dias saem de *Atendimentos* e vêm para cá, para a tela continuar rápida. Nada é apagado; a ficha do paciente e a agenda de datas antigas continuam mostrando tudo.
- **Fila Interna**: pacientes ativos esperando uma especialidade.
- **Bloqueios**: dias sem atendimento.
  - **Feriado ou recesso da clínica:** `Data Início` (e `Data Fim`, se for mais de um dia) com `Profissional ID` **em branco**.
  - **Férias, licença ou congresso de um profissional:** o período e o `Profissional ID` dele.
  - A planilha já tem os feriados nacionais até 2028 e os de Belford Roxo/RJ: Emancipação (3/4), São Jorge (23/4, estadual) e Nossa Senhora da Conceição (8/12, padroeira). O dia 20/6 (morte do ex-prefeito Joca) aparece em alguns calendários, mas não achei a lei; inclua só se a prefeitura confirmar. **Inclua os feriados municipais** e apague Carnaval e Corpus Christi se a clínica abrir nesses dias, porque são ponto facultativo.
  - Num dia bloqueado, a agenda mostra o aviso e marca as sessões como **Sem atendimento**, para a recepção avisar ou remarcar os pacientes. Ninguém consegue agendar nesse dia, e as vagas fixas continuam normalmente depois.
- **Usuários** (obrigatória): uma linha por pessoa da equipe, com `Nome`, `Código` e `Perfil`. Sem nenhum usuário cadastrado, ninguém entra.
  - **Para criar um usuário:** use **CEPRESS → Novo usuário**, digite o nome da pessoa e o sistema gera um código aleatório de 10 caracteres e mostra na tela. Crie primeiro o seu, depois o de cada pessoa da equipe, e entregue o código pessoalmente.
  - **Ou, mais simples:** digite o nome da pessoa na coluna `Nome` da aba Usuários e aperte Enter. O código aparece sozinho na coluna ao lado.
  - Também dá para digitar um código à mão na coluna `Código` (antes do nome), mas ele precisa ter **pelo menos 6 caracteres**. Os códigos gerados automaticamente são mais seguros.
  - `Perfil`: **Direção** vê a tela **Indicadores** e define a meta do mês. Em branco ou **Recepção**: todas as outras telas, sem Indicadores.
  - O nome aparece na coluna `Registrado Por`, então dá para saber quem marcou cada falta ou agendamento.
  - Quando alguém sair da equipe, apague a linha dessa pessoa: o código dela para de funcionar na hora.
  - 5 códigos errados no mesmo aparelho bloqueiam **só aquele aparelho** por 10 minutos. Há também uma trava geral de 100 erros em 10 minutos, contra ataques. Com códigos gerados pelo menu, adivinhar um código é inviável na prática.
  - Como a clínica usa Gmail comum, o Google não informa quem está acessando. Por isso o sistema usa códigos próprios.

## Regras que o sistema aplica

- Paciente inativo não pode ser agendado. Se um paciente com vaga fixa for inativado, a vaga **continua segurando o horário** (aparece na agenda como "Paciente inativo") até alguém clicar em **Encerrar vaga**. O **Painel** lista esses casos, para a recepção decidir.
- O **Painel** mostra também os próximos feriados e os pacientes que mais faltaram.
- Desfazer uma falta ou cancelamento confere de novo se o horário ainda está livre.
- O sistema bloqueia horário ocupado pelo mesmo profissional, pela mesma sala ou pelo mesmo paciente. Duas vagas quinzenais no mesmo horário, em semanas alternadas, **não** são conflito.
- O sistema bloqueia horários fora do expediente, no intervalo ou em dia em que o profissional não atende.
- Ordem de chegada: o paciente não pode entrar duas vezes no mesmo dia e o dia tem limite (`Máx. por Dia`). Quando a lista está lotada, a tela oferece colocar o paciente na fila interna.
- Ordem de chegada também tem **Passar para outro dia** (na Agenda, clicando no paciente): a vaga do dia fica como "Remarcado" e o paciente entra no fim da lista do novo dia, com as mesmas regras de lotação, dia de atendimento e bloqueios.
- Quando um paciente da fila interna é agendado na especialidade que esperava, ele sai da fila automaticamente.
- Cancelar ou remarcar uma sessão libera aquele horário só naquele dia. A vaga fixa continua nas outras semanas.
- Duas pessoas agendando ao mesmo tempo não conseguem pegar o mesmo horário.

## Indicadores (só Direção)

- Filtros por mês, profissional e especialidade. Cards: agendamentos, realizados, faltas, cancelados, agendados sem registro, pacientes distintos atendidos e % de faltas (faltas ÷ presenças + faltas).
- Tabelas: histórico mensal, por profissional, por especialidade e fila interna atual. O acumulado vai do primeiro registro até o fim do 2º mês à frente (vagas fixas não têm fim, então o futuro tem limite).
- Cada sessão conta uma vez. **Não contam:** sessão em feriado/bloqueio sem registro, sessão remarcada (conta no dia novo) e sessão futura de paciente inativo.
- **Meta mensal:** fica na aba Config, na linha `Meta AAAA-MM`. É da unidade inteira e conta só presenças registradas.

## Limitações conhecidas

- Mudar a `Duração` de um profissional muda também a duração das vagas que já existem.
- Quem tem acesso de edição à planilha consegue digitar direto nela e passar por cima das regras. Também consegue ver os códigos de todo mundo. Deixe a planilha compartilhada só com a coordenação; a recepção usa apenas o link da tela.
- Cada ação leva cerca de 1 segundo (medido com 1 ano de dados: 40 mil atendimentos). A ficha do paciente leva uns 2 segundos, porque lê também o Histórico.
- A sincronização de pacientes recusa a planilha de prontuários se ela vier vazia, com status diferente de Ativo/Inativo ou com menos da metade dos ativos de antes, para um erro na origem não inativar todo mundo.
- Limites do Google (conta Gmail comum): a API do Sheets aceita 60 leituras por minuto. Como a tela roda em nome de quem instalou, a equipe toda divide esse limite. Se passar, o sistema não para: lê do jeito antigo, mais lento (uns segundos), até o minuto virar. As rotinas automáticas (sincronização a cada hora e arquivamento às 3h) usam poucos minutos dos 90 diários permitidos.
- WhatsApp fica para a fase 2.
