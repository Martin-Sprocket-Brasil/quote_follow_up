# QUOTE FOLLOW UP — Carteira de cotações, follow-up e meta diária

> **Novo nome (antes "Quote Pipeline")**: o programa agora se chama **Quote Follow Up** e gera o `QuoteFollowUp.html`. Os arquivos de dados passam a se chamar `Quote Follow Up data.xlsx` e `Quote Follow Up reports.xlsx`: na primeira vez que a pasta é aberta na versão nova, os arquivos antigos (`Quote Pipeline data.xlsx` / `reports.xlsx`) são **renomeados sozinhos**, sem perder nada. Atualize o programa de todos ao mesmo tempo — uma versão antiga não acha os arquivos com o nome novo. Os dados guardados no navegador continuam valendo.

Aplicação web (React) feita a partir da tabela **Customer portfolio** do BRAZIL OEM PLAN (pulz). Você importa o **Quote Report** e o **Sales Report**, preenche três campos à mão **em cada cotação** e o programa mostra:

- a carteira **por cotação** (uma linha por Quote ID) com **QT | Customer | Date | FUP | % Win | Thermometer | Quoted | Open | Lost | Sold**;
- na aba **Calendar**, a **meta diária**, a agenda de follow-ups e, ao clicar num dia, o **plano do dia**: a meta diária e as cotações para trabalhar naquele dia;
- na aba **Salespeople**, o **FUP por vendedor**: quantos follow-ups cada Sales Person atualizou no período, quantos depois das 11h e quantos faltam.

Tudo roda no navegador, sem servidor. Os arquivos não saem do computador.

> A interface está em inglês, como no pulz; os comentários do código estão em português.

---

## Telas de cada papel

| Papel | Abas |
|---|---|
| Vendedor externo (Sales Person) e vendedor interno (Quoted By) | **Today** (follow-ups do dia, da maior prioridade para a menor) · **Quotes** (as cotações dele) · **Calendar** (o plano dos dias e a meta) |
| Gerência | **Overview** · **Quotes** · **Team** (FUP por vendedor) · **Goal** · **Analysis** · **Forms** · **Data** (no celular, Forms e Data ficam no menu do usuário) |

**Importar relatórios:** os três perfis (vendedor externo, interno e gerência) importam o Quote Report e o Sales Report: no computador pelos botões do topo; no celular pelo menu do usuário (*Importar o Quote Report* / *Importar o Sales Report*). Cada importação vale para todo mundo. Um relatório filtrado (só as cotações de um vendedor, por exemplo) não apaga as outras: só entram as cotações do arquivo.

**Gerência só visualiza as cotações:** não faz FUP, não muda % Win, não escreve notas, não fecha (Lost/Cancelled), não mexe no plano do dia e não preenche o formulário DCS. Ao abrir uma QT/SO aparece o **Resumo da cotação**: vendedor interno e externo, datas (criada, entrou no programa, último follow-up e quem fez, próximo follow-up ou quando fechou e em quantos dias), valores (cotado, vendido, pedidos, linhas), follow-ups feitos, notas, última atividade, requisição e motivo de fechamento, além do histórico completo e do PDF do DCS. Na tabela e no calendário, FUP e % Win aparecem como texto. O servidor também recusa essas gravações vindas da gerência.

**Overview (gerência):** uma faixa com os números do dia (vendido no mês × meta e o ritmo, em aberto, taxa de conversão, follow-ups atrasados e de hoje), o que precisa de atenção (tocar abre a lista já filtrada), a equipe hoje e as cotações em aberto de maior prioridade. Visual enxuto: cor só para o que é problema (atrasado em vermelho).

**Prioridade (0 a 100):** 70% do % Win + 30% do Thermometer (a mesma conta do Calendar; mude os pesos em `PRIO_WIN` / `PRIO_TEMP` no `core.js`). Aparece em cada cotação, com o Thermometer, e a lista vem da maior para a menor. No computador é uma coluna que dá para ordenar.

**Filtros:** os rápidos ficam à vista (Atrasados, Para hoje, Prioridade alta, Quente ou Morno, Sem FUP; dá para ligar vários) e cada filtro ligado vira um chip com × em cima da lista. No celular, o botão ao lado da busca abre a folha com ordem, período (Este mês, Mês passado, Este ano fiscal…), classificação, linha, vendedor, FUP, % Win, Thermometer e cliente. Os filtros continuam ao trocar de aba.

**Idioma:** English ou Português no menu do usuário (e na tela de entrada). Fica salvo no navegador; sem escolha, segue o idioma do aparelho. As traduções estão no `i18n-pt.js` (texto em inglês → texto em português): para mudar uma, troque só o lado direito e rode o `build.bat`. Os valores gravados (motivos de fechamento, status) continuam em inglês no banco; só a tela traduz.

**Celular:** pelo servidor, dá para instalar na tela inicial (Chrome: menu ⋮ → *Adicionar à tela inicial*; iPhone/Safari: Compartilhar → *Adicionar à Tela de Início*). Abre em tela cheia, com ícone próprio.

## Colunas da tabela

| Coluna | Origem | Regra |
|---|---|---|
| **QT** | automática | Número da cotação (Quote ID). Cotações canceladas ficam de fora. Ao lado vêm o cliente e a data da cotação. |
| **FUP** | manual | A = próximo dia (vermelho) · B = duas vezes em 5 dias (laranja) · C = 2 semanas (azul) · D = 1 mês (verde): as cores vão do mais urgente ao mais folgado. Ao lado aparece a data do próximo follow-up (no B, as duas datas; em vermelho quando atrasado). Trocar a letra recomeça a contagem a partir de hoje. |
| **% Win** | manual | 25% · 50% · 75% · 99%. |
| **Thermometer** | automática | Calculado por cliente (veja abaixo). Ao lado aparecem o Hit Rate e os dias com cotação; passe o mouse para ver a conta. |
| **Quoted** | automática | *Amount* da cotação (soma dos itens). |
| **Open** | automática | Valor da cotação quando ela ainda não foi ganha, perdida nem cancelada. |
| **Lost** | automática | Valor da cotação com status *Lost* / *Perdida* / *Declined*. |
| **Sold** | automática | Cotação ganha (status *Confirmed* ou com *Sales Id*): o **Amount (Type Booked) do pedido no Sales Report**, ligado pelo *Sales Id* da cotação — igual a filtrar o Sales Report pelos Sales IDs das cotações. Pedido que não está no Sales Report (fora do período dele) ou cotação sem Sales Id = sem Sold. Um pedido ligado a duas cotações conta uma vez só. Sem Sales Report importado, usa o *Sales Amt* da cotação. |

As regras de ganha/perdida/cancelada e do valor vendido são as mesmas do pulz. No recorte por período, a cotação ganha conta no mês do **pedido**, e as demais, na data da cotação (igual ao Customer Tendency).

Acima da tabela, os botões **All / Open / Won / Lost / Cancelled** filtram pela situação da cotação. As canceladas só aparecem em **Cancelled**: elas não entram em *All* nem nos totais (mesma regra do pulz), e na linha delas só o *Quoted* tem valor.

No cabeçalho:
- **Ordem padrão**: cliente de A a Z e, dentro do mesmo cliente, a cotação mais recente primeiro.
- **Customer**: clique no nome da coluna para ordenar A–Z / Z–A; o funil ao lado filtra (com busca, em ordem alfabética).
- **Filtro** (ícone de funil): **FUP**, **% Win** e **Thermometer**. Clique no nome da coluna para marcar uma ou mais opções. O número ao lado do funil mostra quantas opções estão marcadas.
- **Ordenação** (setas): **Date**, **Quoted**, **Open**, **Lost** e **Sold**. O primeiro clique põe o mais recente ou o maior primeiro; o segundo inverte.
- **QT** não tem filtro nem ordenação: para achar uma cotação, use a busca.
- **Clear filters** limpa todos os filtros de uma vez.
 A seta de cada linha abre o detalhe: resultado, quem cotou, vendedor, conta AX e os pedidos (SO) com data e valor Booked.

Até 400 linhas são desenhadas de uma vez; o botão **Show all** mostra o resto.

Recortes no topo (valem para as duas abas), no mesmo formato do pulz:
- **Período** (ícone de calendário, só na tabela): *All periods*, um ano fiscal inteiro (julho a junho) ou um mês dele. Cada ano fiscal é um grupo com a opção *whole year* e os meses que têm cotação. O padrão é o ano fiscal atual.
- **Classification**: vem da coluna *Classification* do Quote Report. Na primeira vez, o programa escolhe **OEM** se existir.
- **Line** (botões *All / PT / CP / MH*): pela coluna *Product Group*, com a mesma regra do pulz. Entram as QTs que têm item dessa linha, e os valores passam a ser **só os da linha**: Quoted, Open, Lost, Sold e o Booked do Sales Report. Se a barra mostrar o aviso *quotes without product line*, as cotações foram guardadas por uma versão antiga (ou o relatório não tinha a coluna Product Group): importe o Quote Report de novo. **Product Group que não conta:** itens cujo Product Group não é de nenhuma linha (*No Booking ,Services (99)* sem linha, *Customer Sample (91)*, *Shop Supplies (95/959)*, *Raw materials (89)*, *Machinery maintenance (959)*) são ignorados na leitura do Quote Report e do Sales Report: não somam valor em nenhum lugar (nem em *All*), e a cotação que só tem esses itens sai da base. A mensagem da importação mostra quantas linhas ficaram de fora. Os serviços com linha contam na linha deles: *No Booking ,Services (99 CP)* → CP, *(99 PT)* → PT, *(99 H)* → MH (no Sales Report: *Service Sales CP/PT/MH (99 …)*). A regra fica em `lineOf` e `skipPG` no código.
- **Salesperson**: só a coluna **Sales Person** do Quote Report (o *Quoted By* não entra no filtro e não há separação entre interno e externo). Cada nome mostra o número de cotações não canceladas (todas as datas), igual à tabela em *All periods*. Os nomes seguem a regra do pulz: sem o código SP ("SP0360 - Bruna Marcelino" vira "Bruna Marcelino"), sem diferença de acento ou maiúscula, e grafias da mesma pessoa contam como um nome só. Passe o mouse no nome para ver as grafias agrupadas. O **booked de cada dia** (calendário, meta diária e plano do dia) também segue o vendedor: o programa lê as colunas *Sales Responsible* (SP0178) e *Name* (Julia Pulz Campo) do Sales Report e liga ao vendedor do Quote Report pelo código SP ou pelo nome. Sales Report importado por uma versão antiga não tem o vendedor: importe-o de novo.
- **Clear** limpa a linha e o vendedor.


O botão **Export** baixa a tabela em Excel, já com FUP, próximo FUP, % Win e Thermometer.

### Thermometer automático
Calculado para cada **cliente** e mostrado em todas as cotações dele:

- **Hit Rate** = valor vendido ÷ valor cotado, considerando as cotações do cliente dos **últimos 6 meses a partir de hoje** (constante `THERMO_HR_MONTHS`), na classificação e na linha escolhidas. As canceladas ficam de fora. O vendido segue a regra da coluna Sold: Booked do pedido, ou o Sales Amt da cotação.
- **Dias com cotação** = quantos dias diferentes o cliente cotou nos **últimos 30 dias úteis até a data de referência**. Sábado e domingo não contam no tamanho da janela, que dá umas 6 semanas.
- **Data de referência** = a última data do Quote Report importado (com o relatório do dia, é hoje). Assim, quem parou de cotar vai perdendo dias e esfria, e um relatório de dias atrás não esfria todo mundo de uma vez.

| Thermometer | Regra | Significado |
|---|---|---|
| **Hot** | Hit Rate ≥ mínimo e cotou em 12 dias ou mais | Sempre cota e compra |
| **Warm** | Hit Rate ≥ mínimo e cotou em 11 dias ou menos | Não cota sempre, mas compra quando cota |
| **Cool** | Hit Rate < mínimo e cotou em 12 dias ou mais | Cota, mas nem sempre compra |
| **Cold** | Hit Rate < mínimo e cotou em 11 dias ou menos | Quase não cota e quase não compra |

**Mínimo do Hit Rate:** 50% para clientes **OEM** e 30% para as outras classificações. A classificação do cliente é a da cotação mais recente dele. Passando o mouse no Thermometer, aparece o mínimo usado.

**Atividade recente:** se o cliente cotou em **menos de 3 dias nos últimos 10 dias úteis**, ele **desce um nível** (Hot vira Warm, Cool vira Cold). Na tabela, isso aparece com uma seta vermelha (↓) ao lado do Thermometer; passe o mouse para ver o nível original e a conta.

Exemplo: um cliente cotou todo dia útil até 19/10, parou e só voltou em 30/10. Pela janela longa ele é Cool, mas nos últimos 10 dias úteis (19 a 30/10) cotou só em 2 dias, então cai para Cold.

**Para alterar os números**, edite o bloco `CONFIGURAÇÃO DO THERMOMETER` nas regras (`core.js`, ou no início do `app.jsx` único):

| Constante | Padrão | O que faz |
|---|---|---|
| `THERMO_HR_MONTHS` | 6 | Período do Hit Rate, em meses a partir de hoje |
| `THERMO_HR_MIN` | 0.5 | Hit Rate mínimo (50%) para Hot/Warm nos clientes **OEM** |
| `THERMO_HR_MIN_OTHER` | 0.3 | Hit Rate mínimo (30%) para Hot/Warm nas **outras classificações** (PFP, Distributor, End User, LATAM, Resale, Supplier) |
| `THERMO_WINDOW` | 30 | Janela longa, em dias úteis |
| `THERMO_DAYS_MIN` | 12 | Dias com cotação na janela longa para Hot ou Cool |
| `THERMO_RECENT_WINDOW` | 10 | Janela recente, em dias úteis |
| `THERMO_RECENT_MIN` | 3 | Mínimo de dias com cotação na janela recente; abaixo disso, desce um nível |
| `THERMO_STEP_DOWN` | Hot → Warm, Cool → Cold | Para onde cada nível desce |
| `thermoRefDate` | última data do Quote Report | Data em que as janelas terminam (há uma linha comentada para usar sempre hoje) |

---

## Aba Calendar

A aba tem três partes, nesta ordem:

### 1. Meta diária
Digite a **meta por dia útil**. Ela vale para todos os dias de segunda a sexta; sábado, domingo e feriado ficam sem meta. No plano de um dia dá para trocar a meta **só daquele dia**: digitar o mesmo valor da meta diária volta o dia para a meta padrão, e digitar **0** deixa o dia sem meta. (Correção: versões anteriores gravavam meta 0 num dia só por clicar no campo e sair, e esses dias ficavam presos em 0; ao abrir esta versão esses 0 são apagados e os dias voltam a seguir a meta diária — clique em Save para gravar.) Os ladrilhos mostram:
- **Booked today** e **Missing today**: o que já foi Booked hoje no Sales Report (na classificação escolhida) e quanto falta para a meta de hoje;
- **Days on target**: em quantos dias úteis do mês, até ontem, a meta foi batida;
- **Expected**: cada cotação aberta × % Win dela, somadas.

### 2. Calendário: o plano dos dias
O FUP define **as datas de follow-up** de cada cotação, contadas do dia em que o FUP foi marcado (ou do último ciclo concluído com *Done*):

| FUP | Datas |
|---|---|
| A · Next day | o próximo dia útil |
| B · 2× in 5 days | **duas datas**: o 2º e o 4º dia útil seguintes. Marcado na segunda → quarta e sexta; marcado na terça → quinta e segunda (o fim de semana não conta) |
| C · 2 weeks | 14 dias depois |
| D · 1 month | 1 mês depois |
| E · Ongoing | **automático**: a data que você escolhe no Done (não aparece na lista para escolher à mão) |

C e D que caem em sábado ou domingo passam para a segunda. Trocar a letra recomeça a contagem a partir de hoje.

**Done (no Activities / History da cotação):** o Done fica junto das anotações, no registro da cotação (clique na linha da tabela: a cotação abre no painel lateral, com as setas ‹ › — ou Alt+↑ / Alt+↓ — para ir à cotação anterior / próxima da tabela; Esc fecha). Os botões *Done* do plano do dia e da visão semanal abrem a cotação direto nesse ponto.
- O **Done só libera com uma nota escrita, a data do próximo follow-up escolhida** (a data começa vazia; há atalhos *Tomorrow, +2 days, +1 week, +2 weeks, +1 month*; sábado ou domingo vão para a segunda) **e o % Win revisado**: o passo 3 pergunta *"is it still 50%?"* e é preciso clicar no % atual (*now*, para manter) ou em outro (para trocar). A cada Done a pergunta volta, para ninguém esquecer de olhar a chance.
- Ao confirmar: a nota entra no histórico, a troca do % Win (se houver) também, o Done é registrado e a cotação passa para **E · Ongoing** com o próximo FUP na data escolhida (ex.: *A · Next day → E · Ongoing · Next follow-up 29/10/2026*).
- **Only change the date**: troca só a data do próximo follow-up, sem Done (a nota é opcional). Também passa para E · Ongoing e fica no histórico como *Next follow-up changed*.
- **Customer requisition** (Requisição do cliente): campo de cada cotação no Activities / History (ex.: *RODRIGO - POLIAS*). Digite e clique em **Save** (ou Enter); a mudança entra no histórico e vai para o relatório de FUP e para a aba Follow-up do arquivo de dados.
- **Save only the note**: grava só a nota (em cotação fechada o botão se chama *Add note*).
- O formulário mostra os 3 passos obrigatórios (**1 · Note**, **2 · Next follow-up date** e **3 · % Win**) com o selo *Required* (âmbar) até serem preenchidos e *OK* (verde) depois. O botão do Done diz o que falta (*Done · missing: % Win (3)*) e, pronto, mostra o % e a data: *Mark Done · 50% · next follow-up 29/10*.
- No E, cada Done pede a próxima data de novo. Para sair do E, escolha outra letra no FUP.
- O arquivo de dados guarda a data do E e o passo do ciclo do B (aba Follow-up, colunas *Next date (E)* e *Cycle step*).

**Visão do calendário:** **Month** (padrão, o mês inteiro de **domingo a sábado**, só o número da QT) ou **Week** (segunda a sexta). As setas trocam de mês ou de semana, e **Today** volta para hoje.

Na visão **Week**:
- **resumo da semana**: meta da semana, booked na semana (com % da meta até hoje e dias na meta), valor com FUP de hoje em diante, quantidade de follow-ups (e atrasados) e quantos Done foram feitos;
- **cabeçalho de cada dia**: dia da semana, data e uma barra — nos dias que passaram e hoje, *Booked / meta* (verde = bateu); nos dias à frente, *With FUP / meta* (azul);
- **cartões das cotações**: letra do FUP, QT, valor, cliente, motivo (atrasada, 1º/2º follow-up do B, letra do FUP), % Win e Thermometer. Clique no cartão abre a cotação (registro, Activities History e notas); hoje cada cartão tem o botão **Done**;
- **Done**: embaixo de cada dia, os follow-ups marcados como feitos naquele dia;
- cada dia tem altura máxima (`WEEK_LIST_H`, 420 px): com muitas cotações, a lista do dia ganha barra de rolagem própria;
- clicar no dia (fora de um cartão) abre o plano do dia, como no mês.

O programa monta o plano de hoje até o fim do período mostrado. Para cada dia:
1. entram as cotações com follow-up naquele dia (no dia de hoje, também as atrasadas). A cotação com FUP fica **sempre nas datas do FUP dela**; o B aparece nas duas datas, mas o valor dele só conta uma vez na meta;
2. **(desligado por enquanto)** completar a meta com cotações sem FUP. Hoje cada dia mostra só as cotações com FUP naquela data; as outras aparecem no botão *Show the other … open quotes*, no fim da lista do plano do dia. Para religar, troque `AUTO_FILL = false` para `true` no código: o programa volta a puxar cotações em aberto sem FUP, pela prioridade, até cobrir a meta.

No calendário:
- **cor do FUP**: follow-up naquele dia (o B nas duas datas);
- **borda vermelha tracejada**: follow-up atrasado, trazido para hoje;
- **"+" com borda azul tracejada**: cotação sem FUP puxada para completar a meta do dia (só com `AUTO_FILL` ligado);
- **embaixo de cada dia**: nos dias passados, *vendido / meta* (verde quando bateu); de hoje em diante, *plano / o que falta* (azul quando o plano cobre a meta).

A **prioridade** (0 a 100) usa só o **% Win** e o **Thermometer** — o FUP não entra:
- **70%** do % Win (25% → 0,25 … 99% → 0,99; sem % Win = 0);
- **30%** do Thermometer (Hot = 1 · Warm = 0,67 · Cool = 0,33 · Cold = 0).

Exemplos: 75% + Hot = 83 · 75% + Warm = 73 · 99% + Cold = 69 · 50% + Hot = 65 · sem % Win + Hot = 30. Os pesos ficam em `PRIO_WIN`, `PRIO_TEMP` e `PRIO_TEMP_VAL` no código.

Com `AUTO_FILL` ligado, o preenchimento pode ser por **Priority** (padrão), **Most likely** (maior % Win) ou **Biggest** (maior valor, menos cotações por dia); desligado, esses botões não aparecem.

### 3. Plano do dia (aparece quando você clica num dia)
- **Meta do dia**: a meta diária. Dá para digitar outro valor só para aquele dia, inclusive sábado ou domingo. O link *back to the daily goal* volta ao padrão.
- **Booked on the day**, **Missing**, **Planned** (quantas cotações têm FUP no dia e o valor delas) e **Expected** (planejado × chance).
- **Quotes to work on**: as cotações do dia, com a coluna **Why** explicando o motivo (prazo no dia, atrasado, puxada para a meta, % Win alto, Hot, cotação velha, sem % Win).
- **Filtros e ordenação da lista:** filtro em Customer, FUP, % Win, Thermometer e Why; ordenação em Date, Priority e Value. O 1º clique ordena do maior para o menor, o 2º inverte e o 3º volta à ordem do plano. A coluna **#** sempre mostra a posição no plano.
- **Clique no número da QT** para abrir a gaveta com o registro, o Activities History e as anotações.
- **Done** (só no dia de hoje): o follow-up foi feito. O próximo prazo passa a contar de hoje, e a cotação não volta para o mesmo dia.
- **Caixinha**: tira a cotação do plano, e os dias são planejados de novo.
- **Show the other open quotes** mostra o resto, com o dia em que cada uma foi planejada.
- **Export** baixa o plano do dia em Excel.

Dia que já passou mostra só o vendido contra a meta; o plano começa hoje.

> Um Sales Report importado antes desta versão só tinha o total do mês. Importe de novo para ter o vendido de cada dia.

---

## Importar um relatório antigo (só para completar)
- **Quote Report mais velho** que as cotações guardadas (a última cotação do arquivo é anterior à última guardada): o programa **só completa o que falta** — Product Groups e linhas PT/CP/MH das cotações já guardadas que estavam sem. **Situação, valores, pedidos e o histórico não mudam** (uma QT ganha não volta para Open) e cotações do arquivo que não estão guardadas ficam de fora. Aparece *Older Quote Report … used only to fill what was missing*.
- **Sales Report mais velho** que as vendas guardadas: não substitui nada (as vendas guardadas são mais novas); só entram dias **anteriores** ao começo do que já está guardado, se houver.
- Depois, clique em **Save** (o reports.xlsx é regravado com os Product Groups).

## Formato dos relatórios

**Quote Report**: precisa da coluna `Quote ID` (ou *Quote No* / *Cotação*), `Customer` e `Create Date`. Também usa, quando existem: `Status`, `Amount`, `Sales Amt`, `Sales Id`, `Classification`, `Quoted By` e `Sales Person`. Cada cotação pode ter várias linhas (itens); o programa soma por Quote ID, sem arredondar (o total bate com a soma do Excel até o centavo).

**Sales Report**: as colunas são reconhecidas por palavras-chave, como no pulz (data, cliente, tipo/status, valor, classificação, Sales Id). Só entram as linhas **Booked**.

O cabeçalho pode estar em qualquer uma das 30 primeiras linhas. Datas em texto são reconhecidas como mês/dia/ano ou dia/mês/ano, conforme o próprio arquivo. Se você usar o botão errado, o programa percebe pelo conteúdo e importa no lugar certo.

---

## Com o servidor (MariaDB): celular e várias pessoas ao mesmo tempo

Aberto pelo servidor (`node server.mjs`, direto ou pelo túnel Cloudflare), o programa usa o banco MariaDB em vez da pasta. Não tem botão Save: cada mudança vai sozinha em cerca de 1 segundo, e o que os outros mudam aparece em até 20 segundos. O login é conferido pelo servidor, com as mesmas senhas. Como instalar: `SERVIDOR.md`.

Aberto como arquivo, o programa continua usando a pasta, como descrito abaixo.

## Onde ficam os dados: a pasta com os arquivos de dados

O programa guarda tudo em **dois arquivos Excel numa pasta** (pode ser a pasta de rede), conectada ao programa (Chrome ou Edge, a mesma tecnologia do pulz):

| Arquivo | O que guarda | Tamanho | Quando é gravado |
|---|---|---|---|
| `Quote Follow Up data.xlsx` | FUP, % Win, requisição do cliente, notas, Activities History, metas | ~0,6 MB | Em todo **Save** com mudança |
| `Quote Follow Up reports.xlsx` | Cotações e vendas (dos relatórios importados) | ~8 MB | Só no Save depois de **importar** o Quote Report ou o Sales Report |

Assim o Save do dia a dia grava só o arquivo pequeno (rápido mesmo na rede); o grande só é regravado quando os relatórios mudam. Limpar o cache do navegador não apaga nada: basta abrir a pasta de novo.

**Menu ⋯:** *Open the data folder…* (escolha a pasta onde estão os arquivos) e *Create the data files in a folder…*. O Chrome pede permissão para editar os arquivos da pasta: escolha permitir.

**Arquivo antigo (tudo num arquivo só):** abra a pasta onde ele está. As cotações dele são lidas normalmente; no primeiro Save elas vão para o `Quote Follow Up reports.xlsx` e o `Quote Follow Up data.xlsx` fica pequeno. Quem estava conectado ao arquivo antigo vê o botão **Choose the data folder**.

**No canto de cima:**
- **No data file · create one**: os dados ainda estão só no navegador.
- **Nome da pasta · opened/saved hh:mm**: conectado e tudo salvo.
- **Nome da pasta · N unsaved changes (+ reports)**: há mudanças só no navegador. *N* = quantas **cotações** estão diferentes do arquivo (FUP, % Win, requisição, Done, nota; cada cotação conta uma vez e, se você desfizer, ela sai da conta) + metas alteradas. *history only* = você mudou e desfez: só o histórico tem registros novos. *+ reports* = as cotações/vendas também mudaram (o Save vai gravar o arquivo grande).
- **Reconnect …** / **Choose the data folder**: o navegador pede permissão de novo (ou a conexão antiga era de um arquivo só).
- **Gravação manual, como no pulz**: abrir ou reconectar o arquivo **só lê** (junta com o navegador e não grava nada). As mudanças ficam guardadas no navegador e o botão do arquivo mostra **"N unsaved changes · Save"**; clique nele (ou **Ctrl+S**) para gravar no Excel. Fechar a página com alterações não salvas mostra o aviso do navegador. Criar um arquivo novo grava na hora.
- **O que eu alterei?** Clique no aviso amarelo (*Quote Follow Up · N unsaved changes*): abre a lista do que ainda não foi salvo — por cotação (notas, Done, FUP, % Win, data, requisição, plano, com a hora; clique na QT para abrir), as metas alteradas (de → para) e, à parte, os registros que a importação do Quote Report criou. O botão **Save** da lista grava tudo.
- **Várias pessoas com o programa aberto** (ex.: Julia e Pedro usando a mesma pasta na rede):
  - a cada **1 minuto** e sempre que a janela volta ao foco, o programa confere se alguém salvou os arquivos; se salvou, **traz as mudanças** e avisa (*Someone else saved the data files…*). O que você ainda não salvou continua na tela e no contador do Save;
  - o botão **Refresh** (ao lado do Save) faz essa conferência na hora; se ninguém salvou, avisa *You already have the latest data*;
  - **antes de cada Save** ele confere de novo e junta: fica o que **você** mudou desde a última leitura e, no resto, o que está no arquivo (as mudanças da outra pessoa). Notas, Done e histórico das duas pessoas são somados;
  - se as duas pessoas mudarem **a mesma cotação** (FUP, % Win, data), vale a de quem salvar por último, e as notas/Done das duas ficam no histórico;
  - o único caso que ainda pode perder algo é as duas clicarem em Save **nos mesmos segundos**.
  - **conserto pelo histórico**: toda mudança de FUP, % Win, data e requisição fica no Activities History, e o histórico das pessoas sempre é somado. Se o FUP / % Win / data de uma cotação ficar diferente do último registro do histórico dela (alguém salvou por cima com dados velhos — por exemplo um computador com o app.jsx antigo, ou uma aba aberta desde antes da atualização), o programa volta o campo para o que o histórico diz, avisa e mostra a cotação na lista do aviso amarelo como *Restored from the history*. Clique em Save para gravar;
  - **versão do programa**: cada Save grava a versão do programa na aba Info. Se uma página com uma versão mais velha abrir arquivos salvos por uma mais nova, aparece a faixa *A newer version of Quote Follow Up saved the data files* e o Save fica bloqueado até recarregar a página (F5). Isso vale a partir desta versão: as versões anteriores não fazem essa conferência.
- O arquivo é gravado **compactado** (cerca de 3× menor) e **uma gravação por vez**.
- **Arquivos `.crswap`**: são o temporário que o Chrome usa enquanto grava (ele escreve no `.crswap` e, no fim, troca pelo arquivo). Se a gravação falha, o programa descarta o temporário. Os que sobraram de versões antigas podem ser apagados com o programa fechado.
- Evite deixar o arquivo de dados aberto no Excel enquanto usa o programa: o Excel bloqueia o arquivo e as gravações esperam.
- **Faixa vermelha no topo**: se abrir ou gravar o arquivo falhar, a faixa mostra o que aconteceu, o que fazer e a mensagem exata do navegador (fica até você fechar). Se a leitura deu certo e só a gravação falhou, os dados do arquivo já ficam abertos e o programa tenta gravar de novo.
- **Erro *InvalidStateError … state had changed since it was read from disk***: apesar do texto, é o erro genérico do Chrome quando o Windows recusa a última etapa da gravação (trocar o arquivo pela versão nova). Acontece principalmente com o arquivo numa **unidade de rede**; também pode ser o Excel, um sincronizador ou antivírus. Na rede isso costuma acontecer quando o antivírus está examinando o arquivo recém-escrito (os `.crswap` de 8 MB que sobram na pasta são o sinal disso). O programa tenta até **4 vezes**; da 2ª em diante ele escreve e **espera 3 s, 8 s e 15 s antes de trocar o arquivo**, para o exame terminar. A espera que funcionou fica **guardada neste navegador** (também depois de fechar e abrir de novo): as próximas gravações já começam com ela e passam de primeira, sem deixar `.crswap`. Só se todas falharem aparece a faixa vermelha (as mudanças continuam no navegador; clique em Save/Try again para tentar de novo). No console aparece *save attempt 1 refused by the drive…* e depois *saved on attempt 3 (waited 8s…)* — é esperado. Os `.crswap` que sobraram podem ser apagados com o programa fechado. Se nem assim gravar, a opção é deixar o arquivo de dados numa pasta local (ex.: Documentos). Se o Chrome der esse erro, ele tenta de novo sozinho. Se continuar, a faixa vermelha tem o botão **Save the data file in another folder…** (grava tudo num arquivo novo, por exemplo em Documentos, e passa a usar esse) e **Try again**.
- **Unidade de rede**: gravar um arquivo grande na rede é lento e pode ser bloqueado (antivírus, sincronização, bloqueio do servidor). Se der problema, deixe o arquivo de dados numa pasta local (ex.: Documentos).

**No menu ⋯:** só *Open a data file* e *Create a new data file*. O botão **Save** fica ao lado das abas *Quote portfolio / Calendar*: verde com o número de alterações quando há algo para gravar, *Saved* quando está tudo no arquivo (Ctrl+S também grava). Sem arquivo conectado, o Save cria um. Em navegadores sem essa tecnologia (Firefox), o menu mostra *Load a data file* e *Download the data file*, e o Save baixa o arquivo.

**Importar relatórios novos:**
- **Quote Report**: as **QTs novas entram** no que já estava guardado. As que já existiam são atualizadas, e as mudanças (situação, valor cotado, pedido novo) vão para o Activities History. QTs que não vieram no relatório novo continuam guardadas.
- **Sales Report**: o período do relatório novo (do primeiro ao último dia dele) é substituído; o resto do histórico continua.

**Abas do arquivo:**

| Aba | Conteúdo |
|---|---|
| Quotes | Uma linha por QT: data, cliente, classificação, status, Quoted by, Sales person, valores, pedidos, valor por linha (PT/CP/MH), quando entrou e quando mudou |
| Sales | Linhas Booked somadas por pedido, dia, classificação, linha e cliente |
| Follow-up | FUP, data de início do FUP, % Win, se está fora do plano e os follow-ups feitos |
| Activities | Activities History e anotações, com data e hora |
| Goals | Meta diária e metas digitadas para dias específicos |
| Info | Versão, data em que foi salvo e os últimos relatórios importados |

Não edite o arquivo à mão: o programa o reescreve a cada mudança.

**Ao abrir ou reconectar o arquivo**, ele é juntado com o que estiver no navegador, para nada se perder:
- **QTs:** fica a versão atualizada por último.
- **Histórico e anotações:** entram os dois lados, sem repetir.
- **FUP / % Win e metas:** vale o arquivo.

**Dica:** guarde o arquivo numa pasta sincronizada (OneDrive, por exemplo) para ter cópia de segurança.

O navegador também guarda uma cópia (IndexedDB e localStorage) para o programa abrir rápido.

---

## Activities History e anotações

Clique na seta de uma QT no **Quote portfolio**, ou no número da QT no **plano do dia**, para abrir o registro dela:

- **Resumo:** situação, cliente, valor e idade da cotação; follow-up (FUP, próximo prazo, % Win, Thermometer); vendedores; valor por linha; pedidos.
- **Activities / History**, agrupado por dia (*Today*, *Yesterday*, data), do mais novo para o mais antigo, com a hora:
  - FUP alterado (de → para);
  - % Win alterado (de → para);
  - follow-up feito (Done), com o próximo prazo;
  - tirada ou recolocada no plano do dia;
  - mudanças trazidas pelo Quote Report: situação (ex.: Open → Won), valor cotado e pedido novo;
  - quando a QT entrou no programa e a data de criação dela.
- **Filtros do histórico:** *All*, *Notes*, *Follow-up* e *Quote Report*, cada um com a quantidade.
- **Anotações:** escreva e clique em **Add note** (ou Ctrl+Enter). A nota entra no histórico com data e hora e **não pode ser alterada nem excluída**: não existe botão para isso. Ela também fica salva na aba *Activities* do arquivo de dados.
- **Autor:** o campo de nome foi retirado. A coluna *By* do arquivo de dados continua existindo, para os registros antigos.
- **Contador de notas:** na tabela e no plano do dia, a QT com anotação mostra um balão com o número de notas. Passando o mouse, aparece a última nota.

---

## Formulário DCS / Value Added

No registro da QT (seta na lista ou número da QT no plano do dia) há o quadro **DCS / Value Added**. Ele aparece em qualquer situação da cotação, inclusive Won. Clique em **Preencher**:

1. **O que houve?** Escolha *Valor adicionado (VA)*, *DCS* ou *Deslealdade*.
2. Para VA ou DCS aparecem os **subgrupos** para marcar (um ou vários). VA: repasse de preço, desconto, devolução, frete cortesia, comissão, lead time, prazo de pagamento. DCS: os 19 itens do formulário (tempo inativo evitado, redução de estoque … pesquisa em campo). Deslealdade não tem subgrupos.
3. **Preenchido sozinho:**
   - QT e SO (se a QT tem mais de um pedido, escolha na lista);
   - vendedor interno e externo, pelos nomes da cotação;
   - cliente final e distribuidor: se a classificação do cliente for *Distribuidor*, ele vai como distribuidor; senão, como cliente final;
   - cidade, estado, distribuidor e cliente final são lembrados do último formulário do mesmo cliente (o Quote Report não traz cidade e estado);
   - descrição: o atalho *Usar a última anotação* copia a nota mais recente da QT.
4. Complete o resto: descrição, se foi monetizado (Sim/Não), valor e moeda (BRL/USD), se está documentado. Ao clicar em salvar, o que faltar fica em vermelho.
5. Depois de salvo:
   - **Compartilhar** (celular): abre o WhatsApp, e-mail etc. já com o PDF anexado;
   - **E-mail para jmartins@martinsprocket.com**: baixa o PDF e abre o e-mail com assunto e texto prontos. Pelo computador o anexo precisa ser colocado à mão (o navegador não consegue anexar sozinho). O e-mail abre no programa de e-mail padrão do Windows: para ser o **Outlook clássico**, veja abaixo;
   - **Baixar o PDF**.

O PDF gerado é o formulário oficial, com os campos preenchidos e ainda editáveis.

**Abrir o e-mail no Outlook clássico (e não no novo):** quem decide é o Windows, não o programa. Faça uma vez em cada computador:

- **Windows 11:** Configurações → Aplicativos → Aplicativos padrão → role até *Escolher padrões por tipo de link* → **MAILTO** → escolha **Outlook (clássico)**.
- **Windows 10:** Configurações → Aplicativos → Aplicativos padrão → *Escolher aplicativos padrão por protocolo* → **MAILTO** → **Outlook**.
- Ou, no próprio Outlook clássico: Arquivo → Opções → Geral → marque *Tornar o Outlook o programa padrão para Email, Contatos e Calendário*.

**Gerência → aba Formulários** (no celular: menu do usuário → *Formulários DCS / Value Added*): todos os formulários salvos, a última versão de cada QT (ou *Todas as versões*). Filtros de período (data em que foi salvo), tipo, vendedor e busca (QT, SO, cliente, distribuidor, cidade, descrição). Mostra quantos de cada tipo, o valor monetizado (BRL e USD, sem somar versões antigas), os subgrupos mais usados e, em cada linha, o PDF; clicar abre a QT. **Exportar para Excel** gera duas abas: *Formulários* (todos os campos, um por linha, valor como número) e *Resumo* (por tipo e por subgrupo), sempre com os filtros da tela. O programa sabe quando o formulário foi salvo; se o e-mail foi enviado, não (o envio acontece no Outlook / WhatsApp).

O formulário fica no **histórico** da QT como uma entrada que não pode ser alterada nem excluída. *Corrigir* grava uma versão nova; o quadro e o PDF mostram sempre a última.

Observação sobre o modelo: no PDF original, as listas *Economia (monetizado ou não)* e *Documentado* eram o mesmo campo, ou seja, escolher uma mudava a outra. O `dcs-va-template.pdf` já vem corrigido e vazio.

---

## Instalação e build

```bash
npm install
npm run build   # executa: node build.js
```

No Windows, basta dar dois cliques em `build.bat`, como no pulz. O resultado é **`QuoteFollowUp.html`**, um arquivo único (~1,5 MB, por causa do gerador de PDF e do modelo do formulário DCS / VA; pelo servidor ele vai compactado) que funciona sozinho no Chrome ou Edge.

## Estrutura

| Arquivo | Função |
|---|---|
| `core.js` | Regras sem interface: leitura dos relatórios, carteira por cotação, agenda de FUP, meta diária e prioridade |
| `app.jsx` | Interface: tabela, aba Calendar, importação, backup |
| `entry.jsx` | Ponto de entrada (monta `<App />`) |
| `styles.css` | CSS global, embutido no HTML final |
| `dcs.js` | Formulário DCS / Value Added: tipos, itens, campos obrigatórios e preenchimento do PDF (pdf-lib) |
| `dcs-va-template.pdf` | Modelo oficial do formulário, vazio, embutido no HTML |
| `build.js` / `build.bat` | Empacota tudo em `QuoteFollowUp.html` (esbuild) |

## Stack

| Biblioteca | Versão | Uso |
|---|---|---|
| React / ReactDOM | 18.3.1 | Interface |
| SheetJS (`xlsx`) | 0.18.5 | Leitura dos relatórios e exportação em Excel |
| lucide-react | 0.383.0 | Ícones |
| esbuild (dev) | 0.28.2 | Empacotamento |


## Relatório de FUP (Excel, e-mail ou PDF)
Botão **FUP report** na barra de filtros. Escolha o período (**Today**, **Previous business day**, **This week**, **Last week**, **This month** ou **Custom**) e:
- **Download Excel**: baixa *FUP report <período>.xlsx* já formatado, com as cores do PDF: título, cabeçalho azul (*Comentário novo* rosa, *Termômetro* verde), bordas, linhas zebradas, comentário novo em rosa claro, R$ e %, linha de Total, filtro e cabeçalho fixo. A planilha é montada pelo próprio programa (o SheetJS gratuito não grava cores), sem precisar instalar nada;
- **Copy for e-mail**: copia a tabela com as cores (cabeçalho azul, *Comentário novo* rosa, *Termômetro* verde); é só colar (Ctrl+V) no e-mail;
- **PDF**: abre a página do relatório e o diálogo de impressão (*Salvar como PDF*), A4 deitada.

- Entram as cotações que tiveram **Done** no período, **no recorte da tela** (Salesperson, Classification, Line). Uma linha por cotação (o último Done do período).
- Colunas: **Quotation · New net amount · Name · Comentário antigo · Comentário novo · Requisição do cliente · Termômetro · Follow-up date · Quoted by · Vendedor**.
  - *Comentário novo* = a nota gravada com o Done; *Comentário antigo* = a nota anterior a ela (cada uma com a data);
  - *Termômetro* = o % Win da cotação;
  - *Follow-up date* = a data do próximo follow-up escolhida no Done;
  - *Quoted by* = o código SP do Quoted By; *Vendedor* = o Sales Person.
- No fim, o total de cotações e o valor somado.

## Aba Salespeople — FUP por vendedor
Uma linha por **Responsável** (coluna *Sales Person* do Quote Report; grafias diferentes do mesmo vendedor viram um nome só, como no filtro). Segue os filtros da tela (Classification, Line, Salesperson). Períodos: **Today**, **Previous business day**, **This week**, **Last week**, **This month** ou **Custom**.

| Coluna | Regra |
|---|---|
| Qtd Cotações | Atualizadas + para atualizar (cada cotação conta uma vez) |
| Qtd Atualizada | Cotações com **Done** no período |
| Qtd Atualizada após 11h | O 1º Done da cotação no período foi às **11:00 ou depois** (`LATE_HOUR` em core.js) |
| % Atualizada | Atualizada ÷ Qtd Cotações |
| Qtd para atualizar | Cotações **em aberto**, com FUP, **sem Done no período** e com o follow-up vencendo até o fim do período (as atrasadas de antes também entram, marcadas *Overdue*) |
| % Faltou atualizar | Para atualizar ÷ Qtd Cotações |

A linha **Total** soma todos. Clique num vendedor (a seta) para abrir os detalhes:
- **Lines**: % do valor cotado em PT, CP e MH, nas cotações dele no período escolhido no topo da tela;
- **Product groups**: os grupos que ele está cotando (valor, **QTs** = quantas cotações têm item daquele grupo, e %). Uma QT com itens de vários grupos conta em cada um. Os grupos são lidos do Quote Report a partir desta versão — **importe o Quote Report de novo** uma vez para preencher as cotações já guardadas;
- **Follow-ups**: as cotações do período, com a hora da atualização (laranja = depois das 11h) ou *To update / Overdue*. Clique no número da QT para abrir a cotação.

Botões no topo da aba: **Excel** formatado (aba *FUP by salesperson* com a tabela e o total — % Atualizada e % Faltou atualizar em verde/âmbar/vermelho conforme o valor —, e aba *Details* com as cotações de cada vendedor: status, hora da atualização, data do follow-up e % Win), **Copy for e-mail** (tabela com cores para colar no e-mail) e **PDF**.

> Para períodos passados, "para atualizar" usa a situação atual da cotação: uma cotação que venceu na semana passada e só recebeu Done nesta semana não aparece como pendente da semana passada.

## Feriados
O calendário mostra os feriados de **Artur Nogueira (SP)** e o programa não deixa nenhum follow-up cair neles: a data vai para o próximo dia útil (como sábado e domingo). Isso vale para os ciclos A–D, para a data escolhida no Done (aparece *"02/11 is a holiday: … → next business day"*) e para datas E já gravadas num feriado. Dia de feriado também fica **sem meta** e não conta como dia útil.

| Data | Feriado | Tipo |
|---|---|---|
| 01/01 | Ano Novo | nacional |
| Sexta-feira Santa (2 dias antes da Páscoa) | Paixão de Cristo | nacional |
| 10/04 | Aniversário de Artur Nogueira | municipal |
| 21/04 | Tiradentes | nacional |
| 01/05 | Dia do Trabalhador | nacional |
| 09/07 | Revolução Constitucionalista | estadual (SP) |
| 07/09 | Independência | nacional |
| 15/09 | Nossa Senhora das Dores (padroeira) | municipal |
| 12/10 | Nossa Senhora Aparecida | nacional |
| 02/11 | Finados | nacional |
| 15/11 | Proclamação da República | nacional |
| 20/11 | Consciência Negra | nacional |
| 25/12 | Natal | nacional |

**Carnaval (segunda e terça) e Corpus Christi** são pontos facultativos: aparecem em âmbar no calendário, mas continuam dias úteis. Para bloquear também, troque `block: false` por `block: true` em `HOLIDAYS_EASTER` (core.js / topo do app.jsx). Para acrescentar um feriado fixo, inclua a data em `HOLIDAYS_FIXED`.

## Cancelled e Lost à mão, com o motivo
Quando a vendedora já sabe que a cotação não vai sair (antes do Quote Report mudar), ela escolhe no **FUP** (tabela ou plano do dia) **Lost** ou **Cancelled** — ou, no painel da cotação, os botões *Closing it? Lost / Cancelled*. Abre um quadro para escolher o **motivo** (obrigatório):

| Fechamento | Motivos |
|---|---|
| Cancelled | Budget Only · Project Discontinued · Stalled project · No Response |
| Lost | Lead Time · Price |

(Motivos gravados com nomes antigos — Prazo, Preço, Levantamento de custo, Deadline, Cost Estimate, Desistence, Customer withdrew, Withdrawal… — passam sozinhos para os nomes atuais.)

- A cotação **sai da agenda** (calendário, plano, "para atualizar" da aba Salespeople) e **passa para a aba Lost / Cancelled** na hora: perdida → o valor vai de Open para *Lost*; cancelada → sai de *All* e dos totais, como as canceladas do Quote Report (e também do Hit Rate do Thermometer). No cabeçalho do painel aparece *Lost · by hand* / *Cancelled · by hand*.
- Na tabela, a coluna FUP dessas cotações mostra **Lost · Price**, **Cancelled · Budget Only** etc. Clique para trocar o motivo ou **Reopen** (volta para Open).
- As cotações que **já vêm Lost ou Cancelled no Quote Report** também têm o motivo: aparecem como *Lost · no reason* (tracejado) até alguém escolher — na tabela (clique na pílula) ou no painel (botões do motivo).
- Fechar, reabrir e trocar o motivo entram no Activities / History; o arquivo de dados guarda nas colunas *Closed as* e *Close reason* da aba Follow-up. A exportação da carteira tem a coluna *Close reason*.
- Se o Quote Report depois disser Won, vale o Won.

## Analysis
Segue os filtros do topo (período, classificação, linha e vendedor) e mostra uma parte por vez, escolhida no topo. Em cima, a faixa com cotado, ganho, taxa de conversão (QTs ganhas ÷ QTs cotadas; canceladas ficam fora), em aberto e perdidas + canceladas. O botão **Excel** baixa todas as partes, uma aba para cada.
- **Vendedores**: por vendedor externo (Sales Person) ou interno (Quoted By): QTs, cotado, ganho, conversão, em aberto, follow-ups atrasados, cotações sem FUP e a média de dias da cotação até o pedido. Clique no cabeçalho para ordenar.
- **Mês a mês**: gráfico cotado × ganho pelo mês da cotação (passe o mouse num mês para ver os números) e a tabela com conversão, ainda em aberto, perdidas e canceladas.
- **Carteira em aberto**: valor em aberto, valor ponderado pelo % Win, idade média, atrasados; idade das cotações (0–15 até mais de 90 dias), situação do follow-up (em dia, hoje, atrasado, sem FUP), % Win e as 10 maiores cotações em aberto (clique para abrir).
- **Clientes**: cotado, ganho, conversão, em aberto, perdidas + canceladas e o motivo principal, com busca.
- **Motivos**: motivos das perdidas e das canceladas (quantas ainda estão sem motivo) e os Product Groups com mais perdas.

## Login (usuários e senhas)
Ao abrir o programa aparece a tela de login: escolha o seu nome (os vendedores da coluna **Sales Person** do Quote Report) e digite a senha.
- **Admin é escondido**: não aparece na lista. Para entrar como Admin, na tela de login aperte **Ctrl + Shift + A** (ou dê **duplo clique no ícone verde** do programa) — a opção *Admin · all salespeople* aparece já escolhida; repetir esconde de novo. O programa não lembra o Admin como último usuário, então a opção não fica à vista para quem usar o mesmo PC.
- **Primeira vez** de cada usuário: o programa pede para **criar a senha** (mínimo 4 caracteres, digitada duas vezes). Para isso a pasta de dados precisa estar aberta, porque a senha vai para o arquivo compartilhado — o Save é feito na hora.
- **Vendedor** vê só as próprias cotações: o filtro *Salesperson* fica preso no nome dele (carteira, calendário, Salespeople, FUP report e Analysis). **Admin** vê tudo.
- O nome de quem está logado vai em cada registro do Activities History.
- Topo da tela: o usuário logado; o menu dele tem **Log out** e, para o Admin, **Users and passwords…** — lista quem já criou a senha e o botão **Reset** (a pessoa cria uma nova no próximo login). O Admin deve criar a senha dele antes de todos.
- **Como a senha é guardada**: nunca a senha em si — só um código criptografado (com uma chave aleatória por usuário). E o cadastro inteiro (nomes, Admin/vendedor, códigos das senhas, datas) vai **cifrado** na aba **Config** do *Quote Follow Up data.xlsx*: quem abrir o Excel vê só um texto embaralhado, sem saber qual parte é de quem. O texto muda a cada Save.
- **Se alguém mexer na aba Config**: o programa percebe (o texto tem um código de conferência), avisa e continua com os usuários que o computador já conhece; no próximo Save o cadastro é gravado de novo. Apagar a aba também não tira ninguém: cada PC que já abriu o programa devolve os usuários no próximo Save.
- **Arquivos antigos** com a aba *Users* aberta são lidos normalmente e, no próximo Save, ela vira a aba Config cifrada.
- **Esqueceu a senha do Admin?** Feche o programa em todos os PCs, apague a aba *Config* do data.xlsx, limpe os dados do site no seu navegador e crie a nova senha do Admin antes de os outros abrirem o programa (os vendedores também criam senhas novas).
- **Limite**: o login organiza o acesso dentro do programa, mas não é uma proteção forte — quem tem acesso à pasta da rede consegue abrir os arquivos Excel direto. Para proteger de verdade, a pasta precisa ter permissão de acesso na rede (TI).

