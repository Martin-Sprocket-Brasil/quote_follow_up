# Quote Follow Up no MariaDB

Arquivos:

| Arquivo | Para que serve |
|---|---|
| `schema.sql` | Cria o banco `quote_follow_up` com as tabelas, os índices, as regras e a visão `v_quote_state`. Pode rodar de novo sem apagar nada. |
| `import-xlsx.mjs` | Lê as duas planilhas da pasta de dados **com o próprio `core.js`**, inclusive a aba Config cifrada dos usuários, e gera o `import.sql`. |
| `importar-banco.bat` | Roda o `import-xlsx.mjs` com dois cliques. |
| `.gitignore` | Impede que planilhas, `import.sql` e `node_modules` subam para o GitHub. |

## Passo a passo

1. **Criar o banco** (uma vez):
   ```
   mariadb -u root -p < schema.sql
   ```
2. **Gerar o `import.sql`**: copie os dois arquivos da pasta de dados para a pasta do programa e dê dois cliques em `importar-banco.bat`. Ou rode:
   ```
   node import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx"
   ```
3. **Carregar no banco**:
   ```
   mariadb -u root -p < import.sql
   ```

Pode repetir os passos 2 e 3 sempre que quiser, por exemplo para trazer os dados mais novos antes de trocar de vez. Cotações, vendas e follow-ups são atualizados, e o histórico só ganha os registros que ainda não estavam no banco. Nada é duplicado.

> O `import.sql` tem os dados da empresa e os códigos das senhas. Não mande para o GitHub nem por e-mail: apague depois de usar.

## De onde vem cada tabela

| Planilha · aba | Tabela | Observação |
|---|---|---|
| reports · Quotes | `quotes` | Uma linha por Quote ID. Os valores ficam com 6 casas, sem arredondar, igual ao programa. `customer_key` é o nome normalizado que agrupa o cliente no Thermometer. |
| reports · Quotes · *Sales IDs* | `quote_orders` | Antes era um texto separado por vírgula; agora é uma linha por pedido. |
| reports · Quotes · *Product groups* | `quote_product_groups` | Antes era um JSON na célula; agora é uma linha por grupo. |
| reports · Sales | `sales` | Linhas Booked somadas por pedido, dia, classificação, linha, cliente e vendedor. Aceita valor negativo (devolução). |
| data · Follow-up | `followups` | FUP, % Win, data do E, requisição, fechamento à mão e motivo. |
| data · Follow-up · *Follow-ups done* | `followup_done_days` | Uma linha por dia com Done. |
| data · Activities | `activities` | Histórico e notas. **Não pode ser alterado nem apagado**: o próprio banco recusa. |
| data · Goals | `settings` (`daily_goal`) e `day_goals` | Na meta de um dia, -1 = "sem meta" digitado à mão. |
| data · Config (cifrada) | `users` | Nome, papel, sal e hash da senha. |
| data/reports · Info | `settings` | Últimos relatórios importados e versão do programa. |
| — | `sessions` | Logins do servidor (criada para o servidor; não vem das planilhas). |

## Regras que o banco garante

- **FUP** só A–E. **% Win** só 25, 50, 75 ou 99. **Fechamento à mão** só `lost` ou `cancelled`.
- **Papel do usuário**: `admin`, `manager` (gerência), `outside` (externo), `inside` (interno) ou vazio (sem acesso). O papel antigo `seller` vira sem acesso, como já acontece na tela: o Admin escolhe a tela de cada um.
- **Senhas**: ficam o sal e o hash do jeito que o programa já calcula (SHA-256 repetido 5000 vezes). No login pelo servidor, todo mundo continua com a mesma senha.
- **`followups` e `activities` sem ligação obrigatória com `quotes`**: um Quote Report novo pode tirar uma cotação da base, e o follow-up e o histórico dela não podem sumir junto.
- **`v_quote_state`**: a situação de cada cotação (`won`, `lost`, `cancelled`, `open`) com as mesmas regras do `core.js`, incluindo o fechamento à mão. Exemplo:
  ```sql
  SELECT state, COUNT(*), SUM(amount) FROM v_quote_state GROUP BY state;
  ```

## Conferência feita com as planilhas atuais (MariaDB 10.11)

- **Linhas carregadas:** 33.676 cotações, 14.310 pedidos, 45.808 product groups, 9.517 vendas, 252 follow-ups, 7.562 registros de histórico e 5 usuários.
- **Mesmos números do programa:** soma cotada de R$ 1.403.528.519,41, vendas de R$ 234.270.504,18 e a mesma situação para cada cotação (14.514 won, 2.420 open, 5.169 lost, 11.573 cancelled).
- **Reimportação:** rodar o `import.sql` duas vezes não duplica nada.

## Servidor

O banco é usado pelo servidor do programa (`server.mjs`). Veja `SERVIDOR.md` para ligar o servidor e o túnel Cloudflare.

O servidor usa duas coisas a mais no banco, que o `schema.sql` cria e que ele mesmo cria se faltarem:

- **`sessions`**: quem está logado em cada navegador ou celular. Fica só o SHA-256 do código do cookie.
- **`settings.data_version` / `rep_version`**: sobem 1 a cada gravação. É por eles que as páginas abertas sabem que precisam buscar as mudanças.
