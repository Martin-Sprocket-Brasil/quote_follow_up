-- =====================================================================================
--  QUOTE FOLLOW UP — banco de dados MariaDB
--  Substitui os dois arquivos Excel da pasta de dados:
--    • Quote Follow Up reports.xlsx  → quotes, quote_orders, quote_product_groups, sales
--    • Quote Follow Up data.xlsx     → followups, followup_done_days, activities,
--                                      day_goals, settings, users
--  Testado no MariaDB 10.11 (funciona do 10.5 em diante).
--
--  Como rodar (uma vez):
--    mariadb -u root -p < schema.sql
--  Depois, para trazer os dados das planilhas atuais: veja o import-xlsx.mjs.
--
--  Regras que o banco garante sozinho:
--    • o histórico (activities) e as notas NUNCA são alterados nem apagados (triggers no fim);
--    • o mesmo registro de histórico não entra duas vezes (act_key = a mesma chave do actKey do core.js);
--    • FUP só A–E, % Win só 25/50/75/99, fechamento só "lost" / "cancelled", papel do usuário só os 4 da tela.
-- =====================================================================================

CREATE DATABASE IF NOT EXISTS quote_follow_up
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE quote_follow_up;

-- -------------------------------------------------------------------------------------
-- COTAÇÕES (aba Quotes do reports.xlsx) — uma linha por Quote ID, já somando os itens
-- -------------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS quotes (
  quote_id        VARCHAR(20)    NOT NULL COMMENT 'Quote ID (ex.: QT00210345), sempre em maiúsculas',
  quote_date      DATE           NOT NULL COMMENT 'Create Date da cotação',
  customer        VARCHAR(150)   NOT NULL COMMENT 'Nome do cliente como veio no Quote Report',
  customer_key    VARCHAR(150)   NOT NULL DEFAULT '' COMMENT 'Nome normalizado (normName do core.js): agrupa o cliente no Thermometer',
  account         VARCHAR(20)    NOT NULL DEFAULT '' COMMENT 'Conta AX (ex.: C0002424)',
  classification  VARCHAR(40)    NOT NULL DEFAULT '' COMMENT 'OEM, Distributo, End User, LATAM, PFP, Resale, Supplier…',
  status          VARCHAR(40)    NOT NULL DEFAULT '' COMMENT 'Status do Quote Report em minúsculas (confirmed, created, sent, lost, cancelled…)',
  quoted_by       VARCHAR(100)   NOT NULL DEFAULT '' COMMENT 'Quoted By = vendedor interno (ex.: SP0814 - Vinicius Maestrello)',
  sales_person    VARCHAR(100)   NOT NULL DEFAULT '' COMMENT 'Sales Person = vendedor externo',
  amount          DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Valor cotado (soma dos itens, sem arredondar)',
  sales_amt       DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Sales Amt da própria cotação',
  amount_pt       DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Parte do valor cotado na linha PT',
  amount_cp       DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Parte do valor cotado na linha CP',
  amount_mh       DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Parte do valor cotado na linha MH',
  line_read       TINYINT(1)     NOT NULL DEFAULT 0 COMMENT '1 = a linha dos itens foi lida do Quote Report (tinha Product Group)',
  first_seen      DATETIME       NULL COMMENT 'Quando a cotação entrou no programa',
  updated_at      DATETIME       NULL COMMENT 'Última vez que veio num Quote Report (a mais nova sempre ganha)',
  PRIMARY KEY (quote_id),
  KEY ix_quotes_date (quote_date),
  KEY ix_quotes_customer_key (customer_key),
  KEY ix_quotes_class_date (classification, quote_date),
  KEY ix_quotes_sales_person (sales_person),
  KEY ix_quotes_quoted_by (quoted_by),
  KEY ix_quotes_status (status)
) ENGINE=InnoDB COMMENT='Cotações do Quote Report';

-- Pedidos (Sales Id) ligados à cotação. Hoje cada cotação tem no máximo 1, mas o relatório permite vários.
CREATE TABLE IF NOT EXISTS quote_orders (
  quote_id   VARCHAR(20) NOT NULL,
  sales_id   VARCHAR(20) NOT NULL COMMENT 'Sales Id do pedido (ex.: SO00099598)',
  PRIMARY KEY (quote_id, sales_id),
  KEY ix_quote_orders_sales (sales_id),
  CONSTRAINT fk_quote_orders_quote FOREIGN KEY (quote_id) REFERENCES quotes (quote_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB COMMENT='Cotação → pedidos (Sales IDs)';

-- Valor cotado por Product Group (antes era um JSON na coluna "Product groups")
CREATE TABLE IF NOT EXISTS quote_product_groups (
  quote_id       VARCHAR(20)   NOT NULL,
  product_group  VARCHAR(80)   NOT NULL COMMENT 'Nome já limpo (pgLabel), ex.: Sprockets (11)',
  amount         DECIMAL(18,2) NOT NULL DEFAULT 0,
  PRIMARY KEY (quote_id, product_group),
  KEY ix_qpg_group (product_group),
  CONSTRAINT fk_qpg_quote FOREIGN KEY (quote_id) REFERENCES quotes (quote_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB COMMENT='Cotação → valor por Product Group';

-- -------------------------------------------------------------------------------------
-- VENDAS BOOKED (aba Sales do reports.xlsx) — somadas por pedido + dia + classificação + linha + cliente + vendedor
-- -------------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS sales (
  id              BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  sales_id        VARCHAR(20)    NOT NULL DEFAULT '' COMMENT 'Sales Id do pedido',
  booked_date     DATE           NOT NULL,
  customer        VARCHAR(150)   NOT NULL DEFAULT '',
  classification  VARCHAR(40)    NOT NULL DEFAULT '',
  line            VARCHAR(2)     NOT NULL DEFAULT '' COMMENT 'PT, CP, MH ou vazio',
  booked          DECIMAL(18,6)  NOT NULL DEFAULT 0 COMMENT 'Pode ser negativo (devolução / ajuste)',
  salesperson     VARCHAR(100)   NOT NULL DEFAULT '' COMMENT 'Sales Responsible + Name (ex.: SP0178 - Julia Pulz Campo)',
  PRIMARY KEY (id),
  UNIQUE KEY uq_sales_row (sales_id, booked_date, classification, line, customer, salesperson),
  KEY ix_sales_date (booked_date, classification, line),
  KEY ix_sales_salesperson (salesperson, booked_date),
  CONSTRAINT ck_sales_line CHECK (line IN ('', 'PT', 'CP', 'MH'))
) ENGINE=InnoDB COMMENT='Linhas Booked do Sales Report';

-- -------------------------------------------------------------------------------------
-- FOLLOW-UP DE CADA COTAÇÃO (aba Follow-up do data.xlsx) — os campos preenchidos à mão
-- Sem chave estrangeira para quotes de propósito: um Quote Report novo pode tirar a cotação da base,
-- e o follow-up e o histórico dela não podem sumir junto.
-- -------------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS followups (
  quote_id          VARCHAR(20)  NOT NULL,
  customer          VARCHAR(150) NOT NULL DEFAULT '' COMMENT 'Nome do cliente no momento (para ler o histórico sem a cotação)',
  fup               CHAR(1)      NULL COMMENT 'A = próximo dia · B = 2× em 5 dias · C = 2 semanas · D = 1 mês · E = Ongoing (data do Done)',
  fup_since         DATE         NULL COMMENT 'Âncora do ciclo: dia em que o FUP foi marcado ou do último Done',
  win               TINYINT UNSIGNED NULL COMMENT '% Win: 25, 50, 75 ou 99',
  left_out_of_plan  TINYINT(1)   NOT NULL DEFAULT 0 COMMENT '1 = tirada do plano do dia (caixinha)',
  cycle_step        TINYINT UNSIGNED NULL COMMENT 'B: quantos follow-ups do ciclo já foram feitos',
  next_date         DATE         NULL COMMENT 'E · Ongoing: data do próximo follow-up escolhida no Done',
  requisition       VARCHAR(255) NOT NULL DEFAULT '' COMMENT 'Requisição do cliente (ex.: RODRIGO - POLIAS)',
  closed_as         VARCHAR(10)  NULL COMMENT 'Fechada à mão: lost ou cancelled',
  close_reason      VARCHAR(40)  NOT NULL DEFAULT '' COMMENT 'Price, Lead Time, Budget Only, Project Discontinued, Stalled project, No Response',
  updated_at        TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (quote_id),
  KEY ix_followups_fup (fup, next_date),
  KEY ix_followups_closed (closed_as),
  CONSTRAINT ck_followups_fup    CHECK (fup IS NULL OR fup IN ('A', 'B', 'C', 'D', 'E')),
  CONSTRAINT ck_followups_win    CHECK (win IS NULL OR win IN (25, 50, 75, 99)),
  CONSTRAINT ck_followups_closed CHECK (closed_as IS NULL OR closed_as IN ('lost', 'cancelled'))
) ENGINE=InnoDB COMMENT='Campos manuais de cada cotação (FUP, % Win, requisição, fechamento)';

-- Dias em que houve Done em cada cotação (coluna "Follow-ups done"; o programa guarda os últimos 50)
CREATE TABLE IF NOT EXISTS followup_done_days (
  quote_id  VARCHAR(20) NOT NULL,
  done_day  DATE        NOT NULL,
  PRIMARY KEY (quote_id, done_day),
  KEY ix_done_day (done_day),
  CONSTRAINT fk_done_days_followup FOREIGN KEY (quote_id) REFERENCES followups (quote_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB COMMENT='Dias com Done por cotação';

-- -------------------------------------------------------------------------------------
-- ACTIVITIES / HISTORY e anotações (aba Activities do data.xlsx) — só cresce, nunca muda
-- -------------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS activities (
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  act_key     CHAR(64)     NOT NULL COMMENT 'SHA-256 do actKey (data|QT|tipo|nota|para): impede o mesmo registro duas vezes',
  at          DATETIME     NOT NULL COMMENT 'Data e hora (hora local do computador de quem registrou)',
  quote_id    VARCHAR(20)  NOT NULL,
  customer    VARCHAR(150) NOT NULL DEFAULT '',
  kind        VARCHAR(16)  NOT NULL COMMENT 'note, fup, win, done, resched, req, plan, status, amount, order, close, reopen, why',
  from_value  VARCHAR(255) NULL COMMENT 'De (ex.: FUP A, % Win 50, valor antigo)',
  to_value    VARCHAR(255) NULL COMMENT 'Para (ex.: FUP B, data do próximo follow-up, valor novo)',
  note        TEXT         NULL COMMENT 'Texto da anotação (ou o motivo, no fechamento)',
  created_by  VARCHAR(100) NOT NULL DEFAULT '' COMMENT 'Usuário logado que registrou',
  PRIMARY KEY (id),
  UNIQUE KEY uq_activities_key (act_key),
  KEY ix_activities_quote (quote_id, at),
  KEY ix_activities_kind_at (kind, at),
  KEY ix_activities_by (created_by, at),
  CONSTRAINT ck_activities_kind CHECK (kind IN ('note', 'fup', 'win', 'done', 'resched', 'req', 'plan', 'status', 'amount', 'order', 'close', 'reopen', 'why', 'dcs'))
) ENGINE=InnoDB COMMENT='Activities History e anotações (somente inclusão)';

-- -------------------------------------------------------------------------------------
-- METAS (aba Goals do data.xlsx)
-- -------------------------------------------------------------------------------------
-- Meta digitada só para um dia. -1 = "sem meta" digitado à mão num dia útil (DAY_NO_GOAL do app).
CREATE TABLE IF NOT EXISTS day_goals (
  day         DATE          NOT NULL,
  goal        DECIMAL(18,2) NOT NULL,
  updated_at  TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (day),
  CONSTRAINT ck_day_goals CHECK (goal >= 0 OR goal = -1)
) ENGINE=InnoDB COMMENT='Meta de dias específicos';

-- Configurações gerais: meta diária, últimos relatórios importados, versão do programa (aba Info + "Daily goal")
CREATE TABLE IF NOT EXISTS settings (
  name        VARCHAR(60)  NOT NULL,
  value       VARCHAR(255) NOT NULL DEFAULT '',
  updated_at  TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (name)
) ENGINE=InnoDB COMMENT='Meta diária e informações gerais';

INSERT IGNORE INTO settings (name, value) VALUES
  ('data_version', '1'),     -- o servidor soma 1 a cada gravação de follow-up / histórico / metas / usuários
  ('rep_version', '1'),      -- e a cada importação de relatório (as páginas abertas sabem que precisam buscar)
  ('quote_fill_at', ''),
  ('daily_goal', '0'),
  ('last_quote_report', ''),
  ('quote_report_imported_at', ''),
  ('last_sales_report', ''),
  ('sales_report_imported_at', ''),
  ('program_version', '');

-- -------------------------------------------------------------------------------------
-- USUÁRIOS (aba Config cifrada do data.xlsx)
-- A senha nunca é guardada: só o sal e o hash (SHA-256 repetido 5000 vezes, hashPassword do core.js).
-- Assim as senhas de hoje continuam valendo quando o login passar para o servidor.
-- -------------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS users (
  name        VARCHAR(100) NOT NULL COMMENT 'Nome como aparece no login (Admin, ou o vendedor sem o código SP)',
  role        VARCHAR(10)  NULL COMMENT 'admin · manager (gerência) · outside (externo) · inside (interno) · NULL = sem acesso',
  pass_salt   CHAR(32)     NOT NULL DEFAULT '' COMMENT 'Vazio = ainda vai criar a senha',
  pass_hash   CHAR(64)     NOT NULL DEFAULT '',
  changed_at  DATETIME     NULL COMMENT 'Última mudança de senha ou papel (o mais recente vale)',
  PRIMARY KEY (name),
  CONSTRAINT ck_users_role CHECK (role IS NULL OR role IN ('admin', 'manager', 'outside', 'inside'))
) ENGINE=InnoDB COMMENT='Usuários, papel (tela) e senha';

-- Bancos criados antes do formulário DCS / Value Added: a regra dos tipos de histórico passa a aceitar 'dcs'
-- (o formulário inteiro fica na nota, em JSON). Pode rodar de novo sem problema.
ALTER TABLE activities DROP CONSTRAINT IF EXISTS ck_activities_kind;
ALTER TABLE activities ADD CONSTRAINT ck_activities_kind
  CHECK (kind IN ('note', 'fup', 'win', 'done', 'resched', 'req', 'plan', 'status', 'amount', 'order', 'close', 'reopen', 'why', 'dcs'));

-- Sessões do servidor (quem está logado em cada navegador / celular). O cookie guarda um código aleatório;
-- aqui fica só o SHA-256 dele. Zerar a senha de alguém apaga as sessões dessa pessoa.
CREATE TABLE IF NOT EXISTS sessions (
  token_hash  CHAR(64)     NOT NULL,
  name        VARCHAR(100) NOT NULL,
  created_at  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  last_seen   DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  expires_at  DATETIME     NOT NULL,
  user_agent  VARCHAR(200) NOT NULL DEFAULT '',
  PRIMARY KEY (token_hash),
  KEY ix_sessions_name (name),
  KEY ix_sessions_expires (expires_at)
) ENGINE=InnoDB COMMENT='Sessões de login do servidor';

-- Avisos no celular / PC (Web Push): um registro por aparelho que tocou em "Ativar avisos".
-- endpoint = endereço do serviço de avisos do navegador (Google, Apple, Microsoft, Mozilla); p256dh / auth = chaves
-- do aparelho para cifrar o aviso. O aparelho que desativa (ou troca de navegador) sai sozinho no próximo envio.
CREATE TABLE IF NOT EXISTS push_subs (
  id          INT UNSIGNED NOT NULL AUTO_INCREMENT,
  endpoint    VARCHAR(700) NOT NULL,
  name        VARCHAR(100) NOT NULL COMMENT 'Pessoa logada quando ativou',
  p256dh      VARCHAR(200) NOT NULL,
  auth        VARCHAR(60)  NOT NULL,
  user_agent  VARCHAR(200) NOT NULL DEFAULT '',
  created_at  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  last_ok     DATETIME     NULL COMMENT 'Último aviso entregue',
  PRIMARY KEY (id),
  UNIQUE KEY ux_push_endpoint (endpoint),
  KEY ix_push_name (name)
) ENGINE=InnoDB COMMENT='Aparelhos que recebem os avisos';

-- -------------------------------------------------------------------------------------
-- VISÃO: situação de cada cotação, com as mesmas regras do core.js (quoteState / quoteStateM)
--   ganha     = status com "confirm" ou com pedido (Sales Id)
--   cancelada = status com "cancel"
--   perdida   = status com "lost", "perd" ou "declin"
--   aberta    = o resto; aberta fechada à mão (followups.closed_as) vira lost / cancelled
-- -------------------------------------------------------------------------------------
CREATE OR REPLACE VIEW v_quote_state AS
SELECT
  q.quote_id,
  q.quote_date,
  q.customer,
  q.classification,
  q.quoted_by,
  q.sales_person,
  q.amount,
  CASE
    WHEN q.status LIKE '%confirm%' OR EXISTS (SELECT 1 FROM quote_orders o WHERE o.quote_id = q.quote_id) THEN 'won'
    WHEN q.status LIKE '%cancel%' THEN 'cancelled'
    WHEN q.status LIKE '%lost%' OR q.status LIKE '%perd%' OR q.status LIKE '%declin%' THEN 'lost'
    ELSE 'open'
  END AS report_state,
  CASE
    WHEN q.status LIKE '%confirm%' OR EXISTS (SELECT 1 FROM quote_orders o WHERE o.quote_id = q.quote_id) THEN 'won'
    WHEN q.status LIKE '%cancel%' THEN 'cancelled'
    WHEN q.status LIKE '%lost%' OR q.status LIKE '%perd%' OR q.status LIKE '%declin%' THEN 'lost'
    WHEN f.closed_as IN ('lost', 'cancelled') THEN f.closed_as
    ELSE 'open'
  END AS state,
  f.fup,
  f.win,
  f.next_date,
  f.close_reason
FROM quotes q
LEFT JOIN followups f ON f.quote_id = q.quote_id;

-- -------------------------------------------------------------------------------------
-- O histórico só cresce: nenhuma linha de activities pode ser alterada ou apagada
-- (mesma regra do programa: "notes can't be edited or deleted").
-- -------------------------------------------------------------------------------------
DROP TRIGGER IF EXISTS tr_activities_no_update;
DROP TRIGGER IF EXISTS tr_activities_no_delete;

DELIMITER $$
CREATE TRIGGER tr_activities_no_update BEFORE UPDATE ON activities
FOR EACH ROW
BEGIN
  SIGNAL SQLSTATE '45000' SET MESSAGE_TEXT = 'Activities History cannot be changed';
END$$

CREATE TRIGGER tr_activities_no_delete BEFORE DELETE ON activities
FOR EACH ROW
BEGIN
  SIGNAL SQLSTATE '45000' SET MESSAGE_TEXT = 'Activities History cannot be deleted';
END$$
DELIMITER ;
