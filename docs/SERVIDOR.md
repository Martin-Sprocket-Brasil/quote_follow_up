# Quote Follow Up com servidor (MariaDB + túnel Cloudflare)

```
celular / computador  ──►  túnel Cloudflare (https)  ──►  servidor  (node servidor/server.mjs, porta 8080)  ──►  MariaDB
```

O servidor (`servidor\server.mjs`) entrega a página e a API, faz o login e é o único que fala com o banco. O MariaDB nunca fica exposto na internet.

**O que muda para quem usa**

- Não tem mais pasta nem botão Save: cada mudança vai sozinha para o servidor em cerca de 1 segundo.
- O que os colegas mudam aparece sozinho em até 20 segundos.
- Funciona no celular.
- As senhas continuam as mesmas.
- Sem internet, a mudança fica guardada na página (ícone de nuvem vermelha) e vai sozinha quando a conexão volta.

**A versão antiga continua funcionando:** aberto como arquivo (ou pelo Live Server), o `pagina\QuoteFollowUp.html` usa a pasta dos arquivos Excel, como antes.

---

## Passo a passo (uma vez, no computador que fica ligado)

### 1. Pasta fixa

Copie a pasta do programa para **`C:\QuoteFollowUp`**.

Fora do OneDrive: o "&" do nome da pasta da empresa atrapalha os comandos do Windows, e o OneDrive pode deixar os arquivos só na nuvem.

### 2. Bibliotecas e página

Dê dois cliques em `build.bat`. Ele instala as bibliotecas que faltarem (express, mariadb, compression e, para o formulário DCS / Value Added, pdf-lib) e gera o `pagina\QuoteFollowUp.html`.

### 3. Atualizar o banco

O `banco\schema.sql` novo acrescenta a tabela `sessions` (logins), os contadores de versão e libera o tipo `dcs` no histórico (o formulário DCS / Value Added). Não apaga nada e pode ser rodado de novo quantas vezes quiser.

```
mariadb -u root -p < banco\schema.sql
```

Para trazer os dados mais recentes das planilhas antes de trocar de vez, entre como **Admin** → aba **Dados** → **Importar da versão anterior** e escolha os dois arquivos (ou repita o `importar-banco.bat` e o `mariadb -u root -p < import.sql`). Veja `BANCO-DE-DADOS.md`.

### 4. Usuário do banco só para o programa

Não use o root no servidor. Abra o `mariadb -u root -p` e rode o comando abaixo, trocando a senha:

```sql
CREATE USER 'qfu_app'@'localhost' IDENTIFIED BY 'uma-senha-forte-aqui';
GRANT SELECT, INSERT, UPDATE, DELETE, CREATE ON quote_follow_up.* TO 'qfu_app'@'localhost';
```

### 5. Configuração

Dê dois cliques em `iniciar-servidor.bat`. Na primeira vez, ele cria o `servidor\server.config.json` e abre o arquivo no Bloco de Notas. Coloque a senha do passo 4 em `"password"` e salve.

| Campo | O que é |
|---|---|
| `port` | Porta do servidor (padrão 8080). |
| `host` | `127.0.0.1` = só esta máquina (o túnel roda nela). Use `0.0.0.0` só para testar na rede do escritório. |
| `sessionDays` | Por quantos dias o login dura sem uso (cada uso estende). |
| `db` | Endereço, usuário, senha e nome do banco. |

> O `servidor\server.config.json` tem a senha do banco. Ele já está no `.gitignore`: não mande para o GitHub.

### 6. Ligar e testar

Rode `iniciar-servidor.bat` de novo. Deve aparecer:

```
Quote Follow Up · servidor no ar em http://127.0.0.1:8080
```

Abra **http://localhost:8080** no Chrome e entre com a sua senha de sempre. Para entrar como Admin, use Ctrl+Shift+A, como antes.

### 7. Deixar ligado sempre (inicia com o Windows)

Abra o **Prompt de Comando como administrador** e rode:

```
schtasks /Create /TN "Quote Follow Up" /TR "C:\QuoteFollowUp\iniciar-servidor.bat /servico" /SC ONSTART /RU SYSTEM /RL HIGHEST /F
schtasks /Run /TN "Quote Follow Up"
```

- Nesse modo não tem janela. O registro fica em `C:\QuoteFollowUp\servidor\servidor.log`, e se o servidor cair ele religa sozinho em 10 segundos.
- Para parar: `schtasks /End /TN "Quote Follow Up"`.
- Para tirar do início do Windows: `schtasks /Delete /TN "Quote Follow Up" /F`.
- O MariaDB também precisa iniciar com o Windows. O instalador dele já faz isso.

### 8. Túnel Cloudflare

No painel **Cloudflare Zero Trust → Networks → Tunnels**, abra o seu túnel (o do `cloudflared.exe service install`) e vá em **Public Hostname → Add a public hostname**:

| Campo | Valor |
|---|---|
| Subdomain / Domain | Ex.: `quotes` · `suaempresa.com.br` |
| Service · Type | `HTTP` |
| Service · URL | `localhost:8080` |

Pronto: **https://quotes.suaempresa.com.br** abre o programa de qualquer lugar, inclusive do celular.

### 9. Cloudflare Access (muito recomendado)

Sem o Access, qualquer pessoa com o link vê a tela de login, e nela a lista de nomes dos vendedores. Com o Access, só entra quem receber um código no e-mail da empresa.

1. Abra **Zero Trust → Access → Applications → Add an application → Self-hosted**.
2. Em **Application domain**, use o mesmo endereço do passo 8.
3. Em **Policy**, escolha **Allow** com a regra **Emails ending in** `@suaempresa.com.br`. Se preferir, liste os e-mails um a um.
4. Em **Session duration**, use 1 mês: assim o celular não pede o código toda hora.

É grátis até 50 usuários.

---

## No dia a dia

- **Celular:** abra o endereço no Chrome ou no Safari, toque em **Adicionar à tela inicial** e use como um aplicativo.
- **Ícones no topo:** ✓ verde = tudo salvo (tocar busca as mudanças dos outros na hora) · ↻ girando = salvando · nuvem vermelha = sem conexão, tenta de novo sozinho.
- **Importar relatórios** (gerência), na aba Data ou no topo: o relatório é lido no computador de quem importa, e só as cotações novas ou alteradas vão para o servidor. Os outros recebem em até 20 segundos.
- **Cópia em Excel** (gerência): menu do usuário → *Download a copy (Excel)*. É o mesmo arquivo de dados da versão com pasta, sem as senhas.
- **Senhas:**
  - continuam as mesmas;
  - quem não tem senha cria uma no primeiro acesso, agora com no mínimo 6 caracteres;
  - o Admin zera a senha na aba Data, e a pessoa sai de todos os aparelhos na hora;
  - depois de 8 senhas erradas, o login fica bloqueado por 15 minutos.

## O que o servidor garante

- **Vendedor só grava nas próprias cotações:** externo pelo Sales Person, interno pelo Quoted By, com o mesmo agrupamento de nomes da tela. Tentativas fora da carteira são ignoradas.
- **Os três perfis** (vendedor externo, interno e gerência) importam o Quote Report e o Sales Report. **Só a gerência** muda as metas. **Só o Admin** escolhe a tela de cada pessoa e zera senhas.
- **Gerência só visualiza as cotações:** follow-up, % Win, notas, fechamento, plano do dia e formulário DCS vindos da gerência são recusados (só os vendedores gravam).
- O **histórico** continua só crescendo: o banco recusa alteração e exclusão, e o mesmo registro mandado duas vezes não entra de novo.
- **Duas pessoas gravando ao mesmo tempo:** as gravações entram uma de cada vez, e o que cada uma mudou em cotações diferentes fica. Na mesma cotação, vale a última.
- **Login:**
  - a senha nunca trafega nem fica guardada em texto, só o hash;
  - o cookie de sessão é `HttpOnly` e, pelo túnel, `Secure`;
  - pedidos de outros sites são recusados.

## Limitação desta primeira versão

Todo mundo que entra recebe **todas** as cotações, porque o Thermometer de cada cliente é calculado com elas. A tela de cada vendedor continua mostrando só a carteira dele, e o servidor só deixa gravar nela. Mandar para cada vendedor só as cotações dele é o próximo passo, se a empresa precisar.

## Problemas comuns

| Mensagem | O que fazer |
|---|---|
| `Não foi possível falar com o banco de dados` | Confira o MariaDB ligado e o usuário e a senha no `servidor\server.config.json`. |
| `a tabela sessions não existe…` | Rode o `banco\schema.sql` novo como root (passo 3). |
| `[AVISO] O banco ainda não aceita o formulário DCS / Value Added` ou, na página, *The database needs an update…* | Rode o `banco\schema.sql` novo como root (passo 3). Não precisa reiniciar: em até 1 minuto o servidor percebe e a página reenvia o formulário sozinha. |
| `EADDRINUSE` | A porta 8080 já está em uso, talvez por outro servidor aberto. Feche o outro ou troque o `port`. |
| Página abre, mas pede a pasta de dados | Ela foi aberta como arquivo, não pelo servidor. Use http://localhost:8080 ou o endereço do túnel. |
| Túnel mostra erro 502 | O servidor não está rodando nesta máquina (passo 6 ou 7). |

Depois de mudar o `codigo\app.jsx`, rode o `build.bat`. Não precisa reiniciar o servidor: ele sempre entrega o `pagina\QuoteFollowUp.html` mais novo. Depois de mudar o `codigo\core.js` ou o `servidor\server.mjs`, reinicie o servidor.
