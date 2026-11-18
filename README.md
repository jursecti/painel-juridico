# Painel do Jurídico – SECTI/PE

Site de consulta aos instrumentos e atos acompanhados pelo Setor Jurídico da SECTI.
Uma rotina lê a planilha JURIDICO (somente leitura), publica apenas as colunas públicas,
calcula os prazos e envia um e-mail semanal ao Jurídico.

## Como funciona

- `config.json` – abas lidas, colunas publicadas, regras de prazo e destinatário do e-mail.
- `scripts/build.py` – lê a planilha, filtra, calcula e gera `site/data.json`.
- `site/index.html` – o painel (busca, fichas, servidores, prazos).
- `.github/workflows/painel.yml` – agenda: dias úteis de duas em duas horas; e-mail às segundas, 7h30.

Só entra no site a coluna que estiver listada em `publicas` no `config.json`.
CPF, telefone, e-mail e matrícula são removidos de qualquer texto antes da publicação.

## Instalação (uma única vez)

1. **Repositório.** Criar o repositório público `painel-juridico`, vazio.
2. **Credencial de leitura do Google.**
   - Em console.cloud.google.com, criar um projeto e ativar a "Google Sheets API".
   - Em "IAM e administrador > Contas de serviço", criar uma conta de serviço.
   - Na conta criada, em "Chaves", adicionar uma chave do tipo JSON. Um arquivo será baixado.
   - Compartilhar a planilha JURIDICO com o e-mail da conta de serviço, como **Leitor**.
3. **Chave de envio do Brevo.** Em brevo.com, confirmar o remetente (Remetentes e IP) e, em
   "SMTP e API > Chaves SMTP", gerar uma chave. Anotar também o login SMTP exibido na tela.
4. **Segredos no GitHub.** Em "Settings > Secrets and variables > Actions", criar:
   - `GOOGLE_CREDENTIALS` – o conteúdo inteiro do arquivo JSON;
   - `BREVO_SMTP_LOGIN` – o login SMTP do Brevo;
   - `BREVO_SMTP_KEY` – a chave SMTP gerada;
   - `EMAIL_REMETENTE` – o endereço do remetente confirmado no Brevo.
5. **Publicação.** Em "Settings > Pages", escolher "GitHub Actions" como fonte.
6. **Primeira execução.** Em "Actions > Atualizar painel", clicar em "Run workflow".

## No dia a dia

- **Atualizar na hora:** "Actions > Atualizar painel > Run workflow".
- **Enviar o e-mail fora da segunda-feira:** o mesmo botão, marcando "Enviar também o e-mail".
- **Mudar regra, coluna ou antecedência:** editar o `config.json`.

## Teste local

```
python tests/make_fixture.py
python scripts/build.py --fixture tests/fixture --out site/data.json --previa-email /tmp/email.html
python -m http.server -d site
```
