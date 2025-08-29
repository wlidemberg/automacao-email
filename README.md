# Automação de Emails para Imobiliária

Este projeto é uma automação em **Python** que lê uma planilha do Excel com informações de clientes e envia **emails automáticos** de acordo com o status de pagamento. Ele pode enviar:

- Email de agradecimento para clientes que estão **em dia** com o pagamento.
- Email de **cobrança** quando o pagamento estiver atrasado (2 dias, 5 dias).
- Email de **notificação de protesto** (15 dias de atraso).
- Email de **notificação extra judicial** (35 dias de atraso).

O sistema foi pensado para **imobiliárias**, mas pode ser adaptado para qualquer tipo de cobrança recorrente.

---

## Estrutura de Pastas

automacao_email/
│
├─ templates/
│ ├─ template.html
│ ├─ template_cobranca.html
│ ├─ template_notificacao.html
│ └─ template_protesto.html
│
├─ automacao_email.py
├─ cobranca.xlsx
└─ README.md

yaml
Copiar código

- **templates/**: Contém os arquivos HTML que serão enviados por email.
- **automacao_email.py**: Script principal que processa a planilha e envia emails.
- **cobranca.xlsx**: Planilha com os dados dos clientes.
- **README.md**: Documentação do projeto.

---

## Pré-requisitos

- Python 3.8 ou superior
- Contas de email (Gmail ou outro servidor SMTP) para envio dos emails.
- Bibliotecas Python:
  - `pandas`
  - `openpyxl` (para leitura de arquivos .xlsx)

Você pode instalar as dependências usando:

```bash
pip install -r requirements.txt
requirements.txt sugerido:

nginx
Copiar código
pandas
openpyxl
Estrutura da Planilha
A planilha cobranca.xlsx deve ter as seguintes colunas:

Nome	Email	Telefone	Endereço	Data_Vencimento	Pago

Nome: Nome do cliente

Email: Email do cliente

Telefone: Telefone de contato

Endereço: Endereço do imóvel

Data_Vencimento: Data de vencimento da mensalidade (formato AAAA-MM-DD)

Pago: "Sim" ou "Não"

Configurações do Projeto
No arquivo automacao_email.py você pode alterar:

python
Copiar código
SMTP_SERVER = 'smtp.gmail.com'  # Servidor SMTP
SMTP_PORT = 587                 # Porta SMTP
EMAIL_REMETENTE = 'seu_email@gmail.com'  # Email remetente
EMAIL_SENHA = 'sua_senha_app'           # Senha do email ou App Password (Gmail)
NOME_EMPRESA = 'Assescor Assessoria e Corretagem'  # Nome da sua empresa
TELEFONE_CONTATO = 21968045339                     # Telefone para contato
Substitua pelo email e senha válidos.

Para Gmail, é recomendado criar uma App Password.

Como Funciona
Processamento da Planilha
O script lê cada linha da planilha cobranca.xlsx e verifica:

Se Pago == "Sim" → envia email de agradecimento.

Se Pago == "Não" → calcula o número de dias de atraso e envia:

2 dias → Email de cobrança

5 dias → Email de cobrança

15 dias → Notificação de protesto

35 dias → Notificação extra judicial

Envio de Email

O corpo do email é carregado de um template HTML na pasta templates/.

Placeholders como {{ nome }}, {{ valor }}, {{ data_venc }} são substituídos pelo Python.

Executando o Script Manualmente
No Windows ou Linux, abra o terminal na pasta do projeto e execute:

bash
Copiar código
python automacao_email.py
Certifique-se de que o Excel (cobranca.xlsx) esteja na mesma pasta do script.

O script enviará emails automaticamente conforme as regras de atraso.

Agendamento Automático
Para que a automação rode diariamente sem precisar abrir o script manualmente:

Windows (Agendador de Tarefas)
Abra o Agendador de Tarefas.

Clique em Criar Tarefa.

Aba Geral: Nomeie a tarefa (ex: "Automação de Emails").

Aba Disparadores: Clique em Novo e escolha Diariamente.

Aba Ações:

Ação: Iniciar um programa

Programa/script: python

Adicione argumentos: automacao_email.py

Iniciar em: Caminho da pasta do projeto.

Aba Condições e Configurações: configure conforme necessário.

Salve e teste a tarefa.

Linux (Cron)
Abra o terminal.

Edite o crontab:

bash
Copiar código
crontab -e
Adicione uma linha para executar diariamente às 8h da manhã:

bash
Copiar código
0 8 * * * /usr/bin/python3 /caminho/para/automacao_email.py
Substitua /usr/bin/python3 pelo caminho do seu Python.

Substitua /caminho/para/automacao_email.py pelo caminho completo do script.

Salve e saia.

O cron executará automaticamente o script todos os dias no horário definido.

Personalização
Para alterar nomes, emails ou telefone da empresa, edite as variáveis no início do automacao_email.py.

Para trocar templates de email, edite os arquivos em templates/.

Para ajustar os dias de atraso ou criar novas regras, modifique a função definir_mensagem.

Licença
Este projeto está sob a licença MIT, permitindo uso, cópia e modificação, desde que o crédito ao autor seja mantido.