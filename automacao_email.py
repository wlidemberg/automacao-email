import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# CONFIGURAÇÕES DE EMAIL

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_REMETENTE = 'seuemail@gmail.com'
EMAIL_SENHA = '***************'

NOME_EMPRESA = 'Assescor Assessoria e Corretagem'
NOME_RESPONSAVEL = 'Juan Pablo Montoya'
TELEFONE_CONTATO = 21999999999

# FUNÇÃO DE ENVIO DE EMAIL
def enviar_email(destinatario:str, assunto:str, mensagem:str) -> None:
    try:
        # Monta o envelope/cabeçalhos do e-mail
        msg = MIMEMultipart()
        msg['from'] = EMAIL_REMETENTE
        msg['to'] = destinatario
        msg['subject'] = assunto
        
        # Anexa o corpo da mensagem (Texto Simples)
        msg.attach(MIMEText(mensagem, 'html'))

        # Abre a conexão com servidor SMTP
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.sendmail(EMAIL_REMETENTE, destinatario, msg.as_string())
        print(f'Email enviado para {destinatario} - Assunto {assunto}')
    except Exception as e:
        print(f'Erro ao enviar email para {destinatario}:{e}')  

def carregar_template(caminho_template, **kwargs):
    # Lê um template HTML e substitui os placeholders pelas variáveis passadas.
    # kwargs: dicionário com valores, ex: nome = 'João', valor='200'
    with open(caminho_template, 'r', encoding='utf-8') as tpl:
        conteudo = tpl.read()

    for chave, valor in kwargs.items():
        conteudo = conteudo.replace(f"{{{{{chave}}}}}", str(valor))            
    return conteudo


# REGRAS DE COBRNÇA
def definir_mensagem(nome_empresa, nome_responsavel, email_contato, telefone_contato, nome, data_venc, telefone, endereco, valor, dias_atraso, pago):
    if pago.lower() == 'sim':
        assunto = "Pagamento Confirmado - Parabéns!"
        mensagem = carregar_template(
            "templates/template.html",
             nome_empresa=nome_empresa, email_contato=email_contato, telefone_contato=telefone_contato,
                nome=nome, data_venc=data_venc, valor=valor, endereco=endereco, dias_atraso=dias_atraso)
        
        return assunto, mensagem
    
    else:
        if dias_atraso == 2:
            assunto = "Comunicado de Cobrança – Aluguel em Atraso 2 dias."
            mensagem = carregar_template(
                "templates/template_cobranca.html",
                nome_empresa=nome_empresa, email_contato=email_contato, telefone_contato=telefone_contato,
                nome=nome, data_venc=data_venc, valor=valor, endereco=endereco, dias_atraso=dias_atraso)
            
            return assunto, mensagem
        
        elif dias_atraso == 5:
            assunto = "Comunicado de Cobrança - Aluguel em Atraso 5 dias."
            mensagem = carregar_template(
                "templates/template_cobranca.html",
                nome_empresa=nome_empresa, email_contato=email_contato, telefone_contato=telefone_contato,
                nome=nome, data_venc=data_venc, valor=valor, endereco=endereco, dias_atraso=dias_atraso)

            return assunto, mensagem
        elif dias_atraso == 15:
            assunto = 'Notificação – Encaminhamento a Protesto'
            mensagem = carregar_template(
                "templates/template_protesto.html",
                nome_empresa=nome_empresa, email_contato=email_contato, telefone_contato=telefone_contato,
                nome=nome, data_venc=data_venc, valor=valor, endereco=endereco, dias_atraso=dias_atraso)
            
            return assunto, mensagem
        
        elif dias_atraso == 20:
            assunto = 'Notificação – Extrajudicial'
            mensagem = carregar_template(
                "templates/template_notificacao.html",
                nome_empresa=nome_empresa, nome_responsavel=nome_responsavel, email_contato=email_contato, telefone_contato=telefone_contato,
                nome=nome, data_venc=data_venc, valor=valor, endereco=endereco, dias_atraso=dias_atraso)
            
            return assunto, mensagem
        else:
            return (None, None)
        
# PROCESSAMENTO DA PLANILHA
def processar_planilha(caminho_planilha):
    
    df = pd.read_excel(caminho_planilha) # Carrega o excel para um dataframe
    hoje = datetime.now().date() # Data de hoje

    # Percorre todas as linhas
    for _, row in df.iterrows():
        nome_empresa = NOME_EMPRESA
        nome_responsavel = NOME_RESPONSAVEL
        email_contato = EMAIL_REMETENTE
        telefone_contato = TELEFONE_CONTATO
        nome = row['Nome']
        email = row['Email']
        telefone = row['Telefone']
        endereco = row['Endereço']
        valor = row['Valor']
        valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        data_venc = row['Data_Vencimento'].date()
        data_formatada = data_venc.strftime("%d/%m/%Y")
        pago = str(row['Pago']).strip()

        dias_atraso = (hoje - data_venc).days
        assunto, mensagem = definir_mensagem(nome_empresa, nome_responsavel, email_contato, telefone_contato, nome, data_formatada, telefone, endereco, valor_formatado, dias_atraso, pago)  
        if assunto and mensagem:
            enviar_email(email, assunto, mensagem) 

# EXECUNTANDO O SISTEMA
if __name__ == "__main__":
    processar_planilha('cobranca.xlsx')                    
