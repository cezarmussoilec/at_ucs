import os
import pathlib
import win32com.client as client
from utils import get_logger

logger = get_logger()

def selecionar_conta_outlook(outlook, email_remetente):
    for account in outlook.Session.Accounts:
        smtp = getattr(account, "SmtpAddress", "")
        if smtp and smtp.lower() == email_remetente.lower():
            return account

    raise RuntimeError(f"Conta Outlook não encontrada: {email_remetente}")

def configurar_remetente(message, outlook, email_remetente):
    if not email_remetente:
        return

    try:
        message.SendUsingAccount = selecionar_conta_outlook(outlook, email_remetente)
        logger.info(f"Remetente configurado pela conta Outlook: {email_remetente}")
    except Exception as erro_conta:
        remetente_delegado = outlook.Session.CreateRecipient(email_remetente)
        if not remetente_delegado.Resolve():
            raise RuntimeError(f"Remetente delegado não resolvido pelo Outlook: {email_remetente}") from erro_conta

        logger.info(
            f"{email_remetente} não está adicionado como conta no Outlook; usando permissao de envio delegado: {erro_conta}"
        )
        message.SentOnBehalfOfName = email_remetente

def resolver_destinatarios(message):
    if not message.Recipients.ResolveAll():
        nao_resolvidos = []
        for recipient in message.Recipients:
            if not recipient.Resolved:
                nao_resolvidos.append(recipient.Name)

        raise RuntimeError(f"Destinatario(s) nao resolvido(s) pelo Outlook: {', '.join(nao_resolvidos)}")

def configurar_copia_oculta(message):
    bcc = (os.getenv("INFORMATIVE_EMAIL_BCC") or "").strip()
    if not bcc:
        return

    message.BCC = bcc
    logger.info(f"Copia oculta configurada para informativo: {bcc}")

def anexar_assinatura(message):
    assinatura = pathlib.Path.home() / "Pictures" / "Cezar Mussoi AT.png"
    assinatura = str(assinatura.absolute())
    anexo_assinat = message.Attachments.Add(assinatura)
    anexo_assinat.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "assinatura")

def anexar_acesso_ucs(message):
    acesso_ucs = pathlib.Path.home() / "Pictures" / "acesso_ucs.png"
    acesso_ucs = str(acesso_ucs.absolute())
    anexo_ucs = message.Attachments.Add(acesso_ucs)
    anexo_ucs.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "acesso_ucs"
    )

def lista_cursos_html(cursos):
    if not cursos:
        return "<li>Nenhum curso novo foi atribuido.</li>"

    return "".join(f"<li>{curso}</li>" for curso in cursos)

# Envia email de informativo para o usuário
def envia_email(nome, email, login_cpf, senha_padrao):
    try:
        logger.info(f"Iniciando processo de envio de email para {nome} ({email})")
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        configurar_remetente(message, outlook, os.getenv("OUTLOOK_FROM"))
        message.To = f"{email}"
        configurar_copia_oculta(message)
        anexar_acesso_ucs(message)
        message.Subject = "Universidade Corporativa Senior (UCS)"
        # anexar_assinatura(message)
        html_body = (f"""
            <p>Olá, {nome} ({email})</p>
            <p>Segue seu acesso à Universidade Corporativa Senior (UCS).</p>
            <p>Aproveite o curso e não se esqueça de finalizá-lo em tempo hábil. Estamos acompanhando seu progresso.</p>
            <ol>
                <li>
                    Acesse o endereço: 
                    <a href='https://ucsonline.senior.com.br/lms/index.html#/home' target='_blank'>
                    https://ucsonline.senior.com.br/lms/index.html#/home
                    </a>
                </li>
                <li>Clique em "Login";</li>
                <li>Selecione a opção "Cliente";</li>
                <img src='cid:acesso_ucs'>
                <li>
                Informe seu login e senha:
                <ul>
                    <li>Login: {login_cpf}</li>
                    <li>Senha: {senha_padrao}</li>
                </ul>
                </li>
                <li>No primeiro acesso, será solicitada uma nova senha.</li>
                <li>
                    Ao acessar a plataforma da UCS, você já pode fazer os cursos que estão liberados para você.
                    Em caso de dúvidas entre em contato através de time.progamacao@leclair.com.br.<br>
                </li>
            </ol>
            <p>Caso deseje outro curso, vá na aba 'Catálogo' e escolha o curso que deseja realizar. 
            Mande um e-mail para nós com o nome do curso, que liberamos o acesso para você.</p>
            <p>* Este é um e-mail automático, favor não responder *</p>
            <p>Atenciosamente,</p>
            <p>Equipe de TI Leclair</p>           
        """)
        message.HTMLBody = html_body
        resolver_destinatarios(message)
        message.Save()
        message.Send()
        logger.info(f"Email enviado com sucesso para {nome} ({email})")
    except Exception as e:
        logger.error(f"Erro no envio do email para {nome} ({email}): {e}")
        raise

def envia_email_usuario_existente(nome, email, cursos_atribuidos):
    try:
        logger.info(f"Iniciando envio de email para usuario existente: {nome} ({email})")
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        configurar_remetente(message, outlook, os.getenv("OUTLOOK_FROM"))
        message.To = f"{email}"
        configurar_copia_oculta(message)
        anexar_acesso_ucs(message)
        message.Subject = "Novos cursos liberados na Universidade Corporativa Senior (UCS)"
        # anexar_assinatura(message)
        html_body = (f"""
            <p>Olá, {nome} ({email})</p>
            <p>Seu cadastro na Universidade Corporativa Senior (UCS) já estava ativo.</p>
            <ol>
                <li>
                    Acesse a plataforma pelo endereço: 
                    <a href='https://ucsonline.senior.com.br/lms/index.html#/home' target='_blank'>
                    https://ucsonline.senior.com.br/lms/index.html#/home
                    </a>
                </li>
                <li>Clique em "Login";</li>
                <li>Selecione a opção "Cliente";</li>
                <img src='cid:acesso_ucs'>
                <li>No primeiro acesso, você definiu a senha do seu usuário.</li>
            </ol>
            <p>Os cursos abaixo foram liberados para você:</p>
            <ul>
                {lista_cursos_html(cursos_atribuidos)}
            </ul>
            <p>Caso não lembre sua senha, utilize a opção de recuperação de senha na tela de login. 
            Se não obtiver sucesso, entre em contato através de time.programacao@leclair.com.br</p>
            <p>* Este é um e-mail automático, favor não responder *</p>
            <p>Atenciosamente,</p>
            <p>Equipe de TI Leclair</p>
        """)
        message.HTMLBody = html_body
        resolver_destinatarios(message)
        message.Save()
        message.Send()
        logger.info(f"Email de usuario existente enviado com sucesso para {nome} ({email})")
    except Exception as e:
        logger.error(f"Erro no envio do email para usuario existente {nome} ({email}): {e}")
        raise
