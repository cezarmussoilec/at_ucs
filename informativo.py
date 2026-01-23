import pathlib
import win32com.client as client
from utils import get_logger

logger = get_logger()

def envia_email(nome, email, login_cpf, senha_pad):
    try:
        logger.info("Iniciando processo de envio de email")
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        dir_base = pathlib.Path.home() / "Pictures"
        acesso_ucs = dir_base / "acesso_ucs.png"
        assinatura = dir_base / "Cezar Mussoi AT.png"
        acesso_ucs = str(acesso_ucs.absolute())
        assinatura = str(assinatura.absolute())
        message.To = f"{email}"
        message.Subject = "Universidade Corporativa Senior (UCS)"
        anexo_ucs = message.Attachments.Add(acesso_ucs)
        anexo_ucs.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "acesso_ucs")
        anexo_assinat = message.Attachments.Add(assinatura)
        anexo_assinat.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "assinatura")
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
                <li>Clique na opção "Cliente";</li>
                <img src='cid:acesso_ucs'>
                <li>
                Informe seu login e senha:
                <ul>
                    <li>Login: {login_cpf}</li>
                    <li>Senha: {senha_pad}</li>
                </ul>
                </li>
                <li>No primeiro acesso, será solicitada uma nova senha.</li>
                <li>
                    Ao acessar a plataforma da UCS, você já pode fazer os cursos que estão liberados para você.
                    Em caso de dúvidas entre em contato.<br>
                </li>
            </ol>
            <p>Caso deseje outro curso, vá na aba 'Catálogo' e escolha o curso que deseja realizar. 
            Mande um e-mail para nós com o nome do curso, que liberamos o acesso para você.</p> 
            <p>Atenciosamente,</p>
            <img src='cid:assinatura'>           
        """)
        message.HTMLBody = html_body
        message.Save()
        message.Send()
        logger.info(f"Email enviado com sucesso para {nome} ({email})")
    except Exception as e:
        logger.error(f"Erro no envio do email para {nome} ({email})")
