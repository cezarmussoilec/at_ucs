import os
import traceback

import win32com.client as client

from informativo import configurar_remetente, resolver_destinatarios
from utils import get_logger, get_log_file_path, get_error_screenshots

logger = get_logger()


def env_bool(nome, padrao=True):
    valor = os.getenv(nome)
    if valor is None:
        return padrao

    return valor.strip().lower() in ("1", "true", "sim", "s", "yes", "y")


def destinatarios_alerta():
    valor = os.getenv("ALERT_EMAIL_TO", "")
    return [email.strip() for email in valor.replace(",", ";").split(";") if email.strip()]


def remetente_alerta():
    return os.getenv("ALERT_EMAIL_FROM") or os.getenv("OUTLOOK_FROM")


def formatar_lista(valores):
    if not valores:
        return "Nenhum."

    return "; ".join(str(valor) for valor in valores)


def resumo_linhas(usuarios, limite=10):
    if not usuarios:
        return "Nenhum."

    linhas = []
    for usuario in usuarios[:limite]:
        linha = (
            f"- {usuario.get('nome')} | CPF: {usuario.get('cpf')} | Email: {usuario.get('email')}\n"
            f"    Cursos solicitados: {formatar_lista(usuario.get('cursos_solicitados'))}\n"
            f"    Cursos já atribuidos anteriormente: {formatar_lista(usuario.get('cursos_ja_atribuidos'))}\n"
            f"    Cursos atribuídos nesta execução: {formatar_lista(usuario.get('cursos_atribuidos_execucao'))}\n"
        )
        if usuario.get("erro"):
            linha += f"\n  Erro: {usuario.get('erro')}"
        if usuario.get("screenshot"):
            linha += f"\n  Screenshot: {usuario.get('screenshot')}"
        linhas.append(linha)

    restantes = len(usuarios) - limite
    if restantes > 0:
        linhas.append(f"- ... e mais {restantes} usuário(s).")

    return "\n".join(linhas)


def screenshots_do_resumo(resumo):
    caminhos = []
    for chave in ("ok", "n_ok"):
        for usuario in resumo.get(chave, []):
            screenshot = usuario.get("screenshot")
            if screenshot:
                caminhos.append(screenshot)

    caminhos.extend(get_error_screenshots())
    return list(dict.fromkeys(caminhos))


def enviar_email_alerta(assunto, corpo, anexos_extras=None):
    destinatarios = destinatarios_alerta()
    if not destinatarios:
        logger.info("ALERT_EMAIL_TO não configurado; alerta por e-mail ignorado.")
        return

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    configurar_remetente(message, outlook, remetente_alerta())
    destinatarios_formatados = "; ".join(destinatarios)
    message.To = destinatarios_formatados
    message.Subject = assunto
    message.Body = corpo

    log_path = get_log_file_path()
    if log_path and os.path.exists(log_path):
        message.Attachments.Add(log_path)

    for anexo in anexos_extras or []:
        if anexo and os.path.exists(anexo):
            message.Attachments.Add(anexo)

    resolver_destinatarios(message)
    message.Send()
    logger.info(f"Alerta enviado para: {destinatarios_formatados}")


def enviar_alerta_execucao(resumo):
    if not env_bool("ALERT_ON_SUCCESS", True):
        return

    total_ok = len(resumo.get("ok", []))
    total_n_ok = len(resumo.get("n_ok", []))
    total_pendentes = resumo.get("total_pendentes", 0)
    log_path = get_log_file_path() or "Log nao identificado."

    status = "concluída"
    if total_n_ok:
        status = "concluída com erro(s)"

    assunto = f"Automação UCS {status}: {total_ok} Ok / {total_n_ok} N Ok"
    corpo = f"""Automação UCS {status}.

Resumo:
- Usuarios pendentes: {total_pendentes}
- Processados com sucesso: {total_ok}
- Processados com erro: {total_n_ok}
- Log: {log_path}

Usuários com sucesso:
{resumo_linhas(resumo.get("ok", []))}

Usuários com erro:
{resumo_linhas(resumo.get("n_ok", []))}
"""
    enviar_email_alerta(assunto, corpo, anexos_extras=screenshots_do_resumo(resumo))


def enviar_alerta_erro(erro):
    if not env_bool("ALERT_ON_ERROR", True):
        return

    log_path = get_log_file_path() or "Log não identificado."
    assunto = "Erro fatal na automação UCS"
    corpo = f"""A automação UCS falhou antes de concluir a execução.

Erro:
{erro}

Traceback:
{traceback.format_exc()}

Log: {log_path}
"""
    enviar_email_alerta(assunto, corpo, anexos_extras=get_error_screenshots())
