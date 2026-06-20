"""
Servicio de envío de correos vía Outlook COM (win32com).

Soporta cuentas de:
  - Office 365 / Exchange (AccountType = 0)
  - Hotmail / Outlook.com / Live.com (AccountType = 3, cuentas HTTP personales)
  - Gmail / IMAP configurado en Outlook (AccountType = 1)

FIX principal para Hotmail: cuando SendUsingAccount falla en cuentas personales
Microsoft, se aplica SentOnBehalfOfName como alternativa.
"""
import logging
from typing import List, Optional

import win32com.client as win32

logger = logging.getLogger(__name__)

# Tipo de cuenta Outlook para cuentas personales Microsoft (HTTP)
OUTLOOK_ACCOUNT_TYPE_HTTP = 3


def get_outlook_accounts() -> List[str]:
    """
    Obtiene la lista de direcciones SMTP de cuentas configuradas en Outlook.

    Returns:
        Lista de strings SMTP. Vacía si Outlook no está disponible o no tiene cuentas.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        # Filtra cuentas sin SmtpAddress (pueden aparecer en configuraciones incompletas)
        accounts = [
            acc.SmtpAddress.strip()
            for acc in outlook.Session.Accounts
            if acc.SmtpAddress and acc.SmtpAddress.strip()
        ]
        logger.info("Cuentas Outlook encontradas: %s", accounts)
        return accounts
    except Exception as exc:
        logger.error("Error al obtener cuentas de Outlook: %s", exc)
        return []


def send_email(
    recipients: List[str],
    subject: str,
    body: str,
    sender_account: Optional[str] = None,
) -> None:
    """
    Envía un correo mediante Outlook COM.

    Estrategia de selección de cuenta:
    1. Busca la cuenta por SmtpAddress (case-insensitive + trim).
    2. Intenta asignar con SendUsingAccount (funciona para Exchange/Office 365).
    3. Si falla, usa SentOnBehalfOfName (fallback para Hotmail/Outlook.com).

    Args:
        recipients:      Lista de direcciones destinatarias.
        subject:         Asunto del correo (con placeholders ya resueltos).
        body:            Cuerpo en texto plano (con placeholders ya resueltos).
        sender_account:  Dirección SMTP de la cuenta de envío.
                         None → Outlook usa la cuenta predeterminada.

    Raises:
        ValueError:  Cuenta no encontrada en Outlook, o lista de destinatarios vacía.
        Exception:   Error de Outlook al enviar (incluye mensajes diagnósticos).
    """
    # Obtiene la instancia de Outlook (la lanza si no está abierta)
    outlook = win32.Dispatch("Outlook.Application")

    # 0 = olMailItem: tipo estándar de correo
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body

    if sender_account:
        sender_lower = sender_account.strip().lower()

        # ── Buscar la cuenta por SMTP (sin distinción de mayúsculas) ──────────
        matched_account = None
        matched_type = None
        for acc in outlook.Session.Accounts:
            if acc.SmtpAddress.strip().lower() == sender_lower:
                matched_account = acc
                matched_type = acc.AccountType
                break

        if matched_account is None:
            raise ValueError(
                f"Cuenta '{sender_account}' no encontrada en Outlook.\n"
                "Verifique que la cuenta esté configurada y con sesión activa en Outlook."
            )

        logger.info(
            "Cuenta encontrada: %s (AccountType=%s)", sender_account, matched_type
        )

        # ── Asignar cuenta de envío ────────────────────────────────────────────
        try:
            # Método principal: funciona bien para Exchange y Office 365
            mail.SendUsingAccount = matched_account
            logger.debug("Cuenta asignada vía SendUsingAccount")
        except Exception as exc_primary:
            # Fallback para cuentas personales Microsoft (Hotmail, Outlook.com, Live)
            # que a veces rechazan SendUsingAccount con errores COM
            logger.warning(
                "SendUsingAccount falló (%s). Usando SentOnBehalfOfName como fallback.",
                exc_primary,
            )
            try:
                mail.SentOnBehalfOfName = sender_account.strip()
                logger.debug("Cuenta asignada vía SentOnBehalfOfName")
            except Exception as exc_fallback:
                raise Exception(
                    f"No se pudo asignar la cuenta de envío '{sender_account}'.\n"
                    f"Error primario: {exc_primary}\n"
                    f"Error fallback: {exc_fallback}\n\n"
                    "Para cuentas Hotmail/Outlook.com: asegúrese de que Outlook esté "
                    "abierto y la sesión esté activa (sin ventanas de confirmación pendientes)."
                ) from exc_fallback

        # Excluye al remitente de la lista de destinatarios para no enviarse a sí mismo
        recipients = [r for r in recipients if r.strip().lower() != sender_lower]

    # ── Validar que haya destinatarios ────────────────────────────────────────
    if not recipients:
        raise ValueError(
            "No hay destinatarios válidos.\n"
            "La lista estaba vacía o solo contenía la dirección del remitente."
        )

    # Formato requerido por Outlook: direcciones separadas por punto y coma
    mail.To = "; ".join(r.strip() for r in recipients)

    # ── Enviar ────────────────────────────────────────────────────────────────
    try:
        mail.Send()
        logger.info("Correo enviado a: %s | Cuenta: %s", recipients, sender_account)
    except Exception as exc:
        error_str = str(exc)

        # Error genérico de acceso COM (0x80004005): Outlook bloqueado o sin sesión
        if "0x80004005" in error_str:
            raise Exception(
                "Error de acceso a Outlook (0x80004005).\n"
                "Posible causa: Outlook no tiene la sesión activa o hay un diálogo "
                "de confirmación esperando su atención.\n"
                f"Detalle: {exc}"
            ) from exc

        # Error específico para cuentas personales Microsoft
        is_personal_ms = any(
            domain in (sender_account or "").lower()
            for domain in ("hotmail", "outlook.com", "live.com", "msn.com")
        )
        if is_personal_ms:
            raise Exception(
                f"Error al enviar con cuenta personal Microsoft ({sender_account}).\n"
                "Asegúrese de que:\n"
                "  1. Outlook esté abierto.\n"
                "  2. La cuenta esté autenticada (sin ventanas de inicio de sesión pendientes).\n"
                "  3. No haya reglas de Outlook que bloqueen el envío por COM.\n"
                f"Detalle técnico: {exc}"
            ) from exc

        raise Exception(f"Error al enviar correo: {exc}") from exc
