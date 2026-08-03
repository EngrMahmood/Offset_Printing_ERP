import os

from django.core.mail.backends.smtp import EmailBackend as SMTPEmailBackend


def _get_credentials():
    """DB-stored Gmail credentials (set from Settings) win; environment
    variables are the fallback for servers configured before the UI existed."""
    try:
        from core.models import EmailSettings
        es = EmailSettings.objects.first()
        if es and es.gmail_address and es.gmail_app_password:
            return es.gmail_address, es.gmail_app_password
    except Exception:
        pass
    address = os.environ.get('GMAIL_ADDRESS')
    password = os.environ.get('GMAIL_APP_PASSWORD')
    if address and password:
        return address, password
    return None


class DynamicGmailEmailBackend(SMTPEmailBackend):
    """SMTP backend that reads Gmail credentials at send-time instead of at
    process start, so saving them in Settings works without a server restart.
    Falls back to printing to the console when nothing is configured yet."""

    def __init__(self, *args, **kwargs):
        creds = _get_credentials()
        self._configured = bool(creds)
        if creds:
            # django.core.mail.send_mail()/get_connection() pass username=None,
            # password=None *explicitly* (not simply omitted), so setdefault()
            # would never apply our real credentials — override any falsy value.
            if not kwargs.get('host'):
                kwargs['host'] = 'smtp.gmail.com'
            if not kwargs.get('port'):
                kwargs['port'] = 587
            if not kwargs.get('username'):
                kwargs['username'] = creds[0]
            if not kwargs.get('password'):
                kwargs['password'] = creds[1]
            if not kwargs.get('use_tls'):
                kwargs['use_tls'] = True
            self._from_address = creds[0]
        else:
            self._from_address = None
        super().__init__(*args, **kwargs)

    def send_messages(self, email_messages):
        if not self._configured:
            for message in email_messages:
                print(message.message().as_string())
            return len(email_messages)

        # Gmail rejects/relabels mail whose From doesn't match the
        # authenticated account, so stamp it here rather than trusting
        # whatever DEFAULT_FROM_EMAIL callers used.
        for message in email_messages:
            message.from_email = self._from_address

        return super().send_messages(email_messages)
