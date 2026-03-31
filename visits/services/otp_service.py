from django.conf import settings
from django.core.mail import send_mail

from visits.models import EmailOTP


def send_email_otp(email: str, purpose: str = "email_verification") -> EmailOTP:
    otp = EmailOTP.create_otp(email=email, purpose=purpose, expiry_minutes=10)

    subject = "رمز التحقق"
    message = (
        f"رمز التحقق الخاص بك هو: {otp.code}\n\n"
        "صلاحية الرمز: 10 دقائق.\n"
        "إذا لم تطلب هذا الرمز، يمكنك تجاهل هذه الرسالة.\n\n"
        "هذه رسالة آلية، يرجى عدم الرد عليها."
    )

    send_mail(
        subject=subject,
        message=message,
        from_email=settings.DEFAULT_FROM_EMAIL,
        recipient_list=[email],
        fail_silently=False,
    )

    return otp