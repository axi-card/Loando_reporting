import os
import smtplib
from email.message import EmailMessage


class Mailing:
    def __init__(self):
        """Define constant email contents"""
        self.msg = EmailMessage()

        self.msg['From'] = os.environ.get('EMAIL_USER')
        self.msg['To'] = ['dariusz.giemza@axi-card.pl','anna.rzemek@axi-card.pl']


    def send_error_message(self,elapsedtime,err):

        self.msg['Subject'] = "LOANDO Reporting ERROR"
        self.msg.set_content("""Report creation has failed at creation time: {0}\nError: {1}
                                """.format(round(elapsedtime,0),err))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get('EMAIL_USER'), os.environ.get('EMAIL_PASSWORD'))

            smtp.send_message(self.msg)
            smtp.close()

    def send_success_message(self, elapsedtime):

        self.msg['Subject'] = "LOANDO Reporting Success"
        self.msg.set_content("Report has been successfully created!\n\nTotal creation time: {} seconds.".format(round(elapsedtime,0)))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get('EMAIL_USER'), os.environ.get('EMAIL_PASSWORD'))

            smtp.send_message(self.msg)
            smtp.close()

