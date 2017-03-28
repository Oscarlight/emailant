import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from createEmail import *
 
class sendEmail(object):

	def __init__(self,  **kwds):
		self.__dict__.update(kwds)
		# self.fromaddr = ""
		# self.password = ""
		self.server = smtplib.SMTP('smtp.gmail.com', 587)
		self.server.starttls()
		self.server.login(self.fromaddr, self.password)
		print "Established connect to gmail account: " + self.fromaddr

	def setup(self, **kwds):
		self.__dict__.update(kwds)
		# self.toaddr = ""
		# self.subject = ""
		# self.body = ""

	def send(self, i):
		msg = MIMEMultipart()
		msg['From'] = self.fromaddr
		msg['To'] = self.toaddr
		msg['Subject'] = self.subject
		 
		msg.attach(MIMEText(self.body, 'plain'))

		# reconnect every 10 emails
		if i % 10 == 0:
			server = smtplib.SMTP('smtp.gmail.com', 587)
			server.starttls()
			server.login(self.fromaddr, self.password)

		text = msg.as_string()
		self.server.sendmail(self.fromaddr, self.toaddr, text)

	def __del__(self):
		self.server.quit()
		print "Log out"

if __name__ == "__main__":

	c = createEmail("EXECEL FILE")

	print len(c.getBodyList()[0])

	s = sendEmail(fromaddr = "YOUR GMAIL ACCOUNT", \
				  password = "PASSWORD")

	for i in range(0, len(c.getBodyList()[0])):
		s.setup(toaddr = c.getBodyList()[0][i], \
				  subject = "YOUR SUBJECT", \
				  body = c.getBodyList()[1][i])
		s.send(i)
		print str(i) + ": Finish sending email to: " + self.toaddr
