#coding=utf-8
import sys
import configparser
import codecs
import smtplib
import getpass
from xlsxtool import DiaryXlsxUtil
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

diary_date = datetime.today()
diary_xl = ""
diary_tp = ""

mail_smtp = ""
mail_from = ""
mail_to = ""
mail_cc = ""
mail_daily_subject = ""
mail_weekly_subject = ""
mail_body = u""

def load_config():
	config = configparser.ConfigParser()
	config.read(["diary.properties"], "utf-8")
	
	global diary_xl, mail_smtp, mail_from, mail_to, mail_cc, mail_daily_subject, mail_weekly_subject
	diary_xl = config["DailyReport"]["diaryxl"]
	mail_smtp = config["DailyReport"]["smtp"]
	mail_from = config["DailyReport"]["from"]
	mail_to = config["DailyReport"]["to"]
	mail_cc = config["DailyReport"]["cc"]
	mail_daily_subject = config["DailyReport"]["daily_subject"] + "%s"
	mail_weekly_subject = config["DailyReport"]["weekly_subject"] + "(%s - %s)"

def load_diary_template():
	global diary_tp
	with codecs.open("diary.template", "r", "utf-8") as f:
		diary_tp = f.read()

def compose_daily(p_date):
	global mail_body, mail_daily_subject
	mail_daily_subject = mail_daily_subject % p_date.strftime("%Y-%m-%d")
	util = DiaryXlsxUtil(diary_xl)
	diary = util.get_daily_data(p_date)
	if(diary):
		mail_body = diary_tp % (diary.date, diary.breakfast, diary.lunch, diary.dinner, diary.dessert)

def compose_weekly(start_date, days):
	global mail_body, mail_weekly_subject
	end_date = start_date + timedelta(days = days - 1)
	mail_weekly_subject = mail_weekly_subject % (start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
	util = DiaryXlsxUtil(diary_xl)
	diaries = util.get_weekly_data(start_date, days)
	for diary in diaries:
		mail_body += diary_tp % (diary.date, diary.breakfast, diary.lunch, diary.dinner, diary.dessert)

def send_mail(p_from, p_to, p_cc, p_subject, p_body):
	msg = MIMEMultipart("alternative")
	msg.set_charset("utf-8")
	msg["Subject"] = p_subject
	msg["From"] = p_from
	msg["To"] = p_to
	msg["Cc"] = p_cc
	msg.attach(MIMEText(p_body, "html", "utf-8"))
	
	output_mail_info(p_from, p_to, p_cc, p_subject, p_body)
	
	m_opt = ""
	while(m_opt != "n" and m_opt != "y" and m_opt != "l"):
		m_opt = raw_input("Send diary ? y:yes, n:no, l:login\n").lower()
	if(m_opt == "n"):
		print "User canceled sending diary!"
		return 0
	svr = smtplib.SMTP(mail_smtp)
	if(m_opt == "l"):
		m_user = raw_input("Username: ")
		m_pswd = getpass.getpass()
		try:
			svr.login(m_user, m_pswd)
		except smtplib.SMTPAuthenticationError, err:
			print "Authentication failed!"
			return 1
		else:
			print "Authentication successfully!"
	try:
		svr.sendmail(p_from, [p_to] + p_cc.split(","), msg.as_string())
		svr.quit()
	except smtplib.SMTPRecipientsRefused, err:
		# sys.stderr.write("ERROR: %s\n" % str(err))
		return 1
	else:
		print "Diary was sent successfully!"


def daily_report(p_date):
	load_config()
	load_diary_template()
	compose_daily(p_date)
	send_mail(mail_from, mail_to, mail_cc, mail_daily_subject, mail_body)

def weekly_report(start_date, days):
	load_config()
	load_diary_template()
	compose_weekly(start_date, days)
	send_mail(mail_from, mail_to, mail_cc, mail_weekly_subject, mail_body)

def output_mail_info(p_from, p_to, p_cc, p_subject, p_body):
	print "From: %s" % p_from
	print "Subject: %s" % p_subject
	print "To: %s" % p_to
	print "Cc: %s" % p_cc
	print "-------------------------------------------------------------------------------------------------------------------------------------------------------\n"
	print "%s\n\n" % p_body

def get_monday(p_date):
	return p_date - timedelta(days = p_date.weekday())

def main():
	global diary_date
	argv_len = len(sys.argv)
	if(argv_len >= 2):
		diary_type = sys.argv[1]
		#---------------------------------------------------------------
		# Send daily report
		#---------------------------------------------------------------
		if(diary_type == "daily"):
			if(argv_len >= 3):
				diary_date = datetime.strptime(sys.argv[2], "%Y-%m-%d") 
			daily_report(diary_date)
		#---------------------------------------------------------------
		# Send weekly report
		#---------------------------------------------------------------
		elif(diary_type == "weekly"):
			days = 5
			if(argv_len >= 3):
				diary_date = datetime.strptime(sys.argv[2], "%Y-%m-%d")
				if(argv_len >= 4):
					days = int(sys.argv[3])
			else:
				diary_date = get_monday(diary_date)
			weekly_report(diary_date, days)
		else:
			print "Invalid argument: %s!" % diary_type
			return 1
	else:
		print "Usage: python diary_sdr.py daily/weekly [date [days]]!"
		return 1


if(__name__ == "__main__"):
	sys.exit(main())
