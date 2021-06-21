import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import os,datetime
#COMMASPACE = ', '

CRLF = "\r\n"
login = ""
password = ""
attendees = ["naisen04@gmail.com", "akart.dev@gmail.com", "avpishka@gmail.com"]
organizer = "ORGANIZER;CN=Organizer:mailto:akartsight@gmail.com"

fro = "Andrey Kim <akartsight@gmail.com>"

class Calendar():

    def sendMeeting(self, login, password, CRLF, attendees, fro):

        dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M")

        dtstart = "20210713T183000"
        dtend = "20210713T223000"


        print(dtstart)
        print(dtend)
        print(dtstamp)

        description = "Apple WWDC 2022"
        attendee = ""
        for att in attendees:
            attendee += "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-    PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE" + CRLF + " ;CN=" + att + ";X-NUM-GUESTS=0:" + CRLF + " mailto:" + att + CRLF
        ical = "BEGIN:VCALENDAR" + CRLF + "PRODID:pyICSParser" + CRLF + "VERSION:2.0" + CRLF + "CALSCALE:GREGORIAN" + CRLF
        ical += "METHOD:REQUEST" + CRLF + "BEGIN:VEVENT" + CRLF + "DTSTART:" + dtstart + CRLF + "DTEND:" + dtend + CRLF + "DTSTAMP:" + dtstamp + CRLF + organizer + CRLF
        ical += "UID:FIXMEUID" + dtstamp + CRLF
        ical += attendee + "CREATED:" + dtstamp + CRLF + description + "LAST-MODIFIED:" + dtstamp + CRLF + "LOCATION:" + CRLF + "SEQUENCE:0" + CRLF + "STATUS:CONFIRMED" + CRLF
        ical += "SUMMARY:Meeting " + CRLF + "TRANSP:OPAQUE" + CRLF + "END:VEVENT" + CRLF + "END:VCALENDAR" + CRLF

        eml_body = "by iOS DeV"

        msg = MIMEMultipart('mixed')
        msg['Reply-To'] = fro
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = "Apple WWDC 2013"
        msg['From'] = fro
        msg['To'] = ",".join(attendees)

        part_email = MIMEText(eml_body, "html")
        part_cal = MIMEText(ical, 'calendar;method=REQUEST')

        msgAlternative = MIMEMultipart('alternative')
        msg.attach(msgAlternative)

        ical_atch = MIMEBase('application/ics', ' ;name="%s"' % ("invite.ics"))
        ical_atch.set_payload(ical)
        encoders.encode_base64(ical_atch)
        ical_atch.add_header('Content-Disposition', 'attachment; filename="%s"' % ("invite.ics"))

        eml_atch = MIMEText('', 'plain')
        encoders.encode_base64(eml_atch)
        eml_atch.add_header('Content-Transfer-Encoding', "")

        msgAlternative.attach(part_email)
        msgAlternative.attach(part_cal)

        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(login, password)
        mailServer.sendmail(fro, attendees, msg.as_string())
        mailServer.close()
        print("Sendeed")

calendar = Calendar()
calendar.sendMeeting(login, password, CRLF, attendees, fro)