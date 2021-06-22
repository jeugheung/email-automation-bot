import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import os,datetime

# MARK: - Calendar Variables
startMeeting = "20210713T183000"
endMeeting = "20210714T223000"
attendees = ["naisen04@gmail.com", "akart.dev@gmail.com", "avpishka@gmail.com"]
fro = "Andrey Kim <akartsight@gmail.com>"
meetingSubject = "Say Hello to new iMac"

class Calendar():

    def __init__(self, email, password):
        self.email = email
        self.password = password

    def sendMeeting(self, attendees, fro, startMeeting, endMeeting, meetingSubject):
        CRLF = "\r\n"
        dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M")

        dtstart = startMeeting
        dtend = endMeeting

        organizer = "ORGANIZER;CN=Organizer:mailto:akartsight@gmail.com"

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

        eml_body = "Meeting invitation"

        msg = MIMEMultipart('mixed')
        msg['Reply-To'] = fro
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = meetingSubject
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
        mailServer.login(self.email, self.password)
        mailServer.sendmail(fro, attendees, msg.as_string())
        mailServer.close()
        print("Sendeed")

# MARK: - Calendar Methods
calendar = Calendar(email, password)