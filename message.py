import imaplib
import os
import email
from email.utils import getaddresses

# MARK: - Imap Variables
fromFolder = "Item2"
sender = "naisen04@gmail.com"
subject1 = "1"
subject = subject1.encode('utf-8')
toFolder = "New111"
filePath = "C:/Users/Andrey/Desktop"
fileName = "File.txt"
searchMethod = 2  # 1 - Search by Subject, 2 - Search by Sender, 3 - ALL folder

class Imap:

    def __init__(self, email, password):
        self.email = email
        self.password = password

    def connectImap(self):
        self.imapServer = imaplib.IMAP4_SSL("imap.gmail.com", 993)  # server to connect to
        print("Connecting to mailbox...")
        self.imapServer.login(self.email, self.password)
        return self.imapServer

    def searchGmail(self, searchMethod, sender, subject, fromFolder):
        self.imapServer.list()
        self.imapServer.select(fromFolder, readonly=False)
        if searchMethod == 1:
            self.imapServer.literal = subject
            result, data = self.imapServer.uid("Search", 'CHARSET', 'UTF-8', "SUBJECT")
            print(result, data)
            return result, data
        elif searchMethod == 2:
            result, data = self.imapServer.uid("Search", None, f'FROM "{sender}"')
            print(result, data)
            return result, data
        elif searchMethod == 3:
            result, data = self.imapServer.uid("Search", None, "ALL")
            print(result, data)
            return result, data
        else:
            print("Сhoose the correct method")

    def moveLetter(self, toFolder):
        result, data = self.searchGmail(searchMethod, sender, subject, fromFolder)
        print(result, data)
        if result == "OK":
            for num in data[0].split():
                copy_res = self.imapServer.uid("COPY", num, toFolder)
                if copy_res[0] == "OK":
                    delete_res = self.imapServer.uid("STORE", num, "+FLAGS", "(\Deleted)")
                    self.imapServer.expunge()
                    print("Move completed")

    def moveToStarred(self):
        result, data = self.searchGmail(searchMethod, sender, subject, fromFolder)
        print(result, data)
        if result == "OK":
            for num in data[0].split():
                copy_res = self.imapServer.uid("COPY", num, "[Gmail]/Starred")
                if copy_res[0] == "OK":
                    delete_res = self.imapServer.uid("STORE", num, "+FLAGS", "(\Deleted)")
                    self.imapServer.expunge()
                    print("Move completed")

    def deleteLetter(self):
        result, data = self.searchGmail(searchMethod, sender, subject, fromFolder)
        print(result, data)
        if result == "OK":
            for num in data[0].split():
                delete_res = self.imapServer.uid("STORE", num, "+FLAGS", "(\Deleted)")
                self.imapServer.expunge()
                print("Deleted")

    def mask_an_email(self, markRead):
        result, data = self.searchGmail(searchMethod, sender, subject, fromFolder)
        if result == "OK":
            print(data)
            for num in data[0].split():
                if markRead == False:
                    self.imapServer.uid("STORE", num, '+FLAGS', '\\Seen')  # Mark as read
                    print("SEEN")
                elif markRead == True:
                    self.imapServer.uid("STORE", num, '-FLAGS', '\\Seen')  # Mark as unread
                    print("UNSEEN")
        else:
            print("Nothing to mark.")

        self.imapServer.expunge()  # not need if auto-expunge enabled


class Save(Imap):

    def parseEmail(self):
        result, data = self.searchGmail(searchMethod, sender, subject, fromFolder)
        if result == "OK":
            file = os.path.join(filePath, fileName)
            outFile = open(file, 'w')
            for num in data[0].split():
                result, data = self.imapServer.uid('fetch', num, '(RFC822)')
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)
                my_tuple = email.utils.parseaddr(email_message['From'])
                valueTo = email_message['To']
                valueFrom = my_tuple[1]
                valueDate = email_message['Date']
                msgBody = ('From: ' + valueFrom + '\n' + 'To: ' + valueTo + '\n' + 'Date: ' + valueDate)
                outFile.write(msgBody + '\n')

            outFile.close()

# MARK: - Imap Methods
messages = Imap(email, password)

#MARK: - Save Methods
saving = Save(email, password)