# -*- coding: utf-8 -*-
import os
import sys
import glob
import traceback
from email.parser import Parser as EmailParser
import email.utils
import OleFileIO_PL as OleFile
import email
import random
import string
import mimetypes
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import codecs
import kitchen
from kitchen.text.converters import to_unicode, to_bytes
import email.mime, email.mime.nonmultipart, email.charset

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())


def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')


class Attachment:
    def __init__(self, msg, dir_):
        # Get long filename
        self.longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])

        # Get short filename
        self.shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])

        # Get attachment data
        self.data = msg._getStream([dir_, '__substg1.0_37010102'])



class Message(OleFile.OleFileIO):
    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)
        
    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return self._header

    def _getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
        #unicodeVersion =self._getStream(filename + '001F')
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion


    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')

    @property
    def attachments(self):
        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listdir():
                if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:
                    attachmentDirs.append(dir_[0])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            return self._attachments


def create_attachment(filename,data):
    ctype, encoding = mimetypes.guess_type(filename)
    if ctype is None or encoding is not None:
        # No guess could be made, or the file is encoded (compressed), so
        # use a generic bag-of-bits type.
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    if maintype == 'text':
        
        # Note: we should handle calculating the charset
        msg = MIMEText(data, _subtype=subtype)
        #msg.set_payload(data)
    elif maintype == 'image':
        
        msg = MIMEImage(data, _subtype=subtype)
        #msg.set_payload(data)
    elif maintype == 'audio':
        
        msg = MIMEAudio(data, _subtype=subtype)
        #msg.set_payload(data)
    else:
        
        msg = MIMEBase(maintype, subtype)
        msg.set_payload(data)
        #fp.close()
        # Encode the payload using Base64
        encoders.encode_base64(msg)
    # Set the filename parameter
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    return msg


fn=sys.argv[1]
msg=Message(fn)

emlmsg = MIMEMultipart()
for key in msg.header.keys():
    emlmsg[key]=msg.header[key]
#print(type(msg.body))
#print(kitchen.text.misc.guess_encoding(to_bytes(msg.body)))
payload=(to_bytes(msg.body)).decode('utf-8')
# f=open('temp.txt','w')
# f.write(payload)
# f.close()
# text=open('temp.txt').read()
#print(text)
m=email.mime.nonmultipart.MIMENonMultipart('text', 'plain', charset='utf-8')
cs=email.charset.Charset('utf-8')
#cs.body_encoding = email.charset.QP
#charset.add_charset('utf-8', charset.SHORTEST, charset.QP)
m.set_payload(payload, charset=cs)
print(m)
# body=MIMEText(text,'plain', _charset='utf-8')
# #body.replace_header('Content-Transfer-Encoding', '')
# print(body)
emlmsg.attach(m)

for attachment in msg.attachments:
    filename=attachment.longFilename or attachment.shortFilename or 'noname'+''.join(random.sample(string.digits,5))
    attach=create_attachment(filename,attachment.data)
    emlmsg.attach(attach)
of=os.path.splitext(fn)[0]+'.eml'
f=open(of,'w')
emlstring=emlmsg.as_string()
f.write(emlstring)

    


#if __name__ == '__main__':
#    main()