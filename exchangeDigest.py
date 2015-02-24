import base64
import os, string

import suds
from suds.client import Client
from suds.transport.https import WindowsHttpAuthenticated

#import logging
#logging.basicConfig(level=logging.INFO)
#logging.getLogger('suds.client').setLevel(logging.DEBUG)

import sys
import ntlm
#print sys.modules['suds']
#print sys.modules['ntlm']

import MySQLdb
from datetime import datetime

import time
import binascii
import hashlib

import mimetypes
mimetypes.init()

import urllib2
import re
from htmlentitydefs import codepoint2name

# these are the mailboxes you want to pull from to create the digest
accounts = [('DOMAIN\MAILBOX-1','3'),('DOMAIN\MAILBOX-2','4'),('DOMAIN\MAILBOX-3','5')]

try:
  mb = int(sys.argv[1])
  mailbox = accounts[mb]
except IndexError:
  sys.exit()

# your phpBB3 credentials
try:
  db = MySQLdb.connect("mysql.yourdomain.com","user","password","phpBB3")
  cursor = db.cursor()
except:
  sys.exit()

# path on your webserver to place attachments
path = "/path/to/digestAttachments/"
# modified wsdl since the one provided with Exchange 2010 does not work
wsdl = "file:///path/to/ews/Services.wsdl"
url = "https://mail.yourdomain.com/EWS/Exchange.asmx"

user = mailbox[0]
password = "XXX"
fid = mailbox[1]

class exchangeDigest:

  unichr2entity = dict((unichr(code), u'&%s;' % name) for code,name in codepoint2name.iteritems() if code!=38 and code!=34 and code!=60 and code!=62 and code!=160) # exclude &,<,>,&nbsp,quot
  post_id = 0
  path = path
  user_id = ''

  def __init__(self, user, password, fid, path, wsdl, url):
    self.user = user
    self.password = password
    self.forum_id = fid

    self.wsdl = wsdl
    self.url = url

  # strip the message body of opening and closing <html> and <body> tags and utf-8 encode
  def cleanText(self, text, d=unichr2entity):

    p1 = re.compile(r'<.*html.*<body.*?>|<img.*?>|</body>.*</html>',re.DOTALL)
    body = p1.sub('', text)

    p2 = re.compile(r'&nbsp;|\n')
    body = p2.sub(' ', body)

    for key, value in d.iteritems():
      if key in body:
        body = body.replace(key, value)

    b = unicode(body)
    body = b.encode("utf-8")

    return body

  def connect(self):
    exchangeDigest.ntlm = WindowsHttpAuthenticated(username=self.user,password=self.password)
    try:
      exchangeDigest.c = Client(self.wsdl, transport=self.ntlm, location=self.url)
    except:
      sys.exit()

    # build a list of all email id's and key's, then reverse sort
    xml = self.findItemXML()
    attr = exchangeDigest.c.service.ResolveNames(__inject={'msg':xml})
    try:
      exchangeDigest.id_key_list = attr.FindItemResponseMessage.RootFolder.Items.Message
    except AttributeError:
      exchangeDigest.id_key_list = '99'
      sys.exit()

  def getList(self):
    exchangeDigest.itemArray = []
    for line in exchangeDigest.id_key_list:
      key = str(line.ItemId._ChangeKey)
      id  = str(line.ItemId._Id)
      exchangeDigest.itemArray.append([id,key])
    exchangeDigest.itemArray.reverse()

  def deleteItemXML(self, id, key):
    xml = '''
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <DeleteItem DeleteType="HardDelete" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <ItemIds>
        <t:ItemId Id="''' + id + '''" ChangeKey="''' + key + '''" />
      </ItemIds>
    </DeleteItem>
  </soap:Body>
</soap:Envelope>
'''
    return xml

  def getAttachmentXML(self, id):
    xml = '''
<soap:Envelope 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetAttachment 
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id="''' + id + '''" />
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>
'''
    return xml

  # get an email object -> sans attachments
  def getItemXML(self, id, key):
    xml = '''
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetItem
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:IncludeMimeContent>true</t:IncludeMimeContent>
      </ItemShape>
      <ItemIds>
        <t:ItemId Id="''' + id + '''" ChangeKey="''' + key + '''" />
      </ItemIds>
    </GetItem>
  </soap:Body>
</soap:Envelope>
'''
    return xml

  def findItemXML(self):
    xml = '''
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
              Traversal="Shallow">
      <ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </ItemShape>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>
'''
    return xml

  def processList(self):
    import time
    # begin loop through of email list
    for line in exchangeDigest.itemArray:

      # get individual email
      xml  = self.getItemXML(line[0], line[1])
      attr = exchangeDigest.c.service.ResolveNames(__inject={'msg':xml})
      data = attr.GetItemResponseMessage.Items.Message

      a = str(data.DateTimeSent)

      '''
      try:
        strptime = time.strptime()
      except AttributeError:
        import time
      '''

      dt_obj = time.strptime(a, "%Y-%m-%d %H:%M:%S")
      exchangeDigest.timeSent = str(int(time.mktime(dt_obj)))

      try:
        body = self.cleanText(data.Body.value)
        body_md5 = str(hashlib.md5(body).hexdigest())
      except AttributeError:
        body = ''
        body_md5 = ''

      sql = "select post_id from phpbb_posts where post_time = %s and post_checksum = %s"
      cursor.execute(sql,(exchangeDigest.timeSent, body_md5))
      rows = cursor.fetchall()
      copy = 0
      if (len(rows) > 0):
        copy = 1

    #  print data.From.Mailbox.Name
      if (copy == 0):
        try:
          emailAddress = str(data.From.Mailbox.EmailAddress)
        except UnicodeEncodeError:
          a = data.From.Mailbox.EmailAddress
          b = unicode(a)
          emailAddress = b.encode("utf-8")
        except AttributeError:
          emailAddress = 'none@none.com'

        tmp = emailAddress.partition('@')
        if (tmp[2] == 'yourdomain.com'):
          exchangeDigest.From = tmp[0].lower()
        else:
          exchangeDigest.From = emailAddress

        sql = "select * from phpbb_users where username = '" + exchangeDigest.From + "'";
        cursor.execute(sql)
        rows = cursor.fetchall()
        if (len(rows) > 0):
          #User exists
          for row in rows:
            user_info = row
        else:
          #User does not exist
          try:
            test = urllib2.urlopen("http://www.yourdomain.com/images/staff/" + exchangeDigest.From + ".jpg")
            user_avatar = "http://www.yourdomain.com/images/staff/" + exchangeDigest.From + ".jpg"
          except urllib2.HTTPError:
            user_avatar = "images/avatars/gallery/baby_cute_face.jpg"

          email_hash = str(binascii.crc32(emailAddress))
          salt = str(hashlib.md5(exchangeDigest.timeSent).hexdigest())
          sql = "insert into phpbb_users (user_type, group_id, user_permissions, user_perm_from, user_ip, user_regdate, username, username_clean, user_password, user_passchg, user_pass_convert, user_email, user_email_hash, user_birthday, user_lastvisit, user_lastmark, user_lastpost_time, user_lastpage, user_last_confirm_key, user_last_search, user_warnings, user_last_warning, user_login_attempts, user_inactive_reason, user_inactive_time, user_posts, user_lang, user_timezone, user_dst, user_dateformat, user_style, user_rank, user_colour, user_new_privmsg, user_unread_privmsg, user_last_privmsg, user_message_rules, user_full_folder, user_emailtime, user_topic_show_days, user_topic_sortby_type, user_topic_sortby_dir, user_post_show_days, user_post_sortby_type, user_post_sortby_dir, user_notify, user_notify_pm, user_notify_type, user_allow_pm, user_allow_viewonline, user_allow_viewemail, user_allow_massemail, user_options, user_avatar, user_avatar_type, user_avatar_width, user_avatar_height, user_sig, user_sig_bbcode_uid, user_sig_bbcode_bitfield, user_from, user_icq, user_aim, user_yim, user_msnm, user_jabber, user_website, user_occ, user_interests, user_actkey, user_newpasswd, user_form_salt, user_new, user_reminded, user_reminded_time, user_digest_filter_type, user_digest_format, user_digest_max_display_words, user_digest_max_posts, user_digest_min_words, user_digest_new_posts_only, user_digest_pm_mark_read, user_digest_remove_foes, user_digest_reset_lastvisit, user_digest_send_hour_gmt, user_digest_send_on_no_posts, user_digest_show_mine, user_digest_show_pms, user_digest_sortby, user_digest_type, user_digest_has_ever_unsubscribed, user_digest_no_post_text) values(0, 2, '00000000006xrqeiww\n\n\nzik0zi000000\nzik0zi000000\nzik0zi000000', 0, '10.1.12.70', %s, %s, %s, '', %s, 0, %s, %s, '', %s, %s, 0, '', '', 0, 0, 0, 0, 0, 0, 0, 'en', '-10.00', 0, 'D M d, Y g:i a', 1, 0, '', 0, 0, 0, 0, -3, 0, 0, 't', 'd', 0, 't', 'a', 0, 1, 0, 1, 1, 1, 1, 230271, %s, 2, 90, 90, '', '', '', '', '', '', '', '', '', '', '', '', '', '', %s, 1, 0, 0, 'ALL', 'HTML', 0, 0, 0, 0, 0, 0, 1, 0.00, 0, 1, 1, 'board', 'NONE', 0 ,0)"
          cursor.execute(sql, (exchangeDigest.timeSent, exchangeDigest.From, exchangeDigest.From, exchangeDigest.timeSent, emailAddress, email_hash, exchangeDigest.timeSent, exchangeDigest.timeSent, user_avatar, salt))
          sql = "select * from phpbb_users where username = '" + exchangeDigest.From + "'";
          cursor.execute(sql)
          rows = cursor.fetchall()
          for row in rows:
            user_info = row
        exchangeDigest.user_id =  str(user_info[0])

        try:
          exchangeDigest.subj = str(data.Subject)
        except UnicodeEncodeError:
          a = data.Subject
          b = unicode(a)
          exchangeDigest.subj = b.encode("utf-8")

        if (body != ''):
          has_attachments = '0'
          if (str(data.HasAttachments) == 'True'):
            has_attachments = '1'

          sql = "insert into phpbb_topics (forum_id, icon_id, topic_attachment, topic_approved, topic_reported, topic_title, topic_poster, topic_time, topic_time_limit, topic_views, topic_replies, topic_replies_real, topic_status, topic_type, topic_first_post_id, topic_first_poster_name, topic_first_poster_colour, topic_last_post_id, topic_last_poster_id, topic_last_poster_name, topic_last_poster_colour, topic_last_post_subject, topic_last_post_time, topic_last_view_time, topic_moved_id, topic_bumped, topic_bumper, poll_title, poll_start, poll_length, poll_max_options, poll_last_vote, poll_vote_change) values(%s, 0, %s, 1, 0, %s, %s, %s, 0, 0, 0, 0, 0, 0, %s, %s, '', %s, %s, %s, '', %s, %s, %s, 0, 0, 0, '', 0, 0, 1, 0, 0)"
          cursor.execute(sql, (self.forum_id, has_attachments, exchangeDigest.subj, exchangeDigest.user_id, exchangeDigest.timeSent, exchangeDigest.post_id, exchangeDigest.From, exchangeDigest.post_id, exchangeDigest.user_id, exchangeDigest.From, exchangeDigest.subj, exchangeDigest.timeSent, exchangeDigest.timeSent))
          topic_id = str(cursor.lastrowid)

          sql = "insert into phpbb_posts (topic_id, forum_id, poster_id, icon_id, poster_ip, post_time, post_approved, post_reported, enable_bbcode, enable_smilies, enable_magic_url, enable_sig, post_username, post_subject, post_text, post_checksum, post_attachment, bbcode_bitfield, bbcode_uid, post_postcount, post_edit_time, post_edit_reason, post_edit_user, post_edit_count, post_edit_locked) values(%s, %s, %s, 0, '10.1.12.70', %s, 1, 0, 1, 1, 1, 1, '', %s, %s, %s, %s, '', '', 0, 0, '', 0, 0, 0)"
          cursor.execute(sql, (topic_id, self.forum_id, exchangeDigest.user_id, exchangeDigest.timeSent, exchangeDigest.subj, body, body_md5, has_attachments))
          exchangeDigest.post_id = str(cursor.lastrowid)

          sql = "update phpbb_users set user_posts = user_posts+1 where user_id = %s"
          cursor.execute(sql,(exchangeDigest.user_id))

          if (has_attachments == '1'):
            try:
              for attachment in data.Attachments.FileAttachment:
                try:
                  id = str(line.AttachmentId._Id)
                  self.writeFile(id, exchangeDigest.post_id, topic_id, exchangeDigest.user_id, exchangeDigest.timeSent)
                except AttributeError:
                  id = str(getattr(attachment[1], "_Id", None))
                  mimetype = str(getattr(attachment[1], "ContentType", None))
                  if (id != "None"):
                    self.writeFile(id, exchangeDigest.post_id, topic_id, exchangeDigest.user_id, exchangeDigest.timeSent)
            except AttributeError:
              error = AttributeError

      #xml = self.deleteItemXML(line[0], line[1])
      #attr = exchangeDigest.c.service.ResolveNames(__inject={'msg':xml})

  def sendEmail(self):
    xml = '''
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <CreateItem MessageDisposition="SendAndSaveCopy" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <SavedItemFolderId>
        <t:DistinguishedFolderId Id="sentitems" />
      </SavedItemFolderId>
      <Items>
        <t:Message>
          <t:ItemClass>IPM.Note</t:ItemClass>
          <t:Subject>Project Action</t:Subject>
          <t:Body BodyType="Text">Priority - Update specification</t:Body>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>jeckersley@yourdomain.com</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:IsRead>false</t:IsRead>
        </t:Message>
      </Items>
    </CreateItem>
  </soap:Body>
</soap:Envelope>
'''
    return xml

  def setPID(self):
    if (exchangeDigest.id_key_list != '99'):
      # find the last post_id
      sql = "select post_id from phpbb_posts order by post_id desc limit 1"
      cursor.execute(sql)
      count = cursor.rowcount
      if (count < 1):
        sql = "insert into phpbb_posts (topic_id, forum_id, poster_id, icon_id, poster_ip, post_time, post_approved, post_reported, enable_bbcode, enable_smilies, enable_magic_url, enable_sig, post_username, post_subject, post_text, post_checksum, post_attachment, bbcode_bitfield, bbcode_uid, post_postcount, post_edit_time, post_edit_reason, post_edit_user, post_edit_count, post_edit_locked) values(0, 0, 0, 0, '10.1.12.70', 123456, 1, 0, 1, 1, 1, 1, '', 'none', 'none', 'none', 0, '', '', 0, 0, '', 0, 0, 0)"
        cursor.execute(sql)
        a = str(cursor.lastrowid)
        exchangeDigest.post_id = str(cursor.lastrowid + 1)
        sql = "delete from phpbb_posts where post_id = %s"
        cursor.execute(sql, a)
      else:
        row = cursor.fetchone()
        exchangeDigest.post_id = str(row[0])


  def updateCount(self):
    if (len(exchangeDigest.id_key_list) > 0 and exchangeDigest.user_id != ''):

      sql = "select post_id from phpbb_posts where forum_id = %s";
      cursor.execute(sql, (self.forum_id))
      rows = cursor.fetchall()
      num_posts = cursor.rowcount

      if (num_posts > 0):
        sql = "select topic_id from phpbb_topics where forum_id = %s";
        cursor.execute(sql, (self.forum_id))
        rows = cursor.fetchall()
        num_topics = cursor.rowcount
  
        sql = "update phpbb_forums set forum_posts = %s, forum_topics = %s, forum_topics_real = %s, forum_last_post_id = %s, forum_last_poster_id  = %s, forum_last_post_subject = %s, forum_last_post_time = %s, forum_last_poster_name = %s where forum_id = %s"
        cursor.execute(sql, (num_posts, num_topics, num_topics, exchangeDigest.post_id, exchangeDigest.user_id, exchangeDigest.subj, exchangeDigest.timeSent, exchangeDigest.From, self.forum_id))

      sql = "select post_id from phpbb_posts where post_approved = '1'"
      cursor.execute(sql)
      rows = cursor.fetchall()
      total_num_posts = cursor.rowcount
      sql = "update phpbb_config set config_value = %s where config_name = 'num_posts'"
      cursor.execute(sql, (total_num_posts))

      sql = "select topic_id from phpbb_topics where topic_approved = '1'";
      cursor.execute(sql)
      rows = cursor.fetchall()
      total_num_topics = cursor.rowcount  
      sql = "update phpbb_config set config_value = %s where config_name = 'num_topics'"
      cursor.execute(sql, (total_num_topics))

      sql = "select user_id, username from phpbb_users order by user_id desc";
      cursor.execute(sql)
      rows = cursor.fetchall()
      total_num_users = cursor.rowcount
      sql = "update phpbb_config set config_value = %s where config_name = 'num_users'"
      cursor.execute(sql, (total_num_users))

      sql = "update phpbb_config set config_value = %s where config_name = 'newest_username'"
      cursor.execute(sql, (rows[0][1]))


  def writeFile(self, id, post_id, topic_id, user_id, time):

    xml = self.getAttachmentXML(id)
    attr = exchangeDigest.c.service.ResolveNames(__inject={'msg':xml})
    fileError = '0'

    try:
      filename = str(attr.GetAttachmentResponseMessage.Attachments.FileAttachment.Name)
    except UnicodeEncodeError:
      a = attr.GetAttachmentResponseMessage.Attachments.FileAttachment.Name
      b = unicode(a)
      filename = b.encode("utf-8")
    except AttributeError:
      exchangeDigest.error = AttributeError
      fileError = '99'

    if (fileError != '99'):
      if (os.path.exists(self.path + filename)):
        unique_name = id[0:6] + filename
      else:
        unique_name = filename
      filename = exchangeDigest.path + unique_name

      content = base64.b64decode(attr.GetAttachmentResponseMessage.Attachments.FileAttachment.Content)
      pointer = open(filename,"w")
      pointer.write(content)
      pointer.close()
      size = os.path.getsize(filename)

      splitlist = filename.split(".")
      i = len(splitlist)
      ext = ""
      if (i > 1):
        ext = splitlist[i-1]

      mimetype = mimetypes.guess_type(filename, strict = False)
      if (mimetype[0] == None):
        mimetype = ('text/plain', None)

      sql = "insert into phpbb_attachments (post_msg_id, topic_id, in_message, poster_id, is_orphan, physical_filename, real_filename, download_count, attach_comment, extension, mimetype, filesize, filetime, thumbnail) values(%s, %s, 0, %s, 0, %s, %s, 1, '', %s, %s, %s, %s, 0)"
      cursor.execute(sql, (post_id, topic_id, user_id, unique_name, filename, ext, mimetype[0], size, time))

#---------------------------------------------------
if __name__ == "__main__":
  test = exchangeDigest(user, password, fid, path, wsdl, url)
  test.connect()
  test.getList()

  test.setPID()

  test.processList()
  test.updateCount()
#---------------------------------------------------


cursor.close()
db.close()
