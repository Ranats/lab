require 'mail'
require 'net/imap'
require 'kconv'
require 'dotenv'
require 'fileutils'
require 'tlsmail'
Dotenv.load 'reg.env'

$imap_usessl = true
$imap_host = "outlook.office365.com"
$imap_port = 993

$imap_user=ENV["ID"]
$imap_passwd=ENV["TOKEN"]
$to = ENV["TO"]
$file_path = ENV["FILEPATH"]
addr=/tokai-u/

class Net::IMAP::Envelope
  def mail_address_formatted(value)
    return nil unless ["from", "sender", "reply_to", "to"].include?(value)
    self.__send__(value)[0].mailbox + "@" + self.__send__(value)[0].host
  end
end

def login
 $imap = Net::IMAP.new($imap_host,$imap_port,$imap_usessl)

 $imap.login($imap_user,$imap_passwd)
	#p imap.examine('INBOX')
 $imap.select('INBOX')
end

def deliver(envelope,msg_id,file_path)

  puts "called"
  fetch_attr = '(UID RFC822.SIZE ENVELOPE BODY[HEADER] BODY[TEXT])'

  m = $imap.fetch(msg_id,["BODY[TEXT]"])[0].attr["BODY[TEXT]"]

  body = <<EOF
forwarded from Amethyst.
-------- Forwarded Message --------
Subject: 	#{envelope.subject.toutf8}
Date: 	#{envelope.date}
From: 	#{envelope.from[0].name.toutf8} <#{envelope.from[0].mailbox}@#{envelope.from[0].host}>

#{m}
EOF

  puts "bodyed"

  mail = Mail.new do
    from    $imap_user
    to      $to
    subject "FW:#{envelope.subject.toutf8}"
    body    body
  end

  puts "mailed"

  file_path.each do |file|
    p file
    begin
      mail.add_file(file)
    rescue => e
      puts "添付に失敗 #{e.message}"
    end
  end

  puts "filed"

=begin
  content = <<EOF
From: #{$imap_user}
To: #{$to}
Subject: FW:#{envelope.subject.toutf8}
Date: #{Time.now}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary=\"boundary_string_by_landscape_mail\"
X-Mailer: imput.rb

forwarded from Amethyst.
-------- Forwarded Message --------
Subject: 	#{envelope.subject.toutf8}
Date: 	#{envelope.date}
From: 	#{envelope.from[0].name.toutf8} <#{envelope.from[0].mailbox}@#{envelope.from[0].host}>
  #{$imap.fetch(msg_id,"BODY[TEXT]")[0].attr["BODY[TEXT]"].toutf8}

--boundary_string_by_landscape_mail
Content-Type: application/octet-stream; name=#{ file_path}
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename=#{ File.basename(file_path)}

  #{ encoded}

--boundary_string_by_landscape_mail--
EOF
=end

  Net::SMTP.enable_tls(OpenSSL::SSL::VERIFY_NONE)
  Net::SMTP.start('smtp.office365.com', 587, 'office365.com', $imap_user, $imap_passwd, :login) do |smtp|
    smtp.send_message(mail.encoded, $imap_user, $to)
  end

  puts "deliver"

end

def cron_job
  $imap.search(["SUBJECT","[somelab00",'UNSEEN']).each do |msg_id|
    envelope = $imap.fetch(msg_id, "ENVELOPE")[0].attr["ENVELOPE"]
#    from = envelope.mail_address_formatted("from")
#  puts "#{msg_id}:#{from}:#{envelope.subject.toutf8}"
#    p envelope.from.first.name

    mail = Mail.new($imap.fetch(msg_id,"RFC822")[0].attr["RFC822"])



    files = []
    mail.attachments.each do |attachment|
	    name = envelope.from.first.name.toutf8.gsub(/ /,"_")
    	p name
#      p attachment.content_type
      p attachment.filename

      filename = attachment.filename
      dir = $file_path + "#{name}"#"R:\\Share\\workspace\\添付資料\\#{name}"
      begin
        FileUtils.mkdir_p(dir) unless FileTest.exist?(dir)
        File.open("#{dir}/#{filename}","w+b") do |f|
          f.write attachment.body.decoded
        end
        files << "#{dir}/#{filename}"
      rescue => e
        puts "添付ファイルの保存に失敗 #{e.message}"
      end
    end

    deliver(envelope,msg_id,files)
#    deliver_m(envelope,msg_id,files)

#    $imap.store(msg_id, "+FLAGS", [:Seen])
  end
end

#imap.search(['UNSEEN']).each do |msg_id|
#  p envelope = imap.fetch(msg_id, "ENVELOPE")[0].attr["ENVELOPE"]
#  from = envelope.mail_address_formatted("from")
#  puts "#{msg_id}:#{from}:#{envelope.subject.toutf8}"
#  if from =~ addr
#    puts "delete"
#    imap.store(msg_id, "+FLAGS", [:Deleted])
#  end
#end



#$imap = Net::IMAP.new($imap_host,$imap_port,$imap_usessl)
#$imap.login($imap_user,$imap_passwd)

#p imap.examine('INBOX')
#$imap.select('INBOX')

while true
  login
  begin
    cron_job
  rescue
     login
    cron_job
  end
  puts "checked at #{Time.now}"

  $imap.logout
  $imap.disconnect

  sleep(3600)
#  login
end

