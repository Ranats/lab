require 'mail'
require 'net/imap'
require 'kconv'
require 'dotenv'
require 'fileutils'
require 'tlsmail'
require 'base64'
Dotenv.load 'reg.env'

$imap_usessl = true
$imap_host = 'outlook.office365.com'
$imap_port = 993

$smtp_host = 'smtp.office365.com'
$smtp_port = 587
$smtp_domain = 'office365.com'

$imap_user=ENV["ID"]
$imap_passwd=ENV["TOKEN"]
$from = $imap_user
$to = ENV["TO"]
$file_path = ENV["FILEPATH"]

class Net::IMAP::Envelope
  def mail_address_formatted(value)
    return nil unless ["from", "sender", "reply_to", "to"].include?(value)
    self.__send__(value)[0].mailbox + "@" + self.__send__(value)[0].host
  end
end

def login
  # $imap = Net::IMAP.new($imap_host,$imap_port,$imap_usessl)
  $imap = Net::IMAP.new($imap_host, $imap_port, $imap_usessl)

  puts "connected"

  $imap.login($imap_user, $imap_passwd)

  puts "logined"

  #p imap.examine('INBOX')
  $imap.select('INBOX')
end

def deliver(envelope, body, file_path)
  content = <<EOF
From: #{$imap_user}
To: #{$to}
Subject: FW:#{envelope.subject.toutf8}
Date: #{Time.now}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary=\"boundary_string_by_landscape_mail\"
--boundary_string_by_landscape_mail
Content-Type: text/plain; charset=iso2022-jp
Content-Transfer-Encoding: 7bit

forwarded from #{`hostname`.strip}.

-------- Forwarded Message --------
Subject: 	#{envelope.subject.toutf8}
Date: 	#{envelope.date}
From: 	#{envelope.from[0].name.toutf8} <#{envelope.from[0].mailbox}@#{envelope.from[0].host}>
  #{body.toutf8}

EOF

  file_path.each do |file|
    content += <<EOF
--boundary_string_by_landscape_mail
Content-Type: #{file[:type]} name=#{File.basename(file[:path])}
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename=#{File.basename(file[:path])}

#{Base64.encode64 File.open(file[:path]).read}

EOF
  end

  content << "--boundary_string_by_landscape_mail--\n"

  Net::SMTP.enable_tls(OpenSSL::SSL::VERIFY_NONE)
  Net::SMTP.start($smtp_host, $smtp_port, $smtp_domain, $imap_user, $imap_passwd, :login) do |smtp|
    begin
      smtp.send_message(content, $imap_user, $to)
      puts "delivered"
    rescue => e
      puts "errorror #{e}"
    end
  end

end

def cron_job
  $imap.search(["SUBJECT", "[somelab00", 'UNSEEN']).each do |msg_id|
    envelope = $imap.fetch(msg_id, "ENVELOPE")[0].attr["ENVELOPE"]
    files = []

    $imap.fetch(msg_id, "RFC822").each do |mail|
      m = Mail.new(mail.attr["RFC822"])
      name = envelope.from.first.name.toutf8.gsub(/ /, "_")
      p name
      p m.subject

      body = ''

      if m.multipart?
        if m.text_part
          body = m.text_part.decoded
        elsif m.html_part
          body = m.html_part.decoded
        end

        #添付ファイル
        m.attachments.each do |attachment|
          # 添付ファイルの種類とファイル名
          p attachment.content_type.match(/.*;/)[0]
          p filename = attachment.filename

          # 添付ファイルの保存処理
          dir = $file_path + "#{name}" #"R:\\Share\\workspace\\添付資料\\#{name}"
          begin
            FileUtils.mkdir_p(dir) unless FileTest.exist?(dir)
            File.open("#{dir}/#{filename}", "w+b") do |f|
              f.write attachment.body.decoded
            end
          rescue => e
            puts "添付ファイルの保存に失敗 #{e.message}"
          end
          files << {type: attachment.content_type.match(/.*;/)[0], path: "#{dir}/#{filename}"}
        end

      else
        body = m.body.decoded
      end

      deliver(envelope, body, files)
    end

#    $imap.store(msg_id, "+FLAGS", [:Seen])
  end
end

at_exit do
  quit
end

def quit
  $imap.logout
  $imap.disconnect
end

#while true
  begin
    login
    cron_job
    puts "checked at #{Time.now}"
    quit
  rescue => e
    puts "Err #{e}"
  end
#  sleep(60)
#end

