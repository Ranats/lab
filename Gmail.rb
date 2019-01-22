#coding: utf-8
require 'gmail'
require 'dotenv'
require 'date'
require 'time'
require 'fileutils'
require 'slack'
require 'kconv'

Dotenv.load 'gmail.env'

USERNAME = ENV['ID']
PASSWORD = ENV['PASSb']
$file_path = ENV["FILEPATH"]

agent = Gmail.new(USERNAME, PASSWORD)

p agent

#agent.peek = true
agent.mailbox("[somelab00").emails(:unread).each do |mail|
#agent.mailbox('[somelab00').emails.each do |mail|
  puts "mail"
  p mail

  name = mail.header.fields.select { |field| field.name == "From" }.first.value.gsub(/(.+) <.*/, '\1').gsub(/ /, '_')

  s = File.read("last.txt", :encoding => Encoding::UTF_8)

  puts "subject : #{Kconv.toutf8(mail.subject)}"

  if s.empty?
    s = mail.date
  end
  last_date = s.to_s

  if DateTime.parse(last_date) < DateTime.parse(mail.date)

    File.open("last.txt","w") do |f|
      f.write mail.date
    end

    title = Kconv.toutf8(mail.subject)#subject.toutf8
    body = ""
    if !mail.text_part && !mail.html_part
      body << mail.body.decoded.encode("UTF-8", mail.charset)
    elsif mail.text_part
      body << mail.text_part.decoded
    elsif mail.html_part
      body << mail.html_part.decoded
    end

    text = "*メールを受信しました．*\n"
    text << ">>>" << title << "\n"
    text << name << "\n"
    text << body.gsub(/________________________________(.|\n)*/, "") << "\n"

    puts "text-----"
    p text

    puts "title"
    p title
    puts ""
    Slack.configure do |config|
      config.token = ENV["TOKEN"]
    end

    p Slack.auth_test
    Slack.chat_postMessage(username: "somelaBOT", icon_emoji: ":desktop_computer:", text: text, channel: "#mail")
    File.open("last.txt", "w") do |f|
      md = DateTime.parse(mail.date)

      f.write md
    end
    puts "Subject: #{mail.subject}"
    puts "from: #{mail.from}"

=begin
    mail.attachments.each do |attachment|
      # Attachments is an AttachmentsList object containing a
      # number of Part objects
      filename = attachment.filename
      begin
        dir = $file_path + name

        Slack.files_upload(file: "#{dir}/#{filename}", channel: "#mail", filename: filename)

        unless FileTest.exist?("#{dir}/#{filename}")
          FileUtils.mkdir(dir) unless FileTest.exist?(dir)
          File.open("#{dir}/#{filename}", "w+b", 0644) do |f|
            f.write attachment.body.decoded
          end
        end
      rescue Exception => e
        puts "Unable to save data for #{filename} because #{e.message}"
      end
    end
=end
  end # lastdate < mail.date
end

agent.logout

puts "checked at #{Time.now}"
