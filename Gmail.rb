#coding: utf-8
require 'gmail'
require 'dotenv'
require 'date'
require 'time'
require 'fileutils'
require 'slack'

Dotenv.load 'gmail.env'

USERNAME=ENV['ID']
PASSWORD=ENV['PASS']
$file_path = ENV["FILEPATH"]

p USERNAME
p PASSWORD

agent = Gmail.new(USERNAME,PASSWORD)

agent.peek = true
agent.mailbox("[somelab00").emails(:unread).each do |mail|
#agent.mailbox('[somelab00').emails.each do |mail|
puts "mail"
p mail

name =  mail.header.fields.select{|field| field.name == "From"}.first.value.gsub(/(.+) <.*/,'\1').gsub(/ /,'_')

s = File.read("last.txt",:encoding => Encoding::UTF_8)
if s.empty?
	s = mail.date
	File.open("last.txt","w") do |f|
		f.write s
	end
end
last_date = DateTime.parse(s.to_s)
puts %(last_date:#{last_date})
puts %(mail.date:#{mail.date})

if last_date < mail.date
	
	title = mail.subject
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
	text << body.gsub(/________________________________(.|\n)*/,"") << "\n"
	

	Slack.configure do |config|
		config.token = ENV["TOKEN"]
	end
	Slack.chat_postMessage(username:"somelaBOT(テスト中)",icon_emoji:":desktop_computer:",text: text,channel: "#mail")
	File.open("last.txt","w") do |f|
		md = mail.date

		f.write md
  	end
      	puts "Subject: #{mail.subject}"
  	puts "from: #{mail.from}"


  mail.attachments.each do | attachment |
  # Attachments is an AttachmentsList object containing a
  # number of Part objects
    filename = attachment.filename
    begin
	    dir = $file_path + name
	
	    Slack.files_upload(file:"#{dir}/#{filename}",channel:"#mail",filename:filename)

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
end # lastdate < mail.date
end

agent.logout

puts "checked at #{Time.now}"
