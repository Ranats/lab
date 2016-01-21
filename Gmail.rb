#coding: utf-8
require 'gmail'
require 'dotenv'
Dotenv.load 'reg.env'

USERNAME=ENV['ID']
PASSWORD=ENV['TOKEN']

agent = Gmail.new(USERNAME,PASSWORD)

agent.mailbox('cat').emails.each do |mail|
  puts "Subject: #{mail.subject}"

  if !mail.text_part && !mail.html_part
    puts "body: " + mail.body.decoded.encode("UTF-8", mail.charset)
  elsif mail.text_part
    puts "text: " + mail.text_part.decoded
  elsif mail.html_part
    puts "html: " + mail.html_part.decoded
  end
end

agent.logout