require 'C:\\Users\\1347409\\Documents\\Dropbox\\Projects\\email\\get-email.rb'
require 'C:\\Users\\1347409\\Documents\\Dropbox\\Projects\\email\\get-calendar.rb'

time = Time.new
timeStamp = time.strftime("%Y-%m-%d-%H-%M-%S")
# Open output file
myFile = File.open("c:\\users\\1347409\\documents\\dropbox\\projects\\email\\output.txt", "a")

myFile.print "Starting...", timeStamp, "\n"

time = Time.new
timeStamp = time.strftime("%Y-%m-%d-%H-%M-%S")

process_emails()
process_calendar()

myFile.print "Completed...", timeStamp, "\n"
myFile.print "--------------------------------\n"
