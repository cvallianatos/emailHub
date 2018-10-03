#
# get-calendar.rb 
# Retrieves all emails from Inbox, saves them in a specific folder and saves all the attachments.
# Relies on config.txt to derive localization information.
# Created by Chris Vallianatos on 9/5/18.
# Copyright © 2018 xbinvestments. All rights reserved.
#

require 'win32ole'
require 'parseconfig'

def sanitize(inputName)
  # Remove any character that aren't 0-9, A-Z, or a-z
  x = inputName.gsub(/[^0-9A-Z]/i, '_')
  inputName = x[0..30]
end

def get_calendar(myFolder)

  # Setup section to be changed fby modifying the config.txt file
  config = ParseConfig.new('C:\\Users\\1347409\\Documents\\Dropbox\\Projects\\email\\config.txt')

  # Parameters
  myPath = config['myPath']
  exclude = config['exclude']

  # Calendar
  calendarSubFolder = config['calendarSubFolder']
  calendarAttachmentSubFolder = config['calendarAttachmentSubFolder']
  calendarSaveExtension = config['calendarSaveExtension']

  myCalendar = myPath + calendarSubFolder
  myCalendarAttachments = myCalendar + calendarAttachmentSubFolder

  # Basic common setup for Outlook objects

  outlook = WIN32OLE.new('Outlook.Application')
  mapi = outlook.GetNameSpace('MAPI')

  # Get all calendar entries 

  sourceFolder = mapi.Folders.Item("chris.vallianatos@tcs.com").Folders.Item(myFolder)

  numberOfCalendarEntries = sourceFolder.Items.Count

  print "There are ", numberOfCalendarEntries, " messages in your ", myFolder, "\n"
  print "=====================================================\n"

  	numberOfCalendarEntries.downto(1) do |i|
    	message = sourceFolder.Items.Item(i) 
    	if message.UnRead
    		msgStatus = "Unread"
    	else
    		msgStatus = "Read"
    	end 

    	print i, " - \tStatus: ", msgStatus, " - \tFrom: ", message.Organizer, " - \t", message.Subject, "\n\n"
    	if msgStatus == "Unread"
  			tempSubject = message.Subject
  			tempOrganizer = message.Organizer

  			fileName = myCalendar + message.CreationTime.strftime("%Y-%m-%d %H-%M-%S") + " - " + sanitize(tempOrganizer) + " - " + sanitize(tempSubject) + calendarSaveExtension

  			message.SaveAs(fileName,4)

  			# Save all the attacments of each calendar entry

    		message.Attachments.each do |attachment|
    	  		attachmentFile = attachment.FileName
  		
  	  		# Ignore "ATT00...gif" & "ATT00..img" files 

    	  		if attachmentFile[0..4] != exclude
  	      			attachmentName = myCalendarAttachments + "\\#{attachmentFile}"
  	      			attachment.SaveAsFile(attachmentName)  
  	       		end         
      		end
  		end
  	end
end

def process_calendar()
    get_calendar("Calendar")
end