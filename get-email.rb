#
# get-email.rb 
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

def get_emails(myFolder)

  # Setup section to be changed fby modifying the config.txt file
  config = ParseConfig.new('C:\\Users\\1347409\\Documents\\Dropbox\\Projects\\email\\config.txt')

  # Parameters
  targetStr = config['targetStr']
  myPath = config['myPath']
  exclude = config['exclude']

  # emails
  emailSubFolder = config['emailSubFolder']
  emailAttachmentSubFolder = config['emailAttachmentSubFolder']
  emailSaveExtension = config['emailSaveExtension']

  myEmails = myPath + emailSubFolder
  myEmailAttachments = myEmails + emailAttachmentSubFolder

  # Basic common setup for Outlook objects

  outlook = WIN32OLE.new('Outlook.Application')
  mapi = outlook.GetNameSpace('MAPI')

  # Get all emails 

  sourceFolder = mapi.Folders.Item("chris.vallianatos@tcs.com").Folders.Item(myFolder)

  # Target folder "3-Completed"

  targetFolder = mapi.Folders.Item("chris.vallianatos@tcs.com").Folders.Item("CNV").Folders.Item(targetStr)

  numberOfEmails = sourceFolder.Items.Count

  print "There are ", numberOfEmails, " messages in your ", myFolder, "\n"
  print "=====================================================\n"

  numberOfEmails.downto(1) do |i|
    message = sourceFolder.Items.Item(i) 
    print i, " - \tFrom: ", message.SenderName, " - \t", message.Subject, "\n\n"

  	tempName = message.Subject

  	fileName = myEmails + message.ReceivedTime.strftime("%Y-%m-%d %H-%M-%S") + " - " + message.SenderName + " - " + sanitize(tempName) + emailSaveExtension

  	message.SaveAs(fileName,4)

  	# Save all the attacments of each message

    	message.Attachments.each do |attachment|
    	  attachmentFile = attachment.FileName
  		
  	  # Ignore "ATT00...gif" & "ATT00..img" files

    	  if attachmentFile[0..4] != exclude
  	      attachmentName = myEmailAttachments + "\\#{attachmentFile}"
  	      attachment.SaveAsFile(attachmentName)  
  	  end         
      end

      # Move messages from inbox to "3-Completed" if read
      if !message.UnRead
        message.Move(targetFolder)
      end
  end
end

def process_emails()
    get_emails("Inbox")
    get_emails("Sent Items")

end