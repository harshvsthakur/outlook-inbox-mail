import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetLast()
body_content = message.body

print (body_content)

print ("Subject : ",message.Subject)

#print ("Sender : ",message.Sender)

print ("Sender E-Mail : ",message.SenderEmailAddress)

print ("Sender E-Mail Type : ",message.SenderEmailType)

print ("Sender Name : ",message.SenderName)

print ("Recipients Name : ",message.Recipients)

print ("Sender Name : ",message.ReceivedTime)

## message.CreationTime
## https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?redirectedfrom=MSDN&view=outlook-pia#properties_
