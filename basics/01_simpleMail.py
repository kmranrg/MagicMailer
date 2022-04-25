import win32com.client as client

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "kmranrg@gmail.com"
message.CC = "kmranrg@yahoo.com"
message.Subject = "Happy Birthday2"
message.Body = "Wish you a long and happy life."
message.Save()
message.Send()
