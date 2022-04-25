import win32com.client as client

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "kmranrg@gmail.com"
message.CC = "kmranrg@yahoo.com"
message.Subject = "Happy Birthday2"
message.Body = "Wish you a long and happy life."

message.HTMLBody = "<b>Wishing you all the best on your birthday</b>"
html_body = """
<div>
    <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;">
        Happy Birthday!! 
    </h1>
    <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;">
        Wishing you all the best on your birthday!!
    </span>
</div><br>
<div>
    <img src="https://hips.hearstapps.com/hmg-prod.s3.amazonaws.com/images/cute-birthday-instagram-captions-1584723902.jpg" width=50%>
</div>
"""
message.HTMLBody = html_body
message.Save()
message.Send()
