import win32com.client as client
import pathlib

gift_path = pathlib.Path("gift.jpg")
choco_path = pathlib.Path("choco.jpg")

#absolute path
gift_absolute_path = str(gift_path.absolute())
choco_absolute_path = str(choco_path.absolute())

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.To = "kmranrg@gmail.com"
message.Subject = "Happy Birthday!!"
message.Attachments.Add(gift_absolute_path)
image = message.Attachments.Add(choco_absolute_path)

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
    <img src="cid:choco-img" width=50%>
</div>
"""

# cid -> constant id

image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "choco-img")
message.HTMLBody = html_body
message.Save()
message.Send()
