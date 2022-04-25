from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
from pathlib import Path
import win32com.client as client
import pathlib


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")

window = Tk()

window.title("MagicMailer")
window.geometry("1152x700")
window.configure(bg = "#FFFFFF")

# setting the app logo
icon = PhotoImage(file='appLogo.png')
window.iconphoto(False, icon)


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def outlookMailer(receiverEmail, subject, body, attachment, embeddedImage):

    gift_path = pathlib.Path(attachment)
    choco_path = pathlib.Path(embeddedImage)

    #absolute path
    gift_absolute_path = str(gift_path.absolute())
    choco_absolute_path = str(choco_path.absolute())

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = receiverEmail
    message.Subject = subject
    message.Attachments.Add(gift_absolute_path)
    image = message.Attachments.Add(choco_absolute_path)

    html_body = body
                  
    # cid -> constant id

    image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "choco-img")
    message.HTMLBody = html_body
    message.Save()
    message.Send()
    print("\nCongratulations!!! Mail has been sent successfully...")

canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 700,
    width = 1152,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    487.0,
    700.0,
    fill="#C4C4C4",
    outline="")

image_image_1 = PhotoImage(
    file=relative_to_assets("final_cover.png"))
image_1 = canvas.create_image(
    243.0,
    350.0,
    image=image_image_1
)


entry_image_0 = PhotoImage(
    file=relative_to_assets("entry_1.png"))
entry_bg_0 = canvas.create_image(
    818.5,
    100.0,
    image=entry_image_0
)
entry_0 = Entry(
    bd=0,
    bg="#E8E5E5",
    highlightthickness=0
)
entry_0.place(
    x=603.0,
    y=89.0,
    width=431.0,
    height=50.0
)


entry_image_1 = PhotoImage(
    file=relative_to_assets("entry_1.png"))
entry_bg_1 = canvas.create_image(
    818.5,
    200.0,
    image=entry_image_1
)
entry_1 = Entry(
    bd=0,
    bg="#E8E5E5",
    highlightthickness=0
)
entry_1.place(
    x=603.0,
    y=189.0,
    width=431.0,
    height=50.0
)

entry_image_2 = PhotoImage(
    file=relative_to_assets("entry_2.png"))
entry_bg_2 = canvas.create_image(
    818.5,
    300.0,
    image=entry_image_2
)
entry_2 = Entry(
    bd=0,
    bg="#E8E5E5",
    highlightthickness=0
)
entry_2.place(
    x=603.0,
    y=289.0,
    width=431.0,
    height=50.0
)

entry_image_3 = PhotoImage(
    file=relative_to_assets("entry_2.png"))
entry_bg_3 = canvas.create_image(
    818.5,
    400.0,
    image=entry_image_3
)
entry_3 = Entry(
    bd=0,
    bg="#E8E5E5",
    highlightthickness=0
)
entry_3.place(
    x=603.0,
    y=389.0,
    width=431.0,
    height=50.0
)


entry_image_4 = PhotoImage(
    file=relative_to_assets("entry_2.png"))
entry_bg_4 = canvas.create_image(
    818.5,
    500.0,
    image=entry_image_4
)
entry_4 = Entry(
    bd=0,
    bg="#E8E5E5",
    highlightthickness=0
)
entry_4.place(
    x=603.0,
    y=489.0,
    width=431.0,
    height=50.0
)


canvas.create_text(
    598.0,
    68.0,
    anchor="nw",
    text="Receiver's Email",
    fill="#591E22",
    font=("Roboto Bold", 15 * -1)
)

canvas.create_text(
    598.0,
    168.0,
    anchor="nw",
    text="Subject",
    fill="#591E22",
    font=("Roboto Bold", 15 * -1)
)

canvas.create_text(
    598.0,
    268.0,
    anchor="nw",
    text="Body",
    fill="#591E22",
    font=("Roboto Bold", 15 * -1)
)

canvas.create_text(
    598.0,
    368.0,
    anchor="nw",
    text="Attachment",
    fill="#591E22",
    font=("Roboto Bold", 15 * -1)
)

canvas.create_text(
    598.0,
    468.0,
    anchor="nw",
    text="Embedded Image",
    fill="#591E22",
    font=("Roboto Bold", 15 * -1)
)

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: outlookMailer(str(entry_0.get()),str(entry_1.get()),str(entry_2.get()),str(entry_3.get()),str(entry_4.get())),
    relief="flat"
)
button_1.place(
    x=738.0,
    y=600.0,
    width=161.0,
    height=49.0
)

window.resizable(False, False)
window.mainloop()
