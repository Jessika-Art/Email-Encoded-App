

import tkinter as tk
from styling import *
from tkinter import messagebox, filedialog
import base64
import os
from email.message import EmailMessage # to set the email with all fields required
import smtplib # allows to send the email finding the 'client'
import imghdr # this help you find the type of the file, like jpg, txt, etc...

# MAIN WINDOW ----------------------------------------------------------------------------------

def graphic():
    global messageEntry
    global secret_num

    check = False  # check if the user has attached any file.
    #root = root
    root = Tk()
    root.title('E N C R Y P T')
    root.geometry('850x600')
    root.resizable(0, 0)
    root.config(bg=BG_COLOR_GRAY)

    # BROWSE BUTTON ----------------------------------------------------------------------------------------------
    def browse_1():
        global final_emails
        path = filedialog.askopenfilename(initialdir='Desktop', title='Select Excel File')
        if path == '':
            messagebox.showerror('Error','Please an Excel file')

        else:
            data_xlc = pd.read_excel(path)
            if 'emails' in data_xlc.columns:
                emails_xlc = list(data_xlc['emails'])
                final_emails = []
                for i in emails_xlc:
                    if pd.isnull(i) == False:
                        final_emails.append(i)
                if len(final_emails) == 0:
                    messagebox.showerror('Error', 'No email addresses found in this file.')
                else:
                    emailEntry.config(state=NORMAL)
                    emailEntry.insert(0, os.path.basename(path))
                    emailEntry.config(state='readonly')
                    total.config(text = 'Total: ' + str(len(final_emails)))
                    sent.config(text= 'Sent: ')
                    left.config(text = 'Left: ')
                    failed.config(text = 'Failed: ')
    # BROWSE BUTTON END ----------------------------------------------------------------------------------------------

    # CHECK BUTTON -----------------------------------------------------------------------------------------------
    def button_check():
        if choice.get() == 'multiple':
            browse.config(state=NORMAL)
            emailEntry.config(state='readonly')
        if choice.get() == 'single':
            browse.config(state=DISABLED)
            emailEntry.config(state=NORMAL)
    # CHECK BUTTON END -----------------------------------------------------------------------------------------------

    # ATTACHMENT BUTTON ------------------------------------------------------------------------------------------
    def attach():
        global file_name, file_extension, file_path, check
        check = True # check if the user has attached any file.
        file_path = filedialog.askopenfilename(initialdir='Desktop/',title='Select File')
        file_extension = file_path.split('.')
        file_extension = file_extension[1]
        file_name = os.path.basename(file_path)
        messageEntry.insert(END, f'\n<{file_name}>\n')
    # ATTACHMENT BUTTON END ------------------------------------------------------------------------------------------

    # SENDING EMAIL PROCESS --------------------------------------------------------------------------------------
    def sending_email(to, subject, body):
        f = open('credentials.txt', 'r')
        for i in f:
            credentials = i.split(',')
        message = EmailMessage()
        message['subject'] = subject
        message['to'] = to # this is the receiver
        message['from'] = credentials[0] # this is the sender
        message.set_content(body)
        if check:
            if file_extension == 'png' or file_extension == 'jpg' or file_extension == 'jpeg':
                f = open(file_path, 'rb')
                file_data = f.read()
                subtype = imghdr.what(file_path)  # this tells you 'what' kind of file it is.
                # I need to pass all the methods below for attach the file.
                message.add_attachment(file_data, maintype='image', subtype=subtype, filename=file_name)
            # if I select other files type, the 'else' will be executed.
            else:
                f = open(file_path, 'rb')
                file_data = f.read()
                message.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        s = smtplib.SMTP('smtp.gmail.com', 587) # or 465
        s.starttls()
        s.login(credentials[0], credentials[1])
        s.send_message(message)
        x = s.ehlo()
        if x[0] == 250:
            return 'sent'
        else:
            return 'failed'

    # SENDING EMAIL PROCESS END --------------------------------------------------------------------------------------

    # SEND EMAIL BUTTON -------------------------------------------------------------------------------------------
    def send_email():
        if emailEntry.get() == '' or subjectEntry.get() == '' or messageEntry.get(1.0, END) == '\n':
            messagebox.showerror('Error', ' All fields are required.', parent=root)
        else:
            if choice.get() == 'single':
                result = sending_email(emailEntry.get(), subjectEntry.get(), messageEntry.get(1.0, END))
                if result == 'sent':
                    messagebox.showinfo('Success', 'Email has been sent.')
                if result == 'failed':
                    messagebox.showerror('Error', 'Email not sent.')

            if choice.get() == 'multiple':
                sent_count = 0
                failed_count = 0
                for j in final_emails:
                    result = sending_email(j, subjectEntry.get(), messageEntry.get(1.0, END))
                    if result == 'sent':
                        sent_count += 1
                    elif result == 'failed':
                        failed_count += 1

                    total.config(text='')
                    sent.config(text='Sent: ' + str(sent_count))
                    left.config(text='Left: ' + str(len(final_emails) - (sent_count + failed_count)))
                    failed.config(text=f'Failed: {failed_count}')

                    total.update()
                    sent.update()
                    left.update()
                    failed.update()

                messagebox.showinfo('Success', 'Emails sent successfully.')
    # SEND EMAIL BUTTON END -------------------------------------------------------------------------------------------

    # SETTINGS BUTTON ---------------------------------------------------------------------------------------------
    def settings_window():
        global encrypting_code_entry
        def save():
            if email_address.get() == '' or password_entry.get() == '' or encrypting_code_entry.get() == '':
                messagebox.showerror('Error', ' All fields are required.', parent=root1)
            else:
                with open('credentials.txt', 'w') as f:
                    f.write(email_address.get() + ', ' + password_entry.get())
                    #f.close()
                with open('secret number.txt', 'w') as n:
                    n.write(encrypting_code_entry.get())
                    messagebox.showinfo('Info', 'CREDENTIAL SAVED SUCCESSFULLY')

        root1 = Toplevel()
        root1.title('S E T T I N G S')
        root1.geometry('500x600')
        root1.resizable(0, 0)
        root1.config(bg=BG_COLOR_GRAY)
        # LOGO SETTINGS
        title = Label(root1, text='C R E D E N T I A L   S E T T I N G S', image=settingsLogo, compound=LEFT,
                      fg=TEXT_COLOR_SNOW, font=48, bg=BG_COLOR_GRAY)
        title.grid(row=0, column=0, padx=60, pady=10)
        divider = Label(root1, text='', fg=TEXT_COLOR_SNOW, font=20, bg=BG_COLOR_GRAY)
        divider.grid(row=1, column=0)
        # EMAIL ENTRY BOX
        email_label = Label(root1, text='[From (your email address)]', fg=TEXT_COLOR_SNOW, font=15, bg=BG_COLOR_GRAY)
        email_label.grid(row=2, column=0, padx=60, pady=0)
        email_address = Entry(root1, fg=BG_COLOR_GRAY, font='Calibri', bg=TEXT_COLOR_SNOW, width=40)
        email_address.grid(row=3, column=0)
        # PASSWORD ENTRY BOX
        divider = Label(root1, text='', fg=TEXT_COLOR_SNOW, font=20, bg=BG_COLOR_GRAY)
        divider.grid(row=4, column=0)
        password_label = Label(root1, text='[Set your password]', fg=TEXT_COLOR_SNOW, font=15, bg=BG_COLOR_GRAY)
        password_label.grid(row=5, column=0, pady=10)
        password_entry = Entry(root1, fg=BG_COLOR_GRAY, font='Calibri', bg=TEXT_COLOR_SNOW, show='*', width=40)
        password_entry.grid(row=6, column=0)

        # ENCODE - DECODE CODE
        divider = Label(root1, text='', fg=TEXT_COLOR_SNOW, font=20, bg=BG_COLOR_GRAY)
        divider.grid(row=7, column=0)
        encrypting_code = Label(root1, text='[Set Your Secret Number]', fg=TEXT_COLOR_SNOW, font=15, bg=BG_COLOR_GRAY)
        encrypting_code.grid(row=8, column=0, pady=10)
        encrypting_code_entry = Entry(root1, bg=BG_COLOR_GRAY, fg=TEXT_COLOR_SNOW, justify=CENTER,
                         font=26, bd=3, borderwidth=5)
        encrypting_code_entry.grid(row=9, column=0, pady=10)


        # BUTTON CANCEL
        cancel = Button(root1, text='Close', compound=LEFT, bg='gray18', fg=TEXT_COLOR_SNOW, width=10, height=2, cursor='hand2',
                        font=('Helvetica', 11, 'bold'), activebackground='darkred', command=lambda: root1.destroy())
        cancel.place(x=270, y=450)
        # BUTTON CONFIRM
        confirm = Button(root1, text='Save', compound=RIGHT, bg='gray18', fg=TEXT_COLOR_SNOW, width=10, height=2,
                    font=('Helvetica',11, 'bold'), activebackground='green', command=save, cursor='hand2')
        confirm.place(x=140, y=450)
        # KEEP THE CREDENTIALS DISPLAYED
        with open('credentials.txt', 'r') as f:
            for i in f:
                credentials = i.split(',')
        email_address.insert(0, credentials[0]) # give the index 0.
        password_entry.insert(0, credentials[1]) # give the index 1.

        with open('secret number.txt', 'r') as n:
            for k in n:
                secret_number = k
        encrypting_code_entry.insert(0, secret_number)
    # SETTINGS BUTTON  END ---------------------------------------------------------------------------------------------

    # EXIT BUTTON -------------------------------------------------------------------------------------------------
    def i_exit():
        result = messagebox.askyesno('Notification', 'Do you want to close?')
        if result:
            root.destroy()
        else:
            pass
    # EXIT BUTTON END -------------------------------------------------------------------------------------------------

    # CLEAN BUTTON ------------------------------------------------------------------------------------------------
    def i_clean():
        messageEntry.delete(1.0, END)
        emailEntry.delete(0, END)
        subjectEntry.delete(0, END)
    # CLEAN BUTTON END ------------------------------------------------------------------------------------------------

    # ENCODE ------------------------------------------------------------------------------------------------------
    def encode():
        root1 = Toplevel()
        root1.title('E N C O D E')
        root1.geometry('480x300+250+80')
        root1.resizable(0, 0)
        root1.config(bg=BG_COLOR_GRAY)

        # ENCRYPT
        def encrypt(*args):
            with open('secret number.txt', 'r') as n:
                for k in n:
                    secret_code = k
            entry_code_field = numEntry.get()
            if secret_code in entry_code_field:
                text = messageEntry.get(1.0, END)
                encode_message = text.encode('ascii')
                base64_bytes = base64.b64encode(encode_message)
                encrypt1 = base64_bytes.decode('ascii')
                i_clean()
                messageEntry.insert(END, encrypt1)
                root1.destroy()
            else:
                warning.config(text='Invalid Secret Number!', fg='red', font=10)

        labelFrame = Label(root1, text='E N C O D E \nEncode your message by typing your secret number.',
                           fg=TEXT_COLOR_SNOW,
                           font=18, bg=BG_COLOR_GRAY)
        labelFrame.pack(side=TOP, pady=20)

        warning = Label(root1, text='', fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
        warning.pack(anchor='center')

        numEntry = Entry(root1, bg=BG_COLOR_GRAY, fg=TEXT_COLOR_SNOW, justify=CENTER,
                         font=26, bd=3, borderwidth=5)
        numEntry.pack(padx=100, pady=20, ipady=20, anchor='n')

        confirm = Button(root1, text='Encode Now', activebackground='green', font=('Helveica', 15, 'bold'),
                         command=encrypt)
        confirm.pack(side=BOTTOM, pady=15)
    # ENCODE END ------------------------------------------------------------------------------------------------------

    # DECODE ------------------------------------------------------------------------------------------------------
    def decode():
        root1 = Toplevel()
        root1.title('D E C O D E')
        root1.geometry('480x300+250+80')
        root1.resizable(0, 0)
        root1.config(bg=BG_COLOR_GRAY)

        # DECRYPT
        def decrypt():
            with open('secret number.txt', 'r') as n:
                for k in n:
                    secret_code = k
            entry_code_field = numEntry.get()
            if secret_code in entry_code_field:
                text = messageEntry.get(1.0, END)
                decode_message = text.encode('ascii')
                base64_bytes = base64.b64decode(decode_message)
                decrypt1 = base64_bytes.decode('ascii')
                i_clean()
                messageEntry.insert(END, decrypt1)
                root1.destroy()
            else:
                warning.config(text='Invalid Secret Number!', fg='red', font=10)

        labelFrame = Label(root1, text='D E C O D E \nDecode the message by typing your secret number.',
                           fg=TEXT_COLOR_SNOW, font=18, bg=BG_COLOR_GRAY)
        labelFrame.pack(side=TOP, pady=20)

        warning = Label(root1, text='', fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
        warning.pack(anchor='center')

        numEntry = Entry(root1, bg=BG_COLOR_GRAY, fg=TEXT_COLOR_SNOW, justify=CENTER,
                         font=26, bd=3, borderwidth=5)
        numEntry.pack(padx=100, pady=20, ipady=20, anchor='n')
        confirm = Button(root1, text='Decode Now', activebackground='green', font=('Helveica', 15, 'bold'),
                         command=decrypt)
        confirm.pack(side=BOTTOM, pady=15)
    # DECODE END ------------------------------------------------------------------------------------------------------



    # G R A P H I C   S E T U P -----------------------------------------------------------------------------------

    # TITLE ENCRYPT
    titleFrame = Frame(root, bg=BG_COLOR_GRAY)
    titleFrame.pack(side=TOP, anchor='nw')
    titleImage = PhotoImage(file='images/encode_logo.png')
    titleLabel = Label(titleFrame, image=titleImage, text='E N C R Y P T',fg=BG_COLOR_GRAY, compound=LEFT, font=18)
    titleLabel.pack(side=TOP)

    # SETTINGS
    settingsLogo = PhotoImage(file='images/settings.png')
    settings = Button(root, image=settingsLogo, bg=BG_COLOR_GRAY, text='SETTINGS', font=18, activebackground=BLUE,
                      command=settings_window, cursor='hand2')
    settings.place(x=790, y=10)

    # RADIO BUTTONS
    choice = StringVar()
    oneRadio = Radiobutton(root, text='Send to One', font=16, fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY, variable=choice,
                           value='single', cursor='hand2', command=button_check)
    oneRadio.pack(padx='30', pady='8')
    oneRadio.place(x=450, y=10)

    mult_Radio = Radiobutton(root, text='Send to Many', font=16, fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY, variable=choice,
                           value='multiple', cursor='hand2', command=button_check)
    mult_Radio.pack(padx='30', pady='8')
    mult_Radio.place(x=600, y=10)
    choice.set('single')

    # ADDRESS + BROWSE + SUBJECT
    # LABEL FRAME
    browseImage = PhotoImage(file='images/excel.png')
    emailLabel = LabelFrame(root, text='To', font=6, fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
    emailLabel.pack(pady=7)
    # EMAIL ADDRESS ENTRY
    emailEntry = Entry(emailLabel, font=('Umpush', 12, 'bold'), justify='left')
    emailEntry.pack(pady=10, ipadx=70)
    # BROWSE BUTTON
    browse = Button(emailLabel, image=browseImage, state=DISABLED, text='  BROWSE', compound=LEFT, fg=BG_COLOR_GRAY,
                    width=100, height=18, font=('Helvetica',8, 'bold'), activebackground=BLUE, cursor='hand2', command=browse_1)
    browse.pack()
    # SUBJECT LABEL
    subjectLabel = LabelFrame(root, text='Subject', font=6, fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY,)
    subjectLabel.pack(pady=7)
    # SUBJECT ENTRY
    subjectEntry = Entry(subjectLabel, font=('Umpush',12, 'bold'), justify='left')
    subjectEntry.pack(pady=10, ipadx=70)

    # MESSAGE TEXT FRAME
    messageLabel = LabelFrame(root, text='Compose Email', font=4, fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY,)
    messageLabel.pack(pady=30, ipadx=200)
    # MESSAGE ENTRY
    messageEntry = Text(messageLabel, font=('Umpush' ,11), borderwidth=10)
    messageEntry.pack(pady=10,  side=LEFT)

    # ENCODE
    encode = Button(messageLabel, text='Encode', compound=RIGHT, bg='gray18', fg=TEXT_COLOR_SNOW, width=14, height=2,
                    font=('Helvetica',9, 'bold'), activebackground=BLUE, command=encode, cursor='hand2', relief=FLAT)
    encode.pack(pady=1, anchor='n')
    # DECODE
    decode = Button(messageLabel, text='Decode', compound=RIGHT, bg='gray18', fg=TEXT_COLOR_SNOW, width=14, height=2,
                    font=('Helvetica',9, 'bold'), activebackground=BLUE, command=decode, cursor='hand2', relief=FLAT)
    decode.pack(anchor='n')
    # SEND
    send_image = PhotoImage(file='images/send1.png')
    send = Button(messageLabel, text='Send    ', compound=RIGHT, fg=BG_COLOR_GRAY, width=18, height=2, relief=RAISED,
                    font=('Helvetica',9, 'bold'), activebackground='green', image=send_image, command=send_email)
    send.pack(ipadx=41, ipady=17)
    # ATTACHMENT
    attach_image = PhotoImage(file='images/browse.png')
    attachment = Button(messageLabel, image=attach_image, text='Attach   ', compound=RIGHT, fg=BG_COLOR_GRAY, width=14, height=1,
                    font=('Helvetica',9, 'bold'), activebackground=BLUE, relief=RAISED, command=attach)
    attachment.pack(ipadx=43, ipady=14)
    # CLEAN
    clear_image = PhotoImage(file='images/clear.png')
    clean = Button(messageLabel, image=clear_image, text='Clear   ', compound=RIGHT, fg=BG_COLOR_GRAY, width=14, height=1,
                    font=('Helvetica',9, 'bold'), activebackground=BLUE, command=i_clean, relief=RAISED)
    clean.pack(ipadx=43, ipady=14)
    # EXIT
    exitButton = Button(messageLabel, text='Exit', compound=BOTTOM, fg=BG_COLOR_GRAY, width=10, height=2,
                    font=('Helvetica',12, 'bold'), activebackground='darkred', command=i_exit, relief=RAISED)
    exitButton.pack(pady=10, anchor='s')

    # TOTAL TO SEND
    total = Label(root, text='Total', font=('Helvetica',10, 'bold'), fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
    total.place(x=20, y=572)
    # SENT
    sent = Label(root, text='Sent', font=('Helvetica',10, 'bold'), fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
    sent.place(x=95, y=572)
    # LEFT
    left = Label(root, text='Left', font=('Helvetica',10, 'bold'), fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
    left.place(x=160, y=572)
    # FAILED TO SEND
    failed = Label(root, text='Failed', font=('Helvetica', 10, 'bold'), fg=TEXT_COLOR_SNOW, bg=BG_COLOR_GRAY)
    failed.place(x=220, y=572)

    # R U N
    root.mainloop()
graphic()



