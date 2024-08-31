import re, os, sys, argparse, email.errors, textwrap, json, html
from mailbox import mbox
from argparse import RawTextHelpFormatter
from time import sleep
from datetime import datetime
from pathlib import Path
from email.parser import HeaderParser
from email.header import decode_header
from dateutil import parser as dparser

class color:
    red = '\033[91m'
    gre = '\033[92m'
    yel = '\033[93m'
    blu = '\033[94m'
    rst = '\033[0m'
    synth_yel = '\033[38;5;220m'
    synth_ora = '\033[38;5;208m'
    synth_red = '\033[38;5;197m'
    synth_pin = '\033[38;5;201m'
    synth_pur = '\033[38;5;128m'
    synth_blu = '\033[38;5;56m'
    synth_cya = '\033[38;5;49m'


def getMails(mailbox, key):
    try:
        mail = mailbox[key]
        return mail['date'], mail['from'], mail['to'], mail['reply-to'], mail['subject'], mail['x-attached'], mail.is_multipart(), mail
    except email.errors.MessageParseError as err:
        return False

   
def writeToIndexFile(target, html_content):
    open(target, "a", encoding="utf-8").write(html_content)


def writeHTMLHeader(target):    
    header='''<!doctype html>
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Bootstrap demo</title>
        <link href="../../../css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body class="bg-secondary">
        <script src="../../../js/jquery-3.2.1.slim.min.js"></script>
        <script src="../../../js/bootstrap.min.js"></script>
	    <center>
            <div class="container-fluid vw-75 p-3 vh-100">
                <div class="h-70 w-75 card card-body align-items-center">
                    <div>
                        <pre>
					
  ╒════════════════════════════════╕
  │     ▄███████████████████▄     ╓┘
   │     █                   █     ║├┐
   │     █ ( O&lt; ~~ n00t n00t █     ╚╕│
   │     █  LancreFI mboxer  █      ││
   │     █    HTML-viewer    █     ╓┘│
   │     █                   █     ║├┤
   ├┐    ▀███████████████████▀     ╚╕│
   ╘╧═══════════════════════════════╛│
                                     │
   ╒═════════════════════════════════╛
│ Select a message to view from 
  │ the list or with the pagination.
                    </pre>
                </div>
            </div>
            <div class="card w-75 h-30">
                <div class="card card-body align-items-center d-grid gap-2">'''
    os.makedirs(os.path.dirname(target), exist_ok=True)
    open(target, "w", encoding="utf-8").write(header)


def writeHTMLFooter(target):
    footer='''            </div>
        </div>
    </div>
</center>
</body>
</html>'''
    open(target, "a", encoding="utf-8").write(footer)


def writeMessageFile(target, msg_counter, date, sender='', recipient='', copy_to='', reply_to='', subject='', body='', headers='', attachments=False):
    message_content = '''<!doctype html>
    <html lang="en">
      <head>
      <script>
      function showHide() {
        var x = document.getElementById("headers");
        if (x.style.display === "none") { x.style.display = "block"; } 
        else { x.style.display = "none"; }
       }
       </script>
      </head>
      <body>
        <pre>'''
    message_content += 'Date: ' + str(date) + '</br>Sender: ' + str(sender) + '</br>Recipient: ' + str(recipient) \
    + '</br>Copy To: ' + str(copy_to) + '</br>Reply To: ' + str(reply_to) + '</br>Subject: ' + str(subject) + '</br></br>' \
    + str(body) + '</br></br>'
    message_content += '''					</pre>
					<div>'''

    if attachments:
        for attachment in attachments:
            message_content += 'Attachment: <a href="../attachments/' + attachment + '" class="text-decoration-none" target="_new">' + attachment + '</a><br/>'
    
    message_content += '''					</br></br></div>
    <button type="button" onclick="showHide()">
                        Show/hide headers
                        </button>
    					<div id="headers" style="display: none">'''
    for header in headers:
        message_content += re.sub("\n", "</br>", header)
    message_content += '''</div>
                           </body>
                       </html>'''
    os.makedirs(os.path.dirname(target + '/html/'), exist_ok=True)
    open(target + 'html/message_' + msg_counter + '.html', "w", encoding="utf-8").write(message_content)

    
def decodeSingleField(field):
    decoded = ''
    if field:
        encoding = decode_header(field)[0][1]
        if encoding != None:
            decoded = decode_header(field)[0][0].decode(encoding)
        else:
            if isinstance(decode_header(field)[0][0], str):
                decoded = decode_header(field)[0][0]
            else:
                decoded = decode_header(field)[0][0].decode()
    return decoded


def decodeMultiField(field):
    decoded = ''
    if field:
        for recp in decode_header(field):
            if recp[1] != None:
                decoded += recp[0].decode(recp[1])
            else:
                if isinstance(recp[0], str):
                    decoded += recp[0]
                else:
                    decoded += recp[0].decode()
    return decoded


def printToTerminal(date_p=False, from_p=False, recipient_p=False, copy_to_p=False, reply_to_p=False, subject_p=False):
    print (f"{color.blu}|    |  '--> {color.yel}Content:")
    if date_p:
        print (f"{color.blu}|    |                {color.yel}Date: {date_p}")
    if from_p:
        print (f"{color.blu}|    |                {color.yel}From: {from_p}")
    if recipient_p:
        print_out = re.sub("([^ ]),([^ ])", "\\1; \\2", recipient_p)
        print (f"{color.blu}|    |                {color.yel}To: {print_out}")
    if copy_to_p:
        print_out = re.sub("([^ ]),([^ ])", "\\1; \\2", copy_to_p)
        print (f"{color.blu}|    |                {color.yel}Copy To: {print_out}")
    if reply_to_p:
        print_out = re.sub("([^ ]),([^ ])", "\\1; \\2", reply_to_p)
        print (f"{color.blu}|    |                {color.yel}Reply to: {print_out}")
    if subject_p:
        print (f"{color.blu}|    |                {color.yel}Subject: {subject_p}")
    print (f"{color.blu}|    |{color.rst}")


def saveToFile(target, filetype=False, output=False, data=False, date=False, sender=False, recipient=False, copy_to=False, reply_to=False, subject=False, headers=False, body=False, attachment=False):
    if filetype:
        if output:
            if filetype == 'txt':
                filename = f"{target}/txt_files/{output}.txt"
                print (f"{color.blu}|    |{color.rst}")
                print (f"{color.blu}|    |-----> {color.synth_pin}Saving message to: {color.yel}{output}.txt{color.rst}")
                os.makedirs(os.path.dirname(filename), exist_ok=True)
                with open(filename, "w") as outfile:
                    if date:
                        outfile.write(f"Date: {date} \n")
                    if sender:
                        outfile.write(f"Sender: {sender}\n")
                    if recipient:
                        outfile.write(f"Recipient: {recipient}\n")
                    if copy_to:
                        outfile.write(f"Copy to: {copy_to}\n")
                    if reply_to:
                        outfile.write(f"Reply to: {reply_to}\n")
                    if subject:
                        outfile.write(f"Subject: {subject}\n")
                    if body:
                        outfile.write(f"\n{body}\n")
                    if attachment:
                        for item in attachment:
                            outfile.write (f"Attachment: {attachment[item]}\n")
                    if headers:
                        outfile.write('\n\n\n----------- HEADERS -----------\n')
                        outfile.write(headers)
            elif filetype == 'json':
                filename = f"{target}/json/{output}"
                print (f"{color.blu}|    |{color.rst}")
                print (f"{color.blu}|    |-----> {color.synth_pin}Saving message to: {color.yel}{output}{color.rst}")
                os.makedirs(os.path.dirname(filename), exist_ok=True)
                with open(filename, "w") as outfile:
                    outfile.write(json.dumps(data))
        else:
            sys.exit (f"{color.red}Missing output filename!\n{color.rst}")
    else:
        sys.exit (f"{color.red}Missing output filetype!\n{color.rst}")


##MAIN
if __name__ == '__main__':
    try:
        BANNER = """
        """+color.synth_blu+""" #####################################################
        """+color.synth_cya+""" #######   ###  ####  ###  ###     ##     ##      ####
        """+color.synth_pur+""" ######   ###    ###   ##  #   ###  #  ##  #  ########
        """+color.synth_pin+""" #####   ###  ##  ##    #  #  #######  #  ##     #####
        """+color.synth_red+""" ####   ###  #  #  #  #    #   ###  #  ##  #  ########
        """+color.synth_ora+""" #### fi     ####     ##   ###     ##  ###        ####
        """+color.synth_yel+""" #####################################################
        """+color.rst+"""
                     .--------------------------.
                     |    mbox-extract v0.1     |
                     '--------------------------'
                     .--------------------------.
                     | While doing forensics    |
                     | noticed that Autopsy     |
                     | lacks the ability to     |
                     | extract mails properly.  |
                     '--------------------------'
        """

        parser = argparse.ArgumentParser(
            prog = 'mbox-extract',
            description = BANNER,
            formatter_class = RawTextHelpFormatter,
            epilog = f"{color.synth_cya}n00t n00t{color.rst}")

        parser.add_argument('--mbox', 
            help = 'Define the mbox file to use. Use full path if not in same folder!',
            required = True)

        parser.add_argument('--html', 
            help = 'Output to a HTML bundle readable in your browser.',
            action = 'store_true',
            required = False)

        parser.add_argument('--msg', 
            help = 'Output every mail as a msg conversion.',
            action = 'store_true',
            required = False)

        parser.add_argument('--txt', 
            help = 'Output every mail as a txt file.',
            action = 'store_true',
            required = False)
            
        parser.add_argument('--term', 
            help = 'Output the mail content to terminal.',
            action = 'store_true',
            required = False)
            
        parser.add_argument('--json',
            help = 'Output the mail content to json.',
            action = 'store_true',
            required = False)

        parser.add_argument('--noattachments',
            help = 'By default attachments are saved, use this argument to NOT save them.',
            action = 'store_false',
            required = False)

        args = parser.parse_args()
        mbox_file = args.mbox
        
        ##Check that the mbox file exists
        if os.path.isfile(mbox_file):
            ##Check that the mbox is readable
            if not os.access(mbox_file, os.R_OK):
                sys.exit (f"{color.red}File {mbox_file} is not readable! Exiting!\n")
            else:    
                print (f"\n{color.gre}Starting mbox-extract on {color.yel}{mbox_file}{color.rst}")
                box = mbox(mbox_file, factory=None, create=False)
                mail_keys = box.keys()
                mail_count = len(mail_keys)
                chars_replace = ["\\", "/", "\:", "*", "?", "\"", "<", ">", "|", " "]
                chars_pattern = '[' + re.escape("".join(chars_replace)) + ']'
                output_dir = re.sub(" ", '_', mbox_file)
                
                ##Save to dict for later processing
                box_content = { 'name': Path(mbox_file).name, 'count': mail_count, 'messages': {} }
                
                ##Check that the mbox contains something
                if mail_count > 0:
                    print (f"{color.blu}|  '--> {color.rst}Found {color.gre}{mail_count}{color.rst} messages")
                    ##Iterate through the messages      
                    counter = 0
                    html_buttons = []
                    html_modals = []
                    html_attachments = []
                    body = ''
                    headers = ''
                    
                                        
                    for key in mail_keys:
                        print (f"{color.blu}|    '--> {color.rst}Extracting message {color.yel}{key+1}{color.rst}/{color.gre}{mail_count}{color.rst}")
                        ##Check whether the message is valid or not
                        ##and get some basic data and headers
                        try:
                            mail = box[key]
                            date = mail['date']
                            sender = mail['from']
                            recipient = mail['to']
                            copy_to = mail['cc']
                            reply_to = mail['reply-to']
                            subject = mail['subject']
                            headers = ''
                            headers_decoded = ''
                            for key_value in HeaderParser().parsestr(str(mail)).keys():
                                headers += f"{key_value}: {mail[key_value]}\n"
                                ##Some values are for example utf-8, so need decoding
                                ##there might also be two values, like the recipient as  =?utf-8?q?Näme Fröm Contacts <name@some.com>
                                ##where the name from the contact boot contains umlauts and is utf-8 and the actual address is in binary
                                headers_encoding = decode_header(mail[key_value])[0][1]
                                if headers_encoding != None:
                                    headers_decoded += f"{key_value}: {decode_header(mail[key_value])[0][0].decode(headers_encoding)}"
                                    try:
                                        headers_decoded += decode_header(mail[key_value])[1][0].decode()
                                    except IndexError:
                                        pass
                                    headers_decoded +=  '\n'
                                else:
                                    headers_decoded += f"{key_value}: {mail[key_value]}\n"
                            multipart = mail.is_multipart()
                            ##Here in Finland we have this thing called umlauts messing things up, so:
                            sender_decoded = decodeSingleField(sender)
                            recipient_decoded = decodeMultiField(recipient)
                            copy_to_decoded = decodeMultiField(copy_to)
                            reply_to_decoded = decodeMultiField(reply_to)
                            subject_decoded = decodeSingleField(subject)

                        except email.errors.MessageParseError as err:
                            print (f"{color.blu}|      '--> {color.red}The message is malformed, skipping.{color.rst}")
                            
                        if mail:
                            temp_stamp = dparser.parse(date)       
                            file_stamp = temp_stamp.strftime("%d%m%Y%H%M%S")
                            attachment_counter = 0
                                                        
                            ##If terminal printing was defined, print mail content to terminal
                            if args.term:
                                printToTerminal(date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded)
                                
                            ##Create the HTML content
                            if args.html:
                                if counter == 0:
                                    writeHTMLHeader('output/' + output_dir +'/html/index.html')
                                msg_counter = str(0) * (6-(len(str(counter)))) + str(counter)
                                
                                html_buttons.append('<button type="button" class="btn btn-primary text-start" data-toggle="modal" data-target="#exampleModalLong' \
                                + msg_counter + '">' + msg_counter + ' - [' + date + '] # ' + subject_decoded + ' # From: ' + sender_decoded + '  To: ' + re.sub("([^ ]),([^ ])", "\\1; \\2", recipient_decoded) + '</button>')
                                modal_content = '<div class="modal fade" id="exampleModalLong' + msg_counter + '" tabindex="-1" role="dialog" aria-labelledby="exampleModalLongTitle" aria-hidden="true">'
                                modal_content += '''                 <div class="modal-dialog mw-100 w-50 h-100" role="document"> 
                                                    <div class="modal-content h-75"> 
                                                    <div class="modal-header"> 
                                                    <h5 class="modal-title" id="exampleModalLongTitle">Modal title &nbsp;</h5>
                                                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                    <span aria-hidden="true">&times;</span>
                                                    </button>
                                                    </div>
                                                    <div class="modal-body">'''
                                modal_content += '<iframe id="iframeModalWindow' + msg_counter + '" src="message_' + msg_counter + '.html" name="iframe_modal" class="w-100 h-100"></iframe>'
                                modal_content += '''                    </div>
                                                    <div class="modal-footer">
                                                    <button type="button" class="btn btn-primary" data-dismiss="modal">Close</button>
                                                    </div>
                                                    </div>
                                                    </div>
                                                    </div>'''                                
                                html_modals.append(modal_content)
                                
                            ##Get the attachments
                            if 1 == 1:
                                html_attachments.clear()
                                for part in mail.walk():
                                    content_type = part.get_content_type()
                                    content_disposition = str(part.get("Content-Disposition"))
                                    if 1==1:
                                        ##Get the mail body
                                        if part.get_payload(decode=True) is None:
                                            body = part.get_payload(decode=True)
                                        else:
                                            try:
                                                body = part.get_payload(decode=True).decode()
                                            #If the body can't be parsed with utf-8, it's usually empty and only has a PDF attachment
                                            except UnicodeDecodeError:
                                                body = ''                                            

                                        box_content['messages'][counter] = { 'date': date, 
                                                                     'sender': sender, 
                                                                     'recipient': recipient, 
                                                                     'reply_to': reply_to, 
                                                                     'subject': subject,
                                                                     'body': body,
                                                                     'attachment': {},
                                                                     'attachment_output': {},
                                                                     'headers': headers}

                                        if args.term:
                                            body_term = re.sub(r'\n', f"\n{color.blu}|    |                {color.gre}", f"{color.blu}|    |                {color.gre}{body}")
                                            print (f"{body_term}{color.rst}")
                                            print (f"{color.blu}|    |{color.rst}")
                                                                        
                                        ##Get PNGs and JPGs
                                        if (content_type == "image/png" or content_type == "image/jpeg"):
                                            attachment_name = part.get_filename()
                                            if attachment_name:
                                                output_name = file_stamp + '-' + re.sub("[\\/:*?\"<>| ]", '_', attachment_name)
                                                if args.noattachments:
                                                    print (f"{color.blu}|    |                {color.synth_pin}Saving: {color.rst}{output_name}")
                                                    os.makedirs(os.path.dirname('output/' + output_dir + '/attachments/' + output_name), exist_ok=True)
                                                    open('output/' + output_dir + '/attachments/' + output_name, "wb").write(part.get_payload(decode=True))
                                                    box_content['messages'][counter]['attachment_output'][attachment_counter] = output_name
                                                    if args.html:
                                                        html_attachments.append(output_name)
                                                else:
                                                    print (f"{color.blu}|    |                {color.synth_pin}Attachment: {color.rst}{attachment_name}")
                                                    box_content['messages'][counter]['attachment_output'][attachment_counter] = 'not-saved'
                                                box_content['messages'][counter]['attachment'][attachment_counter] = attachment_name
                                                attachment_counter += 1

                                        ##Get other attachments
                                        elif "attachment" in content_disposition:
                                            attachment_name = part.get_filename()
                                            if attachment_name:
                                                output_name = file_stamp + '-' + re.sub("[\\/:*?\"<>| ]", '_', attachment_name)
                                                if args.noattachments:
                                                    print (f"{color.blu}|    |                {color.synth_pin}Saving: {color.rst}{output_name}")
                                                    os.makedirs(os.path.dirname('output/' + output_dir + '/attachments/' + output_name), exist_ok=True)
                                                    open('output/' + output_dir + '/attachments/' + output_name, "wb").write(part.get_payload(decode=True))
                                                    box_content['messages'][counter]['attachment_output'][attachment_counter] = output_name
                                                    if args.html:
                                                        html_attachments.append(output_name)
                                                else:
                                                    print (f"{color.blu}|    |                {color.synth_pin}Attachment: {color.rst}{attachment_name}")
                                                    box_content['messages'][counter]['attachment_output'][attachment_counter] = 'not-saved'
                                                box_content['messages'][counter]['attachment'][attachment_counter] = attachment_name
                                                attachment_counter += 1
                                    
                                if args.html:
                                    if args.noattachments:
                                        if len(html_attachments) > 0:
                                            writeMessageFile('output/' + output_dir + '/', msg_counter, date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded, body, headers_decoded, html_attachments)
                                        else:
                                            writeMessageFile('output/' + output_dir + '/', msg_counter, date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded, body, headers_decoded)
                                    else:
                                        writeMessageFile('output/' + output_dir + '/', msg_counter, date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded, body, headers_decoded)
                                    except Exception as err:
                                        print(err)
                                        pass        
                                
                                ##If the user chose to save mails to text files
                                if args.txt:
                                    if attachment_counter > 0:
                                        saveToFile('output/' + output_dir + '/', 'txt', file_stamp, False, date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded, headers_decoded, body, box_content['messages'][counter]['attachment'])
                                    else:
                                        saveToFile('output/' + output_dir + '/', 'txt', file_stamp, False, date, sender_decoded, recipient_decoded, copy_to_decoded, reply_to_decoded, subject_decoded, headers_decoded, body)
    
                                counter += 1
                                print (f"{color.blu}|    |{color.rst}")
                        sleep(0.1)
                    if args.json:
                        saveToFile('output/' + output_dir + '/', 'json', 'mailbox.json', box_content)
                    elif args.html:
                        for button in html_buttons:
                            writeToIndexFile('output/' + output_dir + '/html/index.html', button)
                        writeToIndexFile('output/' + output_dir + '/html/index.html', '		<!-- Modal -->')
                        for modal in html_modals:
                            writeToIndexFile('output/' + output_dir + '/html/index.html', modal)
                        writeHTMLFooter('output/' + output_dir + '/html/index.html')
                        
                    print (f"{color.blu}'----'----------------> {color.gre}[ DONE ]{color.rst}")
                else:
                    sys.exit (f"{color.blu}'--> {color.red}Mailbox empty or not a compatible mbox type!{color.rst}\n")
        else:
            sys.exit (f"{color.red}File {color.yel}{mbox_file}{color.red} not found!\n{color.yel}Remember to define full path if not located in the same folder with the script!{color.rst}\n")
    except KeyboardInterrupt:
        sys.exit (f"{color.yel}User interrupted{color.rst}\n")
    except Exception as err:
        sys.exit (f"{color.red}Something borked!\n{color.rst}  {color.yel}Error:{color.rst} {err}\n")
