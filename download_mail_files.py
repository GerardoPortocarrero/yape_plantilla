import os
import win32com.client

# Conectar con la bandeja de entrada de outlook y obtener correos
def get_outlook_mails(
        OUTLOOK,
        DOT,
        APPLICATION,
        MAPI,
):
    outlook = win32com.client.Dispatch(OUTLOOK+DOT+APPLICATION).GetNamespace(MAPI)
    outlook_folder = outlook.GetDefaultFolder(6) # Conectar con la bandeja de entrada
    
    # Obtener y ordenar correos por fecha descendente
    mails = outlook_folder.Items
    mails.Sort("[ReceivedTime]", True) # (mails) Es un objeto lista
    
    return mails

# Obtener archivos de los correos
def get_outlook_files(
        PROJECT_ADDRESS,
        LOCACIONES,
        MAILS,
        MAIL_ITEM,
):
    files_found = {}

    for mail in MAILS:
        # Si no es de tipo mail o no tiene ningun adjunto -> SALTAR
        if mail.Class != MAIL_ITEM or mail.Attachments.Count == 0:
            continue

        email_lower = mail.SenderEmailAddress.lower()
        subject_lower = mail.Subject.lower()
        mail_received_time = mail.ReceivedTime.strftime("%Y-%m-%d")

        # Buscar attachments
        for attachment in mail.Attachments:
            # Si el adjunto no es un excel -> SALTAR
            if not attachment.FileName.endswith((".xlsx", ".xls")):
                continue
            
            # Si el asunto del emisor (remitente) coincide con las keywords
            for locacion in LOCACIONES:
                matches = 0
                subject_keywords = len(locacion['mail_subject'])

                for element in locacion['mail_subject']:
                    if element in subject_lower:
                        matches += 1
                
                if matches == subject_keywords:
                    print(locacion['name'], email_lower, mail_received_time)
                    file_name = attachment.FileName
                    file_address = os.path.join(PROJECT_ADDRESS, attachment.FileName)
                    files_found[locacion['name']] = [file_name, file_address, mail_received_time]
                    attachment.SaveAsFile(file_address)

            # Si el email del emisor (remitente) no se reconoce -> SALTAR
                # if locacion['sender_email'].lower() == email_lower:
                #     file_name = attachment.FileName
                #     file_address = os.path.join(PROJECT_ADDRESS, attachment.FileName)
                #     files_found[locacion['name']] = [file_name, file_address, mail_received_time]
                #     attachment.SaveAsFile(file_address)

            print(f'Files found: {files_found}\n')
            if len(files_found) == 4:
                return files_found
                                        

''' MAIL PROPERTIES
    | (mail.Subject) (mail.ReceivedTime) (mail.SenderName)       |
    | (mail.SenderEmailAddress) (mail.To) (mail.CC)              |
    | (mail.Body) (mail.Attachments.Count) (mail.CreationTime)   |
    | (mail.LastModificationTime) (mail.EntryID)                 |
'''

# Funcion principal
def main(
        PROJECT_ADDRESS,
        LOCACIONES,
        OUTLOOK,
        DOT,
        APPLICATION,
        MAPI,
        MAIL_ITEM,
):
    mails = get_outlook_mails(OUTLOOK, DOT, APPLICATION, MAPI)
    files = get_outlook_files(PROJECT_ADDRESS, LOCACIONES, mails, MAIL_ITEM)

    return files