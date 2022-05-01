import requests
import email

def get_messages(access_token, id_folder_bandeja_entrada, id_folder_alertas_phishing, next_link):
    url = f'https://graph.microsoft.com/v1.0/me/mailfolders/{id_folder_bandeja_entrada}/childfolders/{id_folder_alertas_phishing}/messages'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    if next_link is not None:
        url = next_link
    get_result = requests.get(url, headers=headers)
    return get_result.json()

def get_attachment(access_token, id_folder_bandeja_entrada, id_folder_alertas_phishing, id_message):
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    url = f'https://graph.microsoft.com/v1.0/me/mailfolders/{id_folder_bandeja_entrada}/childfolders/{id_folder_alertas_phishing}/messages/{id_message}/attachments/'
    get_result = requests.get(url, headers=headers)
    id_attachment = get_result.json()['value'][0]['id']
    url_attachment = f'{url}{id_attachment}/$value'
    get_result_attch = requests.get(url_attachment, headers=headers)
    return get_result_attch.text

def manage_header(header):
    decoded_header = email.header.decode_header(header)
    result = ''
    if len(decoded_header) > 1:
        for value in decoded_header:
            try:
                str_header = value[0].decode()
            except:                
                str_header = value[0]
            result += str(str_header)
        return result
    result = decoded_header[0][0]
    try:
        return result.decode()
    except:
        return result

if "__main__" == __name__:
    access_token = 'JWT_Access_Token'
    id_folder_bandeja_entrada = 'ID_inbox_Folder'
    id_subfolder = 'ID_inbox_subfolder'
    messages_alertas_phishing = get_messages(access_token, id_folder_bandeja_entrada, id_subfolder, None)
    next_link = messages_alertas_phishing['@odata.nextLink']
    count = 0
    while next_link is not None:
        for message in messages_alertas_phishing['value']:
            count += 1
            if True:
                id_message = message['id']
                sender = message['sender']
                report_date = message['createdDateTime']
                attachment = get_attachment(access_token, id_folder_bandeja_entrada, id_subfolder, id_message)
                full_message = email.message_from_string(attachment)
                source = manage_header(full_message['From'])
                subject = manage_header(full_message['Subject'])
                print('Correo', str(count), 
                        '\nNombre reporte:', sender['emailAddress']['name'], 
                        '\nCorreo reporte:', sender['emailAddress']['address'],
                        '\nFecha reporte:', report_date,
                        '\nRemitente adjunto:', source, 
                        '\nAsunto adjunto:', subject, '\n\n')        
        try:
            next_link = messages_alertas_phishing['@odata.nextLink']
            messages_alertas_phishing = get_messages(access_token, id_folder_bandeja_entrada, id_subfolder, next_link)
        except:
            next_link = None
