import json
import logging
import os
from typing import List
import requests

logger = logging.getLogger(__name__)


def parse_recipients(recipients: List) -> dict:
    array_recipents_length = len(recipients)
    recipients_array = []
    for recipient in range(array_recipents_length):
        json_data = {}
        json_data_level = {}
        json_data_level['address'] = recipients[recipient]
        json_data['emailAddress'] = json_data_level
        recipients_array.append(json_data)
            
    return (recipients_array)


# This function is used to calculate file size and necessary chunks
def calculate_file_size_and_necessary_chunks(filename, CHUNK_SIZE):
    st = os.stat(filename)
    size = st.st_size
    
    chunks = int(size / CHUNK_SIZE) + 1 if size % CHUNK_SIZE > 0 else 0
    necessaryData = [size, chunks]
    return necessaryData


# This file is used to break up files into the necessary chunks and upload them accordingly
def upload_file_to_upload_url(attatchment_Filenames, attatchemnt, file_size_and_chunks, CHUNK_SIZE, upload_url):
    with open(attatchment_Filenames[attatchemnt], 'rb') as fd:
        start = 0
        for chunk_num in range(file_size_and_chunks[1]):
            chunk = fd.read(CHUNK_SIZE)
            bytes_read = len(chunk)
            upload_range = f'bytes {start}-{start + bytes_read - 1}/{file_size_and_chunks[0]}'
            logger.debug(f'chunk: {chunk_num} bytes read: {bytes_read} upload range: {upload_range}')
            result = requests.put(
                upload_url,
                headers={
                    'Content-Length': str(bytes_read),
                    'Content-Range': upload_range
                },
                data=chunk
            )
            result.raise_for_status()
            start += bytes_read


'''
Attatchment_Filenames should be an array of strings: example Attatchment_Filenames = ["test2.xlsx", "test.xlsx"]
app should be an Application_instance_daemon class object
subject should be the message subject passt in as plain text in a string: example subject = "Hello how are you doing?"
importance defines the imporance of the message The possible values are: low, normal, and high. : example importance = "Low"
body sould be the message body writen in html encased within a stirng: example body = "They were <b>awesome</b>!"
'''


def send_mail_with_attatchments(graph_app, attatchment_Filenames, subject, body, sender, recipients, bccrecipients=None, importance=None):
    # This is the chunk_size limit recomended by Microsoft
    CHUNK_SIZE = 10485760
    
    if (importance is None):
        importance = "normal"
    
    bccrecipients_lists = []
    if (bccrecipients is not None):
        bccrecipients_lists = parse_recipients(bccrecipients)
    recipients_lists = parse_recipients(recipients)
    # Acquiring the access token
    result = graph_app.acquire_token()
      
    if ("access_token" in result):
        # Calling graph using the access token
        access_token = result['access_token']
        # Create the draft file
        graph_data = requests.post(
            f"{graph_app.endpoint}/{sender}/messages",
            headers={'Authorization': 'Bearer ' + result['access_token']},
            json={
                "subject": subject,
                "importance": importance,
                "body": {
                    "contentType": "HTML",
                    "content": body
                },
                "toRecipients": recipients_lists,
                "bccRecipients": bccrecipients_lists
            }).json()
        logger.debug("Graph draft message API call result: ")
        # This is done so we can acces the objects inside the json
        responseJsonString = json.loads(json.dumps(graph_data))
        logger.debug(json.dumps(responseJsonString, indent=2))
        
        if (responseJsonString['id']):
            for attatchemnt in range(len(attatchment_Filenames)):
                file_size_and_chunks = calculate_file_size_and_necessary_chunks(attatchment_Filenames[attatchemnt], CHUNK_SIZE)
                graph_data = requests.post(
                    f"{graph_app.endpoint}/{sender}/messages/{responseJsonString['id']}/attachments/createUploadSession",
                    headers={'Authorization': 'Bearer ' + access_token},
                    json={
                        "AttachmentItem": {
                            "attachmentType": "file",
                            "name": attatchment_Filenames[attatchemnt],
                            "size": file_size_and_chunks[0]
                        }
                    }
                )
                upload_session = json.loads(graph_data.text)
                upload_url = upload_session['uploadUrl']
                logger.debug("File to upload: " + attatchment_Filenames[attatchemnt])
                logger.info("Create upload link status code: ")
                logger.info(graph_data.status_code)

                # Upload the file
                if (upload_url):
                    upload_file_to_upload_url(attatchment_Filenames, attatchemnt, file_size_and_chunks, CHUNK_SIZE, upload_url)
                else:
                    logger.error("no upload url")

            # Sending the draftmessage
            graph_data = requests.post(
                f"{graph_app.endpoint}/{sender}/messages/{responseJsonString['id']}/send",
                headers={'Authorization': 'Bearer ' + access_token}, )
            logger.info("Message send status code:")
            logger.info(graph_data.status_code)
            # check status code
            if (graph_data.status_code != 202):
                raise Exception("status code error: " + str(graph_data.status_code))
                
    else:
        logger.error(result.get("error"))
        logger.error(result.get("error_description"))
        logger.error(result.get("correlation_id"))  # You may need this when reporting a bug


def send_mail(graph_app, subject, body, sender, recipients, bccrecipients=None, importance=None):
   
    if (importance is None):
        importance = "normal"

    bccrecipients_lists = []
    if (bccrecipients is not None):
        bccrecipients_lists = parse_recipients(bccrecipients)
    recipients_lists = parse_recipients(recipients)

    result = graph_app.acquire_token()

    if "access_token" in result:
        # Calling graph using the access token
        graph_data = requests.post(
            f"{graph_app.endpoint}/{sender}/sendMail",
            headers={'Authorization': 'Bearer ' + result['access_token']},
            json= {
                'message': {
                    # recipient list
                    'toRecipients': recipients_lists,
                    "bccRecipients": bccrecipients_lists,
                    # email subject
                    'subject': subject,
                    'importance': importance,
                    'body': {
                        'contentType': 'HTML',
                        'content': body
                    }
                }
            })
        logger.info("Graph API call result: ")
        logger.info(graph_data.status_code)
        logger.debug(graph_data.reason)
        if (graph_data.status_code != 202):
            raise Exception("status code error: " + str(graph_data.status_code))
    else:
        logger.error(result.get("error"))
        logger.error(result.get("error_description"))
        logger.error(result.get("correlation_id"))  # You may need this when reporting a bug


'''
email_address should be the email address you wish to read.
The messages from it should belong to your organitzation: example email_address = "test@test.com"
mail_boxname should be the mail box you wish to read from: example mail_box_name = "Inbox"
'''


def read_mail_from_folder(graph_app, email_address, mail_box_name):
 
    result = graph_app.acquire_token()

    if ("access_token" in result):
        # Calling graph using the access token
        graph_data = requests.get(
            f"{graph_app.endpoint}/{email_address}/mailFolders/{mail_box_name}/messages",
            headers={'Authorization': 'Bearer ' + result['access_token']}, ).json()
        logger.debug("Graph API call result: ")
        logger.debug(json.dumps(graph_data, indent=2))
        return graph_data
    else:
        logger.error(result.get("error"))
        logger.error(result.get("error_description"))
        logger.error(result.get("correlation_id"))  # You may need this when reporting a bug
