# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
import requests
import uuid
import base64
import json
import os

outlook_api_endpoint = 'https://outlook.office.com/api/v1.0{0}'

# Generic API Sending
def make_api_call(method, url, token, user_email, payload = None, parameters = None):
    # Send these headers with all API calls
    headers = { 'User-Agent' : 'python_tutorial/1.0',
                'Authorization' : 'Bearer {0}'.format(token),
                'Accept' : 'application/json',
                'X-AnchorMailbox' : user_email }
                
    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = { 'client-request-id' : request_id,
                        'return-client-request-id' : 'true' }
                        
    headers.update(instrumentation)
    
    response = None
    
    if (method.upper() == 'GET'):
        response = requests.get(url, headers = headers, params = parameters)
    elif (method.upper() == 'DELETE'):
        response = requests.delete(url, headers = headers, params = parameters)
    elif (method.upper() == 'PATCH'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.patch(url, headers = headers, data = json.dumps(payload), params = parameters)
    elif (method.upper() == 'POST'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.post(url, headers = headers, data = json.dumps(payload), params = parameters, verify = False)
        
    return response
    
def get_my_messages(access_token, user_email):
  get_messages_url = outlook_api_endpoint.format('/Me/Messages')
  
  # Use OData query parameters to control the results
  #  - Only first 10 results returned
  #  - Only return the DateTimeReceived, Subject, and From fields
  #  - Sort the results by the DateTimeReceived field in descending order
  query_parameters = {'$top': '10',
                      '$select': 'DateTimeReceived,Subject,From',
                      '$orderby': 'DateTimeReceived DESC'}
                      
  r = make_api_call('GET', get_messages_url, access_token, user_email, parameters = query_parameters)
  
  if (r.status_code == requests.codes.ok):
    return r.json()
  else:
    return "{0}: {1}".format(r.status_code, r.text)
    
def send_message_with_attachment(access_token, user_email, file_path):
  # Simplistic scenario of sending a file to self to demonstrate
  # how to send an email with an send_message_with_attachment
  
  # First let's deal with the attachment. We need it in base64 format
  # to include in the payload to the API call.
  # Open the file in binary mode
  attachment_file = open(file_path, 'rb')
  
  # Read the contents, which will just be the bytes
  raw_attachment_contents = attachment_file.read()
  
  # Base64-encode the data. This returns a bytes object, and we need a 
  # string, so call decode('utf-8') to get a string representation.
  encoded_attachment_contents = base64.b64encode(raw_attachment_contents).decode('utf-8')
  
  # Set up the sendmail URL
  send_message_url = outlook_api_endpoint.format('/Me/SendMail')
  
  # Create the message payload
  message_payload = {
    "Message": {
      "Subject": "Send mail with attachment from API",
      "Body": {
        "ContentType": "HTML",
        "Content": "This was sent using the <strong>Outlook Mail API</strong>."
      },
      "ToRecipients": [
        {
          "EmailAddress": {
            "Address": user_email
          }
        }
      ],
      "Attachments": [
        {
          "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
          "Name": os.path.basename(attachment_file.name),
          "ContentBytes": encoded_attachment_contents
        }
      ]
    },
    "SaveToSentItems": "true"
  }
  
  response = make_api_call('POST', send_message_url, access_token, user_email, message_payload)
  
  if (response.status_code == requests.codes.accepted):
    return "SUCCESS"
  else:
    return "FAILURE"

# MIT License: 
 
# Permission is hereby granted, free of charge, to any person obtaining 
# a copy of this software and associated documentation files (the 
# ""Software""), to deal in the Software without restriction, including 
# without limitation the rights to use, copy, modify, merge, publish, 
# distribute, sublicense, and/or sell copies of the Software, and to 
# permit persons to whom the Software is furnished to do so, subject to 
# the following conditions: 
 
# The above copyright notice and this permission notice shall be 
# included in all copies or substantial portions of the Software. 
 
# THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.