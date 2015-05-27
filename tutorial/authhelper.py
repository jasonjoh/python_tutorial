# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from urllib.parse import quote, urlencode
import requests

# Client ID and secret
client_id = 'YOUR CLIENT ID'
client_secret = 'YOUR CLIENT SECRET'

# Constant strings for OAuth2 flow
# The OAuth authority
authority = 'https://login.microsoftonline.com'

# The authorize URL that initiates the OAuth2 client credential flow for admin consent
authorize_url = '{0}{1}'.format(authority, '/common/oauth2/authorize?{0}')

# The token issuing endpoint
token_url = '{0}{1}'.format(authority, '/common/oauth2/token')

def get_signin_url(redirect_uri):
  # Build the query parameters for the signin url
  params = { 'client_id': client_id,
             'redirect_uri': redirect_uri,
             'response_type': 'code',
             'prompt': 'login',
           }
           
  signin_url = authorize_url.format(urlencode(params))
  
  return signin_url
  
def get_token_from_code(auth_code, redirect_uri):
  # Build the post form for the token request
  post_data = { 'grant_type': 'authorization_code',
                'code': auth_code,
                'redirect_uri': redirect_uri,
                'resource': 'https://outlook.office365.com',
                'client_id': client_id,
                'client_secret': client_secret
              }
              
  r = requests.post(token_url, data = post_data)
  
  try:
    access_token = r.json()['access_token']
    return access_token
  except:
    return 'Error retrieving token: {0} - {1}'.format(r.status_code, r.text)

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