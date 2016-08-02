# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from urllib.parse import quote, urlencode
import requests
import base64
import json
import time

# Client ID and secret
client_id = 'YOUR APP ID HERE'
client_secret = 'YOUR APP PASSWORD HERE'

# Constant strings for OAuth2 flow
# The OAuth authority
authority = 'https://login.microsoftonline.com'

# The authorize URL that initiates the OAuth2 client credential flow for admin consent
authorize_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/authorize?{0}')

# The token issuing endpoint
token_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/token')

# The scopes required by the app
scopes = [ 'openid',
           'offline_access',
           'https://outlook.office.com/mail.read',
           'https://outlook.office.com/calendars.read',
           'https://outlook.office.com/contacts.read' ]

def get_signin_url(redirect_uri):
  # Build the query parameters for the signin url
  params = { 'client_id': client_id,
             'redirect_uri': redirect_uri,
             'response_type': 'code',
             'scope': ' '.join(str(i) for i in scopes)
           }
           
  signin_url = authorize_url.format(urlencode(params))
  
  return signin_url
  
def get_token_from_code(auth_code, redirect_uri):
  # Build the post form for the token request
  post_data = { 'grant_type': 'authorization_code',
                'code': auth_code,
                'redirect_uri': redirect_uri,
                'scope': ' '.join(str(i) for i in scopes),
                'client_id': client_id,
                'client_secret': client_secret
              }
              
  r = requests.post(token_url, data = post_data)
  
  try:
    return r.json()
  except:
    return 'Error retrieving token: {0} - {1}'.format(r.status_code, r.text)

def get_token_from_refresh_token(refresh_token, redirect_uri):
  # Build the post form for the token request
  post_data = { 'grant_type': 'refresh_token',
                'refresh_token': refresh_token,
                'redirect_uri': redirect_uri,
                'scope': ' '.join(str(i) for i in scopes),
                'client_id': client_id,
                'client_secret': client_secret
              }
              
  r = requests.post(token_url, data = post_data)
  
  try:
    return r.json()
  except:
    return 'Error retrieving token: {0} - {1}'.format(r.status_code, r.text)

def get_access_token(request, redirect_uri):
  current_token = request.session['access_token']
  expiration = request.session['token_expires']
  now = int(time.time())
  if (current_token and now < expiration):
    # Token still valid
    return current_token
  else:
    # Token expired
    refresh_token = request.session['refresh_token']
    new_tokens = get_token_from_refresh_token(refresh_token, redirect_uri)

    # Update session
    # expires_in is in seconds
    # Get current timestamp (seconds since Unix Epoch) and
    # add expires_in to get expiration time
    # Subtract 5 minutes to allow for clock differences
    expiration = int(time.time()) + new_tokens['expires_in'] - 300
    
    # Save the token in the session
    request.session['access_token'] = new_tokens['access_token']
    request.session['refresh_token'] = new_tokens['refresh_token']
    request.session['token_expires'] = expiration

    return new_tokens['access_token']

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