# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect
from django.core.urlresolvers import reverse
from tutorial.authhelper import get_signin_url, get_token_from_code
from tutorial.outlookservice import get_my_messages

# Create your views here.

def home(request):
  redirect_uri = request.build_absolute_uri(reverse('tutorial:gettoken'))
  sign_in_url = get_signin_url(redirect_uri)
  return HttpResponse('<a href="' + sign_in_url +'">Click here to sign in and view your mail</a>')
  
def gettoken(request):
  auth_code = request.GET['code']
  redirect_uri = request.build_absolute_uri(reverse('tutorial:gettoken'))
  access_token = get_token_from_code(auth_code, redirect_uri)
  # Save the token in the session
  request.session['access_token'] = access_token
  return HttpResponseRedirect(reverse('tutorial:mail'))
  
def mail(request):
  access_token = request.session['access_token']
  # If there is no token in the session, redirect to home
  if not access_token:
    return HttpResponseRedirect(reverse('tutorial:home'))
  else:
    messages = get_my_messages(access_token)
    context = { 'messages': messages['value'] }
    return render(request, 'tutorial/mail.html', context)
    
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