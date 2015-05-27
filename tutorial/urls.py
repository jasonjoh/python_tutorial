# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.conf.urls import patterns, url 
from tutorial import views 

urlpatterns = patterns('', 
  # The home view ('/tutorial/') 
  url(r'^$', views.home, name='home'), 
  # Explicit home ('/tutorial/home/') 
  url(r'^home/$', views.home, name='home'), 
  # Redirect to get token ('/tutorial/gettoken/')
  url(r'^gettoken/$', views.gettoken, name='gettoken'),
  # Mail view ('/tutorial/mail/')
  url(r'^mail/$', views.mail, name='mail'),
) 

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