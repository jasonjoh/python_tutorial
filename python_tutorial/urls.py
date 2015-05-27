from django.conf.urls import patterns, include, url
from django.contrib import admin

urlpatterns = patterns('',
    # Invoke the home view in the tutorial app by default
    url(r'^$', 'tutorial.views.home', name='home'),
    # Defer any URLS to the /tutorial directory to the tutorial app
    url(r'^tutorial/', include('tutorial.urls', namespace='tutorial')),
    url(r'^admin/', include(admin.site.urls)),
)
