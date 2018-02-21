from django.conf.urls import url, include
from django.contrib import admin
from tutorial import views

urlpatterns = [
    # Invoke the home view in the tutorial app by default
    url(r'^$', views.home, name='home'),
    # Defer any URLS to the /tutorial directory to the tutorial app
    url(r'^tutorial/', include('tutorial.urls', namespace='tutorial')),
    url(r'^admin/', admin.site.urls),
]
