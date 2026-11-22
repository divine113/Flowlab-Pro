from django.contrib import admin
from django.http import JsonResponse
from django.urls import include, path
from django.conf import settings
from django.conf.urls.static import static


def health(request):
    """Lightweight liveness probe for the standalone E&I Lab server."""
    return JsonResponse({'status': 'ok', 'service': 'ei-lab'})


urlpatterns = [
    path('health', health, name='health'),
    path('admin/', admin.site.urls),
    path('api/', include('core.api.urls')),
    path('', include('core.urls')),
]

if settings.DEBUG:
    urlpatterns += [path('__reload__/', include('django_browser_reload.urls'))]
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
