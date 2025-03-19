from django.urls import path, include
from .views import home, redirect_to_streamlit
from rest_framework import routers
from . import views
from django.contrib import admin
from django.urls import path, include
from drf_yasg.views import get_schema_view
from drf_yasg import openapi
from .views import UserCreateView
from rest_framework import permissions
from .views import object_list, object_detail

schema_view = get_schema_view(
   openapi.Info(
      title="My API",
      default_version='v1',
      description="Test description",
      terms_of_service="https://www.google.com/policies/terms/",
      contact=openapi.Contact(email="contact@myapi.local"),
      license=openapi.License(name="BSD License"),
   ),
   public=True,
   permission_classes=(permissions.AllowAny,),
)

urlpatterns = [
    path('admin/', admin.site.urls),
    path('swagger/', schema_view.with_ui('swagger', cache_timeout=0), name='schema-swagger-ui'),
    path('api/users/', UserCreateView.as_view(), name='user-create'),
    path('redoc/', schema_view.with_ui('redoc', cache_timeout=0), name='schema-redoc'),
    path('home/', home, name='home'),
    path('go-to-streamlit/', redirect_to_streamlit, name='redirect_to_streamlit'),
    path('objects/', object_list, name='object_list'),
    path('objects/<int:object_id>/', object_detail, name='object_detail'),
]