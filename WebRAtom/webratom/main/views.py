from django.shortcuts import render, redirect
from rest_framework import viewsets
from rest_framework.permissions import IsAuthenticated, IsAdminUser
from rest_framework import generics
from django.contrib.auth.models import User
from rest_framework import permissions
from .serializers import UserSerializer
from django.shortcuts import render, get_object_or_404
from .models import Object


def object_list(request):
    objects = Object.objects.all()  # Извлечение всех объектов
    return render(request, 'object_list.html', {'objects': objects})

def object_detail(request, object_id):
    obj = get_object_or_404(Object, id_object=object_id)  # Получение объекта по ID
    return render(request, 'object_detail.html', {'object': obj})

class UserCreateView(generics.CreateAPIView):
    queryset = User.objects.all()
    serializer_class = UserSerializer
    permission_classes = [permissions.AllowAny]

def home(request):
    return render(request, 'main/home.html')

def redirect_to_streamlit(request):
    return redirect('http://localhost:8501')