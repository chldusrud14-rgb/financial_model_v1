from django.urls import path

from . import views

app_name = "solar_fm"

urlpatterns = [
    path("", views.financial_model, name="model"),
    # 서버 계산을 안 쓰면 아래 줄은 지워도 된다.
    path("run/", views.run_model, name="run"),
]
