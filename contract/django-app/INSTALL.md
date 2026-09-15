# Django 앱 설치 — 이 폴더를 그대로 복사하면 됩니다

`solar_fm/` 폴더를 Django 프로젝트 루트(`manage.py` 가 있는 곳)에 복사하세요.

```
프로젝트/
├── manage.py
├── config/           (settings.py 가 있는 곳, 이름은 프로젝트마다 다름)
└── solar_fm/         ← 이 폴더를 통째로 복사
```

## 1. 앱 등록 — `settings.py`

```python
INSTALLED_APPS = [
    ...
    "solar_fm",
]
```

## 2. URL 연결 — 프로젝트 `urls.py`

```python
from django.urls import include, path

urlpatterns = [
    ...
    path("financial-model/", include("solar_fm.urls")),
]
```

## 3. 좌측 메뉴에 링크 추가

```django
<a href="{% url 'solar_fm:model' %}">재무모델</a>
```

## 4. `base.html` 블록 이름 맞추기

`templates/solar_fm/model.html` 이 `extra_css` / `content` / `extra_js`
세 블록을 씁니다. 프로젝트의 `base.html` 블록 이름이 다르면 그 파일에서
세 곳만 바꿔주세요. 그 외 수정할 곳은 없습니다.

## 5. 정적 파일

개발 중에는 `django.contrib.staticfiles` 가 알아서 서빙합니다.
배포 시에는 평소처럼 `collectstatic` 하시면 됩니다.

```bash
python manage.py collectstatic
```

## 확인

서버를 띄우고 `/financial-model/` 로 접속해서:

1. 화면이 좌측 메뉴 오른쪽에 뜨는지
2. **좌측 메뉴 모양이 그대로인지** (CSS 가 새지 않았는지)
3. "예시 불러오기 (당진1, 100MW급 PJT)" → "재무모델 생성" → KPI 표시
4. "Excel 다운로드" 클릭 → `.xlsx` 파일 받아지는지

여기까지 되면 연동 완료입니다. **Node 설치나 서버 설정은 필요 없습니다** —
계산과 엑셀 생성이 전부 브라우저에서 이루어집니다.

---

## (선택) 서버에서 계산하기

결과를 DB 에 저장하거나 서버에서 리포트를 만들어야 할 때만 필요합니다.

1. 서버에 **Node 18+** 설치
2. `cd solar_fm && npm install` (exceljs)
3. `views.run_model` 로 POST — 자세한 건 상위 폴더 `DJANGO.md` 참조

쓰지 않을 거면 앱에서 아래 파일들을 지워도 화면은 정상 동작합니다.

```
solar_fm.py  calc.js  engine2.js  xlsxbuild2.js  package.json
```

(주의: `static/solar_fm/` 안의 `engine2.js` `xlsxbuild2.js` 는
**브라우저 계산용이라 지우면 안 됩니다.** 앱 루트에 있는 동명 파일만
서버 계산용입니다.)

## 주의사항

* **`id` 를 바꾸지 마세요.** JS 가 직접 참조합니다(`#run` `#xls` `#core` `#trbox` 등).
  스타일은 class 로 걸어주세요.
* **`data-plant` 요소의 `display` 를 CSS 로 덮어쓰지 마세요.** 태양광/풍력 토글이 깨집니다.
* 색상 변경은 `static/solar_fm/solar-fm.css` 상단 `#solar-fm{...}` 블록의
  변수 36개만 고치면 됩니다. 상위 폴더 `THEMING.md` 참조.
* 호스트 CSS 가 `.card` `.btn` 등에 `!important` 를 쓰면 그것만 스코프를 뚫습니다.
  `scoped/README-scoped.md` 의 확인 방법과 해결책 참조.
