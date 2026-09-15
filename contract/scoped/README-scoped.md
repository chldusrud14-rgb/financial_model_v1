# 스코프 격리판 — 공통 레이아웃(좌측 메뉴) 안에 넣기

좌측 메뉴가 유지되는 화면 오른쪽에 재무모델을 띄우는 구조용입니다.
CSS 규칙 158개 전부가 `#solar-fm` 안에서만 적용되도록 변환돼 있습니다.

## 파일

| 파일 | Django 배치 위치 |
|---|---|
| `solar-fm.css` | `static/solar_fm/` |
| `solar-fm.html` | `templates/solar_fm/` (마크업, `<div id="solar-fm">` 로 감싸져 있음) |
| `exceljs.min.js` | `static/solar_fm/` (925KB — 엑셀 다운로드용) |
| `engine2.js` | `static/solar_fm/` |
| `xlsxbuild2.js` | `static/solar_fm/` |
| `app2.js` | `static/solar_fm/` |
| `dangjin-reference.js` | `static/solar_fm/` (예시 불러오기 데이터, 48KB) |

## 사용법

```django
{% extends "base.html" %}
{% load static %}

{% block extra_css %}
  <link rel="stylesheet" href="{% static 'solar_fm/solar-fm.css' %}">
{% endblock %}

{% block content %}
  {% include "solar_fm/solar-fm.html" %}
{% endblock %}

{% block extra_js %}
  <script src="{% static 'solar_fm/dangjin-reference.js' %}"></script>
  <script src="{% static 'solar_fm/exceljs.min.js' %}"></script>
  <script src="{% static 'solar_fm/engine2.js' %}"></script>
  <script src="{% static 'solar_fm/xlsxbuild2.js' %}"></script>
  <script src="{% static 'solar_fm/app2.js' %}"></script>
{% endblock %}
```

**스크립트 순서를 지켜주세요.** `app2.js` 가 앞의 것들을 참조합니다.
스크립트는 마크업보다 뒤에 와야 합니다.

`{{ }}` `{% %}` `{#` 와 충돌하는 문자열은 전 파일에 0건이라 그대로 넣어도
Django 템플릿 엔진이 건드리지 않습니다.

## 격리 수준 — 실제로 검증한 결과

일부러 충돌하게 만든 호스트 페이지(Georgia 서체, 어두운 배경,
`.card` `.btn` `.hint` 재정의, `button,input,select` 재정의)를 만들어
실제 브라우저에 띄워 확인했습니다.

| 방향 | 결과 |
|---|---|
| 재무모델 CSS → 호스트 | **영향 0.** 좌측 메뉴의 서체·버튼·입력칸 모두 그대로 |
| 호스트 CSS → 재무모델 | **일반 규칙은 차단됨** (`#solar-fm .card` 가 `.card` 를 이김) |
| 호스트의 `!important` → 재무모델 | ⚠️ **통과함** |

기능도 호스트 안에서 정상 동작을 확인했습니다 — 예시 불러오기, 생성,
KPI 산출(단독 실행과 동일), 엑셀 다운로드 151KB, 발전원 토글.

### `!important` 만 남은 구멍

CSS 우선순위 규칙상 `!important` 는 선택자 구체성을 무시하고 이깁니다.
호스트가 아래 이름들에 `!important` 를 쓰면 재무모델 화면이 영향을 받습니다.

```
accent bad badge body btn card cols dim f full ghost good hd hint
in kpi on pill same toast top unit warn wrap
```

**확인 방법** — 호스트 CSS에서 이 이름들에 `!important` 를 쓰는지 검색해보세요.

```bash
grep -rE '\.(card|btn|hint|top|wrap|kpi|badge|pill)\b[^{]*\{[^}]*!important' static/
```

**걸리는 게 없으면 그대로 쓰시면 됩니다.** Bootstrap·Tailwind 기본값은
이 이름들에 `!important` 를 쓰지 않습니다.

**걸리는 게 있으면** 셋 중 하나로 해결하세요.

1. 호스트 쪽 `!important` 를 걷어낸다 (가장 깔끔)
2. 그 항목만 `#solar-fm .card{ ... !important }` 로 되받아친다
3. `iframe` 으로 `standalone.html` 을 띄운다 (완전 격리, 대신 톤 맞추기가 번거로움)

## 색상 변경

`solar-fm.css` 상단 `#solar-fm{ ... }` 블록의 변수 36개만 고치면 됩니다.
CSS 규칙 쪽에 색상 하드코딩은 0곳입니다. 자세한 내용은 상위 폴더의
`THEMING.md` 를 보세요.

## 원본과의 관계

이 폴더는 상위 폴더의 `index2.html` 에서 **자동 변환된 결과물**입니다.
화면 로직을 고칠 일이 생기면 이 파일들을 직접 고치기보다 원본을 고치고
다시 변환하는 편이 안전합니다.
