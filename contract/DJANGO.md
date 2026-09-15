# Django 연동 가이드

엔진은 **JavaScript**이고 Django는 **Python**이라, 둘을 잇는 방법을 먼저 정해야 합니다.
결론부터: **엔진을 Python으로 포팅하지 마세요.** 아래 방법으로 JS 엔진을 그대로 쓰는 게 안전합니다.

## 어떤 방식을 고를지

| | 방식 | Node 필요 | 언제 |
|---|---|---|---|
| **A** | **브라우저에서 계산** (Django는 화면만 서빙) | ❌ | 사용자가 입력→결과를 바로 보는 화면. **가장 쉽고 지금과 동일하게 동작** |
| **B** | **Django가 Node를 서브프로세스로 호출** | ✅ | 결과를 DB에 저장, 서버에서 리포트 생성, 배치 |
| **C** | Node 마이크로서비스 + HTTP | ✅ | 호출량이 많아 프로세스 생성 비용이 부담될 때 |
| **D** | ~~Python으로 포팅~~ | ❌ | **권하지 않습니다** (아래 참조) |

> **D를 권하지 않는 이유**: 엔진 849줄은 원본 재무모델과 대조해 오차
> 0.01 KRWm 미만까지 맞춘 코드입니다. 포팅하면 전체를 다시 검증해야 하고,
> 두 벌을 유지하는 순간 반드시 어긋납니다. 실제로 이전에 "연 단위로 접는"
> 단순화를 했다가 Equity IRR이 6%p 틀어진 적이 있습니다.

대개 **A로 시작해서, 서버 저장이 필요해지면 B를 추가**하는 게 무난합니다.
A와 B는 같은 `engine2.js`를 쓰므로 결과가 항상 일치합니다.

---

## 방식 A — 브라우저에서 계산

Django는 화면만 내려주고 계산·엑셀은 브라우저에서 합니다. Node가 필요 없습니다.

```
myapp/static/solar_fm/   engine2.js  xlsxbuild2.js  app2.js  exceljs.min.js
myapp/templates/solar_fm/model.html
```

```django
{% load static %}
<link rel="stylesheet" href="{% static 'solar_fm/style.css' %}">
<div id="solar-fm"> … index2.html 의 <body> 안쪽 마크업 … </div>
<script src="{% static 'solar_fm/exceljs.min.js' %}"></script>
<script src="{% static 'solar_fm/engine2.js' %}"></script>
<script src="{% static 'solar_fm/xlsxbuild2.js' %}"></script>
<script src="{% static 'solar_fm/app2.js' %}"></script>
```

`index2.html`의 `<style>`을 `style.css`로, `<body>` 안쪽을 템플릿으로 옮기면 됩니다.
`{{ }}` `{% %}` `{#` 와 충돌하는 문자열은 **전 파일에 0건**이라 그대로 붙여도
Django 템플릿 엔진이 건드리지 않습니다.

> ⚠️ CSS 충돌 주의 — `README.md`의 "화면 삽입" 항목과 `THEMING.md` 참조.
> `html,body{}` `button,input,select{}` 세 줄이 기존 페이지에 영향을 줍니다.

## 방식 B — Django가 Node를 호출

포함된 `calc.js`(stdin/stdout JSON)와 `solar_fm.py`(Python 래퍼)를 쓰면 됩니다.

### 설치

```bash
# 서버에 Node 18+ 설치 후
cd /path/to/solar_fm && npm install     # exceljs (엑셀 안 쓰면 생략 가능)
```

`solar_fm.py` `calc.js` `engine2.js` `xlsxbuild2.js` `package.json`을 같은 폴더에 두세요.
Node 경로가 특이하면 환경변수 `SOLAR_FM_NODE`로 지정할 수 있습니다.

### 사용

```python
from . import solar_fm

def run_model(request):
    inp = json.loads(request.body)
    try:
        result = solar_fm.compute(inp, timeout=60)
    except solar_fm.SolarFMError as e:
        return JsonResponse({"error": str(e)}, status=400)

    return JsonResponse({
        "projectIRR": solar_fm.kpi_display(result["kpi"], "projectIRR", pct=True),
        "equityIRR":  solar_fm.kpi_display(result["kpi"], "equityIRR",  pct=True),
        "minDSCR":    solar_fm.kpi_display(result["kpi"], "minDSCRAnnual"),
        "tic":        result["tic"],
    })
```

엑셀까지:

```python
path = os.path.join(settings.MEDIA_ROOT, f"model_{pk}.xlsx")
solar_fm.compute(inp, xlsx_path=path)
```

### 성능

프로세스 생성 때문에 1회 호출에 **약 0.2~0.5초**가 듭니다. 사용자가 슬라이더를
움직일 때마다 부르기엔 부담이라, 그런 화면은 방식 A로 하고 저장 시점에만 B를
쓰는 조합을 권합니다.

---

## Python 쪽에서 반드시 지킬 것

### 1. IRR이 `None`으로 올 수 있습니다 — 0이 아닙니다

현금흐름에 부호전환이 없으면(매출 0, 자본금 0 등) IRR은 수학적으로
존재하지 않습니다. 엔진은 `NaN`을 반환하고 JSON을 거치며 `None`이 됩니다.

```python
# 나쁜 예 — "산출 불가"가 0%로 표시되어 사용자를 오도합니다
irr = result["kpi"]["projectIRR"] or 0

# 좋은 예
irr = solar_fm.kpi_display(result["kpi"], "projectIRR", pct=True)   # → "7.70%" 또는 "—"
```

`None` 가능: `projectIRR` `projectIRRPre` `equityIRR` `equityIRRPre`
`dividendIRR` `investorIRR` `paybackYears`
`equityMultiple` `ebitdaMargin` `minDSCR` `avgDSCR` (부채·매출 없을 때)

### 2. 입력 검증은 Django 쪽 책임입니다

엔진은 예외를 던지지 않습니다. 이상한 입력은 예외 대신 `NaN`/`0`으로 나옵니다.
Django Form이나 DRF Serializer로 막으세요. 특히:

* `constructionMonths >= 1` (0이면 총사업비가 현금유출로 안 잡힘)
* `capacityMW > 0`, `capexEok > 0`, `operationYears >= 1`
* `tranches[].order`는 정수, `amountEok >= 0`
* `ppy`는 **4 고정** — 사용자 입력으로 열지 마세요

### 3. 응답 크기

`compute()` 결과는 86분기 기준 **약 150KB JSON**입니다. 그대로 프런트에 내리면
무거우니, 화면에 필요한 필드만 골라 내려보내고 전체는 DB/파일에 저장하세요.

### 4. 타임아웃과 동시 실행

`compute()`의 기본 타임아웃은 60초입니다(실제 계산은 1초 미만).
동시 요청이 많으면 Node 프로세스가 그만큼 뜨므로, Celery 등 작업 큐로
직렬화하거나 방식 C를 검토하세요.

## 검증

연동 후 `sample-input.json`을 넣어 `sample-output.json`과 비교하면
제대로 붙었는지 바로 확인됩니다.

```python
import json, solar_fm
inp  = json.load(open("sample-input.json", encoding="utf-8"))
want = json.load(open("sample-output.json", encoding="utf-8"))
got  = solar_fm.compute(inp)
bad = [k for k, v in want["kpi"].items()
       if isinstance(v, (int, float))
       and abs((got["kpi"].get(k) or 0) - v) > max(abs(v) * 1e-9, 1e-9)]
assert not bad, bad
print("일치")
```
