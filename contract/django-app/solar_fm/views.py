# -*- coding: utf-8 -*-
"""재무모델 화면.

계산은 기본적으로 **브라우저에서** 이루어진다(static/solar_fm/engine2.js).
서버에서 계산해 DB 에 저장해야 할 때만 아래 run_model 을 쓴다.
"""
import json

from django.http import JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_POST


def financial_model(request):
    """재무모델 화면을 띄운다. 계산은 브라우저에서 한다."""
    return render(request, "solar_fm/model.html")


@require_POST
def run_model(request):
    """(선택) 서버에서 계산해 결과를 돌려준다. Node 18+ 필요.

    이 뷰를 쓰지 않을 거면 solar_fm.py / calc.js / engine2.js /
    xlsxbuild2.js / package.json 을 앱에서 지워도 된다.
    """
    from . import solar_fm  # Node 를 안 쓰는 배포에서는 import 자체를 피한다

    try:
        inp = json.loads(request.body or "{}")
    except json.JSONDecodeError:
        return JsonResponse({"error": "요청 본문이 JSON 이 아닙니다."}, status=400)

    # 엔진은 예외를 던지지 않는다 — 최소한의 입력 검증은 여기서 한다.
    if not inp.get("capacityMW"):
        return JsonResponse({"error": "설비용량(capacityMW)이 필요합니다."}, status=400)
    if (inp.get("constructionMonths") or 0) < 1:
        return JsonResponse({"error": "공사기간(constructionMonths)은 1개월 이상이어야 합니다."}, status=400)

    try:
        result = solar_fm.compute(inp, timeout=60)
    except solar_fm.SolarFMError as e:
        return JsonResponse({"error": str(e)}, status=400)

    kpi = result["kpi"]
    # 전체 결과는 86분기 기준 약 150KB 라 그대로 내리지 않는다.
    # IRR 은 정의되지 않을 수 있으므로 kpi_display 로 문자열화한다("—").
    return JsonResponse({
        "projectIRR": solar_fm.kpi_display(kpi, "projectIRR", pct=True),
        "equityIRR": solar_fm.kpi_display(kpi, "equityIRR", pct=True),
        "dividendIRR": solar_fm.kpi_display(kpi, "dividendIRR", pct=True),
        "minDSCR": solar_fm.kpi_display(kpi, "minDSCRAnnual"),
        "npv": kpi.get("npv"),
        "tic": result.get("tic"),
    })
