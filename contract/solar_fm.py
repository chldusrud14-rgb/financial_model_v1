# -*- coding: utf-8 -*-
"""Django 등 Python 쪽에서 재무모델 엔진(JS)을 호출하는 얇은 래퍼.

계산 로직은 Node 로 돌아가는 engine2.js 안에 그대로 있고, 이 파일은
프로세스를 띄워 JSON 을 주고받기만 한다. 엔진을 Python 으로 포팅하지
않는 이유는 원본 재무모델과 대조 검증된 코드가 JS 쪽에만 있고, 두 벌을
유지하면 반드시 어긋나기 때문이다.

전제: 서버에 Node 18+ 가 설치돼 있고, 이 파일과 같은 폴더에
      calc.js / engine2.js / xlsxbuild2.js 가 있을 것.
      엑셀까지 쓰려면 그 폴더에서 `npm install exceljs` 필요.
"""
import json
import math
import os
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
CALC_JS = os.path.join(HERE, "calc.js")
NODE_BIN = os.environ.get("SOLAR_FM_NODE") or shutil.which("node") or "node"

# 현금흐름 부호전환이 없으면 수학적으로 정의되지 않는 지표들.
# 엔진이 NaN 을 돌려주고 JSON 직렬화 과정에서 None 이 된다.
MAYBE_UNDEFINED = (
    "projectIRR", "projectIRRPre", "equityIRR", "equityIRRPre",
    "dividendIRR", "investorIRR", "paybackYears",
)


class SolarFMError(RuntimeError):
    """엔진 호출 실패 (입력 오류, Node 미설치, 타임아웃 등)."""


def compute(inp, xlsx_path=None, timeout=60):
    """입력 dict 을 넣으면 결과 dict 을 돌려준다.

    inp        : 입력값 (README.md 의 입력 스펙 참조)
    xlsx_path  : 지정하면 그 경로로 엑셀 파일도 생성
    timeout    : 초 단위 상한

    반환 dict 에는 rows / kpi / con / tic 등이 들어있고,
    정의되지 않는 IRR 은 None 으로 온다 (0 이 아님에 주의).
    """
    if not os.path.exists(CALC_JS):
        raise SolarFMError("calc.js 를 찾을 수 없습니다: %s" % CALC_JS)

    cmd = [NODE_BIN, CALC_JS]
    if xlsx_path:
        cmd += ["--xlsx", xlsx_path]

    try:
        proc = subprocess.run(
            cmd,
            input=json.dumps(inp, ensure_ascii=False),
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=timeout,
        )
    except FileNotFoundError:
        raise SolarFMError(
            "Node 를 실행할 수 없습니다 (%s). 서버에 Node 18+ 를 설치하거나 "
            "환경변수 SOLAR_FM_NODE 로 경로를 지정하세요." % NODE_BIN
        )
    except subprocess.TimeoutExpired:
        raise SolarFMError("계산이 %d초 안에 끝나지 않았습니다." % timeout)

    if not proc.stdout.strip():
        raise SolarFMError("엔진이 아무 출력도 내지 않았습니다. stderr=%s"
                           % proc.stderr.strip()[:300])
    try:
        payload = json.loads(proc.stdout)
    except json.JSONDecodeError:
        raise SolarFMError("엔진 출력이 JSON 이 아닙니다: %s"
                           % proc.stdout.strip()[:300])

    if not payload.get("ok"):
        raise SolarFMError(payload.get("error") or "알 수 없는 오류")

    return payload["result"]


def kpi_display(kpi, key, pct=False, digits=2):
    """KPI 를 화면에 표시할 문자열로. 정의되지 않으면 '—' 를 돌려준다.

    IRR 이 None/NaN 인 것은 버그가 아니라 '수학적으로 존재하지 않음'이므로
    0% 로 표시하면 안 된다.
    """
    v = kpi.get(key)
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    if pct:
        return "{:.{d}f}%".format(v * 100, d=digits)
    return "{:.{d}f}".format(v, d=digits)
