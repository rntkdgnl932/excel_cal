# get_gcal_token_device.py
# -*- coding: utf-8 -*-
import json
import time
from pathlib import Path

import requests

SCOPES = ["https://www.googleapis.com/auth/calendar"]

DEVICE_CODE_URL = "https://oauth2.googleapis.com/device/code"
TOKEN_URL = "https://oauth2.googleapis.com/token"

def main():
    # === 여기만 수정 ===
    CLIENT_ID = "여기에_디바이스클라이언트_ID_붙여넣기"
    OUT_TOKEN_PATH = Path(r"C:\my_games\excel_cal\secrets\token.json")
    # ==================

    OUT_TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)

    # 1) device code 발급
    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
    }
    r = requests.post(DEVICE_CODE_URL, data=data, timeout=30)
    r.raise_for_status()
    dc = r.json()

    user_code = dc["user_code"]
    verification_url = dc.get("verification_url") or dc.get("verification_uri")
    verification_url_complete = dc.get("verification_url_complete") or dc.get("verification_uri_complete")
    device_code = dc["device_code"]
    interval = int(dc.get("interval", 5))
    expires_in = int(dc.get("expires_in", 1800))

    print("\n=== 구글 캘린더 토큰 발급 (Device Flow) ===")
    if verification_url_complete:
        print(f"1) 아래 URL을 브라우저에서 여세요:\n{verification_url_complete}\n")
        print("   (보통 자동으로 코드가 입력된 상태로 열립니다)")
    else:
        print(f"1) 아래 URL을 브라우저에서 여세요:\n{verification_url}\n")
        print(f"2) 다음 코드를 입력하세요: {user_code}\n")

    print("3) 권한 허용(Allow) 후, 이 콘솔로 돌아오세요. 토큰 발급을 대기합니다...\n")

    # 2) 승인될 때까지 polling
    start = time.time()
    while True:
        if time.time() - start > expires_in:
            raise RuntimeError("시간 초과: 디바이스 코드가 만료되었습니다. 다시 실행하세요.")

        payload = {
            "client_id": CLIENT_ID,
            "device_code": device_code,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        }
        tr = requests.post(TOKEN_URL, data=payload, timeout=30)
        tj = tr.json()

        if tr.status_code == 200 and "access_token" in tj:
            # 성공: refresh_token이 같이 와야 '다음부터 자동 갱신' 가능
            token_json = {
                "token": tj.get("access_token"),
                "refresh_token": tj.get("refresh_token"),
                "token_uri": "https://oauth2.googleapis.com/token",
                "client_id": CLIENT_ID,
                "scopes": SCOPES,
                "type": "authorized_user",
            }
            # expires_in은 초 단위; google-auth는 expiry를 선호하지만 없어도 동작 가능
            OUT_TOKEN_PATH.write_text(json.dumps(token_json, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"✅ token.json 저장 완료: {OUT_TOKEN_PATH}")

            if not token_json["refresh_token"]:
                print("⚠️ refresh_token이 비어 있습니다.")
                print("   - 이전에 같은 계정으로 이미 승인한 적이 있으면 refresh_token이 안 나올 수 있습니다.")
                print("   - 해결: 브라우저에서 해당 앱 권한을 철회 후 다시 실행하거나, prompt=consent가 필요한 플로우를 써야 합니다.")
            return

        err = tj.get("error")
        if err == "authorization_pending":
            time.sleep(interval)
            continue
        if err == "slow_down":
            interval += 2
            time.sleep(interval)
            continue

        raise RuntimeError(f"토큰 발급 실패: {tj}")

if __name__ == "__main__":
    main()
