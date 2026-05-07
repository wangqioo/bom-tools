# -*- coding: utf-8 -*-
"""测试飞书 API 连通性"""
import requests, json

URL = "https://mcenter.huaqin.com/fs/sheet/v1/spreadsheetsMetainfo"
PARAMS = {
    "origin": "cli_a96ac38049f8d0e5",
    "userId": "100448405",
    "spreadsheetToken": "shthq7d9W17DSo7cwuFhtIg7KPf"
}

print("请求中...")
r = requests.get(URL, params=PARAMS, timeout=15)
print("HTTP:", r.status_code)

d = r.json()
code = d.get("code")
sheets = d.get("data", {}).get("sheets", [])
print("code:", code)
print("sheets:", len(sheets))

if code in (0, 200):
    for s in sheets[:3]:
        print("  -", s.get("title", "?"), "rows:", s.get("rowCount", "?"))
else:
    print("错误:", d.get("msg", d))
