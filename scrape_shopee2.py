import json
import urllib.request
import ssl
import time
import re

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# Try Shopee item search API with different approach
shops = {
    "Sinar Toko Tiga": {"username": "sinartokotigamesinjahit", "shopid": 1175813},
    "TOKO TIGA": {"username": "tokotigamesinjahit", "shopid": 7999868}
}

all_items = []

for shop_name, info in shops.items():
    shopid = info["shopid"]
    offset = 0
    while offset < 200:
        # Try the search_items_by_shop endpoint
        url = f"https://shopee.co.id/api/v4/shop/search_items_by_shop?shopid={shopid}&limit=36&offset={offset}&sort=sold"
        headers = {
            "User-Agent": "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Mobile Safari/537.36",
            "Referer": "https://shopee.co.id/",
            "Accept": "application/json, application/xml",
            "X-Shopee-Language": "ID",
            "X-API-SOURCE": "rweb",
            "af-ac-enc-dat": "",
            "sz-token": "",
        }
        req = urllib.request.Request(url, headers=headers)
        try:
            resp = urllib.request.urlopen(req, context=ctx, timeout=10)
            data = json.loads(resp.read())
            print(f"{shop_name} offset={offset}: keys={list(data.keys())[:5]}")
            if "data" in data:
                items = data.get("data", {}).get("items", [])
                if not items:
                    items = data.get("data", [])
                print(f"  Got {len(items)} items")
                for item in items[:5]:
                    print(f"  - {item.get('name', item.get('item_basic', {}).get('name', 'N/A'))}")
            else:
                print(f"  Response: {json.dumps(data)[:300]}")
            break
        except Exception as e:
            print(f"Error: {e}")
        
        # Try another endpoint
        url2 = f"https://shopee.co.id/api/v4/recommend/recommend?bundle=category_landing_page&limit=36&offset={offset}&shopid={shopid}"
        req2 = urllib.request.Request(url2, headers=headers)
        try:
            resp2 = urllib.request.urlopen(req2, context=ctx, timeout=10)
            data2 = json.loads(resp2.read())
            print(f"  Rec endpoint keys: {list(data2.keys())[:5]}")
        except Exception as e2:
            print(f"  Rec error: {e2}")
        
        offset += 36
        time.sleep(1)
        break

