import json
import urllib.request
import ssl
import time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

shops = {
    "Sinar Toko Tiga": 1175813,
    "TOKO TIGA": 7999868
}

all_items = []

for shop_name, shopid in shops.items():
    for page_type in ["shop_page"]:
        # Try the pcweb endpoint
        url = f"https://shopee.co.id/pc/search_items_by_shop?shopid={shopid}&limit=50&offset=0&sort=sold"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Referer": f"https://shopee.co.id/shop/{shopid}",
            "Accept": "*/*",
        }
        req = urllib.request.Request(url, headers=headers)
        try:
            resp = urllib.request.urlopen(req, context=ctx, timeout=10)
            print(f"{shop_name}: {resp.status}")
            data = json.loads(resp.read())
            print(json.dumps(data)[:500])
        except Exception as e:
            print(f"Error: {e}")
        
        # Try the v4/item endpoint to get individual items
        for itemid in [733103516, 745352035]:
            url = f"https://shopee.co.id/api/v4/item/get?shopid={shopid}&itemid={itemid}"
            req = urllib.request.Request(url, headers=headers)
            try:
                resp = urllib.request.urlopen(req, context=ctx, timeout=10)
                data = json.loads(resp.read())
                item = data.get("data", {})
                if item:
                    all_items.append({
                        "shop": shop_name,
                        "platform": "Shopee",
                        "name": item.get("name", ""),
                        "price": item.get("price", 0) / 100000,
                        "price_min": item.get("price_min", 0) / 100000,
                        "price_max": item.get("price_max", 0) / 100000,
                        "sold": item.get("sold", item.get("historical_sold", 0)),
                        "description": (item.get("description", "") or "")[:500],
                        "brand": item.get("brand", ""),
                        "itemid": itemid,
                        "shopid": shopid,
                    })
                    print(f"  Got item: {item.get('name', 'N/A')} - sold: {item.get('sold', item.get('historical_sold', 'N/A'))}")
                else:
                    print(f"  No data for itemid {itemid}: {json.dumps(data)[:200]}")
            except Exception as e:
                print(f"  Error for itemid {itemid}: {e}")
            time.sleep(0.3)
    break

with open("/root/.openclaw/workspace/ecommerce-research/shopee_raw.json", "w") as f:
    json.dump(all_items, f, indent=2, ensure_ascii=False)
print(f"\nSaved {len(all_items)} items")
