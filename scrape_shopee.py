import json
import urllib.request
import urllib.parse
import time
import ssl

# Disable SSL verification for scraping
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

shops = {
    "Sinar Toko Tiga": 1175813,
    "TOKO TIGA": 7999868
}

all_items = []

for shop_name, shopid in shops.items():
    offset = 0
    limit = 36
    while offset < 200:
        url = f"https://shopee.co.id/api/v4/shop/search_items?shopid={shopid}&limit={limit}&offset={offset}&sort_by=sales"
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Referer": f"https://shopee.co.id/shop/{shopid}",
            "Accept": "application/json",
            "X-Shopee-Language": "ID",
            "X-API-SOURCE": "pc",
        }
        req = urllib.request.Request(url, headers=headers)
        try:
            resp = urllib.request.urlopen(req, context=ctx, timeout=10)
            data = json.loads(resp.read())
            if "items" in data and data["items"]:
                for item_wrapper in data["items"]:
                    item = item_wrapper.get("item_basic", item_wrapper)
                    all_items.append({
                        "shop": shop_name,
                        "platform": "Shopee",
                        "name": item.get("name", ""),
                        "price_min": item.get("price_min", 0) / 100000,
                        "price_max": item.get("price_max", 0) / 100000,
                        "sold": item.get("sold", item.get("historical_sold", 0)),
                        "itemid": item.get("itemid", ""),
                        "description": item.get("description", "")[:200],
                        "categories": json.dumps(item.get("categories", [])),
                        "brand": item.get("brand", ""),
                    })
                if len(data["items"]) < limit:
                    break
                offset += limit
                time.sleep(0.5)
            else:
                print(f"Error for {shop_name} offset={offset}: {json.dumps(data)[:200]}")
                break
        except Exception as e:
            print(f"Error for {shop_name} offset={offset}: {e}")
            break

print(f"Total items scraped: {len(all_items)}")
# Save raw data
with open("/root/.openclaw/workspace/ecommerce-research/shopee_raw.json", "w") as f:
    json.dump(all_items, f, indent=2, ensure_ascii=False)
print("Saved to shopee_raw.json")
