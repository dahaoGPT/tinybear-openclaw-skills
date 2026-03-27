import requests
import json
import time
from datetime import datetime, timezone, timedelta

# 强制使用北京时间 (UTC+8)，避免服务器时区差异导致日期判定错误
CST = timezone(timedelta(hours=8))
import sys
import re
import urllib.parse
import argparse

def fetch_details(article_id):
    """
    获取公告详情页 API，提取公告的实际文本内容。
    """
    url = f"https://zfcg.czt.zj.gov.cn/portal/detail?articleId={urllib.parse.quote(article_id)}&parentId=600007"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8"
    }
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.raise_for_status()
        data = res.json()
        content = data.get("result", {}).get("data", {}).get("content", "")
        # 先移除 <style> 和 <script> 标签及其内容
        clean_text = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
        clean_text = re.sub(r'<script[^>]*>.*?</script>', '', clean_text, flags=re.DOTALL | re.IGNORECASE)
        # 移除 display:none 的隐藏 span（网站防爬虫的干扰字符）
        clean_text = re.sub(r'<span[^>]*display\s*:\s*none[^>]*>.*?</span>', '', clean_text, flags=re.DOTALL | re.IGNORECASE)
        # 处理常见 HTML 实体
        clean_text = clean_text.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        # 移除所有剩余 HTML 标签
        clean_text = re.sub(r'<[^>]+>', '', clean_text).strip()
        # 移除零宽字符（防复制干扰）
        clean_text = re.sub(r'[\u200b\u200c\u200d\u200e\u200f\ufeff]', '', clean_text)
        # Compress consecutive whitespace
        clean_text = re.sub(r'\s+', ' ', clean_text)
        # Limit text length to avoid too many tokens
        if len(clean_text) > 2000:
            clean_text = clean_text[:2000] + "..."
        return clean_text
    except Exception as e:
        return f"获取详情失败: {e}"

def scrape_news(target_date_str=None):
    url = "https://zfcg.czt.zj.gov.cn/portal/category"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://zfcg.czt.zj.gov.cn/site/category?parentId=600007&childrenCode=110-963488"
    }

    if target_date_str:
        target_date = datetime.strptime(target_date_str, "%Y-%m-%d").replace(tzinfo=CST)
    else:
        target_date = datetime.now(tz=CST)

    # 获取目标日期零点和当天结束的时间戳（毫秒），强制使用北京时间
    target_start = target_date.replace(hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000
    target_end = target_date.replace(hour=23, minute=59, second=59, microsecond=999999).timestamp() * 1000

    page_no = 1
    page_size = 15
    matched_results = []
    has_more = True

    while has_more:
        payload = {
            "pageNo": page_no,
            "pageSize": page_size,
            "categoryCode": "110-963488",
            "_t": int(time.time() * 1000)
        }

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=15)
            response.raise_for_status()
            data = response.json()
            items = data.get("result", {}).get("data", {}).get("data", [])
            
            if not items:
                break  # 没有更多数据
        
            has_valid_or_newer_data_in_page = False
            
            for item in items:
                publish_date = item.get("publishDate", 0)
                title = item.get("title", "")
                article_id = item.get("articleId", "")
                district_name = item.get("districtName", "")
                
                if publish_date >= target_start:
                    has_valid_or_newer_data_in_page = True
                
                if publish_date > target_end:
                    # 如果条目时间比目标日期晚（更近的日期），直接跳过
                    continue
                elif target_start <= publish_date <= target_end:
                    # 找到了属于目标日期的条目，检查区划字段 (districtName) 中是否包含关键字
                    if "杭州" in district_name or "浙江" in district_name:
                        detail_url = f"https://zfcg.czt.zj.gov.cn/site/detail?parentId=600007&articleId={urllib.parse.quote(article_id)}"
                        detail_text = fetch_details(article_id)
                        
                        # # Deobfuscate title as well
                        # title_decoded = deobfuscate(title)
                        
                        # 为了和官网视觉对齐，组合展示标题
                        # display_title = f"[{district_name}] {title_decoded}" if district_name else title_decoded
                        display_title = f"[{district_name}] {title}" if district_name else title
                        matched_results.append({
                            "title": display_title,
                            "publishDate": datetime.fromtimestamp(publish_date / 1000, tz=CST).strftime('%Y-%m-%d %H:%M:%S'),
                            "url": detail_url,
                            "content_snippet": detail_text
                        })
                else:
                    # 比目标日期早的数据，直接跳过。不在此立刻 break，以防列表有乱序。
                    continue
            
            # 如果整页的数据都比目标时间早（意味着从这一页开始全部都是老数据了），才停止翻页
            if not has_valid_or_newer_data_in_page:
                has_more = False # 因为条目是按 publishDate 降序排列的
            else:
                # 哪怕本页全是比目标时间更晚的数据，也得继续翻页
                page_no += 1
                time.sleep(1.5)  # 避免请求过快遭到服务端防御或 SSL 断开

        except Exception as e:
            print(json.dumps({"error": str(e)}, ensure_ascii=False))
            sys.exit(1)

    # 将结果以 JSON 格式输出到标准输出，供 LLM 解析
    print(json.dumps({
        "total_matches": len(matched_results),
        "results": matched_results,
        "pages_checked": page_no,
        "target_date": target_date.strftime('%Y-%m-%d')
    }, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="抓取特定日期的浙江政府采购网公告。")
    parser.add_argument("--date", type=str, help="目标日期，格式 YYYY-MM-DD", default=None)
    args = parser.parse_args()
    scrape_news(args.date)
