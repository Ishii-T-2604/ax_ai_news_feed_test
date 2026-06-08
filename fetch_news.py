#!/usr/bin/env python3
import html
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import feedparser
import requests

RSS_FEEDS = [
    {"url": "https://techcrunch.com/category/artificial-intelligence/feed/", "lang": "en", "source": "TechCrunch", "ai_only": True},
    {"url": "https://www.theverge.com/ai-artificial-intelligence/rss/index.xml", "lang": "en", "source": "The Verge", "ai_only": True},
    {"url": "https://venturebeat.com/category/ai/feed/", "lang": "en", "source": "VentureBeat", "ai_only": True},
    {"url": "https://www.wired.com/feed/category/artificial-intelligence/latest/rss", "lang": "en", "source": "Wired", "ai_only": True},
    {"url": "https://rss.itmedia.co.jp/rss/2.0/aiplus.xml", "lang": "ja", "source": "ITmedia AI+", "ai_only": True},
    {"url": "https://gigazine.net/news/rss_2.0/", "lang": "ja", "source": "Gigazine", "ai_only": False},
    {"url": "https://ascii.jp/rss.xml", "lang": "ja", "source": "ASCII.jp", "ai_only": False},
]

AI_KEYWORDS = [
    "AI", "人工知能", "機械学習", "生成AI", "LLM", "ChatGPT", "Claude", "Gemini",
    "OpenAI", "Anthropic", "DeepMind", "深層学習", "ニューラル", "チャットbot",
    "artificial intelligence", "machine learning", "deep learning", "neural",
    "language model", "transformer", "AGI", "generative", "GPT",
    "Mistral", "Llama", "Meta AI",
]

SOURCE_FLAGS = {
    "TechCrunch": "🇺🇸", "The Verge": "🇺🇸", "VentureBeat": "🇺🇸", "Wired": "🇺🇸",
    "ITmedia AI+": "🇯🇵", "Gigazine": "🇯🇵", "ASCII.jp": "🇯🇵",
}


def clean_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def truncate(text: str, max_len: int = 150) -> str:
    return text if len(text) <= max_len else text[:max_len].rstrip() + "…"


def is_ai_related(entry) -> bool:
    text = (entry.get("title", "") + " " + entry.get("summary", "")).lower()
    return any(kw.lower() in text for kw in AI_KEYWORDS)


def get_date(entry) -> datetime | None:
    for attr in ("published_parsed", "updated_parsed"):
        t = getattr(entry, attr, None)
        if t:
            return datetime(*t[:6], tzinfo=timezone.utc)
    return None


def fetch_articles(hours: int = 24) -> dict[str, list]:
    cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)
    articles: dict[str, list] = {"en": [], "ja": []}

    for feed_info in RSS_FEEDS:
        try:
            feed = feedparser.parse(feed_info["url"])
            for entry in feed.entries[:20]:
                date = get_date(entry)
                if date and date < cutoff:
                    continue
                if not feed_info["ai_only"] and not is_ai_related(entry):
                    continue
                summary = clean_html(entry.get("summary", entry.get("description", "")))
                articles[feed_info["lang"]].append({
                    "title": clean_html(entry.get("title", "(タイトルなし)")),
                    "url": entry.get("link", ""),
                    "summary": truncate(summary),
                    "source": feed_info["source"],
                    "date": date,
                })
        except Exception as e:
            print(f"[WARN] {feed_info['source']}: {e}", file=sys.stderr)

    return articles


def select_articles(articles: dict, en_count: int = 2, ja_count: int = 3) -> list:
    seen: set[str] = set()

    def pick(pool: list, count: int) -> list:
        sorted_pool = sorted(
            pool,
            key=lambda x: x["date"] or datetime.min.replace(tzinfo=timezone.utc),
            reverse=True,
        )
        result = []
        for a in sorted_pool:
            if len(result) >= count:
                break
            key = a["title"].lower()
            if key not in seen:
                seen.add(key)
                result.append(a)
        return result

    return pick(articles["en"], en_count) + pick(articles["ja"], ja_count)


def build_card(articles: list) -> dict:
    today = datetime.now().strftime("%Y年%m月%d日")
    en_count = sum(1 for a in articles if SOURCE_FLAGS.get(a["source"]) == "🇺🇸")
    ja_count = sum(1 for a in articles if SOURCE_FLAGS.get(a["source"]) == "🇯🇵")

    body = [
        {
            "type": "TextBlock",
            "text": f"🤖 今日のAIニュース — {today}",
            "weight": "Bolder",
            "size": "Large",
            "wrap": True,
            "color": "Accent",
        },
        {
            "type": "TextBlock",
            "text": f"🇺🇸 英語 {en_count}件  🇯🇵 日本語 {ja_count}件",
            "size": "Small",
            "color": "Good",
            "spacing": "None",
        },
    ]

    for i, article in enumerate(articles):
        flag = SOURCE_FLAGS.get(article["source"], "🌐")
        body.append({
            "type": "Container",
            "separator": i > 0,
            "spacing": "Medium" if i > 0 else "None",
            "items": [
                {
                    "type": "TextBlock",
                    "text": f"{flag} **[{article['title']}]({article['url']})**",
                    "wrap": True,
                    "size": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": article["summary"] or "(サマリーなし)",
                    "wrap": True,
                    "size": "Small",
                    "isSubtle": True,
                    "spacing": "Small",
                },
                {
                    "type": "TextBlock",
                    "text": f"📰 {article['source']}",
                    "size": "Small",
                    "color": "Accent",
                    "spacing": "Small",
                },
            ],
        })

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": body,
        "msteams": {"width": "Full"},
    }


def post_to_teams(webhook_url: str, card: dict) -> None:
    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": None,
                "content": card,
            }
        ],
    }
    resp = requests.post(
        webhook_url,
        headers={"Content-Type": "application/json"},
        data=json.dumps(payload),
        timeout=30,
    )
    resp.raise_for_status()


def main() -> None:
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        sys.exit("Error: TEAMS_WEBHOOK_URL environment variable is not set")

    print("Fetching articles...")
    articles = fetch_articles(hours=24)

    # 24時間で足りなければ48時間に拡張
    if len(articles["en"]) < 2 or len(articles["ja"]) < 3:
        print("Not enough articles in 24h, extending to 48h...")
        articles = fetch_articles(hours=48)

    selected = select_articles(articles, en_count=2, ja_count=3)
    if not selected:
        print("No articles found, skipping.")
        return

    for a in selected:
        print(f"  [{a['source']}] {a['title']}")

    card = build_card(selected)
    post_to_teams(webhook_url, card)
    print("Posted to Teams successfully!")


if __name__ == "__main__":
    main()
