#!/usr/bin/env python3
import html
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import feedparser
import requests

FEED_URL = "https://rss.techplay.jp/event/w3c-rss-format/rss.xml"

AI_KEYWORDS = [
    "AI", "人工知能", "機械学習", "生成AI", "LLM", "ChatGPT", "Claude", "Gemini",
    "OpenAI", "Anthropic", "DeepMind", "深層学習", "ニューラル", "チャットbot",
    "RAG", "プロンプト", "ファインチューニング", "Copilot", "GPT", "MCP",
    "自然言語処理", "NLP", "エージェント", "データサイエンス",
]


def clean_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def is_ai_related(entry) -> bool:
    category_terms = [t.get("term", "") for t in entry.get("tags", [])]
    text = " ".join(category_terms) + " " + entry.get("title", "")
    return any(kw.lower() in text.lower() for kw in AI_KEYWORDS)


def parse_event_start(entry) -> datetime | None:
    raw = entry.get("tp_eventstarttime", "")
    if raw:
        try:
            return datetime.strptime(raw, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    raw = entry.get("tp_eventdate", "")
    if raw:
        try:
            return datetime.strptime(raw, "%Y-%m-%d")
        except ValueError:
            pass
    return None


def format_datetime(entry) -> str:
    start_raw = entry.get("tp_eventstarttime", "")
    end_raw = entry.get("tp_eventendtime", "")
    try:
        start = datetime.strptime(start_raw, "%Y-%m-%d %H:%M:%S")
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        result = start.strftime(f"%Y/%m/%d ({weekdays[start.weekday()]}) %H:%M")
        if end_raw:
            end = datetime.strptime(end_raw, "%Y-%m-%d %H:%M:%S")
            result += f" 〜 {end.strftime('%H:%M')}"
        return result
    except ValueError:
        return entry.get("tp_eventdate", "日時未定")


def fetch_ai_events(max_events: int = 5, within_days: int = 30) -> list:
    feed = feedparser.parse(FEED_URL)
    now = datetime.now()
    cutoff = now + timedelta(days=within_days)

    candidates = []
    for entry in feed.entries:
        if not is_ai_related(entry):
            continue

        start_dt = parse_event_start(entry)
        if start_dt and (start_dt < now or start_dt > cutoff):
            continue

        category_terms = [t.get("term", "") for t in entry.get("tags", [])]
        ai_cats = [c for c in category_terms if any(k.lower() in c.lower() for k in AI_KEYWORDS)]

        candidates.append({
            "title": entry.get("title", "(タイトルなし)"),
            "url": entry.get("tp_eventapplicationurl") or entry.get("link", ""),
            "datetime_str": format_datetime(entry),
            "place": entry.get("tp_eventplace", "") or "会場未定",
            "categories": ai_cats[:3],
            "start_dt": start_dt or now,
        })

    candidates.sort(key=lambda x: x["start_dt"])
    return candidates[:max_events]


def build_card(events: list) -> dict:
    today = datetime.now().strftime("%Y年%m月%d日")

    body = [
        {
            "type": "TextBlock",
            "text": f"🎤 TECH PLAY AI関連イベント — {today}",
            "weight": "Bolder",
            "size": "Large",
            "wrap": True,
            "color": "Accent",
        },
        {
            "type": "TextBlock",
            "text": f"今後30日以内のAI関連イベント {len(events)}件",
            "size": "Small",
            "color": "Good",
            "spacing": "None",
        },
    ]

    for i, event in enumerate(events):
        cats_text = "  ".join(f"#{c}" for c in event["categories"]) if event["categories"] else ""
        body.append({
            "type": "Container",
            "separator": i > 0,
            "spacing": "Medium" if i > 0 else "None",
            "items": [
                {
                    "type": "TextBlock",
                    "text": f"**[{event['title']}]({event['url']})**",
                    "wrap": True,
                    "size": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": f"📅 {event['datetime_str']}　📍 {event['place']}",
                    "size": "Small",
                    "isSubtle": True,
                    "spacing": "Small",
                    "wrap": True,
                },
                *([{
                    "type": "TextBlock",
                    "text": cats_text,
                    "size": "Small",
                    "color": "Accent",
                    "spacing": "None",
                }] if cats_text else []),
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

    print("Fetching AI events from TECH PLAY...")
    events = fetch_ai_events(max_events=5, within_days=30)

    if not events:
        print("No AI-related events found within 30 days, skipping.")
        return

    for e in events:
        print(f"  [{e['datetime_str']}] {e['title']}")

    card = build_card(events)
    post_to_teams(webhook_url, card)
    print("Posted to Teams successfully!")


if __name__ == "__main__":
    main()
