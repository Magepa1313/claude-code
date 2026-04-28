"""
Main scheduler loop.

Usage:
    python scheduler.py              # live mode
    python scheduler.py --dry-run    # preview posts without sending
    python scheduler.py --once       # post one batch now and exit
"""

import argparse
import json
import os
import random
import time
from datetime import datetime, timezone
from pathlib import Path

from config import MAX_POSTS_PER_DAY, PEAK_HOURS, CONTENT_MIX
from reddit_scraper import collect_posts
from content_formatter import make_tweet
from poster import post_tweet

STATE_FILE = Path("state.json")


def load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"seen_ids": [], "posts_today": 0, "last_reset": ""}


def save_state(state: dict):
    STATE_FILE.write_text(json.dumps(state, indent=2))


def reset_daily_counter(state: dict) -> dict:
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    if state.get("last_reset") != today:
        state["posts_today"] = 0
        state["last_reset"] = today
    return state


def seconds_until_next_peak() -> int:
    now = datetime.now(timezone.utc)
    current_hour = now.hour
    for h in sorted(PEAK_HOURS):
        if h > current_hour:
            target = now.replace(hour=h, minute=random.randint(0, 59), second=random.randint(0, 59))
            return max(0, int((target - now).total_seconds()))
    # Next peak is tomorrow
    tomorrow = now.replace(hour=PEAK_HOURS[0], minute=random.randint(0, 59), second=0)
    from datetime import timedelta
    tomorrow += timedelta(days=1)
    return int((tomorrow - now).total_seconds())


def run_batch(state: dict, dry_run: bool = False) -> dict:
    remaining = MAX_POSTS_PER_DAY - state["posts_today"]
    if remaining <= 0:
        print("[scheduler] daily limit reached, skipping batch")
        return state

    # Collect fresh posts
    seen = set(state.get("seen_ids", []))
    posts = collect_posts(seen, count=min(remaining, 10))

    if not posts:
        print("[scheduler] no fresh posts found")
        return state

    print(f"[scheduler] {len(posts)} posts fetched, {remaining} slots remaining today")

    for post in posts:
        if state["posts_today"] >= MAX_POSTS_PER_DAY:
            break

        tweet = make_tweet(post)
        if not tweet.strip():
            continue

        result = post_tweet(tweet, dry_run=dry_run)
        if result:
            state["posts_today"] += 1
            state.setdefault("seen_ids", []).append(post["id"])
            # keep seen_ids from growing unbounded
            state["seen_ids"] = state["seen_ids"][-2000:]

        if not dry_run:
            # randomise delay between posts: 3-8 minutes
            delay = random.randint(180, 480)
            print(f"[scheduler] sleeping {delay}s before next post")
            time.sleep(delay)

    save_state(state)
    return state


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Preview posts without sending")
    parser.add_argument("--once", action="store_true", help="Run one batch then exit")
    args = parser.parse_args()

    state = load_state()
    state = reset_daily_counter(state)

    if args.once or args.dry_run:
        state = run_batch(state, dry_run=args.dry_run)
        return

    print("[scheduler] starting continuous mode")
    while True:
        state = reset_daily_counter(state)

        now_hour = datetime.now(timezone.utc).hour
        if now_hour in PEAK_HOURS:
            state = run_batch(state)
        else:
            wait = seconds_until_next_peak()
            print(f"[scheduler] off-peak, sleeping {wait//60}m until next window")
            time.sleep(wait)


if __name__ == "__main__":
    main()
