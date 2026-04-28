import requests
import random
import time
from config import SUBREDDITS, MIN_SCORE

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; content-bot/1.0)"}


def fetch_hot_posts(subreddit: str, limit: int = 25) -> list[dict]:
    url = f"https://www.reddit.com/r/{subreddit}/hot.json?limit={limit}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()
        posts = r.json()["data"]["children"]
        return [p["data"] for p in posts if not p["data"].get("stickied")]
    except Exception as e:
        print(f"[scraper] failed {subreddit}: {e}")
        return []


def fetch_top_posts(subreddit: str, timeframe: str = "day", limit: int = 25) -> list[dict]:
    url = f"https://www.reddit.com/r/{subreddit}/top.json?t={timeframe}&limit={limit}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()
        posts = r.json()["data"]["children"]
        return [p["data"] for p in posts if not p["data"].get("stickied")]
    except Exception as e:
        print(f"[scraper] failed {subreddit}: {e}")
        return []


def get_post_body(post: dict) -> str:
    """Get the text body of a self post, truncated."""
    text = post.get("selftext", "").strip()
    if text in ("", "[removed]", "[deleted]"):
        return ""
    return text[:1500]


def collect_posts(seen_ids: set, count: int = 20) -> list[dict]:
    """Pull fresh posts from a random mix of subreddits."""
    all_posts = []
    subs = random.sample(SUBREDDITS, min(len(SUBREDDITS), 8))

    for sub in subs:
        posts = fetch_hot_posts(sub, 25) + fetch_top_posts(sub, "day", 10)
        for p in posts:
            if (
                p["id"] not in seen_ids
                and p.get("score", 0) >= MIN_SCORE
                and p.get("title")
            ):
                all_posts.append(p)
        time.sleep(0.5)  # be polite to Reddit

    # dedupe and sort by score
    seen = set()
    unique = []
    for p in all_posts:
        if p["id"] not in seen:
            seen.add(p["id"])
            unique.append(p)

    unique.sort(key=lambda x: x["score"], reverse=True)
    return unique[:count]
