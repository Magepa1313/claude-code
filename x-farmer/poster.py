import tweepy
import os
from config import (
    X_API_KEY, X_API_SECRET,
    X_ACCESS_TOKEN, X_ACCESS_SECRET,
    X_BEARER_TOKEN,
)


def get_client() -> tweepy.Client:
    return tweepy.Client(
        bearer_token=X_BEARER_TOKEN,
        consumer_key=X_API_KEY,
        consumer_secret=X_API_SECRET,
        access_token=X_ACCESS_TOKEN,
        access_token_secret=X_ACCESS_SECRET,
        wait_on_rate_limit=True,
    )


def post_tweet(text: str, dry_run: bool = False) -> dict | None:
    """Post a tweet. Returns response dict or None on failure."""
    if dry_run:
        print(f"[DRY RUN] Would post ({len(text)} chars):\n{text}\n{'─'*60}")
        return {"dry_run": True}

    try:
        client = get_client()
        resp = client.create_tweet(text=text)
        tweet_id = resp.data["id"]
        print(f"[poster] posted tweet {tweet_id}")
        return resp.data
    except tweepy.TweepyException as e:
        print(f"[poster] error: {e}")
        return None


def post_thread(tweets: list[str], dry_run: bool = False) -> list[dict]:
    """Post a thread of tweets, each replying to the previous."""
    results = []
    reply_to = None

    for i, text in enumerate(tweets):
        if dry_run:
            print(f"[DRY RUN] Thread [{i+1}/{len(tweets)}] ({len(text)} chars):\n{text}\n{'─'*60}")
            results.append({"dry_run": True, "index": i})
            continue

        try:
            client = get_client()
            kwargs = {"text": text}
            if reply_to:
                kwargs["in_reply_to_tweet_id"] = reply_to
            resp = client.create_tweet(**kwargs)
            tweet_id = resp.data["id"]
            reply_to = tweet_id
            results.append(resp.data)
            print(f"[poster] thread {i+1}/{len(tweets)} -> {tweet_id}")
        except tweepy.TweepyException as e:
            print(f"[poster] thread error at {i}: {e}")
            break

    return results
