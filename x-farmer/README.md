# X Follower Farmer Bot

Scrapes high-engagement Reddit posts and reformats them as tweets — stories, hot takes, poll bait, and question bait — to maximize engagement and follower growth.

## Setup

```bash
pip install -r requirements.txt
```

Set your X API credentials as environment variables:

```bash
export X_API_KEY=...
export X_API_SECRET=...
export X_ACCESS_TOKEN=...
export X_ACCESS_SECRET=...
export X_BEARER_TOKEN=...
```

Get these from [developer.twitter.com](https://developer.twitter.com) — you need a Free tier app with **Read and Write** permissions.

## Usage

```bash
# Preview what would be posted (no API calls to X)
python scheduler.py --dry-run

# Post one batch now and exit
python scheduler.py --once

# Run continuously (posts during peak hours, sleeps otherwise)
python scheduler.py
```

## How it works

1. **Scrapes** hot/top posts from high-engagement subreddits (AITA, TIFU, unpopularopinion, etc.)
2. **Formats** each post as one of:
   - Reddit story (condensed retelling)
   - Hot take (reframes the post as an opinion)
   - Poll bait ("who's in the wrong?")
   - Question bait (open-ended engagement hook)
3. **Posts** during peak UTC hours with randomised delays to appear organic
4. **Tracks** seen post IDs and daily post count in `state.json`

## Tuning

Edit `config.py` to adjust:
- `SUBREDDITS` — which subreddits to pull from
- `MIN_SCORE` — minimum upvotes to qualify a post
- `MAX_POSTS_PER_DAY` — free tier is ~1500/month, so 45/day is safe
- `CONTENT_MIX` — weights for each content type
- `PEAK_HOURS` — UTC hours to post in
