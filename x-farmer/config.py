import os

# X API credentials - set these as environment variables
X_API_KEY = os.environ.get("X_API_KEY", "")
X_API_SECRET = os.environ.get("X_API_SECRET", "")
X_ACCESS_TOKEN = os.environ.get("X_ACCESS_TOKEN", "")
X_ACCESS_SECRET = os.environ.get("X_ACCESS_SECRET", "")
X_BEARER_TOKEN = os.environ.get("X_BEARER_TOKEN", "")

# Free tier: 1500 posts/month = ~48/day. Stay under to be safe.
MAX_POSTS_PER_DAY = 45
# Spread posts across peak hours (UTC)
PEAK_HOURS = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22]  # 1pm-10pm UTC

# Subreddits to pull from - mix of storytelling + controversy
SUBREDDITS = [
    # AITA / drama / stories
    "AmItheAsshole",
    "tifu",
    "relationship_advice",
    "entitledparents",
    "MaliciousCompliance",
    "pettyrevenge",
    "ProRevenge",
    "raisedbynarcissists",
    # hot takes / debate bait
    "unpopularopinion",
    "changemyview",
    "TrueOffMyChest",
    # viral content
    "interestingasfuck",
    "Damnthatsinteresting",
    "todayilearned",
    "worldnews",
    "news",
]

# Min upvotes to qualify a post
MIN_SCORE = 5000

# Content mix weights (must sum to 1.0)
CONTENT_MIX = {
    "reddit_story": 0.40,    # condensed Reddit story
    "hot_take": 0.30,        # spicy take on a Reddit post
    "poll_bait": 0.15,       # "which side are you on" poll
    "question_bait": 0.15,   # engagement question
}
