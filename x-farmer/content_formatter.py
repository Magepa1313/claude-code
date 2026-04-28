import random
import textwrap

MAX_TWEET = 280


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 3].rsplit(" ", 1)[0] + "..."


# ── Templates ──────────────────────────────────────────────────────────────────

STORY_INTROS = [
    "STORY TIME 🧵",
    "you cannot make this up.",
    "okay this Reddit story is WILD",
    "I can't stop thinking about this:",
    "thread 🧵",
    "this actually happened:",
    "bro what",
    "context matters. here's the full story:",
]

HOT_TAKE_WRAPPERS = [
    "unpopular opinion but {}",
    "controversial take: {}",
    "nobody wants to say it so I will: {}",
    "hot take and I will not apologize: {}",
    "the reason nobody talks about this is because {}",
    "genuine question — {}",
    "actually {}",
    "{}. I said what I said.",
    "{}. fight me.",
]

POLL_TEMPLATES = [
    ("Who's in the wrong here?", ["The OP", "The other person", "Both", "Neither"]),
    ("Real talk:", ["OP is wrong", "OP is right", "It's complicated", "Tell me in replies"]),
    ("Your verdict?", ["NTA", "YTA", "ESH", "NAH"]),
    ("Be honest:", ["I'd do the same", "I'd never", "Depends", "It's not that deep"]),
]

QUESTION_TEMPLATES = [
    "has anyone else dealt with something like this? genuinely asking",
    "what would YOU do in this situation? replying to everyone",
    "is this normal or is it just me??",
    "be honest. am I overreacting?",
    "at what point do you say enough is enough?",
    "the replies to this are going to be interesting.",
    "disagree with me. I dare you.",
    "drop your hottest take below 👇",
    "this needs more attention than it's getting.",
    "am I the only one who thinks this?",
]


def format_reddit_story(post: dict) -> str:
    title = post.get("title", "").strip()
    body = post.get("selftext", "").strip()
    if body in ("", "[removed]", "[deleted]"):
        body = ""

    intro = random.choice(STORY_INTROS)

    if body:
        # Short story: fit in one tweet
        combined = f"{intro}\n\n{title}\n\n{body}"
        if len(combined) <= MAX_TWEET:
            return combined
        # Title + truncated body
        space = MAX_TWEET - len(intro) - len(title) - 4
        if space > 60:
            return f"{intro}\n\n{title}\n\n{_truncate(body, space)}"
    return _truncate(f"{intro}\n\n{title}", MAX_TWEET)


def format_hot_take(post: dict) -> str:
    title = post.get("title", "").strip()
    # strip "AITA for" / "TIFU by" etc. to make it a statement
    for prefix in ("AITA for ", "AITA: ", "TIFU by ", "TIFU: ", "CMV: ", "CMV: I ", "TIL ", "TIL: "):
        if title.lower().startswith(prefix.lower()):
            title = title[len(prefix):]
            break

    take = title[0].lower() + title[1:] if title else title
    wrapper = random.choice(HOT_TAKE_WRAPPERS)
    tweet = wrapper.format(take)
    return _truncate(tweet, MAX_TWEET)


def format_poll_bait(post: dict) -> str:
    title = _truncate(post.get("title", "").strip(), 140)
    question, options = random.choice(POLL_TEMPLATES)
    lines = [title, "", question]
    for i, opt in enumerate(options):
        lines.append(f"{'🔴🔵🟡🟢'[i]} {opt}")
    return "\n".join(lines)[:MAX_TWEET]


def format_question_bait(post: dict) -> str:
    title = _truncate(post.get("title", "").strip(), 200)
    question = random.choice(QUESTION_TEMPLATES)
    tweet = f"{title}\n\n{question}"
    return _truncate(tweet, MAX_TWEET)


# ── Dispatcher ─────────────────────────────────────────────────────────────────

FORMATTERS = {
    "reddit_story": format_reddit_story,
    "hot_take": format_hot_take,
    "poll_bait": format_poll_bait,
    "question_bait": format_question_bait,
}


def pick_format(weights: dict) -> str:
    keys = list(weights.keys())
    vals = [weights[k] for k in keys]
    return random.choices(keys, weights=vals, k=1)[0]


def make_tweet(post: dict, fmt: str | None = None, weights: dict | None = None) -> str:
    from config import CONTENT_MIX
    if fmt is None:
        fmt = pick_format(weights or CONTENT_MIX)
    return FORMATTERS[fmt](post)
