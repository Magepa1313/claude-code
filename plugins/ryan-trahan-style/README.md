# Ryan Trahan Style Generator Plugin

A deep-research, production-grade skill for generating YouTube scripts, hooks, outlines, and video concepts in the **exact style of Ryan Trahan** — one of YouTube's most studied storytellers with 22+ million subscribers.

## What This Plugin Does

When you say **"in the style of Ryan Trahan"** or ask Claude to write content inspired by his approach, this skill activates and gives Claude:

- Ryan's **6-step psychological hook formula** (analyzed by George Blackman)
- His **three-act narrative structure** for challenge vlogs
- His **voice patterns**, verbal tics, and sentence rhythm
- The **"Kevin" character technique** for making strangers into recurring characters
- Complete **video format playbooks** for all his major content types
- His **charity integration strategy**
- Editing cues and production notes to write toward

## Trigger Phrases

This skill auto-activates when you use phrases like:
- "in the style of Ryan Trahan"
- "Ryan Trahan style"
- "like Ryan Trahan would write it"
- "write a Ryan Trahan hook for..."
- "create a Ryan Trahan script for..."
- "outline a Ryan Trahan-style challenge video about..."

## Content Types Supported

| Type | Description |
|------|-------------|
| **Hook** | 60-second opener using the 6-step formula |
| **Full Script** | Complete 10–20 minute video script with editing cues |
| **Video Outline** | Three-act structure with chapter markers |
| **Video Concept** | Title + premise + format recommendation |
| **Narration Passage** | A specific scene in Ryan's voice |
| **Challenge Design** | Full challenge rules, stakes, and twist mechanic |

## Skill Files

```
skills/ryan-trahan-style/
├── SKILL.md                          # Core skill — Ryan's complete style guide
└── references/
    ├── hook-formula-deep-dive.md     # Deep dive on the 6-step hook formula
    ├── voice-and-script-patterns.md  # Language patterns, verbal tics, pacing
    └── video-format-playbooks.md     # Full playbooks for each major format
```

## Example Usage

```
User: Write a hook in the style of Ryan Trahan for a video about
      surviving 30 days using only items found at dollar stores.

Claude: [Activates ryan-trahan-style skill]
        [Produces 6-step hook with his voice, proof shot, and Grand Payoff tease]
```

```
User: Give me a full video outline in Ryan Trahan style for
      "I stayed in every Airbnb tier from $10/night to $10,000/night"

Claude: [Produces three-act structure, tier progression format,
        character moments, Grand Payoff setup, charity integration option]
```

## What Makes This Skill Deep

This skill is built from:
- Analysis of Ryan's video structure by scriptwriting experts (George Blackman's breakdown)
- The penny series storytelling anatomy (30 daily episodes, $1.4M raised)
- His "100 Days in GTA" three-act narrative deconstruction
- His speech pattern analysis (Texas drawl, comedic timing, self-deprecation formula)
- His editing philosophy ("low-effort" aesthetic that takes weeks to achieve)
- The "Kevin technique" for character building with strangers
- His charity integration approach across multiple series

## Installation

This plugin is included in the Claude Code plugins directory. No additional setup required.

To use in your own project, install Claude Code and reference this plugin via `.claude/settings.json`.
