# Teaching & discussion decks

Frontier Commons sessions (for example the **Fall Fellows "Plug In"** weeks) are taught and discussed, not read. There are three deck styles:

| `"style"` | Looks like | Use for |
|---|---|---|
| `teaching` (default) | Welcoming and illustrated, built on the Plug In structure. Section openers and pauses are Papaya with an illustration cluster (Periwinkle disc, Deep Blue tile with a line icon, Amber dot, dot texture). Bullet slides have a Papaya side panel with the title and a big Amber-disc illustration. Pillars have a spot illustration above each label. Reflection questions sit beside a speech-bubble panel. Definitions and quotes have illustrations too, and pauses use calm concentric "breathe" rings. Spark bullets, uppercase labels and hairlines stay. The title (default `bold`) is Deep Blue with a Periwinkle dot texture and "PLUG IN" in Amber | Fellows sessions, trainings, studies, anything taught and discussed |
| `workshop` | Warm and playful: a Papaya title with rounded shapes, icon rows, speech bubbles and time chips | Hands-on workshops, youth or volunteer events, very interactive sessions |
| `formal` | Brand-guide standard: Deep Blue title, stats, charts, footers | Partner or board briefings, reports, pitches |

`examples/decks/teaching-template.json` is a **blank session** in this shape, with every text a `[placeholder]`. **Copy it for any teaching deck**, replace every bracket with the real content, and delete the slides you don't need. Never leave a `[bracket]` in a finished deck unless you tell the user it still needs their input.

## The session shape (from the Plug In weeks)

| # | Slide | Layout | Notes |
|---|---|---|---|
| 1 | Program title | `title` | `kicker` (e.g. "Frontier Commons"), `title` (program name), `series` (Amber italic line), `info` ("Week N · Topic · Date"), `org` |
| 2 | Recap of last week | `recap` | 3–5 points, drawn out of the room before they're shown |
| 3 | Pause | `break` | "Let's take a\npause". Quiet, prayer or settling in |
| 4 | What it is | auto → `definition` | One sentence, big and centred |
| 5 | Where it shows up | auto → `pillars` | 2–4 labelled columns, e.g. three books or people where the topic appears. A `ref` gives a scripture reference and `note` a footnote |
| 6 | What it looks like | auto → `content` | 3–5 spark-bulleted full sentences |
| 7 | Reflection questions | auto → `questions` | 2–4 open questions, numbered, with hairlines between |
| 8 | Why it matters | auto → `pillars` | 3–4 reasons, each a LABEL plus one or two sentences |
| 9 | Part two opener | `section` | Two short lines; the last line is highlighted automatically ("[Part two\ntitle]") |
| 10–11 | Application | auto → `content` | applying the idea to everyday life |
| 12 | Key quote + question | `quote` | Big centred quote; `question` goes under a hairline in italics |
| 13 | Part three opener | `section` | "[Part three\ntitle]" (mark a different highlighted word with `**…**` if you like) |
| 14–17 | Building | `content` / `pillars` | Points that turn the topic into action ("your [topic] is your edge / your why"). `lead` adds a line under the title ("As Christ-followers, we should ask ourselves:") |
| 18–19 | Tools | `tool` | Empathy map or persona: `rows` of Purpose / Goal / Keep in mind plus an example `image` |
| 20 | Deliverables | `content` | What participants produce this week |

Adapt the shape to the topic, but keep its rhythm: **open (recap, pause), teach (definition, pillars, points), reflect (questions), apply (sections, quote), build (tools), commit (deliverables)**. Never go more than 3–4 teaching slides without a reflection, quote or question.

## Illustrations
- **Title slide options:** set `"title_variant"` on the deck (or `"variant"` on the title slide) to `bold` (default: Deep Blue, capitals, dot texture), `illustrated` (Papaya, sentence case, info pills, brand shapes with the topic icon) or `classic` (skewed band, dots, optional photo). Add `"icon"` to the bold title to show a topic icon on the right.
- Every illustration is built from brand shapes plus one line icon picked from the slide's title and points. Set `"icon": "<name>"` on the slide (or on a pillar) when the automatic pick misses. Watch for words with two meanings, like *Job* (the book): use `book-heart`, not a briefcase.
- The deck-level `"kicker"` (e.g. "Week 3 · Hospitality") appears above every bullet-slide title.
- Every bullet-slide illustration is the same size (1.4in), so the deck feels consistent. Keep panel titles under about 90 characters so there is room for it.

## Writing like the Plug In decks
- **Titles are the idea in plain words:** applying the idea to everyday life, "Why this matters for your idea". Use sentence case, and let titles run to two lines when needed.
- **Points are complete, warm sentences** (usually 60–140 characters), spoken to "you" and "we". Three or four per slide.
- **Pillar labels are 1–3 words** ("Healthy, not weak", "An act of trust") with one or two sentences under each.
- **Section openers** are 2–5 words over two lines. The builder highlights the last line; mark a different word with `**…**`.
- **Scripture:** give the reference (`"ref": "Revelation 21:4"`) and quote only text the user supplied. Never write verse text from memory. If the user hasn't provided it, add `"note": "[Scripture references and translation to be added]"` and flag it in the hand-over.
- **Faith and innovation go together.** Tie design-thinking ideas (empathy, users, personas, "problems to fix" → "people to serve") back to the week's spiritual theme.

## Discussion questions that work
- **Open, not yes/no.** Start with *What, How, Why, When, Where, Who,* or *Tell about a time…*. The copy check flags *Is/Do/Can/Would…* questions; set `"closed": true` when a closed question is intended (Kingdom-impact style self-check questions are fine as spark bullets).
- **Start from experience before opinion:** "What is something that grieves or deeply saddens your heart?"
- Keep reflection lists to 2–4 questions. Use `question` for one question to the whole room, with `format` / `time` / `prompts`.
- Be culturally aware: don't assume a Western experience, and invite international voices to lead.

## Facilitation slides (all styles)

| Layout | Use it to | Fields (* required) |
|---|---|---|
| `question` | One open question for the room | `question`*, `format`, `time`, `prompts` (≤3), `tag` |
| `questions` | Reflection questions (numbered list in teaching style) | `questions`*, `title`, `format`, `time` |
| `think_pair_share` | Think, then pairs, then the room | `question`*, `think`/`pair`/`share`, `think_time`/`pair_time`/`share_time` |
| `scenario` | Story, then "What would you do?" with A–D options | `story`*, `question`, `options` |
| `activity` | Hands-on practice | `title`*, `steps`* (≤5), `time`, `group`, `materials`, `output` |
| `recap` | Last week's or today's learnings, plus one next step | `points`* (≤5), `next_step`, `title` |
| `break` | Pause | `title` (default "Let's take a\npause"), `time` |
| `pillars` | 2–4 labelled ideas | `items`* `[{label, text, ref, quote}]`, `note`, `lead` |
| `definition` | Title plus one big sentence | `title`*, `text`* |
| `tool` | A method or template with an example | `title`*, `rows` `[{label, text}]` (or `purpose` / `goal` / `keep_in_mind`), `image`, `image_caption` |

With `auto`:
- a single point under a title becomes a `definition`
- 2–4 titled points (no subtitles) become `pillars`
- points ending in "?" become `questions`
- plain sentences become spark bullets

## Speaker notes
Write notes for every reflection or discussion slide:
```
Ask: <the question, word for word>
Give 1 minute of quiet before anyone answers.
Listen for: <the ideas you hope surface>
If it's quiet: <a simpler follow-up>
Transition: "<the sentence that leads into the next slide>"
```
