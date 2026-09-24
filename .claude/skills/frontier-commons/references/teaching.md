# Teaching & discussion decks

The default deck style (`"style": "teaching"`) is for sessions where people **talk, think and practise**, not just listen: volunteer trainings, workshops, small-group studies, classes and conferences. Slides support the facilitator; they aren't a script.

## Session rhythm

| Phase | Share of time | Slides |
|---|---|---|
| **Open**: connect people to the topic from their own experience | 10% | `title` → `agenda` (with minutes) → a `question` |
| **Teach**: one idea at a time | 30–40% | `section`, then `auto` slides (icons, steps, cards, stats); 1–3 per idea |
| **Discuss**: make sense of it together | 25–30% | `question`, `questions`, `think_pair_share` |
| **Practise**: try it | 20–25% | `scenario`, `activity` |
| **Close**: remember and act | 10% | `recap` with one `next_step` → `closing` |

- **Never more than 3–4 teaching slides in a row** without an interaction.
- Add a `break` for sessions over 75 minutes.
- Plan about 1–2 minutes per teaching slide, plus the time printed on each interaction. Check that the agenda minutes add up to the session length.

## Writing for the room
- **Conversational titles:** "What the first weeks can feel like", not "Phases of International Student Adjustment". Use a question or a plain sentence where it helps.
- **One idea per slide, 3–5 short points.** The builder turns 4 short points into friendly rows, each with an icon.
- **Talk to the audience as "you" and "we".** Keep examples concrete, and use names and places only when they are real or clearly fictional.
- Keep words on the slide short; put the explanation in the **speaker notes**.

## Discussion questions that work
- **Open, not yes/no.** Start with *What, How, Why, When, Where, Who,* or *Tell about a time…*. The copy check flags questions that start with *Is/Do/Can/Would…*. Set `"closed": true` only when a quick show of hands is the point.
- **Start from experience before opinion:** "Think back to a time you were the newcomer…" works better than "What is hospitality?".
- **One question per slide** for the whole room (`question`). Use `questions` (2–4) when groups each pick one.
- **Say how to discuss:** `"format": "Pairs"`, `"time": "4 min"`. Add up to 3 `"prompts"` as follow-ups.
- Be culturally aware. Don't assume a Western experience, and invite international voices in the room to lead.

## Facilitation slides

| Layout | Use it to | Fields (* required) |
|---|---|---|
| `question` | One big open question for everyone | `question`*, `format`, `time`, `group`, `prompts` (≤3), `tag`, `theme` (`periwinkle` default, or `papaya`) |
| `questions` | 2–4 questions: groups choose one, or work through them in order | `questions`*, `title`, `format`, `time`, `lettered` |
| `think_pair_share` | Quiet thinking, then pairs, then the room | `question`*, and optionally `think` / `pair` / `share` (instructions) and `think_time` / `pair_time` / `share_time` |
| `scenario` | A short realistic story, then "What would you do?" with A–D options for a show of hands | `story`* (≤ 450 chars), `question`, `options` (2–4), `tag` |
| `activity` | Hands-on practice | `title`*, `steps`* (≤5 short instructions), `time`, `group`, `materials`, `output` (what groups bring back) |
| `recap` | What we learned plus one personal next step | `points`* (≤4), `next_step`, `next_label`, `title` |
| `break` | Pause | `title`, `time`, `text`, `icon` |

With `auto`: a single point ending in "?" becomes a `question`, several become `questions`. `kind` can be `think_pair_share`, `activity` (points = steps), `recap` (points = learnings), or `scenario` (`story` plus points = options).

## Speaker notes
Write notes for every interaction slide in this pattern:
```
Ask: <the question, word for word>
Give 1 minute of quiet before anyone answers.
Listen for: <the ideas you hope surface>
If it's quiet: <a simpler follow-up question>
Transition: "<the sentence that leads into the next slide>"
```
For teaching slides, one or two sentences of what to **say** and a real example are enough.

## Look and feel (the builder does this)
- A warm Papaya title slide with rounded brand shapes and a topic icon, plus a session chip ("Volunteer training · 60 min").
- Content points sit in rounded white rows with Periwinkle icon discs. The agenda is a run-sheet with minutes.
- Discussion slides use Periwinkle or Papaya grounds with a white speech bubble. Timing and format appear as chips.
- Footers are lighter (page number and icon only). The closing slide invites questions.
- Everything stays inside the brand: only palette colours, Inter + DM Sans, rounded corners, and Amber used sparingly.
