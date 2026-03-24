# ALTT Weekly Agenda Tool

**Live:** [hak29.github.io/altt-agenda](https://hak29.github.io/altt-agenda/)

Weekly check-in agenda for the AI Leadership Think Tank (ALTT) at MMA Global.

## Structure

| Tab | Focus | ~Time |
|-----|-------|-------|
| **Labs** | CAP, A3, AURA, SIFT, ACE, ARC experiments | 25 min |
| **Knowledge Center** | DAMit, MAC, MAF, Research, Guides | 8 min |
| **Community** | ALC, CATS, Great Debates, Slack, Awards | 7 min |
| **Action Items** | Hassan's items, team items, live notes | 5 min |

Each item follows Greg Stuart's template:
- **Accomplished** since last meeting
- **What's next**
- **Help needed** / blockers

## Weekly Workflow

### Before the Call (Monday)

Tell Claude: **"Update my ALTT agenda for this week"**

Claude will:
1. Pull latest from Pipeline MCP, Fireflies transcripts, and Outlook
2. Update the HTML with fresh data for each brand/experiment
3. Push to GitHub — live in ~1 minute

### During the Call

1. Open [hak29.github.io/altt-agenda](https://hak29.github.io/altt-agenda/) on phone or laptop
2. Use **tabs** to navigate between Labs / Knowledge / Community / Actions
3. Click section headers to expand/collapse
4. Type notes in any text field — **auto-saves every 600ms**
5. Use the **Action Items > Live Notes** textarea for quick capture

### After the Call

- Click **Export** to download a Markdown file of the full agenda + your notes
- Tell Claude: **"Summarize this week's ALTT meeting and update the pipeline"**

## Features

- **Week navigation**: Arrow buttons to browse past/future weeks
- **"Today" button**: Jump back to current week
- **Auto-save**: Notes persist in browser localStorage (per week, per device)
- **Read-only past weeks**: Navigating to old weeks shows notes but prevents edits
- **Attention alerts**: Red banner on Labs tab for flagged experiments
- **Source citations**: Color-coded (green=call, amber=email, red=action)
- **Mobile responsive**: Works on phone for on-the-go review

## Data Sources

| Source | What it provides |
|--------|-----------------|
| Pipeline MCP | Experiment stages, deadlines, brand context, attention flags |
| Fireflies | Meeting transcripts, action items, key decisions |
| Outlook | Email communications, calendar events |

## Notes on localStorage

- Notes are stored **per device** (phone notes =/= laptop notes)
- Each week gets its own storage key (`altt-YYYY-MM-DD`)
- Clearing browser data deletes notes — always **Export** after meetings
- Static content (experiment details, statuses) is baked into the HTML by Claude

## Tech

Single HTML file, zero dependencies, GitHub Pages hosted. Navy blue + MMA Gold (#FFA400) theme.
