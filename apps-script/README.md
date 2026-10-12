# Player ratings backend

The site is static, so anonymous voting needs a tiny backend. This is a Google
Apps Script web app attached to a **private** Google Sheet (not the attendance
sheet). Setup takes about ten minutes, once.

## One-time setup

1. Create a new Google Sheet called something like `W9-10 Ratings` (private, just you).
2. **Extensions → Apps Script**. Delete the default code, paste in `Code.gs`, save.
3. In the editor, pick the `setup` function and press **Run**. Approve the permissions.
   This creates the tabs `Voters`, `Votes`, `Windows` and `Seeds`, and opens the first
   quarterly window (e.g. `2026-Q4`).
4. **Deploy → New deployment → Web app**. Execute as **Me**, who has access **Anyone**.
   Copy the web-app URL (ends in `/exec`).
5. Put that URL into `index.html` as `VOTE_API` and push. The **Player Cards** tab goes live.

## Seeding indicative stats

Open the `Seeds` tab. Every player is listed with 65 across the board. Put in your
indicative numbers (40–99) for PAC / SHO / PAS / DRI / DEF / PHY. These are only the
starting point voters see; they never show on the public cards.

## Sending the links

1. Run `generateLinks`. A `Links` tab appears with one personal link per player.
2. Send each person their own link (WhatsApp). Tell them not to forward it — whoever
   has the link votes as them.
3. Run `burnLinks`. The tab is deleted and the raw links no longer exist anywhere but in
   people's chats. From this point nobody, including you, can tell who cast which votes.
4. New player joins later? Run `generateLinks` again; only missing players get links.
   Someone lost or leaked theirs? Run `rotateLink('Their Name')`.

## Windows

- `openWindow()` opens a new quarter (or `openWindow('2026-Q4')` to name it).
- `closeWindow()` freezes it. Cards keep showing the last window until the next opens.

## Rules the script enforces

- One ballot per link; re-rating someone replaces the earlier rating.
- You can only rate teammates you've shared at least **3** sessions with, never yourself.
- Values are clamped to 40–99; comments to 280 characters.
- A player's card only shows a rating once **3** ballots have rated them.
- Aggregation: each voter's generosity is half-corrected toward the group average,
  the top and bottom 20% of votes are trimmed once there are 5+, and votes are weighted
  by how well the voter knows the player (stored only as a low/medium/high bucket).
- Comments are private: only the player they are about can read them, via their own link.
