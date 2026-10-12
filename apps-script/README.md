# Player ratings backend

The site is static, so anonymous voting needs a tiny backend. This is a Google
Apps Script web app attached to a **private** Google Sheet (not the attendance
sheet). Setup takes about ten minutes, once.

## One-time setup

1. Create a new Google Sheet called something like `W9-10 Ratings` (private, just you).
2. **Extensions → Apps Script**. Delete the default code, paste in `Code.gs`, save.
3. In the editor, pick the `setup` function and press **Run**. Approve the permissions.
   This creates the tabs `Voters`, `Votes`, `Windows` and `Seeds`, opens the first
   quarterly window (e.g. `2026-Q4`), and on a first run also creates everyone's links
   (the `Links` tab — see below).
4. **Deploy → New deployment → Web app**. Execute as **Me**, who has access **Anyone**.
   Copy the web-app URL (ends in `/exec`).
5. Put that URL into `index.html` as `VOTE_API` and push. The **Player Cards** tab goes live.

## Seeding indicative stats

Open the `Seeds` tab. Every player is listed with 65 across the board. Put in your
indicative numbers (40–99) for PAC / SHO / PAS / DRI / DEF / PHY. These are only the
starting point voters see; they never show on the public cards.

## Letting people claim their own link (recommended)

Post one message to the group with the **group code** from the `Settings` tab of the ratings
sheet (key `claim_code`; change it any time). Each player opens the site, taps **Get your
personal link**, picks their name, enters the code, and the site stores their link on that
device and shows it so they can save it. One claim per name; five wrong codes locks a name
for 15 minutes. If someone claims the wrong name, delete that row in the `Voters` tab and
they can claim again.

## Sending the links by hand (alternative)

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
- The admin sets an exact baseline per player (Seeds tab, or the Starting points editor on
  the admin's You tab). Cards show the baseline from day one. Each vote is a level (Needs
  work … Elite → 52/62/72/82/92) and the shown figure is a weighted mean in which the
  baseline counts like 3 votes, so votes trim the baseline rather than replace it. A player
  with no baseline needs 3 votes before anything shows.
- Each voter's generosity is half-corrected toward the group average, the top and bottom
  20% of votes are trimmed once there are 5+, and votes are weighted by how well the voter
  knows the player (stored only as a low/medium/high bucket).
- Comments are private: only the player they are about can read them, via their own link.

## Current deployment (2026-10-12)

- Sheet: `W9-10 Ratings` in Raji's Drive. Script project of the same name.
- Web app URL is set as `VOTE_API` in `index.html`. Redeploying a new version keeps the same URL only if you **Manage deployments → edit → new version**; a brand-new deployment gets a new URL.

## Hosting (2026-10-12)

The public address is **https://wednesday910.vercel.app** (Vercel project `wednesday910`). It is a
Vercel rewrite that proxies https://rajichalla11.github.io/wednesday-hof/, so deploys still happen
by pushing to `main` on GitHub; nothing on Vercel needs touching. Do not add a redirect on the
GitHub Pages side, or the proxy would loop. Personal links use whichever address the player opened,
so links claimed on the Vercel address work there.
