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

Each player opens the site, taps **Get your personal link**, picks their name, and the site
stores their link on that device and shows it so they can save it. No code is needed while the
`claim_code` cell in the `Settings` tab is blank; put a word there to require it. One claim per name; five wrong codes locks a name
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


## Photos (backend Version 6, 13 Oct 2026)

Players can add their own photo from the You tab (Add a photo / Change photo / Remove photo). The site shrinks it to 720px, posts it to the web app, and the script saves it in a Drive folder called **W9-10 Photos** (created on first upload, shared view-only by link) and records the file id in the **Photos** tab. A photo set this way wins over the Images column in the attendance sheet. To remove someone's photo as admin, delete their row in the Photos tab (and trash the file in Drive if you like).

The first upload needs the Drive permission: run `setup` once from the editor after deploying and approve the permissions prompt.


## Claims need approval (backend Version 7, 13 Oct 2026)

Claiming is now "Find your card": a grid of faces with a search box, taken names dimmed. A claim creates a Voters row with status **pending**. The person's device remembers them, but their page stays locked (no ballot, balance or photo) until the admin approves it on the You tab under **Claims waiting** (Approve / Not them). Rejecting deletes the row so the name can be claimed again. Rows from before this version with a `claimed` date and a blank status count as approved. If the admin's own row is ever reset, set its status to `approved` by hand in the Voters tab, since nobody else can approve it.


## Instant claims and "Not me" (backend Version 8, 13 Oct 2026)

Claims are approved on the spot again. To go back to admin approval, add a row `require_approval` = `1` on the Settings tab (the Claims waiting section on the You tab still works). Anyone who picked the wrong card taps **Not me** on their page (twice, to confirm): the POST action `unclaim` deletes their Voters row so the name is free again, the device forgets them, and the picker opens.


## Speed (backend Version 9, 13 Oct 2026)

Every tab the web app reads (Voters, Votes, SelfRatings, Windows, Seeds, Settings, Photos) is cached in CacheService for 5 minutes and busted by the write that changes it, so a call does one or two sheet reads instead of five. If you edit Settings, Seeds or Voters by hand and want it live at once, run `clearCaches` from the editor. The floor for any call is about 1.2 s (Apps Script start-up plus its redirect); the site hides that by showing the last answer it saw on that device and repainting only if the fresh one differs.


## Hidden ratings (backend Version 10, 13 Oct 2026)

Players listed in Settings → `hide_stats` (comma-separated names) show **To be determined** instead of numbers, on the cards, the profile overlay and the Players intro count. Their starting point is also left out of what voters receive, so nothing leaks through the ballot. If the `hide_stats` row is missing, the admin's own rating is hidden by default. To publish Raji's rating later, add a Settings row `hide_stats` with an empty value (or list other names), then run `clearCaches` or wait up to 5 minutes.


## Votes decide, open counter (backend Version 12, 13 Oct 2026)

- The site is open to everyone (`LOCKED = false` in index.html).
- Only players with a row in the Seeds tab have a starting point. Everyone else is decided by votes alone (their card shows dashes until 3 ballots). In the admin Starting points editor a player with no row reads "votes decide"; typing a number creates a row (the other five stats fill from the position defaults), and the **Votes decide** button on a row deletes it (POST `clearSeed`). `setup` no longer pre-fills the Seeds tab.
- Every page open sends one anonymous ping per browser session (`hit`: a random device id, the claimed name if any, phone or not, which tab). Rows land in the **Visits** tab. The admin's You tab shows Site activity: people and opens today / 7 days / all time, share on phones, who has opened and when, and claimed names not seen yet.


## Game day (backend Version 17, 14 Oct 2026)

Set it on your You tab under **Game day**: date, kick-off (Nairobi time) and gate code, then Save; **No game** clears it. Stored in Settings as `game_date`, `game_time`, `gate_code`.

- On the day, before kick-off, the site shows a Game day banner: countdown, gate code (only to players with a claimed card; it never goes into the public data), who's in tonight, and storylines (milestones and streaks). Players tap **I'm in** to check in (Checkins tab). If the attendance sheet already has tonight's column, those players count as in.
- During the hour it shows **Playing now**. From kick-off for 36 hours, players can **Rate tonight**: each teammate out of 10 (4–10 on the buttons) plus one man of the match. Anonymous like the season ballot (MatchVotes tab, ballot hash). A resubmission replaces the old one. Rating someone adds them to the squad.
- After the game the banner shows the live match-ratings leaderboard once 3 players have rated, and the man-of-the-match leader. Each night's man of the match (most votes, at least 2, ties share) earns a gold **Man of the match** medal, counted across games.


## Lineups and scores (backend Version 18, 14 Oct 2026)

Sign-up stays on WhatsApp. In the Game day panel, paste the BLACK / WHITE message exactly as posted (asterisks, dots and spacing are fine); the preview shows each side, and names it can't match to a player are marked as guests. After the game, fill in Black goals and White goals and Save again. Each game is kept in the **Games** tab (date, black, white, goals), and players' profiles show their won / drawn / lost record from games with a score. The I'm in button is gone; tonight's squad is the two lineups plus anyone already on the attendance sheet's column for the day.

## Keeping ballots anonymous (Version 20)

The admin's claims table used to show exactly how many teammates each player had rated and when they last voted. Watching that number tick up while a new ballot appeared would have matched the ballot to the player. Now:

- The Voters tab keeps only a rough band (1–4, 5–9, 10+) and no last-vote time. It's refreshed when the player next opens the site, not at the moment they vote.
- Ballot rows (Votes, MatchVotes) and Checkins store the day only, never the time, so they can't be lined up against the Visits tab.
- New ballot rows go in at a random position, so row order says nothing about who voted when.
- `anonymise()` (run once from the editor after deploying) cleans the data written before this: bands instead of counts, last-vote times cleared, times stripped, and existing ballot rows shuffled. It's safe to run again.

The ballot id is still a one-way hash of the player's private link. Nobody, including the sheet owner, can work back from a ballot to a name.

## Named votes for the admin (Version 21)

From 14 Oct 2026, Raji chose to see who voted what, while other players still only see totals. The ballot and the match-rating sheet say so plainly.

- New season votes are stored under a different ballot id (`id3_`, a separate hash from the old `id2_`) with the voter's name in a `voter` column at the end of the Votes tab. MatchVotes gets a `voter` column too. Because the hash is different, a named vote can't be used to find the same person's earlier anonymous ballot.
- Ballots made before the switch keep the old id and no name, and they still count. A voter can change one like any other rating (no separate warning; the ballot intro says Raji can see who gave what). `rate` overwrites the old row in place with a named one, and `unrate` removes either kind. Raji decided on 14 Oct that this link was fine.
- `admin` returns `named` (voter, target, stats, traits, comment), `anon` (per player, how many earlier anonymous ballots there are) and `matchNamed`. The You tab's "Who voted what" panel shows them, one player at a time.
- Public results never include names. The ballot count on the Players tab counts people who have voted (from the Voters band) rather than ballot ids, since one person can now hold one of each.

## Match voting window and the next game (Version 23)

- Match ratings open at kick-off and close at **11:00 the next morning** (Nairobi). `closes_(date)`.
- The man of the match is whoever got the most stars; a tie goes to the better average match rating, and if that's tied too, they share it (`motmFor_`). The gold medal counts a night only once its voting has closed.
- The last game played is read from the Games tab (which now also stores the kick-off `time`), so Raji can enter next week's game straight after the match without cutting short the voting on the last one. `results` returns `last` (the game played, with `closes`, `final` and `motm`) and `next` (the game in Settings, if it hasn't kicked off). `me` returns `last.mine`. `matchRate` always applies to the last game.
- The banner shows, in order: live (first hour), voting open until 11am (man of the match so far, a "Next game" line if one is set), then the next game as soon as it's entered (with last game's man of the match underneath), or the final result for up to 6 days if no next game is set.
- Admin: "Set up the next game →" in the Game day form moves the date on a week and clears the gate code, lineup and score. Changing the date also clears the score fields.

## Lists, standby and format (Version 24)

- The Game day lineup box takes either the BLACK / WHITE teams or, before teams are made, the plain list of who's in, as posted in WhatsApp (numbers, ticks and headings like "Monday 9 aside" are ignored). Names under a "Standby" / "Reserves" / "Waiting" line are standby. In a plain list, names beyond the spots for the format go to standby in posted order.
- Format (5- to 11-a-side) is Settings `game_size`, defaulting to 6. `game_teams` JSON now holds `{ black, white, list, standby }`. The Games tab gains `list` and `size` columns.
- The banner shows the list as "In · 16 of 18 · 2 spots left · teams to follow", standby with each player's place in line ("you're 2nd in line"), and "9-a-side" next to the kick-off when it isn't 6. Once the teams are pasted over the list, the banner switches to Black vs White. List players count as playing for match ratings; standby players don't.

## Man of the match override (Version 26)

- The admin You tab has a "Man of the match" panel listing the last six games played, with each night's stars and average match ratings. Its picker defaults to "By the vote"; choosing a player saves an override in column 9 (`motm`) of the Games tab (`setMotm`), and choosing "By the vote" again clears it.
- `motmFor_` returns the override when one is set, otherwise the vote (`motmAuto_`). The banner, the gold medal and the tally all follow it, and a picked winner shows as "Man of the match" even before voting closes.
- The lineup preview flags names that another player shares a first name with (Karan / Karan N) in gold, so a wrong match is easy to spot before saving.

## Fixed cards (Version 27)

- Settings → `fixed_stats` (comma-separated names, none by default): those players' cards show their starting point exactly, and votes don't move the six stats. Their traits still come from teammates' votes. Results mark them `fixed: true`.
- POST `setFlag` (admin) sets `hide_stats` or `fixed_stats` from the site.
- 17 Oct 2026: Raji's card was unhidden (`hide_stats` empty) and fixed to 81 · 82 · 92 · 94 · 68 · 82, his own numbers.

## Matches, freeing claims (Version 28)

- `results.games` now carries each game's time, format, list (for games with no teams), number of raters, and man of the match once voting has closed or the admin picked one. The site's **Matches** tab lists every played game, newest first: score row (Black/White, winner bright), man of the match, and the teams or who played.
- The game-day banner shows the score row for games entered as a plain list (no teams) too.
- POST `freeClaim` (admin, `{ name }`) deletes that player's Voters row so the right person can claim the card again; their earlier ballot stays and still counts. The admin claims table has a two-tap **Free** button per row.
- The earlier-ballot naming flow (ackNamed) was not shipped; the pre-14 Oct ballots stay anonymous.

## Form guide, sturdier calls (Version 30)

- `results.games[].ratings` holds each player's average match rating for that night (players with 2+ votes). The site builds a form guide from it: the last 5 games as chips in the player's strip on Home (rating, W/D/L for their side, ★ for man of the match), and the last 8 as a table on their profile. Numbers follow the rate-10-teammates lock.
- The site's API helper retries when Google answers with an HTML error page instead of JSON (it happens when several calls land at once), and the admin panels show a Try again button instead of hanging on Loading….

## Sign-ups on the site (Version 31)

Replaces the copy-pasted WhatsApp list. The admin opens games ("+ Open a game for sign-ups" on Home: date, kick-off, format), which adds a Games row with no lineup. Several can be open at once.

- **Signups tab**: date, ts, name, kind (player | guest), by. Order by ts is the list; the first 2 × size are in, the rest standby, so when someone drops out the next standby moves up automatically (their card says "You're in now").
- Players with a claimed card tap **I'm in** / **I'm out**, add another player ("added by …") or a guest ("guest · via …"). Only the person, whoever added them, or the admin can take someone off. No lock.
- The list feeds the game: `allGames_()` uses sign-ups as the game's list and standby (unless teams are set), so the banner, team maker and match ratings follow it. `results.fixtures` lists open games with their sign-ups.
- Admin: move one game's standby onto another game's list, remove a game, and `importList` to load a list taken elsewhere.
- **Copy for WhatsApp** produces the usual "Monday 9-10pm / 1. … / Standby:" text with a sign-up link (`/#in`).

