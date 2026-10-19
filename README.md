# Pilot Schedule Bid

A browser app that runs the yearly line bid from the bid packet spreadsheet and writes the results back into the same workbook.

## Running it

Open `index.html` in Chrome, Edge or Firefox. There's nothing to install. It needs an internet connection the first time to load the zip library (JSZip, from cdnjs).

## How a bid works

1. **Setup**
   - Choose the bid packet `.xlsx` and the year's sheet (for example `2026`).
   - The app reads from that sheet:
     - the numbered lines (rows under the date header, down to the Chief Pilot row)
     - holidays (the row of `H` marks)
     - event days (any mark in the `BOARD DATES` or `MT` rows)
     - safety and company training days (dates marked "Safety Day" or "training" below the lines, outside the recurrent TRAINING DATES row), which use the "Pilots Required on Safety & Training" minimum (everyone, if the sheet has no number)
     - staffing minimums (the "Pilots Required on …" cells in column B)
     - training sessions (yellow blocks in the `TRAINING DATES` row, labelled with initials)
   - It flags template errors, such as an off block of the wrong length or `"OFF "` with a trailing space.
   - Enter the pilots in seniority order with their initials, or copy names from an earlier year's sheet. Set V and HAC allowances (a default for everyone, or per pilot) and the rounds.
2. **Bidding** goes in seniority order every round.
   - **Rounds:** round 1 is always the line pick on its own, so all training is on the calendar before anyone picks a day off. Each later round gives every pilot a number of V, HAC and holiday days. By default the holiday round comes right after the line pick, followed by last year's two rounds of 7 V + 2 HAC and one round of 5 V. Rounds after round 1 can be put in any order with the ↑ ↓ buttons. The yearly V and HAC allowances (19 and 4 by default, or set per pilot) cap the total.
   - **Line pick:** training goes on the line as soon as it's picked: `P` the day before, `T` on training days, `P` the day after. A training day that lands on a line day off is written `OFF T` / `OFF P`, as in 2025.
   - **Holidays (ranking round):** on their turn each pilot ranks the holidays they want, most wanted first. Holidays they're already off or in training for aren't offered. Once every pilot has ranked, holidays are awarded one at a time: the most senior pilot gets their top remaining choice, then the next pilot, round robin. A pilot's choice is skipped if it would leave fewer pilots available than the holiday minimum. Awards stop when each pilot has the round's "Holidays per pilot" number or nothing more can be granted. Results lists who got what, with each pilot's choice number.
   - **Days off:** the pilot switches between V and HAC and clicks a first and last day on their line's calendar. Each V/HAC pick must make a group of at least 2 days off in a row (a single day is fine when it's added next to days off the pilot already has), and must either touch other days off or leave at least 2 working days before the next group of days off. Both numbers are set per round, along with an optional "must touch an existing block of days off".
   - The app blocks any day that would drop availability below the minimum (safety/training day, event, holiday, weekday or weekend, checked in that order). That check includes training still to be placed for pilots who haven't picked a line yet.
   - "Undo last" steps back one pick.
3. **Results:** download the filled spreadsheet. The app writes pilot names in column A of their lines and the P/T/V codes into the day cells. HAC days are written as `X` by default (change it under Pilots → "HAC written as"). All other sheets, formatting and formulas stay unchanged, and Excel recalculates the counts when the file is opened.

## Practice mode

- **View as** (top right) shows the app as the admin or as any pilot. A pilot sees only the bidding screen: their own turn with the pickers, or a waiting screen showing whose turn it is, how many pilots are ahead of them, and their own line.
- In the admin view, the bidding screen has **Auto-pick this turn / rest of round / whole bid**, which makes reasonable picks so you can jump ahead. **Start over** clears all picks but keeps settings and pilots.

## Saving

Progress is saved automatically in this browser. **Save bid (Excel)** downloads an Excel workbook with everything picked so far written into the sheet, so it can be opened in Excel. The workbook also carries the bid itself (settings, pilots, picks) in a hidden sheet, so **Open saved bid** (or loading it under Setup) carries on exactly where it left off, on this or another computer. Older `.json` bid files still open. The final **Download spreadsheet** on Results leaves the hidden sheet out.

## Files

- `engine.js`: bid rules (turn order, allowances, staffing checks, training placement)
- `xlsx.js`: reads the template and writes results into the sheet XML
- `app.js`, `index.html`: the interface
