# VCMC Tabulation Software

*Score tabulation for a mock DECA conference*

> **This is not a new project.** It is a copy of a private repository, published here so my own contributions to it are readable in one place. The original tabulator was written by Grace Xu in 2024 for VCMC, our DECA chapter's mock conference. My work on it was done in November 2026 in the original repository, before this copy existed. The original is private, so its history can't be linked, and the commit history here is just the import. The changes are described under [What I contributed](#what-i-contributed).

The program reads three spreadsheets (teams, judge scores and exam scores) and writes a ranked results workbook. Events are weighted as in the 2026-27 DECA Guide, and a judge who scored much harder or easier than the others has their scores adjusted toward the rest.

## What I contributed

The 2024 tabulator was a pandas script that read fixed column positions, used a 60/40 written and presentation split, and wrote one sheet per category. For the November 2026 conference I rewrote it.

### 2026-27 DECA scoring

DECA changed its weighting for 2026-27. I read the event guidelines for every event VCMC runs and matched the formulas to them.

- Roleplay events (Principles, Individual Series, Team Decision Making, Personal Financial Literacy) count the exam as 1/3 and the roleplay as 2/3. DECA gives the exam and each of two roleplays a third. VCMC has one roleplay, so it carries the full two thirds, which is also how DECA weights its final round. Team Decision Making averages both members' exams.
- Operations research and business plan events now split 50/50 between the written entry and the presentation. In 2024 the split was 60/40. VCMC mentors grade an executive summary out of 50, which is doubled to fill the 100-point written score.
- Pitch deck events (Innovation Plan, Start-Up Business Plan), the marketing campaign events, project management and professional selling are scored on the presentation minus penalty points. DECA judges the pitch deck inside the presentation rubric. DECA also counts a cluster exam for the marketing campaign and selling events, but VCMC students only take the exam for their roleplay event, so it isn't counted here.

### Judge curve

At the 2026 conference most judges scored about five teams each, and their averages ranged from 35 to 92 out of 100 on the same rubric. Students' exam scores showed the harsh judges hadn't simply drawn weaker students. Which room a team was sent to mattered more than how it performed.

I added a per-judge shift toward the pool average, shrunk by how many teams the judge saw:

```
shift = (pool average - judge average) * n / (n + lambda)
```

`n` is the number of teams the judge scored. `lambda` is the ratio of within-judge variance to between-judge variance, estimated from the scores themselves. A judge who scored five teams keeps most of the gap, and one who scored two keeps less, since two scores say little about a judge. The order inside a judge's room never changes and the margins between students are kept. Roleplay judges and written-event judges are curved as separate pools, and curved scores are capped at 0 to 100.

I chose this over z-scores and within-room ranks. With five scores per judge, z-scores swing on a single outlier and force every judge to the same spread. Ranks throw away the size of the gaps and penalize a room full of strong students.

### Bug fixes

- The exam lookup started at the second data row, so the first student in the exam file always got a 0.
- A student missing from the exam file silently scored 0. Missing exams are now flagged on the team's row.

### Output

The results workbook has these tabs:

- Event winners: top 3 in each event
- Category winners: top 3 in each award category
- Best overall participants: average of each student's event scores, ranked for students who competed in both a roleplay and a written event
- Event rankings: every team ranked within its event
- One tab per category: every team ranked within its category
- Judge curve: each judge's average and the points added or removed

Every ranking tab shows the score and rank without the curve and how many places the team moved, so the adjustment can be checked by hand.

The script now needs only openpyxl.

### Building the 2026 inputs

The conference data came in five spreadsheets: roleplay scores, written scores, executive summary scores, exam results and the schedule. I matched them to the schedule by name, with an alias table for nicknames and misspellings. Twenty forms where the judge's total didn't match the sum of the rubric items were flagged for checking by hand. Those scripts are not in this repository because they contain student names.

### A note on the data in this repository

The files in `sample-data` are generated. The real conference data has student and judge names and is not published here. The sample keeps the same shape as the real files, including the same 179 teams, the same event mix, the same team sizes and the same number of teams per judge. It also includes one very harsh judge in each pool, so the curve has something to correct. Scores are random and the names are made up.

## Run

```
pip install -r requirements.txt
python3 tabulate.py sample-data
```

This writes `sample-data/2026_VCMC_results.xlsx`. To choose the output file, pass it as a second argument.

## Input files

The folder needs one file ending in each of `_teams.xlsx`, `_judging_data.xlsx` and `_exam_scores.xlsx`.

Teams, by column position:
- A: Team number
- B: Event
- C, D: Captain name and ID
- E, F: Partner 1 name and ID
- G, H: Partner 2 name and ID
- I: Number of members
- Optional `Type` column (Roleplay or Written)

Judging:
- B: Judge name
- C: Team number
- E: Score 1. The roleplay, PSE or FCE judge score, or the written score out of 100 for research and business plan events
- F: Score 2. The presentation score for written events
- Optional `Exec Summary (/50)` and `Exec Summary Penalty Pts` columns

Exam scores:
- A: Student ID
- D: Score as a fraction (0.82 means 82%)
