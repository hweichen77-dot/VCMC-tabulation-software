import glob
import os
import statistics
import sys
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

SERIES = {'ACT', 'AAM', 'ASM', 'BFS', 'BSM', 'ENT', 'FMS', 'HLM', 'HRM', 'MCS', 'QSRM', 'RFSM', 'RMS', 'SEM'}
TDM = {'BLTDM', 'BTDM', 'ETDM', 'FTDM', 'HTDM', 'MTDM', 'STDM', 'TTDM'}
PRINCIPLES_A = {'PEN', 'PMK'}
PRINCIPLES_B = {'PBM', 'PFN', 'PHT'}
PFL = {'PFL'}
RESEARCH = {'BOR', 'BMOR', 'BSOR', 'FOR', 'HTOR', 'SEOR'}
BUSINESS_PLANS = {'EBG', 'EFB', 'EIB', 'IBP'}
PITCH_DECKS = {'EIP', 'ESB'}
IMC = {'IMCE', 'IMCP', 'IMCS'}
SELLING = {'FCE', 'HTPS', 'PSE'}
PROJECT_MANAGEMENT = {'PMBS', 'PMCD', 'PMCA', 'PMCG', 'PMFL', 'PMSP'}

ROLEPLAY_EVENTS = SERIES | TDM | PRINCIPLES_A | PRINCIPLES_B | PFL
WRITTEN_PLUS_PRESENTATION = RESEARCH | BUSINESS_PLANS
JUDGE_SCORE_IN_SCORE_1 = ROLEPLAY_EVENTS | SELLING

CATEGORIES = [
    ('Principles (PEN, PMK)', PRINCIPLES_A),
    ('Principles (PBM, PFN, PHT)', PRINCIPLES_B),
    ('Individual Series', SERIES),
    ('Team Decision Making', TDM),
    ('Personal Financial Literacy', PFL),
    ('Operations Research', RESEARCH),
    ('Business Plans', BUSINESS_PLANS),
    ('Pitch Decks (EIP, ESB)', PITCH_DECKS),
    ('Marketing Campaigns', IMC),
    ('Professional Selling', SELLING),
    ('Project Management', PROJECT_MANAGEMENT),
]

EXAM_WEIGHT = 1 / 3
ROLEPLAY_WEIGHT = 2 / 3
WRITTEN_WEIGHT = 0.5
PRESENTATION_WEIGHT = 0.5
MIN_EVENTS_FOR_OVERALL = 2


def find_input(folder, suffix):
    matches = sorted(glob.glob(os.path.join(folder, f'*{suffix}')))
    if not matches:
        sys.exit(f'No *{suffix} file in {folder}')
    return matches[-1]


def first_sheet_rows(path):
    sheet = load_workbook(path, data_only=True).worksheets[0]
    rows = list(sheet.iter_rows(values_only=True))
    return [str(h).strip() if h is not None else '' for h in rows[0]], rows[1:]


def as_number(value):
    if value is None or value == '':
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def column(header, name):
    return header.index(name) if name in header else None


def load_teams(path):
    header, rows = first_sheet_rows(path)
    type_col = column(header, 'Type')
    teams = {}
    for row in rows:
        if row[0] is None:
            continue
        count = int(row[8])
        members = [(row[2 + 2 * i], int(row[3 + 2 * i])) for i in range(count) if row[2 + 2 * i]]
        event = str(row[1]).strip()
        kind = row[type_col] if type_col is not None else ('Roleplay' if event in ROLEPLAY_EVENTS else 'Written')
        teams[int(row[0])] = dict(number=int(row[0]), event=event, members=members, kind=kind)
    return teams


def load_exams(path):
    _, rows = first_sheet_rows(path)
    return {int(r[0]): float(r[3]) * 100 for r in rows if r[0] is not None and as_number(r[3]) is not None}


def load_judging(path, teams):
    header, rows = first_sheet_rows(path)
    written_col = column(header, 'Exec Summary (/50)')
    penalty_col = column(header, 'Exec Summary Penalty Pts')
    for row in rows:
        if row[2] is None:
            continue
        team = teams[int(row[2])]
        team['judge'] = str(row[1]).strip() if row[1] else None
        score1, score2 = as_number(row[4]), as_number(row[5])
        if team['event'] in JUDGE_SCORE_IN_SCORE_1:
            team['judge_raw'] = score1
        else:
            team['judge_raw'] = score2
            team['written'] = score1
        if written_col is not None and as_number(row[written_col]) is not None:
            team['written'] = as_number(row[written_col]) * 2
        team['penalty'] = as_number(row[penalty_col]) if penalty_col is not None else None


def judge_curve(scores):
    by_judge = defaultdict(list)
    for judge, score in scores:
        by_judge[judge].append(score)
    overall = statistics.mean(s for _, s in scores)
    groups = list(by_judge.values())
    within = sum((x - statistics.mean(g)) ** 2 for g in groups for x in g) / (len(scores) - len(groups))
    spread_of_means = statistics.variance([statistics.mean(g) for g in groups])
    between = max(spread_of_means - within * statistics.mean(1 / len(g) for g in groups), 1e-9)
    shrink = within / between
    table = {}
    for judge, g in by_judge.items():
        weight = len(g) / (len(g) + shrink)
        table[judge] = dict(n=len(g), mean=statistics.mean(g), weight=weight, shift=weight * (overall - statistics.mean(g)))
    return dict(overall=overall, within=within, between=between, shrink=shrink, judges=table)


def clamp(value):
    return max(0.0, min(100.0, value))


def score_team(team, exams, curves):
    judge_raw = team.get('judge_raw')
    if judge_raw is None:
        team['status'] = 'No judge score'
        return
    curve = curves[team['kind']]['judges'][team['judge']]
    team['shift'] = curve['shift']
    team['judge_curved'] = clamp(judge_raw + curve['shift'])
    event = team['event']

    def combine(judge_score):
        if event in ROLEPLAY_EVENTS:
            exam_scores = [exams.get(sid) for _, sid in team['members']]
            missing = [name for (name, _), s in zip(team['members'], exam_scores) if s is None]
            if missing:
                team.setdefault('flags', []).append('No exam for ' + ', '.join(missing) + ' (counted as 0)')
            team['exam'] = statistics.mean(s or 0 for s in exam_scores)
            return team['exam'] * EXAM_WEIGHT + judge_score * ROLEPLAY_WEIGHT
        if event in WRITTEN_PLUS_PRESENTATION:
            written = team.get('written')
            if written is None:
                team.setdefault('flags', []).append('No exec summary score (counted as 0)')
                written = 0
            return written * WRITTEN_WEIGHT + judge_score * PRESENTATION_WEIGHT - (team.get('penalty') or 0)
        return judge_score - (team.get('penalty') or 0)

    team['final'] = round(combine(team['judge_curved']), 2)
    team['final_no_curve'] = round(combine(judge_raw), 2)
    team['status'] = 'Scored'


def rank(teams, key):
    ordered = sorted(teams, key=lambda t: -t[key])
    ranks = {}
    for i, t in enumerate(ordered):
        ranks[t['number']] = ranks[ordered[i - 1]['number']] if i and t[key] == ordered[i - 1][key] else i + 1
    return ranks


def write_sheet(workbook, title, header, rows, number_cols=()):
    sheet = workbook.create_sheet()
    sheet.title = title[:31]
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    for line in sheet.iter_rows(min_row=2):
        for i in number_cols:
            line[i].number_format = '0.00'
    for i, col in enumerate(sheet.columns, start=1):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        sheet.column_dimensions[get_column_letter(i)].width = min(max(8, width + 1), 50)
    return sheet


def blank(value):
    return None if value is None else round(value, 2)


def team_row(t, place, place_no_curve):
    names = ', '.join(name for name, _ in t['members'])
    moved = place_no_curve - place if place and place_no_curve else None
    return [place, t['event'], t['number'], names, blank(t.get('final')),
            blank(t.get('exam')), blank(t.get('written')), blank(t.get('judge_raw')), blank(t.get('shift')),
            blank(t.get('judge_curved')), blank(t.get('penalty')), blank(t.get('final_no_curve')),
            place_no_curve, moved, t.get('judge'), '; '.join(t.get('flags', [])) or None]


DETAIL_HEADER = ['Team #', 'Members', 'Final Score', 'Exam %', 'Written (/100)', 'Judge Raw', 'Judge Curve',
                 'Judge Curved', 'Penalty Pts', 'Final (no curve)', 'Rank (no curve)', 'Places Moved', 'Judge', 'Flags']
CATEGORY_HEADER = ['Category Rank', 'Event'] + DETAIL_HEADER
EVENT_HEADER = ['Event', 'Event Rank'] + DETAIL_HEADER


def ranked_rows(group):
    scored = [t for t in group if t['status'] == 'Scored']
    places = rank(scored, 'final')
    places_no_curve = rank(scored, 'final_no_curve')
    ordered = sorted(scored, key=lambda t: (places[t['number']], t['number'])) + [t for t in group if t['status'] != 'Scored']
    return [team_row(t, places.get(t['number']), places_no_curve.get(t['number'])) for t in ordered]


def overall_rows(teams, exams):
    per_student = defaultdict(lambda: dict(Roleplay=[], Written=[]))
    for t in teams.values():
        if t['status'] == 'Scored':
            for name, sid in t['members']:
                per_student[(name, sid)][t['kind']].append(t)

    def label(events):
        return '; '.join(f"{t['event']} {t['final']:.2f}" for t in events) or None

    students = []
    for (name, sid), kinds in per_student.items():
        finals = [t['final'] for events in kinds.values() for t in events]
        students.append(dict(name=name, sid=sid, score=round(statistics.mean(finals), 2), events=len(finals),
                             roleplay=label(kinds['Roleplay']), written=label(kinds['Written']), exam=exams.get(sid)))
    eligible = sorted((s for s in students if s['events'] >= MIN_EVENTS_FOR_OVERALL), key=lambda s: (-s['score'], s['name']))
    others = sorted((s for s in students if s['events'] < MIN_EVENTS_FOR_OVERALL), key=lambda s: (-s['score'], s['name']))
    rows = []
    for i, s in enumerate(eligible):
        place = rows[-1][0] if i and s['score'] == eligible[i - 1]['score'] else i + 1
        rows.append([place, s['name'], s['sid'], s['score'], s['events'], s['roleplay'], s['written'], blank(s['exam']), None])
    for s in others:
        rows.append([None, s['name'], s['sid'], s['score'], s['events'], s['roleplay'], s['written'], blank(s['exam']),
                     'Only one event scored, not ranked overall'])
    return rows


def judge_rows(curves):
    rows = []
    for kind, curve in sorted(curves.items()):
        for judge, j in sorted(curve['judges'].items(), key=lambda kv: kv[1]['shift']):
            rows.append([kind, judge, j['n'], round(j['mean'], 2), round(curve['overall'], 2), round(j['weight'], 3), round(j['shift'], 2)])
    rows.append([])
    for kind, curve in sorted(curves.items()):
        rows.append([kind, f"lambda = {curve['shrink']:.2f}", f"within-judge variance {curve['within']:.1f}",
                     f"between-judge variance {curve['between']:.1f}"])
    return rows


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else '.'
    teams = load_teams(find_input(folder, '_teams.xlsx'))
    exams = load_exams(find_input(folder, '_exam_scores.xlsx'))
    load_judging(find_input(folder, '_judging_data.xlsx'), teams)

    unknown = sorted({t['event'] for t in teams.values()} - set().union(*(events for _, events in CATEGORIES)))
    if unknown:
        sys.exit('Events with no category: ' + ', '.join(unknown))

    pools = defaultdict(list)
    for t in teams.values():
        if t.get('judge_raw') is not None and t.get('judge'):
            pools[t['kind']].append((t['judge'], t['judge_raw']))
    curves = {kind: judge_curve(scores) for kind, scores in pools.items()}
    for t in teams.values():
        score_team(t, exams, curves)
        if t['status'] != 'Scored':
            t.setdefault('flags', []).append('Absent / no judge score, not ranked')

    category_tabs = []
    category_winners = []
    for name, events in CATEGORIES:
        group = [t for t in teams.values() if t['event'] in events]
        if group:
            rows = ranked_rows(group)
            category_tabs.append((name, rows))
            category_winners += [[name, r[0], r[1], r[3], r[4], len(group)] for r in rows if r[0] is not None and r[0] <= 3]

    event_rows = []
    event_winners = []
    for event in sorted({t['event'] for t in teams.values()}):
        group = [t for t in teams.values() if t['event'] == event]
        rows = ranked_rows(group)
        event_rows += [[r[1], r[0]] + r[2:] for r in rows]
        event_winners += [[event, r[0], r[3], r[4], len(group)] for r in rows if r[0] is not None and r[0] <= 3]

    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, 'Event winners', ['Event', 'Place', 'Members', 'Final Score', 'Teams in Event'], event_winners, number_cols=(3,))
    write_sheet(workbook, 'Category winners', ['Category', 'Place', 'Event', 'Members', 'Final Score', 'Teams in Category'],
                category_winners, number_cols=(4,))
    write_sheet(workbook, 'Best overall participants',
                ['Rank', 'Name', 'Student ID', 'Overall Score', 'Events Scored', 'Roleplay', 'Written', 'Exam %', 'Note'],
                overall_rows(teams, exams), number_cols=(3, 7))
    write_sheet(workbook, 'Event rankings', EVENT_HEADER, event_rows, number_cols=range(4, 12))
    for name, rows in category_tabs:
        write_sheet(workbook, name, CATEGORY_HEADER, rows, number_cols=range(4, 12))
    write_sheet(workbook, 'Judge curve', ['Pool', 'Judge', 'Teams Scored', 'Judge Average', 'Pool Average', 'Weight n/(n+lambda)', 'Points Added'],
                judge_rows(curves), number_cols=(3, 4, 6))

    year = os.path.basename(find_input(folder, '_teams.xlsx')).split('_')[0]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(folder, f'{year}_VCMC_results.xlsx')
    workbook.save(out)
    print('Saved', out)
    for kind, curve in sorted(curves.items()):
        print(f"{kind}: {len(curve['judges'])} judges, lambda {curve['shrink']:.2f}, pool average {curve['overall']:.1f}")


if __name__ == '__main__':
    main()
