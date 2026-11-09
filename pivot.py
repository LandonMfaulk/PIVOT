"""PIVOT | A Faulk Product. Offline pile verification and deviation reporting."""
# Heyyy, whatcha doin'? Are you supposed to be here?
import csv
import math
import os
import re
import sys
import zipfile
import tempfile
import uuid
import traceback
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pile_centers import CenterEngine, normalize_description, selected_center_elevation
from pile_center_ui import CenterWindow
from direct_shots import TaggedShotTracker, parse_tagged_description
from coverage import raw_key, merge_raw, reconcile, untagged_highspot_candidates

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
COL_RE = re.compile(r'^([A-Z]+)')

def colnum(ref):
    letters = COL_RE.match(ref).group(1)
    n = 0
    for x in letters: n = n * 26 + ord(x) - 64
    return n

def value(cell, strings):
    kind = cell.get('t', '')
    v = cell.find(NS+'v')
    if kind == 'inlineStr':
        return ''.join(t.text or '' for t in cell.iter(NS+'t')).strip()
    if v is None: return ''
    if kind == 's': return strings[int(v.text)]
    return v.text or ''

def _standards_sheet_metadata(archive):
    """Return worksheet names and visibility from workbook metadata.

    Files frequently contain hidden copies of the pile log; do not treat
    those copies as authoritative when the visible All Areas sheet exists.
    """
    meta={}
    try:
        root=ET.fromstring(archive.read('xl/workbook.xml'))
        relations=ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        relns='{http://schemas.openxmlformats.org/package/2006/relationships}'
        rid_to_target={r.get('Id'):r.get('Target','') for r in relations.findall(relns+'Relationship')}
        for order,sh in enumerate(root.iter(NS+'sheet')):
            rid=sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target=rid_to_target.get(rid,'').replace('\\','/')
            if not target:continue
            if target.startswith('/'): target=target.lstrip('/')
            elif not target.startswith('xl/'):target='xl/'+target
            # Most OpenXML writers store the target as worksheets/sheetN.xml.
            from posixpath import normpath
            target=normpath(target)
            meta[target]={'name':sh.get('name',target),
                          'hidden':sh.get('state','visible')!='visible',
                          'order':order}
    except (KeyError,ET.ParseError,ValueError):
        pass
    return meta


def read_standards(path, with_diagnostics=False):
    """Read design piles, prioritizing the primary visible design worksheet.

    The standards may include area tabs plus an older hidden calculation copy.
    Do not fail the whole import on conflicting duplicates: provide diagnostics
    and require explicit operator confirmation in the UI before using them.
    """
    with zipfile.ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            strings=[''.join(t.text or '' for t in si.iter(NS+'t')) for si in root.findall(NS+'si')]
        sheet_meta=_standards_sheet_metadata(z)
        sheet_names=sorted(
            (n for n in z.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)),
            key=lambda n:int(re.search(r'sheet(\d+)',n).group(1)))
        all_piles=[]
        for sheetname in sheet_names:
            root=ET.fromstring(z.read(sheetname))
            rows=[]
            for row in root.iter(NS+'row'):
                cells={colnum(c.get('r')):value(c,strings) for c in row.findall(NS+'c')}
                rows.append((row.get('r'),cells))
            groups=[]
            for _,cells in rows[:15]:
                for col,v in cells.items():
                    if (v.strip().upper() in ('PILE ID','PILE #','PILE NUMBER')
                        and cells.get(col+1,'').strip().upper()=='NORTHING'
                        and cells.get(col+2,'').strip().upper()=='EASTING'):
                        groups.append((col,col+1,col+2,col+4))
            info=sheet_meta.get(sheetname,{'name':sheetname,'hidden':False,
                                          'order':sheet_names.index(sheetname)})
            for row_number,cells in rows:
                for idc,nc,ec,zc in groups:
                    pid=str(cells.get(idc,'')).strip()
                    if not pid or pid.upper() in ('PILE ID','PILE #','PILE NUMBER'):
                        continue
                    try:
                        north=float(cells[nc]);east=float(cells[ec])
                        if not (math.isfinite(north) and math.isfinite(east)):
                            continue
                    except (KeyError,ValueError,TypeError):
                        continue
                    try:
                        elev=float(cells[zc])
                        if not math.isfinite(elev):elev=None
                    except (KeyError,ValueError,TypeError):
                        elev=None
                    all_piles.append({'id':pid,'north':north,'east':east,'elev':elev,
                        'source':f"{info['name']} row {row_number}",
                        '_sheet':info['name'],'_hidden':info['hidden'],
                        '_sheet_order':info['order'],
                        '_master':('all areas' in info['name'].lower()
                                   or 'all_areas' in info['name'].lower())})
        if not all_piles:
            raise ValueError('No pile standards found. Expected PILE ID / NORTHING / EASTING / TOP OF PILE / TOP OF CAGE columns.')
        # Primary sheet first; visible area tabs next; hidden worksheets last.
        # Stable ordering ensures adding an extra hidden copy cannot silently
        # replace a coordinate found in the official main design sheet.
        all_piles.sort(key=lambda p:(0 if p['_master'] and not p['_hidden'] else
                                      1 if not p['_hidden'] else 2,
                                      p['_sheet_order']))
        seen={};unique=[];conflicts=[];secondary_only=[]
        conflicting_ids=set()
        for pile in all_piles:
            key=pile['id'].upper()
            if key not in seen:
                seen[key]=pile;unique.append(pile)
                if pile['_hidden']:
                    secondary_only.append(pile)
                continue
            selected=seen[key]
            if (abs(selected['north']-pile['north'])>0.000001
                or abs(selected['east']-pile['east'])>0.000001):
                if key not in conflicting_ids:
                    conflicts.append({'pile_id':selected['id'],
                                      'chosen':selected.copy(),'alternate':pile.copy()})
                    conflicting_ids.add(key)
        # Report all anomalous IDs, but avoid overwhelming users with hundreds
        # of duplicate copies that agree with the master.
        diagnostics={'conflicts':conflicts,'secondary_only':secondary_only,
                     'primary_sheet':next((p['_sheet'] for p in unique if p['_master']
                                           and not p['_hidden']),None),
                     'total_occurrences':len(all_piles)}
        for pile in unique:
            pile.pop('_hidden',None);pile.pop('_master',None)
            pile.pop('_sheet',None);pile.pop('_sheet_order',None)
        return (unique,diagnostics) if with_diagnostics else unique

def read_field(path):
    result=[]
    with open(path,encoding='utf-8-sig',newline='') as f:
        for line,row in enumerate(csv.reader(f),1):
            if len(row)<4: continue
            desc=','.join(row[4:]).strip() if len(row)>4 else ''
            try:
                north=float(row[1]); east=float(row[2]); elev=float(row[3])
                if not all(math.isfinite(v) for v in (north,east,elev)):
                    raise ValueError('non-finite survey value')
            except (ValueError,TypeError):
                # Never silently discard a named pile: a bad coordinate is an
                # actionable import error, not permission to omit the pile.
                if described_pile_id(desc):
                    raise ValueError(f'CSV line {line}, pile {desc}: invalid '
                                     'Northing, Easting, or Elevation. '
                                     'Correct the source record and re-import.')
                continue
            result.append({'point':row[0].strip(),'north':north,'east':east,'elev':elev,'desc':desc,'line':line})
    if not result: raise ValueError('No numeric survey points found in field CSV.')
    return result

def match_piles(field, standards):
    """Global shortest-first unique pairing. Each field point and design pile is used at most once.
    Both northing and easting contribute to distance; no tolerance-based exclusion.
    """
    # Each survey point considers all piles, and assignments are sorted globally by distance.
    pairs=[]; alternatives=[]
    for i,f in enumerate(field):
        candidates=sorted(((math.hypot(f['north']-p['north'],f['east']-p['east']),j) for j,p in enumerate(standards)))
        alternatives.append(candidates[:3])
        for d,j in candidates: pairs.append((d,i,j))
    pairs.sort()
    assignments={}; used=set()
    for d,i,j in pairs:
        if i not in assignments and j not in used:
            assignments[i]=(j,d);used.add(j)
            if len(assignments)==min(len(field),len(standards)):break
    results=[]
    for i,f in enumerate(field):
        nearest=alternatives[i]
        if i not in assignments:
            results.append({'field':f,'design':None,'status':'UNMATCHED','distance':None,'second':None});continue
        j,dist=assignments[i];p=standards[j]
        second=next((d for d,k in nearest if k!=j),None)
        # review if assignment isn't the nearest candidate, or if nearest two are close
        status='MATCHED'
        if nearest[0][1]!=j: status='REVIEW - COLLISION'
        elif second is not None and second-dist < max(0.25,dist*0.2): status='REVIEW - CLOSE CANDIDATES'
        results.append({'field':f,'design':p,'status':status,'distance':dist,'second':second})
    return results

def fmt(x): return '' if x is None else f'{x:.3f}'

def direction(delta, positive, negative): return 'On design' if delta == 0 else positive if delta>0 else negative


DATE_PATTERN = re.compile(r'_(\d{2})(\d{2})(\d{2})(?:_|$)')

def shot_info(filename, csv_path=None):
    """Survey convention JOB_MMDDYY_CREW. Use header Job if filename was renamed."""
    candidates=[Path(filename).stem]
    if csv_path:
        with open(csv_path,encoding='utf-8-sig',errors='replace') as fp:
            first=fp.readline()
        m=re.search(r'Job:([^,\s]+)',first,re.I)
        if m: candidates.append(m.group(1))
    for name in candidates:
        match=DATE_PATTERN.search(name)
        if match:
            try:
                date=datetime.strptime(''.join(match.groups()),'%m%d%y').date()
                return date.isoformat(),name
            except ValueError: pass
    raise ValueError('No MMDDYY shot date found in the filename or CSV Job header. Example: 260170_102126_DM.csv')

def _coordinate_identity(value):
    """Stable Excel round-trip identity at sub-millimeter precision in feet.

    Excel can serialize large Easting/Northing floats with a few low-order binary
    digits changed. Seven coordinate decimals avoid session-ID-dependent duplicate
    center records without conflating distinct survey observations at report scale.
    """
    if value in ('', None):
        return None
    return round(float(value), 7)


def record_identity(record):
    """Identify the measurement, not the app session or imported workbook.

    PIVOT can regenerate the same polygon centroid under a new CENTER_SESSION ID.
    A center is identical only if pile/date, original perimeter source signature,
    calculated XY and selected Z all agree. Different shot dates, source lists,
    coordinates or selected elevations must remain independent historical records.
    """
    f=record['field']; pile=report_pile_id(record)
    shot_date=str(record.get('shot_date') or '').strip()
    point=str(f.get('point','')).strip().upper()
    coordinates=tuple(_coordinate_identity(f.get(k)) for k in ('north','east','elev'))
    description=' '.join(str(f.get('desc') or '').split()).casefold()
    if point.startswith('CENTER-') or str(record.get('status','')).startswith('CENTER -'):
        return ('CALCULATED_CENTER',pile,shot_date,coordinates,description)
    return ('SURVEY_OBSERVATION',pile,shot_date,point,coordinates,description)


def normalize_drill_date(value):
    """Read editable Excel drill dates (including Excel serial dates) safely.

    Drill Date is distinct from Date Shot. Never infer drilling dates from survey
    filenames or the date a report was generated.
    """
    text=str(value or '').strip()
    if not text:
        return ''
    try:
        serial=float(text)
        if 30000 <= serial <= 90000:
            return (datetime(1899,12,30)+timedelta(days=serial)).strftime('%m/%d/%Y')
    except (ValueError,OverflowError):
        pass
    for fmt_ in ('%Y-%m-%d','%m/%d/%Y','%m/%d/%y','%m-%d-%Y','%m-%d-%y'):
        try:
            return datetime.strptime(text,fmt_).strftime('%m/%d/%Y')
        except ValueError:
            pass
    # Permits non-date placeholders (e.g. 'Pending') without inventing a value.
    return text

def deduplicate_records(records):
    """Remove true duplicate observations while keeping all distinct survey shots.

    Preserve the earliest matching record as the canonical history entry. Carry
    forward later user comments and a live-import marker when present. Source
    files and unrelated survey revisions remain untouched.
    """
    out=[]; seen={}
    for record in records:
        key=record_identity(record)
        prior=seen.get(key)
        if prior is None:
            out.append(record)
            seen[key]=record
            continue
        if record.get('comment') and not prior.get('comment'):
            prior['comment']=record['comment']
        elif record.get('comment') and record['comment']!=prior.get('comment'):
            # Retain both comments rather than silently losing field review notes.
            comments=[str(prior['comment']),str(record['comment'])]
            prior['comment']=' | '.join(dict.fromkeys(comments))
        for key in ('drill_date','rig_number'):
            if str(record.get(key) or '').strip():
                # Explicitly supplied newer report metadata wins over blank data.
                prior[key]=record[key]
        if record.get('_live_import'):
            prior['_live_import']=True
    return out

def accepted_center_record(engine, pile_id, standards, elevation_mode='Blank', manual_elevation=None,
                           batch='CENTER_SESSION'):
    """Convert an explicitly approved center into a derived deviation record.

    No survey point is changed or replaced; only exact pile-ID design matches are used.
    A horizontal centroid has no inherent surveyed elevation, so blank is the default.
    """
    approved=engine.accepted.get(pile_id)
    if approved is None:
        raise ValueError(f'{pile_id}: approve the center before adding it to the deviation table.')
    calc=engine.calculations[pile_id]
    # The chosen elevation is derived only from included exterior shots.
    # No horizontal center calculation is allowed to fabricate a Z coordinate.
    elev=selected_center_elevation(calc,elevation_mode,manual_elevation)
    pile_key=pile_id.strip().upper()
    design=next((p for p in standards if p['id'].strip().upper()==pile_key), None)
    distance=(math.hypot(approved['north']-design['north'],approved['east']-design['east'])
              if design else None)
    sources=calc.get('sources',[])
    known_dates=set()
    jobs=set()
    unknown_sources=0
    for source in sources:
        try:
            try:
                date,job=shot_info(source)
            except ValueError:
                date,job=shot_info(source,source)
            known_dates.add(date);jobs.add(job)
        except (ValueError,OSError):
            unknown_sources+=1
            jobs.add(Path(source).stem)
    if len(known_dates)==1 and unknown_sources==0:
        shot_date=next(iter(known_dates))
    elif len(known_dates)>1 and unknown_sources==0:
        shot_date='Multiple dates'
    else:
        shot_date='Unknown (center)'
    source_ids=', '.join(str(o['point']) for o in calc['included'])
    field={'point':f'CENTER-{pile_key}', 'north':approved['north'],
           'east':approved['east'], 'elev':elev, 'desc':
           f'Derived center ({approved["method"]}); elevation: {elevation_mode}; source points: {source_ids}', 'line':0}
    if not design:
        status='CENTER - STANDARD NOT FOUND'
    elif calc['status']=='Review Required':
        status='CENTER - QC REVIEWED'
    else:
        status='CENTER - MATCHED'
    return {'field':field,'design':design,'batch':batch,'shot_date':shot_date,
            'shot_job':', '.join(sorted(jobs)) or 'Perimeter survey', 'status':status,
            'distance':distance,'second':None,'comment':'',
            '_center_pile':pile_key,'_center_session':batch}


def update_center_design(record, standards):
    """Attach an exact-ID standard after a center was approved without standards loaded."""
    name=record.get('_center_pile')
    if not name:
        return
    design=next((p for p in standards if p['id'].strip().upper()==name),None)
    record['design']=design
    record['distance']=(math.hypot(record['field']['north']-design['north'],
                                   record['field']['east']-design['east']) if design else None)
    if record.get('_direct_center'):
        if not design:
            record['status']='CENTER/HS - STANDARD NOT FOUND'
        elif record.get('status')=='CENTER/HS - REVIEW HIGH SPOT':
            record['status']='CENTER/HS - REVIEW HIGH SPOT'
        else:
            record['status']='CENTER/HS - MATCHED'
    else:
        record['status']=('CENTER - STANDARD NOT FOUND' if not design else
                          'CENTER - QC REVIEWED' if record['status']=='CENTER - QC REVIEWED' else 'CENTER - MATCHED')


def combine_records(old, newer):
    # Also clean duplicates already present in an imported historical report.
    return deduplicate_records([*old, *newer])


def report_pile_id(record):
    """Return the design pile ID, not the field-point number, for revisions."""
    pile=record.get('design')
    text=(pile.get('id') if pile else record.get('_center_pile')) or ''
    return str(text).strip().upper() or None


def current_pile_records(records):
    """Show one latest dated as-built per pile; retain all other shots in history.

    Older survey results remain intact, and a newly imported earlier-date survey
    cannot supersede a later date. Live CSV/center results win same-day ties over
    prior-report data. Unmatched field points without a pile ID are preserved.
    """
    winners={}
    for idx,record in enumerate(records):
        pile=report_pile_id(record)
        if not pile:
            continue
        date=str(record.get('shot_date') or '')
        valid_date=bool(re.fullmatch(r'\d{4}-\d{2}-\d{2}',date))
        rank=(int(valid_date),date if valid_date else '',
              int(bool(record.get('_live_import'))),idx)
        if pile not in winners or rank>winners[pile][0]:
            winners[pile]=(rank,idx)
    keep={index for _,index in winners.values()}
    # Drill information belongs to the pile, not to the particular survey date.
    # A later as-built shot shouldn't erase values entered in an earlier report.
    # However, metadata explicitly entered for the latest shot has priority.
    for pile,(_,winner_idx) in winners.items():
        winner=records[winner_idx]
        candidates=sorted(
            ((str(r.get('shot_date') or ''),idx,r) for idx,r in enumerate(records)
             if report_pile_id(r)==pile), reverse=True)
        for field in ('drill_date','rig_number'):
            if not str(winner.get(field) or '').strip():
                for _,_,r in candidates:
                    if str(r.get(field) or '').strip():
                        winner[field]=r[field]
                        break
    return [record for idx,record in enumerate(records)
            if report_pile_id(record) is None or idx in keep]


# Pile coverage is a first-class part of the running report.  A recognized
# pile must NEVER disappear solely because its perimeter cannot be solved.
# Pending rows have no as-built coordinates: they are not fabricated centers.
def is_pending_record(record):
    return str(record.get('status', '')).startswith('PENDING -')


def described_pile_id(description, pattern=None):
    tagged=parse_tagged_description(description)
    if tagged:
        return tagged[0]
    return normalize_description(description, pattern or CenterEngine().pattern)


def pending_record(pile_id, observations, batch, standards, reason):
    """Represent an observed pile needing work without inventing survey data."""
    design=next((p for p in standards if str(p['id']).strip().upper()==pile_id), None)
    numbers=list(dict.fromkeys(str(p['point']).strip() for p in observations))
    descriptions=list(dict.fromkeys(str(p.get('desc','')).strip() for p in observations))
    return {
        'field':{'point':'PENDING-'+pile_id, 'north':None,'east':None,
                 'elev':None, 'desc':('Source points: '+', '.join(numbers)
                    + ' | Descriptions: '+', '.join(descriptions)), 'line':0},
        'design':design, 'batch':batch.get('batch',''),
        'shot_date':batch.get('shot_date',''), 'shot_job':batch.get('shot_job',''),
        'status':'PENDING - '+reason, 'distance':None,'second':None,
        '_center_pile':pile_id, 'comment':'',
    }


def add_pending_coverage(records, observations, batch, standards, pattern=None):
    """Store every recognized pile from an imported file, even if not matched.

    Previously reported identical source shots are not made pending again.
    Later shots for the same pile/date do re-open review, rather than quietly
    leaving a stale current measurement in place.
    """
    grouped=defaultdict(list)
    for p in observations:
        pile=described_pile_id(p.get('desc'),pattern)
        if pile:grouped[pile].append(p)
    updated=list(records)
    for pile,shots in grouped.items():
        date=batch.get('shot_date','')
        def already_reported(record):
            if (is_pending_record(record) or report_pile_id(record)!=pile or
                record.get('shot_date')!=date):
                return False
            field=record['field']
            # Direct C/HS pairs preserve their center point's raw XY.
            if any(str(field.get('point',''))==str(p['point']) and
                   field.get('north')==p.get('north') and
                   field.get('east')==p.get('east') for p in shots):
                return True
            # A derived polygon center stores the contributing field point IDs
            # in its description. Reimporting those SAME raw shots must not
            # create a false missing-center warning. New shot IDs reopen review.
            if str(field.get('point','')).upper().startswith('CENTER-'):
                source_text=str(field.get('desc',''))
                match=re.search(r'source points:\s*(.*)$',source_text,re.I)
                if match:
                    existing_ids={t.strip() for t in match.group(1).split(',')}
                    if {str(p['point']).strip() for p in shots}<=existing_ids:
                        return True
            return False
        if any(already_reported(r) for r in updated):
            continue
        tagged=[parse_tagged_description(p.get('desc')) for p in shots]
        has_center=any(t and t[1]=='C' for t in tagged)
        has_hs=any(t and t[1]=='HS' for t in tagged)
        if has_center or has_hs:
            reason='C/HS PAIR OR SHOT NEEDS MATCHING'
        elif len({(p['north'],p['east']) for p in shots})>=3:
            reason='PERIMETER CENTER NOT APPROVED'
        elif len(shots)==2:
            reason='TWO SHOTS; CENTER NOT ESTABLISHED'
        else:
            reason='POINT NOT MATCHED'
        # Merge sources, don't accumulate pending records for the same pile/date.
        earlier=[r for r in updated if is_pending_record(r)
                 and report_pile_id(r)==pile and r.get('shot_date')==date]
        for r in earlier:
            # Source-point lists are informational. The new observation set is
            # also retained independently by the center/tagged-shot engines.
            for point in str(r['field'].get('desc','')).split(' | ')[0].replace('Source points: ','').split(', '):
                if point and point not in [str(x['point']) for x in shots]:
                    shots.append({'point':point,'desc':'previously imported',
                                  'north':None,'east':None,'elev':None})
        updated=[r for r in updated if r not in earlier]
        replacement=pending_record(pile,shots,batch,standards,reason)
        for old in earlier:
            if old.get('comment'):replacement['comment']=old['comment']
        replacement['_live_import']=True
        updated.append(replacement)
    return updated


def cover_piles_with_results(records, newly):
    """Replace placeholders only for piles actually resolved by matching."""
    resolved={(report_pile_id(r),str(r.get('shot_date','')))
              for r in newly if report_pile_id(r) and not is_pending_record(r)}
    previous=[r for r in records if not (
        is_pending_record(r) and
        (report_pile_id(r),str(r.get('shot_date',''))) in resolved)]
    return combine_records(previous,newly)


def missing_pile_coverage(records, observations, shot_date, pattern=None):
    """Exact detected-ID coverage assertion for every imported CSV."""
    detected={described_pile_id(p.get('desc'),pattern) for p in observations}
    detected.discard(None)
    represented={report_pile_id(r) for r in records
                 if r.get('shot_date')==shot_date}
    return sorted(detected-represented)


def match_explicit_pile_ids(points, standards):
    """Never silently assign a B057-labelled shot to another design ID.

    Exact pile IDs win; unlabelled shots retain existing global spatial matching.
    """
    by_id={str(p['id']).strip().upper():p for p in standards}
    explicit=[]; unlabeled=[]
    for f in points:
        name=described_pile_id(f.get('desc'))
        if name:
            p=by_id.get(name)
            if p is None:
                explicit.append({'field':f,'design':None,'_center_pile':name,
                                 'status':'REVIEW - STANDARD NOT FOUND',
                                 'distance':None,'second':None})
            else:
                distance=math.hypot(f['north']-p['north'],f['east']-p['east'])
                explicit.append({'field':f,'design':p,'status':'MATCHED',
                                 'distance':distance,'second':None})
        else:
            unlabeled.append(f)
    reserved={r['design']['id'] for r in explicit if r['design']}
    available=[p for p in standards if p['id'] not in reserved]
    return explicit+(match_piles(unlabeled,available) if unlabeled else [])

def xlsx_sheet_rows(z, name):
    strings=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        root=ET.fromstring(z.read('xl/sharedStrings.xml'))
        strings=[''.join(t.text or '' for t in si.iter(NS+'t')) for si in root.findall(NS+'si')]
    root=ET.fromstring(z.read(name))
    result=[]
    for row in root.iter(NS+'row'):
        result.append({colnum(c.get('r'))-1:value(c,strings) for c in row.findall(NS+'c')})
    return result

def load_previous_report(path):
    """Load cumulative PIVOT history; migrate old 2-sheet reports when possible."""
    with zipfile.ZipFile(path) as z:
        sheets=z.namelist()
        if 'xl/worksheets/sheet3.xml' in sheets:
            rows=xlsx_sheet_rows(z,'xl/worksheets/sheet3.xml')
            if rows and rows[0].get(0)=='PIVOT_HISTORY_V1':
                results=[]
                def number(x):return float(x) if x not in ('',None) else None
                for row in rows[1:]:
                    if not row.get(0):continue
                    f={'point':str(row.get(3,'')),'north':number(row.get(4)),'east':number(row.get(5)),
                       'elev':number(row.get(6)),'desc':str(row.get(7,'')),'line':0}
                    if any(f[k] is None for k in ('north','east')) and not str(row.get(12,'')).startswith('PENDING -'):continue
                    p=None
                    if row.get(8):
                        p={'id':str(row[8]),'north':number(row.get(9)),'east':number(row.get(10)),
                           'elev':number(row.get(11)),'source':'Prior PIVOT report'}
                    record={'batch':str(row.get(0,'')),'shot_date':str(row.get(1,'')),
                        'shot_job':str(row.get(2,'')),'field':f,'design':p,
                        'status':str(row.get(12,'')),'distance':number(row.get(13)),'second':None,
                        'comment':str(row.get(14,'') or ''),
                        'drill_date':normalize_drill_date(row.get(17,'')),
                        'rig_number':str(row.get(18,'') or '')}
                    if f['point'].startswith('CENTER-'):
                        record['_center_pile']=f['point'][7:]
                    # Optional backward-compatible extension to the hidden history.
                    # Enables newly imported C/HS pairs to retain their design pile ID
                    # even if standards are loaded only after re-import.
                    if row.get(15):
                        record['_center_pile']=str(row[15])
                    if str(record['status']).startswith('CENTER/HS'):
                        record['_direct_center']=True
                    results.append(record)
                # The visible Comments column can be edited after export.
                # V2 maps each printed row to its full-history record so a
                # current-only printable report can still retain older shots.
                # Older PIVOT workbooks used one printed row per history row.
                report_rows=xlsx_sheet_rows(z,'xl/worksheets/sheet1.xml')
                mapped=rows[0].get(16)=='VISIBLE_ROW_INDEX_V2'
                # Previous formats used the old 19-column sheet. New exports
                # identify the shifted fields by their header labels, rather than
                # assuming the same Excel column positions forever.
                headers=report_rows[5] if len(report_rows)>5 else {}
                has_drill_fields=(str(headers.get(6,'')).strip().lower()=='drill date'
                                  and str(headers.get(7,'')).strip().lower()=='rig number')
                comment_col=next((i for i,h in headers.items() if str(h).strip().lower()=='comments'),
                                 20 if has_drill_fields else 18)
                for idx, record in enumerate(results):
                    visible_index=(int(float(rows[idx+1].get(16,-1)))
                                   if mapped else idx)
                    if visible_index<0:
                        continue
                    report_idx=6+visible_index
                    if report_idx<len(report_rows):
                        visible=report_rows[report_idx]
                        record['comment']=str(visible.get(comment_col,'') or '')
                        if has_drill_fields:
                            # Sheet 1 contains editable values. They override
                            # the hidden history, including intentional clears.
                            record['drill_date']=normalize_drill_date(visible.get(6,''))
                            record['rig_number']=str(visible.get(7,'') or '')
                # Clean pre-fix exported centers after restoring visible comments.
                return deduplicate_records(results), False
        # Legacy PIVOT reports were undated. Preserve them, explicitly marking missing shot dates.
        if 'xl/worksheets/sheet1.xml' not in sheets or 'xl/worksheets/sheet2.xml' not in sheets:
            raise ValueError('Not a PIVOT report: missing the two report worksheets.')
        design_rows=xlsx_sheet_rows(z,'xl/worksheets/sheet1.xml')
        verification_rows=xlsx_sheet_rows(z,'xl/worksheets/sheet2.xml')
        out=[]
        for n,vr in enumerate(verification_rows[5:]):
            ix=n+6
            if ix>=len(design_rows):break
            dr=design_rows[ix]
            def number(x):
                try:return float(str(x).replace(',',''))
                except (ValueError,TypeError):return None
            pn=str(vr.get(1,'') or dr.get(0,'')).strip()
            fn=str(vr.get(0,'')).strip()
            if not pn and not fn:continue
            f={'point':fn,'north':number(dr.get(6)),'east':number(dr.get(7)),
               'elev':number(dr.get(8)),'desc':str(vr.get(2,'')),'line':0}
            if any(f[k] is None for k in ('north','east')):continue
            p=None
            if pn:
                p={'id':pn,'north':number(dr.get(1)),'east':number(dr.get(2)),
                   'elev':number(dr.get(3)),'source':'Previous PIVOT report'}
            out.append({'batch':'legacy:'+Path(path).stem,'shot_date':'Unknown (legacy)',
                        'shot_job':Path(path).stem,'field':f,'design':p,
                        'status':str(vr.get(3,'MATCHED')),'distance':number(vr.get(4)),'second':None,
                        'comment':str(dr.get(18,'') or ''),
                        'drill_date':'','rig_number':''})
        if not out:raise ValueError('The previous report contains no readable matching records.')
        return deduplicate_records(out), True

HEADERS=['Pile #','Northing','Easting','Elevation','','Pile #','Northing','Easting','Elevation','Date Shot','Pile #','Δ Northing','N / S','Δ Easting','E / W','Δ Elevation','High / Low','Radial Offset','', 'Field Point','Description','Match Status','Match Distance']

def refresh_pending_group_records(records, raw_rows, standards):
    """Keep every dated multi-shot pile in review instead of matching one vertex.

    This is intentionally conservative: an incomplete perimeter is never a
    surveyed center. Matching is resumed by explicitly approving a centroid.
    """
    out=list(records)
    for (date,pile),shots in group_current_raw(raw_rows).items():
        if len(shots)<2:
            continue
        source_points={str(s.get('point','')).strip() for s in shots}
        def covers_shots(record):
            if (report_pile_id(record)!=pile or str(record.get('shot_date',''))!=date
                    or is_pending_record(record)):
                return False
            field=record['field']
            if not str(field.get('point','')).upper().startswith('CENTER-'):
                return False
            match=re.search(r'source points:\s*(.*)$',str(field.get('desc','')),re.I)
            if not match:
                return False
            ids={v.strip() for v in match.group(1).split(',')}
            return source_points<=ids
        if any(covers_shots(r) for r in out):
            continue
        already=[r for r in out if is_pending_record(r)
                 and report_pile_id(r)==pile and str(r.get('shot_date',''))==date]
        for prev in already:out.remove(prev)
        source=shots[-1]
        batch={'batch':source.get('shot_job',''), 'shot_date':date,
               'shot_job':source.get('shot_job','')}
        note=('MULTIPLE PERIMETER SHOTS — CENTER APPROVAL REQUIRED'
              if len(shots)>=3 else 'TWO SHOTS — CENTER NOT ESTABLISHED')
        pending=pending_record(pile,shots,batch,standards,note)
        pending['_live_import']=True
        pending['comment']=next((r.get('comment','') for r in already if r.get('comment')), '')
        out.append(pending)
    return deduplicate_records(out)


def load_report_sources(path):
    """Read optional hidden raw-shot ledger from recent PIVOT reports."""
    with zipfile.ZipFile(path) as z:
        if 'xl/worksheets/sheet4.xml' not in z.namelist():
            return []
        rows=xlsx_sheet_rows(z,'xl/worksheets/sheet4.xml')
        if not rows or rows[0].get(0)!='PIVOT_RAW_V1':
            return []
        out=[]
        for row in rows[1:]:
            try:
                point=str(row.get(2,'')).strip()
                if not point:continue
                data={'shot_date':str(row.get(0,'') or ''),
                      'shot_job':str(row.get(1,'') or ''),
                      'point':point, 'north':float(row[3]),'east':float(row[4]),
                      'elev':float(row[5]),'desc':str(row.get(6,'') or ''),
                      'source':str(row.get(7,'') or ''),
                      '_sources':set(str(row.get(8,'') or '').split(' | '))}
                out.append(data)
            except (ValueError,KeyError,TypeError):
                continue
        return merge_raw([],out)


def group_current_raw(raw_rows):
    """Group independent dated pile observations; never merge survey dates."""
    groups=defaultdict(list)
    for row in raw_rows:
        name=described_pile_id(row.get('desc'))
        if name and not parse_tagged_description(row.get('desc')):
            groups[(str(row.get('shot_date','')),name)].append(row)
    return groups


def export_csv(path,results):
    with open(path,'w',newline='',encoding='utf-8-sig') as out:
        w=csv.writer(out);w.writerow(HEADERS)
        for r in results:
            f=r['field'];p=r['design']
            if p is None:
                w.writerow(['']*19+[f['point'],f['desc'],r['status'],'']);continue
            dn=round(f['north']-p['north'],3);de=round(f['east']-p['east'],3)
            dz=round(f['elev']-p['elev'],3) if p['elev'] is not None and f['elev'] is not None else None
            offset=round(math.hypot(dn,de),3)
            w.writerow([p['id'],fmt(p['north']),fmt(p['east']),fmt(p['elev']),'',p['id'],fmt(f['north']),fmt(f['east']),fmt(f['elev']),'',p['id'],fmt(dn),direction(dn,'North','South'),fmt(de),direction(de,'East','West'),fmt(dz),direction(dz,'High','Low') if dz is not None else '',fmt(offset),'',f['point'],f['desc'],r['status'],fmt(r['distance'])])


def editable_report_comment(record):
    """Keep report Comments for user-written notes, not software QC notices.

    Previous PIVOT exports sometimes inserted a literal automatic 'REVIEW
    REQUIRED: <status>' into Comments for pending piles. Migrate that exact
    generated placeholder to blank without discarding actual user annotations.
    Review/status still appear on the Match Verification sheet.
    """
    value=str(record.get('comment') or '')
    if is_pending_record(record) and value.strip()==('REVIEW REQUIRED: '+str(record.get('status',''))):
        return ''
    return value


def display_shot_date(value):
    # Store ISO dates in PIVOT history; show survey dates as MM-DD-YYYY.
    if value and re.fullmatch(r'\d{4}-\d{2}-\d{2}', str(value)):
        return datetime.strptime(str(value), '%Y-%m-%d').strftime('%m-%d-%Y')
    return str(value or '')

def export_excel(path, results, history_records=None, raw_observations=None):
    """Create a standards-compliant Excel workbook with XlsxWriter (bundled)."""
    import xlsxwriter
    workbook = xlsxwriter.Workbook(path, {'strings_to_urls': False, 'nan_inf_to_errors': True})
    if history_records is None:
        history_records=results
    # Normalize existing pre-fix history too. This is a non-destructive report
    # cleanup: measurements from different dates or source shots are retained.
    history_records=deduplicate_records(history_records)
    canonical={record_identity(record):record for record in history_records}
    results=deduplicate_records([canonical.get(record_identity(r),r) for r in results])
    try:
        navy='#17365D';blue='#0B4D91';red='#C8102E';white='#FFFFFF'
        def style(**kwargs): return workbook.add_format(kwargs)
        title=style(bold=True,font_color=white,bg_color=navy,font_size=16,align='left',valign='vcenter')
        subtitle=style(bold=True,font_color=navy,font_size=11,valign='vcenter')
        info=style(font_color='#596779',font_size=9,valign='vcenter')
        caution=style(bold=True,font_color=red,font_size=9,align='left')
        group=style(bold=True,font_color=white,bg_color=navy,font_size=12,align='center',valign='vcenter')
        header=style(bold=True,font_color=navy,bg_color='#DFE9F6',align='center',valign='vcenter',border=1,border_color='#BCC8D8',text_wrap=True)
        fieldheader=style(bold=True,font_color=navy,bg_color='#FFF2CC',align='center',valign='vcenter',border=1,border_color='#BCC8D8',text_wrap=True)
        editable_header=style(bold=True,font_color=navy,bg_color='#DDF0E1',align='center',valign='vcenter',border=1,border_color='#B6CCBA',text_wrap=True)
        design_text=style(font_color=blue,bg_color='#FFF2CC',border=1,border_color='#CCD1D8',align='center')
        design_num=style(font_color=blue,bg_color='#FFF2CC',border=1,border_color='#CCD1D8',num_format='#,##0.000')
        asbuilt_text=style(font_color=blue,bg_color='#EFF5FF',border=1,border_color='#CCD1D8',align='center')
        asbuilt_num=style(font_color=blue,bg_color='#EFF5FF',border=1,border_color='#CCD1D8',num_format='#,##0.000')
        normal_text=style(border=1,border_color='#CCD1D8',align='center')
        normal_num=style(border=1,border_color='#CCD1D8',num_format='#,##0.000')
        signed=style(border=1,border_color='#CCD1D8',num_format='+0.000;-0.000;0.000')
        # Direction labels use neutral text; red is reserved for out-of-range numerical deviations.
        negative_text=style(border=1,border_color='#CCD1D8',align='center')
        warning=style(font_color=red,bg_color='#FFF0F0',border=1,border_color='#CCD1D8',align='center',bold=True)
        comment_input=style(font_color='#17365D',bg_color='#F0F8F1',border=1,border_color='#B6CCBA',align='left',valign='vcenter',text_wrap=True,locked=False)
        drill_input=style(font_color='#17365D',bg_color='#F0F8F1',border=1,border_color='#B6CCBA',align='center',valign='vcenter',num_format='mm/dd/yyyy',locked=False)
        rig_input=style(font_color='#17365D',bg_color='#F0F8F1',border=1,border_color='#B6CCBA',align='center',valign='vcenter',num_format='@',locked=False)

        # Sheet 1: main printable deviation report only.
        sheet = workbook.add_worksheet('Pile Deviation Report')
        sheet.hide_gridlines(2)
        sheet.freeze_panes(6, 0)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
        sheet.set_margins(0.25, 0.25, 0.45, 0.45)
        sheet.set_row(0,34);sheet.set_row(1,28);sheet.set_row(2,28)
        sheet.set_row(3,18);sheet.set_row(4,26);sheet.set_row(5,32)
        widths=[12,19,20,14,2.5,12,15,13,19,20,14,16,12,14,11,14,11,14,11,30]
        for i,w in enumerate(widths): sheet.set_column(i,i,w)
        sheet.merge_range('A1:C3','',style(bg_color='#FFFFFF'))
        sheet.insert_image('A1',asset_path('PIVOT_Report_Logo.png'),{
            'x_scale':0.35,'y_scale':0.35,'x_offset':105,'y_offset':4,
            'object_position':1,'description':'Southern Survey logo', 'decorative':False})
        sheet.merge_range('D1:T1','PILE DEVIATION REPORT',title)
        sheet.merge_range('D2:T2','Southern Survey  |  Lafayette, LA',subtitle)
        sheet.merge_range('D3:T3','Generated: '+datetime.now().strftime('%B %d, %Y at %I:%M %p')+'  |  Dates shot: '+', '.join(sorted({str(r.get('shot_date','')) for r in results if r.get('shot_date')})),info)
        pending_count=sum(is_pending_record(r) for r in results)
        sheet.merge_range('A4:T4',
            f'Prepared with PIVOT  |  A Faulk Product  |  {len(results)} piles  |  {pending_count} pending review  |  Light green cells can be edited',info)
        for rng,text_ in [('A5:D5','DESIGN LOCATION'),('F5:L5','AS-BUILT LOCATION'),('M5:S5','DEVIATION (As-Built - Design)')]:
            sheet.merge_range(rng,text_,group)
        sheet.write(4,19,'COMMENTS',group)
        headers=['Pile #','Northing','Easting','Elevation','','Pile #','Drill Date','Rig Number','Northing','Easting','Elevation','Date Shot','Pile #','Δ Northing','N / S','Δ Easting','E / W','Δ Elevation','High / Low','Comments']
        for i,h in enumerate(headers):
            if i == 4:
                continue
            sheet.write(5,i,h,editable_header if i in (6,7,19) else
                        fieldheader if i in (0,1,2,3,5,8,9,10,11) else header)
        for index,record in enumerate(results):
            row=6+index; field=record['field']; pile=record['design']
            if pile and not is_pending_record(record):
                dn=round(field['north']-pile['north'],3)
                de=round(field['east']-pile['east'],3)
                dz=round(field['elev']-pile['elev'],3) if pile['elev'] is not None and field['elev'] is not None else None
                values=[pile['id'],pile['north'],pile['east'],pile['elev'],None,pile['id'],record.get('drill_date',''),record.get('rig_number',''),field['north'],field['east'],field['elev'],display_shot_date(record.get('shot_date','')),pile['id'],dn,direction(dn,'North','South'),de,direction(de,'East','West'),dz,direction(dz,'High','Low') if dz is not None else '',editable_report_comment(record)]
            else:
                values=[None]*20
                pile_name=report_pile_id(record)
                if pile:
                    values[0:4]=[pile['id'],pile['north'],pile['east'],pile['elev']]
                if pile_name:
                    if not pile:values[0]=pile_name
                    values[5]=pile_name;values[12]=pile_name
                if not is_pending_record(record):
                    values[8]=field['north'];values[9]=field['east'];values[10]=field['elev']
                # The Comments column belongs to the survey team. Do not
                # prefill it with automated pending/review text; that belongs
                # on the Match Verification sheet instead.
                values[19]=editable_report_comment(record)
                values[11]=display_shot_date(record.get('shot_date',''))
                values[6]=record.get('drill_date','')
                values[7]=record.get('rig_number','')
            for i,value_ in enumerate(values):
                if i in (4,6,7,19) or value_ is None:
                    continue
                if i<=3:
                    fmt_=design_num if isinstance(value_,(float,int)) else design_text
                elif i<=11:
                    fmt_=asbuilt_num if isinstance(value_,(float,int)) else asbuilt_text
                elif i in (13,15,17):
                    fmt_=signed
                elif i in (14,16,18) and value_ in ('South','West','Low'):
                    fmt_=negative_text
                elif isinstance(value_,(float,int)):
                    fmt_=normal_num
                else:
                    fmt_=normal_text
                sheet.write(row,i,value_,fmt_)
            # Drill Date, Rig Number and Comments can be edited by Southern Survey.
            # All other cells (including coordinates/deviations) stay locked.
            sheet.write(row,6,normalize_drill_date(values[6]),drill_input)
            sheet.write(row,7,str(values[7] or ''),rig_input)
            sheet.write(row,19,values[19] or '',comment_input)
            sheet.set_row(row,21)
        # Southern Survey tolerances: three inches = 0.250 feet PER AXIS.
        # Horizontal offset: if |Delta Northing| or |Delta Easting| exceeds
        # 0.250 ft, highlight all PROTECTED report-data cells in that row.
        # The 3 editable columns G (Drill Date), H (Rig Number), and T
        # (Comments) stay pale green so users can always identify inputs.
        # Vertical-only offset: if |Delta Elevation| exceeds 0.250 ft,
        # highlight BOTH the deviation value (R) and High / Low label (S).
        # An exact +/-0.250 is in tolerance. Radial offset is informational
        # and appears only on Match Verification (not used for highlighting).
        outside_tolerance=style(font_color='#B00020',bg_color='#FFD9DD',bold=True)
        last_row=max(6,5+len(results))
        horizontal_outside='=OR(AND(ISNUMBER($N7),ABS($N7)>0.25),AND(ISNUMBER($P7),ABS($P7)>0.25))'
        vertical_outside='=AND(ISNUMBER($R7),ABS($R7)>0.25)'
        # A:F and I:S together cover all non-editable report-data columns.
        for first_col,last_col in ((0,5),(8,18)):
            sheet.conditional_format(6,first_col,last_row,last_col,{
                'type':'formula','criteria':horizontal_outside,
                'format':outside_tolerance,
            })
        # The High / Low direction is tied to the same vertical deviation,
        # so it receives the identical warning without coloring the whole row.
        sheet.conditional_format(6,17,last_row,18,{
            'type':'formula','criteria':vertical_outside,
            'format':outside_tolerance,
        })
        # Filter dropdowns on every header, including design/as-built/deviation Pile #.
        # Filtering rows is allowed on a protected Excel sheet; sorting locked
        # report values is constrained by Excel's worksheet protection behavior.
        sheet.autofilter(5,0,max(6,5+len(results)),19)
        sheet.print_area(0,0,max(6,5+len(results)),19)
        sheet.repeat_rows(4,5)
        sheet.set_header('&CSouthern Survey | Pile Deviation Report')
        sheet.set_footer('&LSouthern Survey&CPIVOT - A Faulk Product&RPage &P of &N')
        sheet.protect('PIVOT2026', {
            'select_locked_cells': True, 'select_unlocked_cells': True,
            'format_cells': False, 'format_columns': False, 'format_rows': False,
            'insert_columns': False, 'insert_rows': False, 'delete_columns': False,
            'delete_rows': False, 'sort': True, 'autofilter': True,
        })

        # Sheet 2: match verification only.
        verify = workbook.add_worksheet('Match Verification')
        verify.hide_gridlines(2)
        verify.freeze_panes(5,0)
        verify.set_landscape()
        verify.set_margins(0.25,0.25,0.45,0.45)
        for r,h in enumerate([22,20,18,24,28]): verify.set_row(r,h)
        verify_widths=[14,14,16,24,16,16,22,24]
        for i,w in enumerate(verify_widths): verify.set_column(i,i,w)
        verify.merge_range('A1:A3','',style(bg_color='#FFFFFF'))
        # The icon (not the wide wordmark) fits within the A1:A3 header area.
        verify.insert_image('A1',asset_path('SouthernSurvey_Compact.png'),{
            'x_scale':0.65,'y_scale':0.65,'x_offset':17,'y_offset':3,
            'object_position':1,'description':'Southern Survey logo', 'decorative':False})
        verify.merge_range('B1:H1','MATCH VERIFICATION',title)
        verify.merge_range('B2:H2','Southern Survey  |  Lafayette, LA',subtitle)
        verify.merge_range('B3:H3','Generated: '+datetime.now().strftime('%B %d, %Y at %I:%M %p'),info)
        verify.merge_range('A4:H4','Prepared with PIVOT  |  A Faulk Product  |  Print sheet 1 for deviation results.',info)
        verify_headers=['Field Point','Pile ID','Date Shot','Source Job','Match Distance','Radial Offset','Match Status','Description']
        for i,h in enumerate(verify_headers): verify.write(4,i,h,header)
        for index,record in enumerate(results):
            row=5+index; field=record['field']; pile=record['design']; flag=not is_match_without_review(record)
            radial=(math.hypot(field['north']-pile['north'],field['east']-pile['east'])
                    if pile and not is_pending_record(record)
                    and field.get('north') is not None and field.get('east') is not None
                    else None)
            values=[field['point'], pile['id'] if pile else record.get('_center_pile',''), record.get('shot_date',''),record.get('shot_job',''), record['distance'],round(radial,3) if radial is not None else None,record['status'],field['desc']]
            for i,value_ in enumerate(values):
                if value_ is None:
                    continue
                if i in (4,5) and isinstance(value_,(float,int)):
                    fmt_=normal_num
                elif i==6 and flag:
                    fmt_=warning
                else:
                    fmt_=normal_text
                verify.write(row,i,value_,fmt_)
            verify.set_row(row,21)
        verify.print_area(0,0,max(5,4+len(results)),7)
        verify.repeat_rows(4,4)
        verify.set_header('&CSouthern Survey | Match Verification')
        verify.set_footer('&LSouthern Survey&CPIVOT - A Faulk Product&RPage &P of &N')
        verify.protect('PIVOT2026', {
            'select_locked_cells': True, 'select_unlocked_cells': True,
            'format_cells': False, 'format_columns': False, 'format_rows': False,
            'insert_columns': False, 'insert_rows': False, 'delete_columns': False,
            'delete_rows': False, 'sort': False, 'autofilter': False,
        })

        # Re-importable machine-readable history: unlike a print sheet it survives
        # multiple survey batches, preserves all as-built coordinates and dates.
        history=workbook.add_worksheet('_PIVOT_HISTORY')
        history.hide()
        history.write(0,0,'PIVOT_HISTORY_V1')
        history.write(0,16,'VISIBLE_ROW_INDEX_V2')
        history.write(0,17,'DRILL_DATE_V1')
        history.write(0,18,'RIG_NUMBER_V1')
        # Exact Python object identity is safe here: sorted_results() returns the
        # same dictionaries contained in self.results, not copied dictionaries.
        visible_position={id(record):index for index,record in enumerate(results)}
        for idx,record in enumerate(history_records,1):
            f=record['field'];p=record['design']
            vals=[record.get('batch',''),record.get('shot_date',''),record.get('shot_job',''),
                f['point'],f['north'],f['east'],f['elev'],f['desc'],
                p['id'] if p else '',p['north'] if p else None,p['east'] if p else None,
                p['elev'] if p else None,record['status'],record['distance'],editable_report_comment(record),record.get('_center_pile','')]
            for col,v in enumerate(vals):
                if v is not None:history.write(idx,col,v)
            history.write_number(idx,16,visible_position.get(id(record),-1))
            if record.get('drill_date'):
                history.write(idx,17,normalize_drill_date(record['drill_date']))
            if record.get('rig_number'):
                history.write(idx,18,str(record['rig_number']))

        # Hidden provenance ledger survives report re-import; it is not a CSV
        # or a third visible report tab. Copies of the same CSV row are unique.
        if raw_observations:
            sources=workbook.add_worksheet('_PIVOT_SOURCES')
            sources.hide()
            sources.write(0,0,'PIVOT_RAW_V1')
            for i,shot in enumerate(merge_raw([],raw_observations),1):
                vals=[shot.get('shot_date',''),shot.get('shot_job',''),
                      str(shot.get('point','')),shot.get('north'),
                      shot.get('east'),shot.get('elev'),shot.get('desc',''),
                      shot.get('source',''),
                      ' | '.join(sorted(shot.get('_sources') or []))]
                for col,val in enumerate(vals):
                    if val is not None:sources.write(i,col,val)

        workbook.set_properties({'title':'PIVOT Pile Deviation Report',
                                 'subject':'As-built versus design pile deviation report',
                                 'author':'Faulk Products', 'company':'Faulk Products'})
    finally:
        workbook.close()


def classify_pile_shots(points, engine):
    """Classify newly imported pile-coded observations without altering the engine.

    One observation is NOT enough for a polygon centroid. Instead, a single
    pile-coded shot is matched normally and checked against the 0.250-foot
    deviation threshold, without a separate confirmation popup. Two shots
    also cannot form a polygon, so neither case should trigger center reminders.
    Three or more shots for a pile remain eligible for center review.
    """
    groups=engine.groups()
    existing=set(engine._seen)
    counts={pile:len(shots) for pile,shots in groups.items()}
    new_by_id={}
    pile_coded=[]
    for point in points:
        pile=normalize_description(point.get('desc'),engine.pattern)
        if not pile:
            continue
        pile_coded.append(point)
        key=(pile,str(point.get('point','')).strip(),point.get('north'),
             point.get('east'),point.get('elev'))
        if key not in existing:
            new_by_id.setdefault(pile,{})[key]=point
    for pile,items in new_by_id.items():
        counts[pile]=counts.get(pile,0)+len(items)
    singles=[next(iter(items.values())) for pile,items in new_by_id.items()
             if counts[pile]==1]
    perimeter={pile:count for pile,count in counts.items()
               if pile in new_by_id and count>=3}
    insufficient={pile:count for pile,count in counts.items()
                  if pile in new_by_id and count==2}
    return pile_coded, singles, perimeter, insufficient


def field_shot_key(point):
    return (str(point.get('point','')).strip(),point.get('north'),
            point.get('east'),point.get('elev'),str(point.get('desc','')).strip())


# Southern Survey tolerance is 3 inches (0.250 ft) on EACH axis:
# Northing, Easting, and Elevation evaluated separately. A diagonal horizontal
# radial offset may exceed 0.250 even when both individual axes comply.
# These QC thresholds do not prove a single exterior shot is a true center.
SINGLE_SHOT_TOLERANCE = 0.250
SINGLE_SHOT_ACCEPTED = 'SINGLE SHOT - ASSUMED CENTER'


def classify_single_shot_match(record, pattern, tolerance=SINGLE_SHOT_TOLERANCE):
    """Evaluate one observed pile shot relative to its proposed design pile.

    A shot within the N, E, and Z limits is permitted to stand in for
    the design center *for this report*; measured N/E/Z must never be replaced
    with the standard values. Missing design elevations and ambiguous matches
    remain reviewable instead of being auto-approved.
    """
    design = record.get('design')
    if not design:
        return 'UNMATCHED - SINGLE SHOT'
    observed = record['field']
    pile = normalize_description(observed.get('desc',''), pattern)
    if pile and pile != str(design['id']).strip().upper():
        return 'SINGLE SHOT - REVIEW PILE ID'
    if record.get('status') != 'MATCHED':
        return 'SINGLE SHOT - REVIEW MATCH'
    dn=observed['north']-design['north']
    de=observed['east']-design['east']
    if not (math.isfinite(dn) and math.isfinite(de)):
        return 'SINGLE SHOT - REVIEW COORDINATES'
    if abs(dn)>tolerance+1e-9:
        return 'SINGLE SHOT - REVIEW NORTHING'
    if abs(de)>tolerance+1e-9:
        return 'SINGLE SHOT - REVIEW EASTING'
    if observed.get('elev') is None or design.get('elev') is None:
        return 'SINGLE SHOT - REVIEW ELEVATION'
    if abs(observed['elev'] - design['elev']) > tolerance + 1e-9:
        return 'SINGLE SHOT - REVIEW ELEVATION'
    return SINGLE_SHOT_ACCEPTED


def is_match_without_review(record):
    return record.get('status') in ('MATCHED','CENTER - MATCHED',
                                   'CENTER/HS - MATCHED',SINGLE_SHOT_ACCEPTED)


def center_candidate_queue(engine, ignored_groups, results, center_batch):
    """Return perimeter groups and their center-approval state.

    Single- and two-shot groups remain visible in the optional review queue,
    but are NOT outstanding center calculations. A polygon requires three
    distinct perimeter observations. Single shots are offered for ordinary
    coordinate matching with an explicit review warning instead.
    The ignored-state signature changes if new observations arrive,
    causing a formerly ignored group to be reviewed again. Neither acceptance
    nor an ignored group modifies the source observations.
    """
    approved_in_main={r.get('_center_pile') for r in results
                      if r.get('_center_session') == center_batch}
    rows=[]
    for pile, points in engine.groups().items():
        if not points:
            continue
        signature=frozenset(p['key'] for p in points)
        if ignored_groups.get(pile)==signature:
            state='Ignored'
        elif pile in engine.accepted and pile in approved_in_main:
            state='Accepted'
        elif len({(p['north'],p['east']) for p in points})<3:
            state=('Single shot (no center)' if len(points)==1 else
                   'Insufficient shots (2)' if len(points)==2 else
                   'Insufficient distinct shots (need 3)')
        else:
            state='Needs Review'
        rows.append({'pile':pile,'shots':len(points),'state':state,
                     'points':', '.join(str(p.get('point','')) for p in points),
                     'signature':signature,'sources':len({p.get('source','') for p in points})})
    def natural(pile):
        return tuple((1,int(x)) if x.isdigit() else (0,x.casefold())
                     for x in re.split(r'(\d+)',pile))
    return sorted(rows,key=lambda r:natural(r['pile']))


def asset_path(filename):
    return os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), filename)

class App:
    def __init__(self,root):
        self.root=root;root.title('PIVOT | A Faulk Product');root.geometry('1250x780');root.minsize(890,600)
        self.standards=[];self.field=[];self.single_shot_keys=set();self.results=[];self.pending_batch=None;self.latest_batch=None;self.sort_column='pile';self.sort_reverse=False
        self.center_engine=CenterEngine()
        self.tagged_shots=TaggedShotTracker()
        self.center_batch='CENTER_SESSION:'+uuid.uuid4().hex
        self.center_window=None
        self.raw_observations=[]; self.live_raw_observations=[]
        self.source_files=set()
        # A user can explicitly ignore irrelevant perimeter groups without
        # deleting the raw shots. Ignore is scoped to this PIVOT session and
        # expires automatically if new observations are imported for the pile.
        self.ignored_center_groups={}
        s=ttk.Style(root);s.theme_use('clam');s.configure('Treeview',rowheight=26);s.configure('Treeview.Heading',font=('Segoe UI',9,'bold'))
        # Compact Excel-inspired title block, separate from the results workspace.
        banner=tk.Frame(root,bg='#FFFFFF',height=104,highlightbackground='#D4DEEA',highlightthickness=1)
        banner.pack(fill='x');banner.pack_propagate(False)
        self.logo_images=[]
        try:
            photo=tk.PhotoImage(file=asset_path('PIVOT_Header_Logo.png'))
            self.logo_images.append(photo)
            tk.Label(banner,image=photo,bg='#FFFFFF').pack(side='left',padx=(18,24),pady=6)
        except Exception:
            tk.Label(banner,text='PIVOT',bg='#FFFFFF',fg='#17365D',font=('Segoe UI',24,'bold')).pack(side='left',padx=18)
        title_area=tk.Frame(banner,bg='#FFFFFF');title_area.pack(side='left',fill='both',expand=True,pady=(10,8),padx=(0,18))
        tk.Label(title_area,text='PILE DEVIATION REPORT',bg='#17365D',fg='#FFFFFF',
                 font=('Segoe UI',17,'bold'),anchor='w',padx=12).pack(fill='x')
        tk.Label(title_area,text='A Faulk Product  |  Pile Inspection, Verification & Offset Tracking',
                 bg='#FFFFFF',fg='#17365D',font=('Segoe UI',10,'bold'),anchor='w').pack(fill='x',pady=(6,0))
        try:
            self.icon=tk.PhotoImage(file=asset_path('PIVOT_Icon.png'))
            root.iconphoto(True,self.icon)
        except Exception: pass
        menu=tk.Menu(root)
        center_menu=tk.Menu(menu,tearoff=False)
        center_menu.add_command(label='Review Center Queue...',command=self.review_center_candidates)
        center_menu.add_command(label='Calculate Pile Centers',command=self.center_action)
        menu.add_cascade(label='Pile Centers',menu=center_menu)
        menu.add_command(label='Pile Coverage / Missing Piles...',command=self.show_coverage)
        menu.add_command(label='Import Multiple CSVs...',command=self.import_multiple_csvs)
        root.configure(menu=menu)
        top=tk.Frame(root,padx=14,pady=10);top.pack(fill='x')
        # Present the normal survey workflow in the order field crews use it.
        actions=(
            ('1  Load Standards',self.load_standards),
            ('2  Import CSV',self.load_field),
            ('3  Calculate Centers',self.center_action),
            ('4  Match Points',self.calculate),
            ('5  Import Previous Report',self.import_previous),
            ('6  Export Excel',self.export),
        )
        for title,command in actions:
            ttk.Button(top,text=title,command=command).pack(side='left',padx=4)
        filters=tk.Frame(root,padx=18,pady=5);filters.pack(fill='x')
        ttk.Button(filters,text='Pile Coverage / Missing Piles',command=self.show_coverage).pack(side='right',padx=4)
        ttk.Label(filters,text='Show:').pack(side='left')
        self.view_mode=tk.StringVar(value='Current piles')
        ttk.Combobox(filters,textvariable=self.view_mode,state='readonly',width=21,
                     values=['Current piles','All shot history','Most recent import']).pack(side='left',padx=(6,14))
        self.view_mode.trace_add('write',lambda *_:self.refresh_table())
        ttk.Label(filters,text='Date shot:').pack(side='left')
        self.date_filter=tk.StringVar(value='All dates')
        self.date_selector=ttk.Combobox(filters,textvariable=self.date_filter,state='readonly',width=20,values=['All dates'])
        self.date_selector.pack(side='left',padx=(6,14))
        self.date_filter.trace_add('write',lambda *_:self.refresh_table())
        self.summary=tk.StringVar(value='Load standards and import a dated field CSV, or import a previous PIVOT report.')
        tk.Label(root,textvariable=self.summary,anchor='w',padx=20,font=('Segoe UI',10)).pack(fill='x',pady=(0,8))
        # Original compact workspace, with Date Shot immediately after Field Point.
        area=tk.Frame(root,padx=14);area.pack(fill='both',expand=True)
        cols=('point','shot_date','pile','dn','de','dz','offset','distance','status')
        self.tree=ttk.Treeview(area,columns=cols,show='headings',selectmode='browse')
        headings=('Field Point ⇅','Date Shot ⇅','Pile ID ⇅','Δ Northing','Δ Easting',
                  'Δ Elevation','Radial Offset','Match Distance','Status')
        self.headings=dict(zip(cols,headings))
        for key,label in zip(cols,headings):
            self.tree.heading(key,text=label,command=lambda k=key:self.sort_by(k))
            self.tree.column(key,width=130,anchor='center',stretch=True)
        self.tree.column('status',width=215)
        scrolly=ttk.Scrollbar(area,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrolly.set)
        self.tree.pack(side='left',fill='both',expand=True)
        scrolly.pack(side='right',fill='y')
        self.tree.tag_configure('review',background='#fff0d6')
        self.tree.tag_configure('unmatched',background='#ffe4e4')
        bottom=tk.Frame(root,bg='#f0f3f7',padx=18,pady=12);bottom.pack(fill='x')
        tk.Label(bottom,text='Matching uses combined Northing + Easting distance, checks all standard piles, and prevents duplicate assignments. Review flagged matches before using results.',bg='#f0f3f7',fg='#34445a',wraplength=1100,justify='left').pack(anchor='w')
        tk.Label(bottom,text='Elevation comparison: As-Built Top of Cage − Design Top of Cage  |  Units: same as imported files (field header indicates US survey feet)',bg='#f0f3f7',fg='#34445a',wraplength=1100,justify='left').pack(anchor='w')
    def coverage_snapshot(self, expected=None):
        return reconcile(getattr(self,'raw_observations',[]),self.results,self.standards,
                         described_pile_id,report_pile_id,is_pending_record,expected)

    def show_coverage(self):
        """Always accessible accounting dashboard; no external audit file needed."""
        window=tk.Toplevel(self.root)
        window.title('PIVOT | Imported Pile Coverage')
        window.geometry('1100x650');window.minsize(780,420)
        title=tk.Frame(window,bg='#17365D',padx=14,pady=10);title.pack(fill='x')
        tk.Label(title,text='PILE COVERAGE & IMPORT RECONCILIATION',
                 font=('Segoe UI',15,'bold'),bg='#17365D',fg='white').pack(anchor='w')
        tk.Label(title,text='Every recognized pile is shown; pending does not mean deleted.',
                 bg='#17365D',fg='white').pack(anchor='w')
        top=tk.Frame(window,padx=12,pady=9);top.pack(fill='x')
        tk.Label(top,text='Expected piles (optional):').pack(side='left')
        expected_var=tk.StringVar(value='')
        ttk.Entry(top,textvariable=expected_var,width=9).pack(side='left',padx=8)
        summary=tk.StringVar()
        # Keep counts on their own line so a long reconciliation summary is not
        # clipped by the optional expected-count entry on smaller screens.
        tk.Label(window,textvariable=summary,font=('Segoe UI',10,'bold'),
                 anchor='w',padx=12,pady=3).pack(fill='x')
        frame=tk.Frame(window,padx=12,pady=5);frame.pack(fill='both',expand=True)
        keys=('pile','shots','dates','status','issues','field_points')
        tree=ttk.Treeview(frame,columns=keys,show='headings')
        headings={'pile':'Pile ID','shots':'Raw Shots','dates':'Date(s) Shot',
                  'status':'Report Status','issues':'Checks / Suspected Label Problems',
                  'field_points':'Source Point Numbers'}
        for key in keys:
            tree.heading(key,text=headings[key])
            tree.column(key,width=(280 if key=='issues' else 200 if key=='field_points'
                                   else 160 if key=='dates' else 120),anchor='w')
        tree.column('pile',width=100)
        ys=ttk.Scrollbar(frame,orient='vertical',command=tree.yview)
        xs=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
        tree.configure(yscrollcommand=ys.set,xscrollcommand=xs.set)
        ys.pack(side='right',fill='y')
        xs.pack(side='bottom',fill='x')
        tree.pack(side='left',fill='both',expand=True)
        tree.tag_configure('pending',background='#FFF1D6')
        tree.tag_configure('missing',background='#FFD9DD')
        tree.tag_configure('issue',background='#FFE9BD')
        info=tk.StringVar()
        tk.Label(window,textvariable=info,anchor='w',justify='left',wraplength=1000,
                 padx=12,pady=6).pack(fill='x')
        def refresh():
            entered=expected_var.get().strip()
            try: expected=int(entered) if entered else None
            except ValueError:
                messagebox.showerror('Expected pile count','Enter a whole number or leave the field blank.',parent=window)
                return
            data=self.coverage_snapshot(expected)
            for item in tree.get_children():tree.delete(item)
            for row in data['items']:
                status=row['status']
                tag=('missing' if status=='MISSING FROM REPORT' else
                     'pending' if status=='PENDING REVIEW' else
                     'issue' if row['issues'] else '')
                tree.insert('', 'end',values=tuple(row[k] for k in keys),tags=(tag,) if tag else ())
            gap=(f'  |  Expected difference: {data["difference"]:+d}'
                 if data['difference'] is not None else '')
            summary.set(f'{data["pile_count"]} unique IDs | '
                        f'{data["completed_count"]} completed | '
                        f'{data["pending_count"]} pending | '
                        f'{data["missing_count"]} missing | '
                        f'{data["untagged_highspot_count"]} HS labels to verify'+gap)
            unknown=data['unidentified']
            excluded='; '.join(f'{r.get("point", "?")}: {r.get("desc", "")}' for r in unknown[:12])
            daily=', '.join(f'{date[5:]}: {count}' for date,count in
                            data['date_pile_counts'].items())
            info.set(f'{data["raw_count"]} unique source observations | '
                     f'{data["pile_date_count"]} pile/date observations | '
                     f'{len(data["repeat_ids"])} re-shot piles ('+
                     (', '.join(data['repeat_ids']) or 'none')+').\n'
                     f'Piles by date: {daily or "None"}.\n'
                     f'{len(unknown)} non-pile descriptions: {excluded or "None"}. '
                     'Double-click a pile to inspect source labels.')
            self._coverage_rows=data['items']
        def details(event=None):
            selected=tree.selection()
            if not selected:return
            values=tree.item(selected[0],'values')
            pile=values[0]
            row=next((r for r in self._coverage_rows if r['pile']==pile),None)
            if row:
                messagebox.showinfo('Pile '+pile,
                    f'Pile: {pile}\nField points: {row["field_points"]}\n'
                    f'Date(s): {row["dates"]}\nStatus: {row["status"]}\n'
                    f'Sources: {row["source"]}\nChecks: {row["issues"] or "None"}',parent=window)
        tree.bind('<Double-1>',details)
        bottom=tk.Frame(window,padx=12,pady=9);bottom.pack(fill='x')
        ttk.Button(bottom,text='Refresh / Reconcile',command=refresh).pack(side='left')
        ttk.Button(bottom,text='Close',command=window.destroy).pack(side='right')
        refresh()

    def update_date_filter(self):
        days=sorted({r.get('shot_date','') for r in self.results if r.get('shot_date')})
        choices=['All dates']+days
        self.date_selector['values']=choices
        if self.date_filter.get() not in choices:self.date_filter.set('All dates')

    def center_queue(self):
        return center_candidate_queue(self.center_engine,self.ignored_center_groups,
                                      self.results,self.center_batch)

    def ignore_center_candidates(self, pile_ids):
        """Remove only current-session center approvals; preserve all survey observations."""
        groups=self.center_engine.groups()
        dropped=[]
        for pile in pile_ids:
            points=groups.get(pile,[])
            if not points:
                continue
            self.ignored_center_groups[pile]=frozenset(p['key'] for p in points)
            self.center_engine.accepted.pop(pile,None)
            dropped.append(pile)
        if not dropped:
            return
        removed=[r for r in self.results if
                 r.get('_center_session')==self.center_batch and r.get('_center_pile') in dropped]
        self.results=[r for r in self.results if r not in removed]
        # Deliberately ignored piles remain traceable rather than disappearing
        # from the report with no indication of the operator's decision.
        for pile in dropped:
            previous=next((r for r in removed if report_pile_id(r)==pile),None)
            if previous:
                batch={k:previous.get(k,'') for k in ('batch','shot_date','shot_job')}
            else:
                shots=groups.get(pile,[])
                try:
                    date,job=shot_info(shots[0].get('source',''))
                except (ValueError,IndexError):
                    date,job='', ''
                batch={'batch':job,'shot_date':date,'shot_job':job}
            if not any(report_pile_id(r)==pile and
                       r.get('shot_date')==batch.get('shot_date') and
                       not is_pending_record(r) for r in self.results):
                ignored=pending_record(pile,groups.get(pile,[]),batch,
                                       getattr(self,'standards',[]),'CENTER IGNORED BY OPERATOR')
                ignored['_live_import']=True
                self.results=[r for r in self.results if not (
                    is_pending_record(r) and report_pile_id(r)==pile and
                    r.get('shot_date')==batch.get('shot_date'))]
                self.results.append(ignored)
        self.update_date_filter();self.refresh_table()
        if self.center_window and self.center_window.window.winfo_exists():
            self.center_window.refresh()
        self.summary.set(f'Ignored {len(dropped)} center group(s) for this session; raw survey shots preserved.')

    def review_center_candidates(self, for_export=False):
        """Modal pre-export reminder plus reversible ignore/trash list.

        Returns 'export' only when the user explicitly confirms proceeding.
        Center calculations are never silently accepted or removed.
        """
        candidates=self.center_queue()
        if not candidates:
            if not for_export:
                messagebox.showinfo('Center queue',
                    'No pile-coded exterior observations are loaded. Import perimeter survey shots to review center candidates.',
                    parent=self.root)
            return 'export' if for_export else 'close'
        dlg=tk.Toplevel(self.root)
        dlg.title('PIVOT | Review Pile Centers Before Export' if for_export else 'PIVOT | Center Queue')
        width=min(830,max(600,self.root.winfo_screenwidth()-90))
        height=min(530,max(380,self.root.winfo_screenheight()-140))
        dlg.geometry(f'{width}x{height}')
        dlg.minsize(min(590,width),min(360,height))
        dlg.transient(self.root)
        dlg.grab_set()
        response={'action':'cancel'}
        header=tk.Frame(dlg,bg='#17365D',padx=16,pady=13)
        header.pack(fill='x')
        tk.Label(header,text='PILE CENTER REVIEW',font=('Segoe UI',16,'bold'),
                 bg='#17365D',fg='white').pack(anchor='w')
        lead=tk.StringVar()
        tk.Label(dlg,textvariable=lead,justify='left',anchor='w',wraplength=780,
                 font=('Segoe UI',10),padx=16,pady=9).pack(fill='x')
        frame=tk.Frame(dlg,padx=16)
        frame.pack(fill='both',expand=True)
        cols=('pile','points','shots','sources','state')
        tree=ttk.Treeview(frame,columns=cols,show='headings',selectmode='extended',height=9)
        for col,label,w in [('pile','Pile ID',105),('points','Field Point(s)',180),
                            ('shots','Shots',70),('sources','Source Files',100),
                            ('state','Center Status',190)]:
            tree.heading(col,text=label)
            tree.column(col,width=w,anchor='center',stretch=True)
        scroll=ttk.Scrollbar(frame,orient='vertical',command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left',fill='both',expand=True)
        scroll.pack(side='right',fill='y')
        note=tk.Label(dlg,text=('Trash / Ignore removes the selected pile from center reminders for this session. '
                                'It does NOT delete any original CSV shots. Importing new shots for that pile '
                                'will bring it back for review.'),
                      font=('Segoe UI',9),fg='#526477',wraplength=780,justify='left',padx=16,pady=6)
        note.pack(fill='x')
        buttons=tk.Frame(dlg,padx=14,pady=10,bg='#EDF2F7')
        buttons.pack(side='bottom',fill='x')

        def refresh():
            old=tree.selection()
            for item in tree.get_children():tree.delete(item)
            rows=self.center_queue()
            for item in rows:
                status=item['state']
                tree.insert('', 'end',iid=item['pile'],values=(item['pile'],item['points'],
                           item['shots'],item['sources'],status))
            for item in old:
                if tree.exists(item):tree.selection_add(item)
            pending=sum(r['state']=='Needs Review' for r in rows)
            ignored=sum(r['state']=='Ignored' for r in rows)
            accepted=sum(r['state']=='Accepted' for r in rows)
            prefix=('Before exporting: ' if for_export else '')
            lead.set(f'{prefix}{pending} pile group(s) have pile-coded observations without an approved center.\n'
                     f'{accepted} accepted  |  {ignored} ignored for this session. '
                     'Choose Calculate / Review, ignore unnecessary groups, or continue the export. '
                     'Single- and two-shot piles do not require a centroid.')
        def selected_ids():
            ids=list(tree.selection())
            if not ids:messagebox.showwarning('Select piles','Select one or more pile IDs from the list first.',parent=dlg)
            return ids
        def ignore_selected():
            ids=selected_ids()
            if not ids:return
            states={r['pile']:r['state'] for r in self.center_queue()}
            change=[pile for pile in ids if states.get(pile)!='Ignored']
            if not change:return
            if not messagebox.askyesno('Ignore center candidate(s)?',
                'Ignore the selected pile center group(s)?\n\n'
                'They will not appear in the pre-export reminder for this session. '
                'If already accepted in this session, their derived center rows will be removed '
                'from the main results. Original survey shots and previous reports will NOT be deleted.',
                parent=dlg):return
            self.ignore_center_candidates(change)
            refresh()
        def restore_selected():
            ids=selected_ids()
            if not ids:return
            self.restore_center_candidates(ids)
            refresh()
        def finish(action):
            response['action']=action
            dlg.destroy()
        def open_calculator():
            ids=list(tree.selection())
            finish('review')
            self.show_center_window()
            if ids:
                center=self.center_window
                # The review window can calculate any previously imported pile.
                if ids[0] not in self.center_engine.calculations:
                    self.center_engine.calculate_all()
                    self.reconcile_center_acceptances()
                    center.refresh()
                if center.tree.exists(ids[0]):
                    center.tree.selection_set(ids[0])
                    center.tree.see(ids[0])
                    center.select_pile()
        ttk.Button(buttons,text='Calculate / Review Centers',command=open_calculator).pack(side='left',padx=3)
        ttk.Button(buttons,text='Trash / Ignore Selected',command=ignore_selected).pack(side='left',padx=3)
        ttk.Button(buttons,text='Restore Ignored',command=restore_selected).pack(side='left',padx=3)
        if for_export:
            ttk.Button(buttons,text='Export Anyway',command=lambda:finish('export')).pack(side='right',padx=3)
            ttk.Button(buttons,text='Cancel Export',command=lambda:finish('cancel')).pack(side='right',padx=3)
        else:
            ttk.Button(buttons,text='Close',command=lambda:finish('close')).pack(side='right',padx=3)
        dlg.protocol('WM_DELETE_WINDOW',lambda:finish('cancel'))
        refresh()
        dlg.wait_window()
        return response['action']

    def load_standards(self):
        path=filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx')])
        if not path:return
        try:
            candidates,diagnostics=read_standards(path,with_diagnostics=True)
            conflicts=diagnostics['conflicts']
            hidden_only=diagnostics['secondary_only']
            if conflicts or hidden_only:
                primary=diagnostics['primary_sheet'] or 'the first visible design sheet'
                lines=[f"Primary design sheet: {primary}",
                       f"Conflicting pile IDs: {len(conflicts)}",
                       f"Pile IDs found only on hidden sheets: {len(hidden_only)}"]
                for entry in conflicts[:8]:
                    chosen=entry['chosen'];alternate=entry['alternate']
                    lines.append(f"{entry['pile_id']}: {chosen['source']} "
                                 f"({chosen['north']:.3f}, {chosen['east']:.3f}) "
                                 f"vs {alternate['source']} "
                                 f"({alternate['north']:.3f}, {alternate['east']:.3f})")
                if len(conflicts)>8:
                    lines.append(f"... and {len(conflicts)-8} additional conflicts.")
                lines.append('\nContinue using the primary visible design sheet '
                             'for conflicting pile IDs? The underlying workbook '
                             'must be verified by the survey/design team.')
                if not messagebox.askyesno('Review conflicting pile standards',
                                           '\n'.join(lines)):
                    return
            self.standards=candidates
            self.standard_diagnostics=diagnostics
            for record in self.results:
                if record.get('_center_session')==self.center_batch or record.get('_direct_center'):
                    update_center_design(record,self.standards)
                elif is_pending_record(record):
                    record['design']=next((p for p in self.standards
                        if p['id'].strip().upper()==report_pile_id(record)),None)
            self.refresh_table()
            self.summary.set(f'Standards loaded: {len(self.standards)} unique piles | {len(diagnostics["conflicts"])} coordinate conflicts reviewed | {os.path.basename(path)} | {len(self.results)} previous shot records retained')
        except Exception as e:messagebox.showerror('Standards error',str(e))

    def restore_center_candidates(self, pile_ids):
        """Undo a current-session ignore without deleting original measurements."""
        changed=False
        for pile in pile_ids:
            if pile in self.ignored_center_groups:
                self.ignored_center_groups.pop(pile,None)
                changed=True
        if changed:
            for record in self.results:
                if (report_pile_id(record) in pile_ids and is_pending_record(record)
                    and 'IGNORED BY OPERATOR' in record['status']):
                    record['status']='PENDING - CENTER NEEDS REAPPROVAL'
            self.summary.set('Restored pile center group(s) to the center review queue.')
            if self.center_window and self.center_window.window.winfo_exists():
                self.center_window.refresh()
        return changed

    def is_center_candidate_ignored(self, pile_id):
        """An ignored group becomes active if its source observations change."""
        points=self.center_engine.groups().get(pile_id,[])
        return (bool(points) and
                self.ignored_center_groups.get(pile_id)==frozenset(p['key'] for p in points))

    def center_action(self):
        # Center calculation is OPTIONAL. Never require a centroid for a pile
        # with only one or two exterior observations.
        center_groups=[pile for pile,points in self.center_engine.groups().items()
                       if len({(p['north'],p['east']) for p in points})>=3
                       and not self.is_center_candidate_ignored(pile)]
        if not center_groups:
            note=('No pile has three or more exterior observations, so there are '
                  'no perimeter centers ready to calculate.\n\n'
                  'You can skip this step and click 4  Match Points. '
                  'Single-shot pile locations can be matched as measured points '
                  'with a review warning. They are not calculated centers.')
            if self.tagged_shots.has_shots():
                note += ('\n\nTagged C/HS shots already contain surveyed center '
                         'coordinates and do not require polygon calculations.')
            messagebox.showinfo('No center calculations needed',note)
            return
        self.show_center_window()

    def show_center_window(self):
        if self.center_window and self.center_window.window.winfo_exists():
            self.center_window.window.lift()
            self.center_window.window.focus_set()
        else:
            self.center_window=CenterWindow(
                self.root,self.center_engine,
                on_accept=self.add_accepted_center,
                on_change=self.reconcile_center_acceptances,
                on_ignore=self.ignore_center_candidates,
                on_restore=self.restore_center_candidates,
                is_ignored=self.is_center_candidate_ignored,
                on_import_field=self.load_field)

    def reconcile_center_acceptances(self):
        # An altered/excluded perimeter invalidates its previous acceptance.
        # Never silently delete that pile from the report: restore a pending
        # review record until the corrected center is approved again.
        invalidated=[r for r in self.results if (
            r.get('_center_session')==self.center_batch and
            r.get('_center_pile') not in self.center_engine.accepted)]
        if not invalidated:return
        self.results=[r for r in self.results if r not in invalidated]
        groups=self.center_engine.groups()
        for old in invalidated:
            pile=report_pile_id(old)
            observations=groups.get(pile,[])
            batch={k:old.get(k,'') for k in ('batch','shot_date','shot_job')}
            pending=pending_record(pile,observations,batch,self.standards,
                                   'CENTER NEEDS REAPPROVAL')
            pending['_live_import']=True
            pending['comment']=old.get('comment','')
            self.results=cover_piles_with_results(self.results,[pending])
        self.update_date_filter();self.refresh_table()
        self.summary.set('A previously accepted center changed: it is now PENDING REVIEW '
                         'until you approve a new center. Original survey shots are preserved.')

    def add_accepted_center(self, pile_id, elevation_mode='Blank', manual_elevation=None):
        # Called only after explicit acceptance in the center review workspace.
        self.ignored_center_groups.pop(pile_id,None)
        record=accepted_center_record(self.center_engine,pile_id,self.standards,
                    elevation_mode,manual_elevation,self.center_batch)
        # Replace the current-session version for this pile, preserving all prior
        # survey batches and the original perimeter observations.
        self.results=[r for r in self.results if not (
            r.get('_center_session')==self.center_batch and r.get('_center_pile')==pile_id)]
        record['_live_import']=True
        self.results=cover_piles_with_results(self.results,[record])
        self.latest_batch=self.center_batch
        self.sort_column='pile';self.sort_reverse=False
        self.view_mode.set('Current piles');self.date_filter.set('All dates')
        self.update_date_filter();self.refresh_table()
        description='Design standard matched by pile ID' if record['design'] else 'No exact pile ID in standards; load/check standards'
        self.summary.set(f'{pile_id} approved center added to main PIVOT table | {description} | {len(self.results)} total records')

    def import_multiple_csvs(self):
        paths=filedialog.askopenfilenames(title='Import All Survey CSV Files',
                                          filetypes=[('Survey CSV files','*.csv')])
        if not paths:return
        successes=0
        for path in paths:
            if self.load_field(path):successes+=1
        self.summary.set(f'{successes}/{len(paths)} CSV files imported. '
                         f'{self.coverage_snapshot()["pile_count"]} unique pile IDs identified. '
                         'Click Match Points, then Pile Coverage to review unresolved piles.')

    def rebuild_live_match_queue(self):
        """Re-evaluate all current-session shots by Pile ID *and* survey date.

        This fixes an import-order bug: two observations of the same pile in
        separate CSVs must not let the first observation get matched as a center.
        Conversely, one shot on each of two different dates is two independent
        valid one-shot surveys, not an invented two-point perimeter.
        """
        groups=group_current_raw(self.live_raw_observations)
        selected=[]; single_keys=set(); seen=set()
        for point in self.live_raw_observations:
            date=str(point.get('shot_date',''))
            if parse_tagged_description(point.get('desc')):
                continue
            pid=described_pile_id(point.get('desc'))
            if pid:
                if len(groups[(date,pid)])!=1:
                    continue
            elif not ('ASB' in str(point.get('desc','')).upper() and
                      'CAGE' in str(point.get('desc','')).upper()):
                # Controls and other unidentified descriptions are kept in the
                # source ledger and Coverage panel, not assigned to random piles.
                continue
            key=(date,field_shot_key(point))
            if key in seen:continue
            seen.add(key)
            matchpoint=dict(point)
            matchpoint['_pivot_batch']={'batch':point.get('shot_job',''),
                                      'shot_date':date,'shot_job':point.get('shot_job','')}
            selected.append(matchpoint)
            if pid:single_keys.add(field_shot_key(matchpoint))
        self.field=selected
        self.single_shot_keys=single_keys
        self.pending_batch=selected[-1]['_pivot_batch'] if selected else None
        self.results=refresh_pending_group_records(
            self.results,self.live_raw_observations,self.standards)

    def load_field(self, path=None):
        """Import through the same pipeline from either main UI or center workspace.

        A CSV with isolated pile IDs is ordinary matchable survey data, not a
        list of center candidates. Centralizing imports avoids routing all
        those points into the polygon calculator as invalid geometry.
        """
        if path is None:
            path=filedialog.askopenfilename(filetypes=[('CSV files','*.csv')])
        if not path:return False
        try:
            points=read_field(path)
            tagged=[p for p in points if parse_tagged_description(p['desc'])]
            survey=[p for p in points if 'ASB' in p['desc'].upper() and 'CAGE' in p['desc'].upper()
                    and not parse_tagged_description(p['desc'])]
            pile_coded,singles,perimeter_groups,two_shot_groups=classify_pile_shots(
                points,self.center_engine)
            # Direct C/HS observations are handled only by the paired-shot tracker.
            # True perimeter groups cannot be treated as measured centers.
            # A single pile-coded point is matchable as a raw survey location,
            # only after an explicit warning at Match Points.
            fallback=not tagged and not survey and not pile_coded
            # Do not match explicit control/check shots to random design piles.
            controls={'CHK','CHECK','CTRL','CONTROL','BM','BENCHMARK','CP','CONTROL POINT'}
            spatial=[p for p in points if str(p.get('desc','')).strip().upper() not in controls]
            selected=list(survey)+singles+(spatial if fallback else [])
            selected=list({field_shot_key(p):p for p in selected}.values())
            shot_date=job=None
            if selected or tagged or pile_coded:
                shot_date,job=shot_info(path,path)
            if selected or tagged:
                date_note=(datetime.strptime(shot_date,'%Y-%m-%d').strftime('%m-%d-%Y')
                           if shot_date else 'Unknown')
                description=(f'Source: {os.path.basename(path)}\nJob: {job}\n\n'
                             f'Date shot: {date_note}\n\n'
                             f'{len(selected)} survey observation(s) available for Match Points.')
                if singles:
                    description += (f'\n\n{len(singles)} pile-coded single shot(s) detected. '
                        'Match Points will check Northing, Easting and elevation individually against '
                        'the ±0.250 tolerance without requiring a center calculation.')
                if perimeter_groups:
                    description+=f'\n{len(perimeter_groups)} multi-shot pile group(s) available for optional center review.'
                if two_shot_groups:
                    description+=f'\n{len(two_shot_groups)} two-shot group(s) need more perimeter observations.'
                if tagged:
                    description+=f'\n{len(tagged)} C/HS measurement(s) also stored for pairing.'
                description+='\n\nContinue importing?'
                if not messagebox.askokcancel('Confirm Date Shot',description):
                    return False
            # Every numeric row is retained for reconciliation, including check
            # points, incomplete centers and paired C/HS observations.
            raw=[]
            for row in points:
                entry=dict(row)
                entry.update(shot_date=shot_date or '',shot_job=job or '',
                             source=os.path.basename(path))
                raw.append(entry)
            self.raw_observations=merge_raw(getattr(self,"raw_observations",[]),raw)
            self.live_raw_observations=merge_raw(getattr(self,"live_raw_observations",[]),raw)
            getattr(self,"source_files",set()).add(os.path.abspath(path))
            added_tagged=duplicate_tagged=0
            if tagged:
                added_tagged,duplicate_tagged=self.tagged_shots.add(tagged,shot_date,job,path)
            added=duplicates=0
            if pile_coded:
                added,ignored,duplicates=self.center_engine.add_observations(points,path)
                self.reconcile_center_acceptances()
                if self.center_window and self.center_window.window.winfo_exists():
                    self.center_engine.calculate_all(eligible_only=True)
                    self.reconcile_center_acceptances()
                    self.center_window.refresh()
            if selected or tagged or pile_coded:
                batch={'batch':job,'shot_date':shot_date,'shot_job':job}
                self.results=add_pending_coverage(
                    self.results,points,batch,self.standards,self.center_engine.pattern)
                missing=missing_pile_coverage(self.results,points,shot_date,
                                            self.center_engine.pattern)
                if missing:
                    raise ValueError('Pile coverage failed at import: '+', '.join(missing))
                self.update_date_filter();self.refresh_table()
            # Match queue is reconstructed from all dated live survey shots.
            # No pile-coded raw shot can be silently dropped due to import order.
            self.rebuild_live_match_queue()
            # Tagged observations are reported in the status line; no extra popup.
            # No automatic opening of the center workspace. This was forcing
            # users into invalid center calculations for one-shot piles.
            if tagged or selected or pile_coded:
                recognized={report_pile_id(r) for r in self.results}
                recognized.discard(None)
                self.summary.set(f'Imported CSV: {os.path.basename(path)} | '
                                 f'{len(recognized)} distinct pile IDs accounted for | '
                                 f'{len(self.field)} matchable survey points '
                                 f'({len(singles)} single-shot piles) | '
                                 f'{added_tagged} C/HS shots | '
                                 f'{len(perimeter_groups)} perimeter groups. '
                                 'Calculate Centers is optional; use Match Points for regular shots.')
            if fallback:
                messagebox.showwarning('Field point filter',
                    'No ASB TOP CAGE or pile-code descriptions found; all numeric survey '
                    'points were included. Verify before calculating.')
            return True
        except Exception as e:
            messagebox.showerror('Field data error',str(e))
            return False

    def import_previous(self):
        path=filedialog.askopenfilename(filetypes=[('PIVOT Excel report','*.xlsx')])
        if not path:return
        try:
            imported,legacy=load_previous_report(path)
            self.raw_observations=merge_raw(self.raw_observations,load_report_sources(path))
            existing=len(self.results)
            self.results=combine_records(self.results,imported)
            if self.latest_batch is None and imported:self.latest_batch=imported[-1].get('batch')
            self.update_date_filter();self.refresh_table()
            self.summary.set(f'Imported previous report: {len(self.results)-existing} historical shots added | {len(current_pile_records(self.results))} current piles | {len(self.results)} total shots')
            if legacy:messagebox.showwarning('Older PIVOT report','Imported this older report, but it did not record the date shot. Its rows are labeled Unknown (legacy).')
            else:messagebox.showinfo('Running Total',f'{len(self.results)-existing} prior records imported. Previous dates and pile matches are retained.')
        except Exception as e:messagebox.showerror('Import previous report error',str(e))
    @staticmethod
    def natural_key(value):
        # Sort alphanumeric IDs like B2, B15, B35 in numeric sequence.
        return tuple((1,int(piece)) if piece.isdigit() else (0,piece.casefold())
                     for piece in re.split(r'(\d+)',str(value or '')))

    def filtered_results(self):
        mode=self.view_mode.get()
        if mode=='All shot history':
            rows=self.results
        elif mode=='Most recent import' and self.latest_batch:
            # Show the actual most recently imported batch, including when its
            # survey date is older than the most recent dated pile observation.
            rows=current_pile_records([r for r in self.results
                                       if r.get('batch')==self.latest_batch])
        else:
            rows=current_pile_records(self.results)
        if self.date_filter.get()!='All dates':
            # Historical records can be viewed by switching to All shot history.
            rows=[r for r in rows if r.get('shot_date')==self.date_filter.get()]
        return rows

    def sorted_results(self, for_export=False):
        rows=list(current_pile_records(self.results) if for_export
                  else self.filtered_results())
        def key(r):
            f=r['field'];p=r['design'];col=self.sort_column or 'pile'
            dn=f['north']-p['north'] if p and f['north'] is not None else None
            de=f['east']-p['east'] if p and f['east'] is not None else None
            dz=f['elev']-p['elev'] if p and p['elev'] is not None and f['elev'] is not None else None
            values={'pile':p['id'] if p else r.get('_center_pile',''), 'point':f['point'],
                    'shot_date':r.get('shot_date',''), 'design_n':p['north'] if p else None,
                    'design_e':p['east'] if p else None,'design_z':p['elev'] if p else None,
                    'field_n':f['north'],'field_e':f['east'],'field_z':f['elev'],
                    'dn':dn,'de':de,'dz':dz,'offset':math.hypot(dn,de) if dn is not None and de is not None else None,
                    'ns':direction(dn,'North','South') if dn is not None else '',
                    'ew':direction(de,'East','West') if de is not None else '',
                    'hl':direction(dz,'High','Low') if dz is not None else '',
                    'distance':r['distance'],'status':r['status']}
            v=values.get(col,'')
            if col=='pile':return (p is None,self.natural_key(v),r.get('shot_date',''))
            return (v is None,v if isinstance(v,(int,float)) else self.natural_key(v))
        return sorted(rows,key=key,reverse=self.sort_reverse)

    def sort_by(self,column):
        if self.sort_column==column:self.sort_reverse=not self.sort_reverse
        else:self.sort_column=column;self.sort_reverse=False
        self.refresh_table()

    def refresh_table(self):
        for column,label in self.headings.items():
            if self.sort_column==column:
                self.tree.heading(column,text=label.replace(' ⇅','')+(' ▼' if self.sort_reverse else ' ▲'))
            else:self.tree.heading(column,text=label)
        for item in self.tree.get_children():self.tree.delete(item)
        for r in self.sorted_results():
            f=r['field'];p=r['design'];tag='unmatched' if p is None else 'review' if not is_match_without_review(r) else ''
            dn=round(f['north']-p['north'],3) if p and f['north'] is not None else None
            de=round(f['east']-p['east'],3) if p and f['east'] is not None else None
            dz=round(f['elev']-p['elev'],3) if p and p['elev'] is not None and f['elev'] is not None else None
            vals=(f['point'],display_shot_date(r.get('shot_date','')),p['id'] if p else r.get('_center_pile',''),
                  fmt(dn),fmt(de),fmt(dz),fmt(math.hypot(dn,de)) if dn is not None and de is not None else '',
                  fmt(r['distance']),r['status'])
            self.tree.insert('', 'end',values=vals,tags=(tag,) if tag else ())

    def calculate(self):
        ordinary_ready=bool(self.field and self.pending_batch)
        if ordinary_ready and not self.standards:
            messagebox.showwarning('Standards needed',
                'Load the pile standards before matching ordinary field points.')
            return
        # Ask for the elevation source once, BEFORE matching any records.
        # C always supplies the XY center. HS/H is optional for the Z value.
        direct, issues=self.tagged_shots.build_records(self.standards)
        # One simple Yes/No prompt for paired C + H/HS shots. No extra
        # single-shot confirmation or post-match review popup is shown.
        # If all tagged pairs were matched earlier and no new pair has arrived,
        # do not ask again merely because Match Points was clicked twice.
        def previously_matched_pair(r):
            return any(
                existing.get('_direct_center') or
                str(existing.get('status','')).startswith('CENTER/HS')
                for existing in self.results
                if report_pile_id(existing) == report_pile_id(r)
                and str(existing.get('shot_date','')) == str(r.get('shot_date',''))
                and str(existing['field'].get('point','')) == str(r['field'].get('point',''))
                and _coordinate_identity(existing['field'].get('north')) == _coordinate_identity(r['field'].get('north'))
                and _coordinate_identity(existing['field'].get('east')) == _coordinate_identity(r['field'].get('east'))
                and (str(existing['field'].get('desc','')).find('HS shot point') >= 0 or
                     str(existing['field'].get('desc','')).find('high-spot Z') >= 0)
            )
        fresh_pairs=[r for r in direct if not previously_matched_pair(r)]
        if fresh_pairs:
            use_hs=messagebox.askyesno(
                'Center / High Spot Elevation',
                'For center (C) and high-spot (H/HS) shot pairs, would you like '
                'to use the H/HS elevation for the C shot?\n\n'
                'Yes = use the H/HS elevation.\n'
                'No = keep the C-shot elevation.\n\n'
                'The center shot supplies Northing and Easting in either case.',
                default=messagebox.YES)
            source='HS' if use_hs else 'C'
            selected, issues=self.tagged_shots.build_records(
                self.standards,elevation_source=source)
            # Only newly imported pairs should be added. Previously processed
            # pairs cannot change silently due to a later ordinary CSV import.
            fresh_keys={(report_pile_id(r),str(r.get('shot_date','')),
                         str(r['field'].get('point',''))) for r in fresh_pairs}
            direct=[r for r in selected
                    if (report_pile_id(r),str(r.get('shot_date','')),
                        str(r['field'].get('point',''))) in fresh_keys]
        else:
            direct=[]
        if not ordinary_ready and not direct:
            if any(is_pending_record(r) for r in self.results):
                count=sum(is_pending_record(r) for r in current_pile_records(self.results))
                self.summary.set(f'{count} pile(s) are pending review. '
                                 'Those piles remain listed in the main table and '
                                 'in the Excel report; no center has been fabricated.')
                self.refresh_table()
                return
            if issues:
                self.summary.set('No complete C/HS pairs available: '
                                 + '; '.join(issues[:2]))
            elif self.tagged_shots.has_shots():
                self.summary.set('All complete C/HS pairs have already been matched. '
                                 'Import new field observations to continue.')
            else:
                messagebox.showwarning('Files needed',
                    'Import a dated field CSV and load standards before matching.')
            return
        candidates=list(self.field) if ordinary_ready else []
        # Single-shot pile-coded points are now automatically matched, not
        # silently discarded and not sent to the polygon center calculator.
        # Each result is evaluated against design after matching below.
        match_fields=candidates
        skipped=[]
        if not match_fields and not direct:
            messagebox.showinfo('No new observations',
                                'No new field points or unprocessed C/HS pairs to match.')
            return
        newly=[]
        if match_fields:
            batches=defaultdict(list)
            for point in match_fields:
                meta=point.get('_pivot_batch',self.pending_batch or {})
                batches[(meta.get('shot_date',''),meta.get('batch',''))].append(point)
            for group in batches.values():
                batch=group[0].get('_pivot_batch',self.pending_batch or {})
                matches=match_explicit_pile_ids(group,self.standards)
                for r in matches:
                    r.update(batch)
                    if field_shot_key(r['field']) in self.single_shot_keys:
                        r['status']=classify_single_shot_match(r,self.center_engine.pattern)
                newly.extend(matches)
        # An HS shot alongside one untagged observation (e.g. B262_HS + B262)
        # is NOT a verified C/HS pair. Retain the measured point but flag it
        # rather than silently treating its elevation as an approved C/HS value.
        possible_untagged=untagged_highspot_candidates(
            getattr(self,'live_raw_observations',[]), described_pile_id)
        ambiguous_keys={(x['pile'],x['date'],str(x['shot']['point']).strip())
                        for x in possible_untagged}
        for record in newly:
            if (report_pile_id(record),str(record.get('shot_date','')),
                str(record['field'].get('point','')).strip()) in ambiguous_keys:
                record['status']='SINGLE SHOT - VERIFY UNTAGGED C / HS'
        # Paired records use exact pile ID design lookup: never approximate-match
        # an explicit A218_C center to the wrong pile just because it is closer.
        newly.extend(direct)
        for r in newly:
            r['_live_import']=True
        # A user may re-run Match Points and choose a different elevation source
        # for the SAME surveyed C/HS pair. Replace that derived measurement
        # instead of storing two competing versions in the running history.
        paired_by_key={
            (report_pile_id(r),str(r.get('shot_date','')),
             str(r['field'].get('point','')).strip()):r
            for r in direct
        }
        prior=[]
        for existing in self.results:
            key=(report_pile_id(existing),str(existing.get('shot_date','')),
                 str(existing['field'].get('point','')).strip())
            replacement=paired_by_key.get(key) if existing.get('_direct_center') else None
            if replacement and record_identity(existing)!=record_identity(replacement):
                if existing.get('comment') and not replacement.get('comment'):
                    replacement['comment']=existing['comment']
                continue
            prior.append(existing)
        old_size=len(prior)
        self.results=cover_piles_with_results(prior,newly)
        if match_fields:
            self.latest_batch=self.pending_batch['batch']
        elif direct:
            self.latest_batch=direct[-1]['batch']
        self.pending_batch=None;self.field=[];self.single_shot_keys=set()
        self.sort_column='pile';self.sort_reverse=False
        self.update_date_filter();self.refresh_table()
        added=len(self.results)-old_size
        inventory=self.coverage_snapshot()
        self.summary.set(
            f'Matched new measurements | {inventory["pile_count"]} unique piles identified | '
            f'{inventory["completed_count"]} with reported measurements | '
            f'{inventory["pending_count"]} pending center review | '
            f'{inventory["untagged_highspot_count"]} untagged C/HS candidates. '
            'Open Pile Coverage for detailed count reconciliation.')
        # Pairing warnings remain in the summary. They no longer force a
        # second modal dialog immediately after matching.

    def export(self):
        coverage=self.coverage_snapshot()
        if coverage['missing_count']:
            names=', '.join(r['pile'] for r in coverage['items']
                            if r['status']=='MISSING FROM REPORT')
            messagebox.showerror('Cannot export — missing pile IDs',
                 f'{coverage["missing_count"]} recognized piles are not in the current report:\n'
                 f'{names}\n\nUse Match Points and Pile Coverage to resolve them. '
                 'PIVOT will not silently export a report that omits recognized pile IDs.')
            return
        unresolved=[r for r in current_pile_records(self.results) if is_pending_record(r)]
        if unresolved:
            names=', '.join(report_pile_id(r) or '?' for r in unresolved[:20])
            if len(unresolved)>20:names+=f' ... (+{len(unresolved)-20} more)'
            proceed=messagebox.askyesno('Unresolved Piles — Export Check',
                f'{len(unresolved)} recognized pile(s) have no reliable matched center:\n\n'
                f'{names}\n\n'
                'They WILL appear in the report with blank as-built deviations '
                'and PENDING REVIEW status on Match Verification.\n\n'
                'Continue exporting this INCOMPLETE report?\n'
                'Yes = export with visible pending piles; No = return to PIVOT.',
                icon='warning')
            if not proceed:return
        outstanding=[r for r in self.center_queue() if r['state']=='Needs Review']
        if outstanding and not unresolved and self.review_center_candidates(for_export=True)!='export':
            return
        if not self.results:messagebox.showwarning('No results','Run Match Piles or approve pile centers before exporting.');return
        path=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile='PIVOT_Pile_Deviation_Report.xlsx',filetypes=[('Excel workbook','*.xlsx')])
        if path:
            try:
                # Never overwrite the intended report with a partial/broken ZIP.
                dest=os.path.abspath(path)
                folder=os.path.dirname(dest)
                if not os.path.isdir(folder):
                    raise FileNotFoundError('The selected folder does not exist: '+folder)
                fd,temp_path=tempfile.mkstemp(prefix='.pivot_report_',suffix='.xlsx',dir=folder)
                os.close(fd)
                try:
                    self.results=deduplicate_records(self.results)
                    export_excel(temp_path,self.sorted_results(for_export=True),history_records=self.results,raw_observations=self.raw_observations)
                    # Validate the file and spreadsheet XML before replacing anything.
                    with zipfile.ZipFile(temp_path,'r') as workbook:
                        bad=workbook.testzip()
                        if bad: raise ValueError('Excel workbook has a damaged component: '+bad)
                        required=('xl/workbook.xml','xl/styles.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/worksheets/sheet3.xml')
                        for member in required:
                            ET.fromstring(workbook.read(member))
                        if 'xl/worksheets/sheet4.xml' in workbook.namelist():
                            ET.fromstring(workbook.read('xl/worksheets/sheet4.xml'))
                    os.replace(temp_path,dest)
                finally:
                    if os.path.exists(temp_path): os.remove(temp_path)
                messagebox.showinfo('Export complete','Excel report saved successfully:\n'+dest)
            except PermissionError:
                messagebox.showerror('Export error','Windows prevented PIVOT from saving the report.\n\nClose the Excel file if it is already open, or choose a different file name/folder (such as Desktop) and export again.')
            except Exception as e:
                details=traceback.format_exc()
                try:
                    logfile=os.path.join(os.path.expanduser('~'),'PIVOT_export_error.log')
                    with open(logfile,'w',encoding='utf-8') as out:out.write(details)
                    extra='\n\nDiagnostic details saved to: '+logfile
                except OSError: extra=''
                messagebox.showerror('Export error',type(e).__name__+': '+str(e)+extra)

if __name__=='__main__':
    root=tk.Tk();App(root);root.mainloop()
