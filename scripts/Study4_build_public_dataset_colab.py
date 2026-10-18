#!/usr/bin/env python3
"""Study 4 CheckMate: Colab/Drive -> claim-linked GitHub-review Excel export.

RUN IN GOOGLE COLAB (or locally with --input/--output):
  %pip install -q XlsxWriter
  !python /content/drive/MyDrive/PhD/Study4/Study4_build_public_dataset_colab_v3.py

This does NOT upload to GitHub, publish private CheckMate data, or grant release approval.
Only the REVIEW .xlsx belongs in the output folder. It contains ORIGINAL
CheckMate message IDs and readable TESA targets for joins and interpretation.
Unusually long, URL-containing, or suspicious target strings are withheld for
manual review to prevent accidental release of claim text or private identifiers.
No original claim text column is exported. Original IDs may increase linkability;
obtain explicit CheckMate release approval before putting the workbook on GitHub.
Keep PRIVATE_DO_NOT_PUBLISH outside the repo; never share the source JSONL.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import hmac
import ipaddress
import json
import math
import re
import secrets
from collections import Counter
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

try:
    import xlsxwriter
except ImportError as exc:
    raise SystemExit("Install XlsxWriter first: %pip install -q XlsxWriter") from exc

# Actual Step 16 Drive source verified in PhD/Study4/s4_ml/modelling_dataset.
DRIVE_STUDY4 = Path('/content/drive/MyDrive/PhD/Study4')
SOURCE_RELATIVE = Path('CheckMate Raw Data/Checkmate_full_export_set/s4_ml/modelling_dataset/checkmate_s4_message_inspection_view.jsonl')
EXPECTED = {'Claims': 1069, 'Reddit_Posts': 590, 'Reddit_Comments': 17416, 'Web_Evidence': 3568}
EVIDENCE_STAGES = [('0m',0), ('1m',1), ('15m',15), ('20m',20),
                   ('60m',60), ('90m',90), ('24h',1440), ('48h',2880)]
BASE36 = re.compile(r'^[a-z0-9]{3,18}$', re.I)
# Screening is conservative, not a substitute for human release/privacy review.
TARGET_SENSITIVE_PATTERN = re.compile(
    r'https?://|www\.|[A-Za-z0-9_.%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}|'
    r'original[_\s-]*text|message[_\s-]*body|forwarded\s+message', re.I)

def target_requires_review(raw):
    return (not raw or len(raw) > 80 or bool(TARGET_SENSITIVE_PATTERN.search(raw))
            or any(ord(ch) < 32 or ord(ch) == 127 for ch in raw))


def utc_iso(value):
    if value is None or value == '':
        return ''
    try:
        if isinstance(value, (int, float)):
            return dt.datetime.fromtimestamp(float(value), dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        s = str(value).strip()
        x = dt.datetime.fromisoformat(s.replace('Z', '+00:00'))
        if x.tzinfo is None:
            x = x.replace(tzinfo=dt.timezone.utc)
        return x.astimezone(dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    except (TypeError, ValueError, OverflowError):
        return ''


def stage_at(source_utc, submission_utc):
    if not source_utc or not submission_utc:
        return 'Full'
    delta = (dt.datetime.fromisoformat(source_utc.replace('Z','+00:00'))
             - dt.datetime.fromisoformat(submission_utc.replace('Z','+00:00'))).total_seconds() / 60
    for label, minutes in EVIDENCE_STAGES:
        if delta <= minutes:
            return label
    return 'Full'


def nested(value, field):
    if value is None or value == '':
        return []
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError as exc:
            raise ValueError(f'{field} cannot be parsed as JSON') from exc
    if not isinstance(value, list):
        raise TypeError(f'{field} must be a list, got {type(value).__name__}')
    if any(not isinstance(x, dict) for x in value):
        raise TypeError(f'{field} contains a non-object entry')
    return value


def reddit_id(v, kind):
    s = str(v or '').strip().lower()
    if s.startswith(('t1_','t3_')):
        if not s.startswith(kind + '_'):
            return ''
        s=s[3:]
    return s if BASE36.fullmatch(s) else ''


def safe_web_url(raw):
    """Conservative sanitation; even resulting URLs REQUIRE human approval."""
    if not isinstance(raw, str) or not raw.strip():
        return '', 'MISSING'
    try:
        u = urlsplit(raw.strip())
        if u.scheme.lower() not in {'https','http'} or not u.hostname or u.username or u.password:
            return '', 'BLOCKED_SCHEME_OR_CREDENTIALS'
        host = u.hostname.lower().strip('.')
        if host in {'localhost','local'} or host.endswith(('.local','.internal','.lan')):
            return '', 'BLOCKED_PRIVATE_HOST'
        try:
            if not ipaddress.ip_address(host).is_global:
                return '', 'BLOCKED_NONPUBLIC_IP'
        except ValueError:
            pass
        netloc = host + (f':{u.port}' if u.port and u.port not in (80,443) else '')
        clean = urlunsplit((u.scheme.lower(), netloc, u.path or '/', '', ''))
        if len(clean) > 1024:
            return '', 'BLOCKED_LENGTH'
        return clean, 'QUERY_REMOVED_REVIEW' if u.query or u.fragment else 'REVIEW_REQUIRED'
    except (ValueError, TypeError):
        return '', 'INVALID'


def digest(*parts, prefix='R'):
    s='\x1f'.join(str(p or '') for p in parts)
    return prefix + '-' + hashlib.sha256(s.encode('utf-8')).hexdigest()[:16]


def key_for(private_dir):
    path=private_dir/'private_target_key.hex'
    if path.exists():
        return bytes.fromhex(path.read_text(encoding='utf8').strip())
    key=secrets.token_bytes(32)
    path.write_text(key.hex()+'\n', encoding='utf8')
    try: path.chmod(0o600)
    except OSError: pass
    return key


def approved_strings(path, key):
    if not path.exists():
        with path.open('w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow([key,'approved_text'])
        return {}
    with path.open(newline='', encoding='utf-8-sig') as f:
        return {r.get(key,'').strip():r.get('approved_text','').strip()
                for r in csv.DictReader(f) if r.get(key,'').strip() and r.get('approved_text','').strip()}


def write_csv(path, columns, records):
    with path.open('w',newline='',encoding='utf-8') as f:
        w=csv.writer(f);w.writerow(columns);w.writerows(records)


def rows_from_jsonl(path):
    with path.open(encoding='utf-8') as f:
        for lineno,line in enumerate(f,1):
            if not line.strip():continue
            v=json.loads(line)
            if not isinstance(v,dict):
                raise ValueError(f'Row {lineno} is not an object')
            yield v


def as_excel_value(val):
    # Write as strings, never Excel formulas, URLs-as-formulas or hyperlinks.
    if val is None:return ''
    if isinstance(val,(int,float)) and not isinstance(val,bool):
        return val if math.isfinite(float(val)) else ''
    return str(val).replace('\x00','')[:31000]


def write_workbook(path, sheets):
    # XlsxWriter is available via pip in standard Google Colab (unlike the prior
    # internal-only workbook runtime); constant-memory limits resource use.
    wb=xlsxwriter.Workbook(str(path), {'constant_memory': True, 'strings_to_formulas': False,
                                      'strings_to_urls': False, 'default_date_format':'yyyy-mm-dd'})
    wb.set_properties({'title':'Study 4 claim-linked reconstruction dataset — REVIEW',
                       'comments':'Restricted source text withheld. Check release approvals before publishing.'})
    header=wb.add_format({'bold':True,'bg_color':'#203A5A','font_color':'#FFFFFF',
                          'text_wrap':True,'valign':'vcenter','border':0})
    for name, records in sheets.items():
        sh=wb.add_worksheet(name)
        head, body=records[0],records[1:]
        sh.freeze_panes(1,0)
        sh.set_row(0,32)
        sh.set_column(0,max(0,len(head)-1),23)
        for idx,col in enumerate(head):
            if col in {'source_url','reddit_post_url','reddit_comment_url','target','note','description'}:
                sh.set_column(idx,idx,53)
            if col.endswith('_json'):
                sh.set_column(idx,idx,48)
            sh.write_string(0,idx,str(col),header)
        for r,record in enumerate(body,1):
            for c,value in enumerate(record):
                v=as_excel_value(value)
                if isinstance(v, (int,float)) and not isinstance(v,bool):
                    sh.write_number(r,c,v)
                else:
                    sh.write_string(r,c,str(v))
        if body:
            sh.autofilter(0,0,len(body),len(head)-1)
    wb.close()


def build(input_file, public_dir, private_dir, allow_other_count=False):
    input_file=input_file.expanduser().resolve()
    public_dir=public_dir.expanduser().resolve()
    private_dir=private_dir.expanduser().resolve()
    if not input_file.is_file():
        raise FileNotFoundError(f'Source JSONL not found: {input_file}')
    if private_dir == public_dir or public_dir in private_dir.parents or private_dir in public_dir.parents:
        raise ValueError('Private directory must be OUTSIDE the public-review directory tree')
    public_dir.mkdir(parents=True,exist_ok=True)
    private_dir.mkdir(parents=True,exist_ok=True)

    records=list(rows_from_jsonl(input_file))
    mids=[str(r.get('message_id') or '').strip() for r in records]
    if any(not mid for mid in mids) or len(set(mids))!=len(mids):
        raise ValueError('Missing or duplicate internal message_id in source')
    if len(records)!=1069 and not allow_other_count:
        raise ValueError(f'Expected exactly 1,069 claims; found {len(records)}')
    hmac_key=key_for(private_dir)
    approved_targets=approved_strings(private_dir/'approved_targets.csv','target_ref')
    approved_queries=approved_strings(private_dir/'approved_queries.csv','query_ref')

    claims=[['checkmate_message_id','submission_month_utc','label_binary','split','root_tesa_present',
             'web_evidence_count','reddit_post_count','reddit_comment_count']]
    tesa=[['checkmate_message_id','tesa_ref','source_type','reddit_post_id','reddit_comment_id',
           'target_ref','target','target_review_status','entity_type','emotion','stance','n_mentions']]
    posts=[['checkmate_message_id','reddit_post_id','reddit_post_url','subreddit','created_utc',
            'first_available_stage','relevance_basis_REDACTED','relevance_confidence','n_comments_in_view']]
    comments=[['checkmate_message_id','reddit_post_id','reddit_comment_id','parent_reddit_id',
               'reddit_comment_url','created_utc','first_available_stage',
               'comment_rank','tesa_annotation_count']]
    web=[['checkmate_message_id','web_ref','rank','source_url','url_review_status','source_domain',
          'source_date_utc','date_source','first_available_stage','relevance_score','tavily_score','query_ref']]
    queries=[['checkmate_message_id','query_ref','approved_search_query','query_privacy_status','web_ref']]
    private_times=[];private_targets=[];private_queries=[];private_urls=[]
    bad=Counter();counts=Counter()

    def add_tesa(cid, annotations, kind, post_id='', comment_id=''):
        items=nested(annotations,'dict_tesa')
        for index,a in enumerate(items):
            orig=str(a.get('target') or '').strip()
            tref='TARGET-'+hmac.new(hmac_key,orig.casefold().encode('utf8'),hashlib.sha256).hexdigest()[:16]
            ref=digest(cid,kind,post_id,comment_id,index,prefix='TESA')
            needs_review = target_requires_review(orig)
            if needs_review:
                # Private explicit whitelist only; an approval should contain a vetted
                # public-facing target name, never a copied original claim body.
                target_text = approved_targets.get(tref, '')
                target_status = 'APPROVED_MANUALLY' if target_text else 'WITHHELD_REVIEW'
                if not target_text:
                    counts['targets_withheld_review'] += 1
            else:
                target_text = orig
                target_status = 'VISIBLE'
                counts['targets_visible'] += 1
            tesa.append([cid,ref,kind,post_id,comment_id,tref,target_text,target_status,
                         a.get('spacy_label') or '',a.get('emotion') or '',
                         a.get('stance') or '',a.get('n_mentions') or ''])
            private_targets.append([ref,cid,kind,post_id,comment_id,tref,orig,target_status])
            counts['tesa']+=1
        return len(items)

    for record in records:
        mid=str(record['message_id']).strip()
        cid=mid  # Original CheckMate message ID is the cross-sheet join key.
        submitted=utc_iso(record.get('submission_time'))
        p_list=nested(record.get('dict_reddit_threads'),'dict_reddit_threads')
        w_list=nested(record.get('dict_kb'),'dict_kb')
        root_list=nested(record.get('dict_tesa'),'dict_tesa')
        n_comments=sum(len(nested(p.get('dict_comments'),'dict_comments')) for p in p_list)
        claims.append([cid,submitted[:7] if submitted else '',record.get('label_binary'),
                       record.get('split') or record.get('train_test_split') or '',
                       bool(root_list),len(w_list),len(p_list),n_comments])
        private_times.append([cid,submitted])
        add_tesa(cid,root_list,'claim')

        for post in p_list:
            pid=reddit_id(post.get('submission_id'),'t3')
            if not pid:bad['posts']+=1
            created=utc_iso(post.get('created_datetime') or post.get('created_utc'))
            comm=nested(post.get('dict_comments'),'dict_comments')
            conf=post.get('relevance_confidence')
            # relevance_basis is free text and may repeat restricted claim text.
            # Only a fixed enum is released; otherwise leave blank.
            basis=str(post.get('relevance_basis') or '').strip().lower()
            allowed_basis=basis if basis in {'strict','direct','contextual','exact'} else ''
            posts.append([cid,pid,f'https://www.reddit.com/comments/{pid}/' if pid else '',
                          post.get('subreddit') or '',created,stage_at(created,submitted),
                          allowed_basis,conf if isinstance(conf,(int,float,bool)) else '',len(comm)])
            add_tesa(cid,post.get('dict_thread_tesa'),'reddit_post',pid)
            for comment in comm:
                comment_id=reddit_id(comment.get('comment_id'),'t1')
                if not comment_id:bad['comments']+=1
                parent=str(comment.get('parent_id') or '').strip().lower()
                if parent.startswith(('t1_','t3_')) and reddit_id(parent,parent[:2]):
                    pass
                else:parent=''
                ctime=utc_iso(comment.get('created_datetime') or comment.get('created_utc'))
                tcount=add_tesa(cid,comment.get('dict_comment_tesa'),'reddit_comment',pid,comment_id)
                url=f'https://www.reddit.com/comments/{pid}/_/{comment_id}/' if pid and comment_id else ''
                comments.append([cid,pid,comment_id,parent,url,ctime,stage_at(ctime,submitted),
                                 comment.get('comment_rank_within_thread') or '',tcount])

        for index,kb in enumerate(w_list,1):
            raw_url=kb.get('kb_url') or ''
            url,status=safe_web_url(raw_url)
            rank=kb.get('kb_rank_within_message') or index
            wref=digest(cid,rank,raw_url,prefix='WEB')
            qref=digest(cid,wref,prefix='QUERY')
            # SOURCE URLs need line-by-line privacy/ToS review before GitHub release.
            private_urls.append([cid,wref,raw_url,url,status])
            raw_query=str(kb.get('q_combined') or '')
            private_queries.append([cid,qref,wref,raw_query,str(kb.get('q_priority') or '')])
            date=utc_iso(kb.get('url_datetime'))
            domain=urlsplit(url).hostname or ''
            web.append([cid,wref,rank,url,status,domain,date,kb.get('url_datetime_source') or '',
                        stage_at(date,submitted),kb.get('kb_relevance_score') or '',
                        kb.get('tavily_score') or '',qref])
            queries.append([cid,qref,approved_queries.get(qref,''),
                            'APPROVED' if qref in approved_queries else
                            ('REVIEW_REQUIRED' if raw_query else 'NOT_RECORDED'),wref])

    computed={name:len(rows)-1 for name,rows in [('Claims',claims),('Reddit_Posts',posts),
                         ('Reddit_Comments',comments),('Web_Evidence',web),('TESA',tesa)]}
    if not allow_other_count:
        for key,expected in EXPECTED.items():
            if computed[key]!=expected:
                raise ValueError(f'{key}: found {computed[key]}, expected {expected}. '
                                 'STOP: verify Step 16 snapshot; no release workbook written.')
    if bad['posts'] or bad['comments']:
        raise ValueError(f'Invalid Reddit IDs: {dict(bad)}; no release workbook written')

    # IDs already appear in the review workbook; withhold exact submission timestamps.
    write_csv(private_dir/'PRIVATE_submission_times.csv',
              ['checkmate_message_id','submission_time_utc_PRIVATE'],private_times)
    write_csv(private_dir/'PRIVATE_target_review.csv',
              ['tesa_ref','checkmate_message_id','source_type','reddit_post_id','reddit_comment_id',
               'target_ref','original_target_PRIVATE','export_status'],private_targets)
    write_csv(private_dir/'PRIVATE_query_review.csv',
              ['checkmate_message_id','query_ref','web_ref','original_query_PRIVATE','original_priority_PRIVATE'],private_queries)
    write_csv(private_dir/'PRIVATE_web_url_review.csv',
              ['checkmate_message_id','web_ref','original_url_PRIVATE','sanitized_url','status'],private_urls)
    (private_dir/'DO_NOT_PUBLISH.txt').write_text(
        'This folder contains exact CheckMate timestamps, original search queries, URLs and a TESA target audit, including withheld entries.\n'
        'Keep it OUTSIDE GitHub and shared/public folders. Preserve private_target_key.hex for stable IDs.\n',encoding='utf8')

    codebook=[['sheet','field','definition'],
      ['All','checkmate_message_id','Original CheckMate message ID, used in all six data sheets as the cross-sheet join key. NO original claim text is exported. Public release requires explicit CheckMate approval.'],
      ['Claims','submission_month_utc','Year and month only; exact CheckMate timestamps withheld.'],
      ['Claims','label_binary','Research outcome label; CheckMate approval required for release.'],
      ['TESA','tesa_ref','Unique identifier per annotation; join using checkmate_message_id, source_type and Reddit post/comment IDs where relevant.'],
      ['TESA','target_ref','Stable keyed reference for target grouping and review of flagged values.'],
      ['TESA','target','Original extracted target text for ordinary entity names; suspicious/long/URL/email-like values withheld for manual review. Never a whole claim text column.'],
      ['TESA','target_review_status','VISIBLE / WITHHELD_REVIEW / APPROVED_MANUALLY. A withheld target may be added after explicit review with a private approved_targets.csv.'],
      ['TESA','emotion / stance','Stored target-directed TESA labels; not manually relabelled.'],
      ['Reddit_Posts','reddit_post_id','Public submission base36 ID; not WhatsApp propagation.'],
      ['Reddit_Comments','reddit_comment_id','Public comment base36 ID; no body or author released.'],
      ['Web_Evidence','source_url','HTTP(S) URL with query and fragment stripped; review before release.'],
      ['Web_Evidence','url_review_status','REVIEW_REQUIRED even if sanitized; linkability risk remains.'],
      ['Queries','approved_search_query','Blank until manually reviewed; original query is PRIVATE.'],
      ['All','first_available_stage','Earliest Step 17-style time window relative to submission; missing dates -> Full.']]
    provenance=[['category','source / rule','description'],
      ['source','Step 16 / checkmate_s4_message_inspection_view.jsonl','1,069 nested CheckMate claim records, not public.'],
      ['labels','Step 11 / original research labels','Original split and binary labels retained; release needs approval.'],
      ['reddit','Step 13 / strict relevance','Related Reddit discussion, NOT observed forwarding history.'],
      ['web','Step 12 / relevance-screened evidence','Up to five retrieved pages per claim; no veracity labels.'],
      ['TESA','Step 15b / target-based emotion and stance','Original extracted entity target strings shown except suspicious targets withheld for source/privacy review.'],
      ['stages','0m, 1m, 15m, 20m, 60m, 90m, 24h, 48h, Full','For missing timestamps, only Full considered available.'],
      ['privacy','No original CheckMate claim-text column, author data, exact submission time or raw query','Original CheckMate IDs and most extracted targets are explicitly included, increasing linkability. Manual review and CheckMate approval required.'],
      ['reproduction','Reddit post/comment IDs and sanitized URLs','Actual source content must be recollected lawfully if still accessible.'],
      ['gate','CheckMate written clearance','Required before public redistribution of original CheckMate message IDs and claim-linked labels.'],
      ['gate','Manual URL/target/linkability screening','Required before posting to GitHub.'],
      ['status','HOLD: REVIEW ONLY','Not publication-approved; no GitHub upload performed.']]
    qa=[['test','value','expected / decision'],
       ['Claims',computed['Claims'],'1069'],
       ['Reddit_Posts',computed['Reddit_Posts'],'590'],
       ['Reddit_Comments',computed['Reddit_Comments'],'17416'],
       ['Web_Evidence',computed['Web_Evidence'],'3568'],
       ['TESA',computed['TESA'],'Claim, post and comment labels combined'],
       ['Distinct CheckMate message IDs',len(set(row[0] for row in claims[1:])),'Must match Claims'],
       ['Original claim text exported',0,'MUST remain 0'],
       ['Invalid Reddit IDs',sum(bad.values()),'0'],
       ['TESA targets visible',counts['targets_visible'],'Plain extracted target names, not pseudonymised'],
       ['TESA targets withheld for review',counts['targets_withheld_review'],'Potential claim text / URL / email / unusually long target; inspect PRIVATE_target_review.csv'],
       ['Reviewed URL approvals',0,'NOT yet verified live or authorised'],
       ['Release approval','HOLD','CheckMate, privacy, platform and source checks remain']]
    sheets={'Claims':claims,'TESA':tesa,'Reddit_Posts':posts,
            'Reddit_Comments':comments,'Web_Evidence':web,'Queries':queries,
            'Codebook':codebook,'Provenance':provenance,'QA':qa}
    # Assert that all claim-linked data sheets use original IDs and that no
    # source claim-content columns have crept into the output schema.
    mid_set=set(mids)
    for name,rows in sheets.items():
        if name in {'Claims','TESA','Reddit_Posts','Reddit_Comments','Web_Evidence','Queries'}:
            assert rows[0][0]=='checkmate_message_id', (name, rows[0])
            if not all(row[0] in mid_set for row in rows[1:]):
                raise ValueError(f'{name} has a join key not in original CheckMate IDs')
            forbidden={'claim_text','original_claim_text','original_text','message_text','detection_text','raw_claim_text'}
            if forbidden.intersection(set(rows[0])):
                raise ValueError(f'{name} would export a claim-text column')
    target=public_dir/'Study4_reconstruction_dataset_TARGETS_VISIBLE_REVIEW.xlsx'
    write_workbook(target,sheets)
    if not target.exists() or target.stat().st_size < 10000:
        raise RuntimeError('Excel file missing or implausibly small')
    import zipfile
    with zipfile.ZipFile(target) as z:
        badzip=z.testzip()
        if badzip:raise RuntimeError('Workbook ZIP integrity test failed: '+badzip)
    manifest={'output':str(target),'status':'HOLD_REVIEW_ONLY',
              'identifier_policy':'ORIGINAL_CHECKMATE_MESSAGE_ID_EXPLICIT_RELEASE_APPROVAL_REQUIRED',
              'original_claim_text_column_exported':False,
              'targets_visible':counts['targets_visible'],
              'targets_withheld_review':counts['targets_withheld_review'],
              'target_policy':'VISIBLE_BY_DEFAULT_EXCEPT_SUSPICIOUS_LONG_URL_EMAIL_OR_SOURCE_TEXT',
              'counts':computed,
              'sheets':{k:len(v)-1 for k,v in sheets.items()},
              'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
              'private_dir':str(private_dir)}
    (public_dir/'RELEASE_MANIFEST_REVIEW.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf8')
    (public_dir/'REVIEW_BEFORE_GITHUB.txt').write_text(
        'DO NOT UPLOAD YET. This is a review candidate, not permission to redistribute.\n'
        '1) Obtain CheckMate written permission to publish ORIGINAL message IDs and claim-linked metadata/labels.\n'
        '2) Review every Web_Evidence.source_url (even sanitized URLs).\n'
        '3) Readable TESA target names are present. Review all PERSON entities, withheld target values and claim-linkability risks.\n'
        '4) Check Reddit and webpage terms and any third-party restrictions.\n'
        '5) Verify all source counts, labels and stage cutoffs against Study 4.\n'
        '6) Verify NO original claim text appears in any published sheet, including via TESA target values.\n'
        '7) Unflagged target names are automatically included, but that is NOT privacy clearance. Manually review before release.\n'
        '8) After approval, publish ONLY the approved workbook and documentation.\n'
        'NEVER upload source JSONL or PRIVATE_DO_NOT_PUBLISH/ to GitHub.\n',encoding='utf8')
    print(json.dumps(manifest,indent=2))
    return target


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--input',type=Path,default=DRIVE_STUDY4/SOURCE_RELATIVE)
    p.add_argument('--public-dir',type=Path,default=DRIVE_STUDY4/'GitHub_Release_REVIEW')
    p.add_argument('--private-dir',type=Path,default=DRIVE_STUDY4/'PRIVATE_DO_NOT_PUBLISH')
    p.add_argument('--allow-other-count',action='store_true',help='Override snapshot counts ONLY after investigating differences')
    args=p.parse_args()
    build(args.input,args.public_dir,args.private_dir,args.allow_other_count)

if __name__=='__main__':main()
