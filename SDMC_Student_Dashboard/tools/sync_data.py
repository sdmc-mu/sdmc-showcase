"""Regenerate public student data from the Excel workbook and local student_photos folder."""
from pathlib import Path
from openpyxl import load_workbook
import json,re,shutil
ROOT=Path(__file__).resolve().parents[1]
BOOK=ROOT/'source'/'SDMC_Student_Portfolio_Track_Sheet_2026.xlsx'
PHOTO_DIR=ROOT/'student_photos'; PROFILE=ROOT/'assets'/'profiles'
PROGRAMMES=[('BTCM',4,'B.Tech Computational Media'),('BJMC',3,'BA Journalism & Mass Communication'),('MJMC',2,'MA Journalism & Mass Communication')]
def norm(s): return re.sub(r'[^a-z0-9]','',str(s or '').lower())
def slug(s): return re.sub(r'[^a-z0-9]+','-',str(s).lower()).strip('-')
old=json.loads((ROOT/'data/students.json').read_text(encoding='utf-8'))['students']; old_by_name={norm(x['name']):x for x in old}
wb=load_workbook(BOOK,data_only=True); students=[]
for code,duration,pname in PROGRAMMES:
 sheet=next((s for s in wb.sheetnames if s.startswith(code)),None)
 if not sheet: continue
 ws=wb[sheet]; header_row=None; cols={}
 for n,row in enumerate(ws.iter_rows(values_only=True),1):
  if any(str(v).strip().lower()=='student name' for v in row if v is not None):
   header_row=n; cols={str(v).strip().lower():i for i,v in enumerate(row) if v is not None}; break
 if not header_row: continue
 for rn,row in enumerate(ws.iter_rows(min_row=header_row+1,values_only=True),header_row+1):
  def val(k):
   i=cols.get(k); return row[i] if i is not None and i<len(row) else None
  name=str(val('student name') or '').strip()
  if not name: continue
  prev=old_by_name.get(norm(name),{})
  sid=str(val('btech id') or prev.get('id') or slug(name)).strip().lower()
  batch=str(val('batch/year') or prev.get('batch') or '').strip()
  admission=prev.get('admissionYear')
  if not admission:
   m=re.search(r'(20\d\d)',sid); admission=int(m.group(1)) if m else None
  raw=str(val('portfolio link') or '').strip(); url=raw if re.match(r'^https?://',raw,re.I) else prev.get('portfolioUrl','')
  skills=val('skills') or ''
  skill_list=[x.strip() for x in re.split(r'[,;\n]',str(skills)) if x.strip()] if skills else prev.get('skills',[])
  photo=prev.get('photo','')
  # New photo drop: folder/file name should contain the student ID; common image types supported.
  candidates=[]
  for f in PHOTO_DIR.rglob('*'):
   if f.is_file() and f.suffix.lower() in ('.jpg','.jpeg','.png','.webp') and sid in norm(str(f.parent.name)+' '+f.stem): candidates.append(f)
  if candidates:
   src=sorted(candidates)[0]; dest=PROFILE/(slug(sid)+src.suffix.lower()); shutil.copy2(src,dest); photo='assets/profiles/'+dest.name
  students.append({'id':sid,'name':name,'programme':code,'programmeName':pname,'batch':batch,'admissionYear':admission,'graduatingYear':admission+duration if admission else None,'skills':skill_list,'portfolioUrl':url if re.match(r'^https?://',url,re.I) else '', 'photo':photo})
(ROOT/'data/students.json').write_text(json.dumps({'updatedAt':__import__('datetime').date.today().isoformat(),'students':students},ensure_ascii=False,indent=2),encoding='utf-8')
payload={'updatedAt':__import__('datetime').date.today().isoformat(),'students':students}
(ROOT/'data/students.js').write_text('window.SDMC_STUDENTS = '+json.dumps(payload,ensure_ascii=False,separators=(',',':'))+';\\n',encoding='utf-8')
print(f"Updated dashboard data: {len(students)} students. Photos found: {sum(bool(x['photo']) for x in students)}")
