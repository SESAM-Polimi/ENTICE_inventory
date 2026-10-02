"""Prepare one consolidated classification view from maintained and source tables.

Run from the repository root. The JSON is a build intermediate, not a second
editable mapping source. The companion JS builder produces the Excel view.
"""
import argparse,json,sys
from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from entice_inventory.core.registry import Registry
from entice_inventory.core.paths import project_data_root
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--output-dir',type=Path,default=ROOT/'build/classifications')
args=parser.parse_args()
root=project_data_root(required=True);work=args.output_dir
work.mkdir(parents=True,exist_ok=True)
reg=Registry.load();gtap=reg.adapters['GTAP12'];historical={r['sector_id']:r for r in gtap['records']}
def rows(path,sheet):
 w=load_workbook(path,read_only=True,data_only=True)
 a=[list(r) for r in w[sheet].values];w.close();return a
sheets=[]
def add(name,headers,data,widths,note=''):
 sheets.append({'name':name,'headers':headers,'rows':data,'widths':widths,'note':note})
records=[];scopes=[]
for s in reg.sectors.values():
 code=s['id'];c=s['classification'];h=historical[code];p=reg.parent(code)
 records.append([code,s['name'],'Published' if s['lifecycle']=='published_d24' else 'Inactive',', '.join(p['nace_codes']),p['parent'] or '',p['status'],', '.join(dict.fromkeys(r['parent'] for r in h['legacy_matching_rows'])),h['published_parent'] or '',', '.join(h['declared_nace_codes']),', '.join(dict.fromkeys(r['pipeline'] for r in h['legacy_matching_rows'] if r['pipeline'])),c['note'],'; '.join(c['evidence'])])
 for scope in c['scopes']:
  pr=reg.parent(code,scope=scope['id'])
  for nace in scope['codes']:scopes.append([code,scope['id'],scope['label'],nace,pr['parent'] or '',pr['status'],'; '.join(c['evidence'])])
add('ENTICE sectors',['Code','Activity','D2.4 status','Reviewed NACE Rev. 2','NACE-derived GTAP12 parent','Resolution','Legacy pipeline parents','D2.4 parent','Declared NACE (unreviewed)','Source pipeline','Scope notes','Classification evidence'],records,[12,65,15,22,24,25,24,16,26,21,95,75], 'Review snapshot. Blank reviewed mappings remain unresolved. Pipeline mappings are preserved separately in Sector; changing a reviewed parent requires rebuilding dependent data.')
add('Activity scopes',['Code','Scope ID','Activity scope','NACE Rev. 2','GTAP12 parent','Resolution','Evidence'],scopes,[12,30,85,18,20,27,90])
add('GTAP sectors',['GTAP12 Power code','Sector name'],[[r['id'],r['name']] for r in gtap['parent_catalogue']],[22,95], '76 sector identifiers from the D2.4 baseline. Case is significant for electricity sectors.')
bridgepath=root/'Classifications/_sources/Purdue/GTAP_EXIOBASE_maps.xlsx'
geo_map={str(r[0]).upper():r[1] for r in rows(bridgepath,'Regions_map')[1:] if r[0]}
add('Regions',['GTAP12 code','Region name','EXIOBASE source mapping','Mapping status'],[[r['id'],r['name'],geo_map.get(r['id'],''),'Inherited crosswalk; scope not certified'] for r in reg.regions()],[19,65,30,50], '163 canonical GTAP regions, including MRT. EXIOBASE mappings reproduce the Purdue source and are not a validated EXIOBASE adapter.')
g=reg.geographies['GTAP12'];members=[]
for group in g['groups']:
 for member in group['members']:members.append([group['id'],member,group['semantics'],'Complete target universe' if group['id']=='GLOBAL' else 'Inherited membership; not independently reviewed'])
add('Region groups',['Group','GTAP12 member','Meaning','Review status'],members,[35,20,35,60])
ex=[]
for i,r in enumerate(rows(bridgepath,'Sectoral_maps')[1:],2):
 if r[2]:ex.append(['GTAP-ENTICE',str(r[2]),r[4],r[1],r[3],f'GTAP_EXIOBASE_maps:Sectoral_maps row {i}'])
 if len(r)>13 and r[12]:ex.append(['EXIOBASE',str(r[12]),r[10],r[13],r[11],f'GTAP_EXIOBASE_maps:Sectoral_maps row {i}'])
add('EXIOBASE bridge',['System','Code','Description','Mutual sector group','Cost group or original code','Source'],ex,[20,20,85,30,32,65], 'Inherited common-group crosswalk. Membership in the same group does not establish a one-to-one sector mapping or a NACE classification.')
legacy=root/'Classifications/_sources/Sectors list GTAP-NACE-ENTICE.xlsx'
lr=rows(legacy,'entice new sectors');add('Legacy concordance',lr[0],lr[1:],[20,16,18,17,17,22,18,18,20,20,80], 'Historical source labels explicitly refer to GTAP9/11/11Power. NACE entries here are declarations, not reviewed GTAP12 classifications.')
reviewed_by_parent=defaultdict(set)
for rule in gtap['bridge']:reviewed_by_parent[rule['parent']].add(rule['nace_code'])
legacy_by_parent=defaultdict(set)
for row in lr[1:]:
 if row[0] is not None and row[5]:legacy_by_parent[str(row[5]).upper()].add(str(row[0]))
target_sheet=next(s for s in sheets if s['name']=='GTAP sectors')
target_sheet['headers']+=['Reviewed NACE scopes (partial)','Legacy NACE declarations (GTAP11 Power)']
target_sheet['widths']+=[40,65]
for row in target_sheet['rows']:
 row.extend([', '.join(sorted(reviewed_by_parent[row[0]])),', '.join(sorted(legacy_by_parent[row[0].upper()]))])
target_sheet['note']+=' Reviewed scopes are a partial bridge, not a complete NACE concordance; GTAP11 declarations require review before reuse.'

# Partner templates also contain classifications and custom input/region groups.
# Preserve distinct rows with source IDs rather than selecting one template.
vocabulary=defaultdict(list);partner_sources=[]
for index,p in enumerate(sorted((root/'Data collection/Inventory cleaning/Partner inventories').glob('*.xlsx')),1):
 source_id=f'P{index:02d}'
 partner_sources.append([source_id,'Partner inventories/'+p.name,'Preserved D2.4 template dictionaries; current partner submissions may differ'])
 w=load_workbook(p,read_only=True,data_only=True)
 for kind in ['Sector','Region','Primary input','Emissions','Energy']:
  if kind not in w.sheetnames:continue
  for line,row in enumerate(w[kind].values,1):
   if line<3 or not any(v is not None for v in row):continue
   padded=list(row)+[None]*5
   group,label,code,description,notes=padded[:5]
   if kind!='Sector':notes=description;description=None
   key=tuple('' if v is None else str(v) for v in [kind,group,label,code,description,notes])
   if source_id not in vocabulary[key]:vocabulary[key].append(source_id)
 w.close()
add('Partner vocabulary',['Dictionary','Input or group','Activity or member','Declared target code','Target description','Original note','Source IDs'],[list(k)+[', '.join(v)] for k,v in vocabulary.items()],[22,35,70,25,65,70,80], 'Union of the five classification/group dictionaries in 36 preserved partner templates. Custom examples and historical labels are retained as evidence, not activated mappings.')
working=root/'Classifications/_sources/New sectors classification.xlsx'
catalogue=[]
for row in rows(working,'New GTAP sectors list')[1:]:
 if not row[0]:continue
 code=str(row[0]);status=reg.parent(code)['status'] if code in reg.sectors else 'Outside current registry'
 catalogue.append(list(row[:5])+[status])
add('Working catalogue',['Code','Activity','Split in D2.4 (source)','GTAP parent (source)','Parent description','Registry review status'],catalogue,[15,85,24,24,65,35], 'Current working classification, distinct from the published release catalogue. Source split flags and parent declarations are preserved, not activated or reclassified.')
declarations=[]
for name in ['New GTAP - CPC concordances','New GTAP - NACE-ISIC levels','GTAP-CPC','GTAP-ISIC-NACE']:
 for row in rows(working,name)[1:]:
  if not any(v is not None for v in row):continue
  row=list(row)+[None]*4
  if name=='New GTAP - CPC concordances':out=[name,row[1],row[2],row[0],None,None,row[3]]
  elif name=='New GTAP - NACE-ISIC levels':out=[name,row[1],row[2],None,row[0],row[0],row[3]]
  elif name=='GTAP-CPC':out=[name,None,row[0],row[1],None,None,row[2]]
  else:out=[name,None,row[0],None,row[1],row[3],row[2]]
  declarations.append([None if v is None else str(v) for v in out])
add('Source concordances',['Source sheet','ENTICE activity','GTAP code (source)','CPC code','ISIC or shared level','NACE or shared level','Source description'],declarations,[35,85,25,20,25,25,85], 'Inherited declarations from the working classification. Shared NACE/ISIC levels reproduce the source claim; codes, decimal separators and partial scopes have not been normalised or reviewed. Not the live NACE-to-GTAP bridge.')
for name,widths in [('Sector',[75,25,18,25,20]),('Factor of production',[55,25,25])]:
 rr=rows(ROOT/'data/GTAP12_matching.xlsx',name)
 if name=='Factor of production':rr[0][0]='Factor'
 add(name,rr[0],rr[1:],widths, 'Exact compatibility table from the versioned GTAP12_matching.xlsx. Existing pipeline behaviour is preserved; this table does not certify NACE mappings.')
add('Sources',['Source','Location or reference','Role'],[
 ['Canonical registry','GitHub: data/registry/{sectors,gtap12,regions}.json','Maintained activity identities, reviewed NACE scopes and target geography'],
 ['Pipeline mappings','GitHub: data/GTAP12_matching.xlsx','Legacy operational sector/factor lookup; reproduced without alteration'],
 ['Historical concordances','_sources/Sectors list GTAP-NACE-ENTICE.xlsx','Original GTAP9/11 declarations; retained for provenance'],
 ['Purdue EXIOBASE crosswalk','_sources/Purdue/GTAP_EXIOBASE_maps.xlsx','Mutual sector grouping and regional crosswalk; not an implemented adapter'],
 ['Detailed Purdue mappings','_sources/Purdue/EXIOBASE_maps.xlsx','Additional working mappings retained in original layout'],
 ['Trade concordances','_sources/Purdue/GTAP sectors H5.xlsx; _sources/Purdue/GTAP_HS_mapping.xlsx','HS versions and trade product mappings, retained as specialist sources'],
 ['GTAP-HS reference','_sources/GTAP HS concordance.xlsx','Versioned commodity concordances, retained in their original workbook'],
 ['Working sector classification','_sources/New sectors classification.xlsx','Working catalogue and source CPC/NACE/ISIC declarations; distinct from the published snapshot used by the canonical registry'],
 ['Region grouping source','_sources/region_clustering.xlsx','Historical clusters; canonical GLOBAL restores the complete target universe'],
 ['NACE-ISIC and other source files','_sources/','Supporting classification evidence retained without overwriting'],
 ['Updating this workbook','Regenerate from the versioned registry and preserved sources','Review corrections in the registry first. This workbook is a consolidated view, not a second independently editable mapping source.']
]+partner_sources,[32,90,105])
for s in sheets:
 n=len(s['headers']);s['rows']=[(r+[None]*n)[:n] for r in s['rows'] if any(x is not None for x in r)]
(work/'classification_tables.json').write_text(json.dumps({'sheets':sheets},indent=2,default=str)+'\n')
print([(s['name'],len(s['rows'])) for s in sheets])
