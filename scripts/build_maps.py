import json,math,subprocess,concurrent.futures
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).resolve().parents[1];d=json.loads((R/'data.json').read_text());p={x['id']:x for x in d['places']};cache=R/'.map-cache';cache.mkdir(exist_ok=True)
font=ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf',20)
small=ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf',15)
def xy(lat,lon,z):return ((lon+180)/360*2**z*256,(1-math.asinh(math.tan(math.radians(lat)))/math.pi)/2*2**z*256)
def tile(t):
 z,x,y=t;f=cache/f'{z}-{x}-{y}.png'
 if not f.exists():subprocess.run(['curl','-L','--max-time','20','-sS','-A','HKFieldTripHandbook/1.0 (educational static map)','https://tile.openstreetmap.org/%s/%s/%s.png'%(z,x,y),'-o',str(f)],capture_output=True)
 try:return x,y,Image.open(f).convert('RGB')
 except:return x,y,None
def make(name,ids,z=None):
 points=[p[i] for i in dict.fromkeys(ids) if i and p[i]['coordinates']];w,h=1200,540
 if not z:
  for z in range(17,7,-1):
   pts=[xy(*a['coordinates'],z) for a in points]
   if max(a[0] for a in pts)-min(a[0] for a in pts)<w-210 and max(a[1] for a in pts)-min(a[1] for a in pts)<h-160:break
 pts=[xy(*a['coordinates'],z) for a in points];cx=(min(a[0] for a in pts)+max(a[0] for a in pts))/2;cy=(min(a[1] for a in pts)+max(a[1] for a in pts))/2;x0=int(cx-w/2);y0=int(cy-h/2)
 tiles=[(z,x,y) for x in range(x0//256,(x0+w)//256+1) for y in range(y0//256,(y0+h)//256+1)]
 im=Image.new('RGB',(w,h),'#e3e9e7');missing=0
 with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
  for x,y,t in ex.map(tile,tiles):
   if t:im.paste(t,(x*256-x0,y*256-y0))
   else:missing+=1
 draw=ImageDraw.Draw(im)
 used={}
 for i,(point,(x,y)) in enumerate(zip(points,pts),1):
  x-=x0;y-=y0; key=(round(x),round(y))
  if key in used:continue
  used[key]=i
  draw.ellipse((x-20,y-20,x+20,y+20),fill='#173d50',outline='white',width=3);draw.text((x,y),str(i),font=font,fill='white',anchor='mm')
 draw.rectangle((0,h-26,w,h),fill='white');draw.text((12,h-22),'Geographic reference points only | No road route implied',font=small,fill='#213e50');draw.text((w-360,h-22),'Map data © OpenStreetMap contributors',font=small,fill='#213e50')
 out=R/'assets/maps'/f'{name}.png';im.save(out);print(name,'tiles missing',missing,flush=True)
 return dict(image=str(out.relative_to(R)),legend=[dict(number=i+1,id=a['id'],name=a['name']) for i,a in enumerate(points)])
# An overview then actual daily geography; park day gets separate district and building detail.
d['overviewMap']=make('overview',['airport','hotel','park'])
for day in d['days']:
 ids=[e['place'] for e in day['events'] if e['place']]
 if day['day']==15: ids=['ai','regen','park']
 day['map']=make('day-'+str(day['day']),ids)
 if day['day']==14: day['map']['legend']=day['map']['legend'][:1];day['map']['note']='校园参考点；系所交流与校园导览的实际起点待确认。'
(R/'data.json').write_text(json.dumps(d,ensure_ascii=False,indent=2))
