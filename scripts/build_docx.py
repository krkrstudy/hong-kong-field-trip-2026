"""Generate the editable handbook from the shared data.json. Requires python-docx, Pillow."""
import json, math, re
from pathlib import Path
from urllib.parse import quote
from PIL import Image,ImageOps
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
R=Path(__file__).resolve().parents[1]; D=json.loads((R/'data.json').read_text());DOC=Document();S=DOC.sections[0]
S.page_width=Cm(21);S.page_height=Cm(29.7);S.top_margin=Cm(1.9);S.bottom_margin=Cm(1.7);S.left_margin=Cm(2);S.right_margin=Cm(2);S.header_distance=Cm(.8);S.footer_distance=Cm(.8);S.different_first_page_header_footer=True
INK='17394D';GREEN='376B63';MUTED='596973';LINE='D9DEDC';WHITE='FFFFFF'
styles=DOC.styles
for name in ['Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3','Caption']:
 st=styles[name];st.font.name='Arial';st._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'STSong' if name in ['Normal','Caption'] else 'Arial Unicode MS');st.font.color.rgb=RGBColor.from_string(INK);st.paragraph_format.line_spacing=1.5
styles['Normal'].font.size=Pt(10.5);styles['Normal'].paragraph_format.space_after=Pt(6)
styles['Title'].font.size=Pt(30);styles['Title'].font.bold=True
styles['Subtitle'].font.italic=False
styles['Caption'].font.bold=False;styles['Caption'].font.italic=False
styles['Heading 1'].font.size=Pt(21);styles['Heading 1'].paragraph_format.space_before=Pt(0);styles['Heading 1'].paragraph_format.space_after=Pt(12)
styles['Heading 2'].font.size=Pt(14);styles['Heading 2'].paragraph_format.space_before=Pt(12);styles['Heading 2'].paragraph_format.space_after=Pt(5)
styles['Heading 3'].font.size=Pt(11);styles['Heading 3'].font.color.rgb=RGBColor.from_string(GREEN)
styles['Caption'].font.size=Pt(8);styles['Caption'].font.color.rgb=RGBColor.from_string(MUTED);styles['Caption'].paragraph_format.space_after=Pt(7)
for name in ['Heading 1','Heading 2','Heading 3']:styles[name].paragraph_format.keep_with_next=True
for name in ['Normal','Caption']:styles[name].paragraph_format.widow_control=True
for st in styles:
 for fonts in st.element.xpath('.//w:rFonts'):
  for key in list(fonts.attrib):
   if 'Theme' in key:del fonts.attrib[key]
head=S.header.paragraphs[0];head.text='香江翱翔  /  HONG KONG FIELD NOTES                              2026.10.11—10.17';head.paragraph_format.line_spacing=1;head.runs[0].font.size=Pt(8);head.runs[0].font.color.rgb=RGBColor.from_string(MUTED)
foot=S.footer.paragraphs[0];foot.alignment=WD_ALIGN_PARAGRAPH.RIGHT;r=foot.add_run('V4 · 2026-10-07     /     ');r.font.size=Pt(8)
f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');foot._p.append(f)
DOC.core_properties.title='香港访学精美路书';DOC.core_properties.subject=D['title'];DOC.core_properties.author='香港访港研学团';DOC.core_properties.keywords='香港 访学 香江翱翔 V4'
settings=DOC.settings.element;upd=OxmlElement('w:updateFields');upd.set(qn('w:val'),'true');settings.append(upd)
TMP=R/'.doc-build';TMP.mkdir(exist_ok=True)
page=1;toc=[];bk=0;pending_break=False

def para(text='',style=None,size=None,color=None,bold=False):
 p=DOC.add_paragraph(text,style);p.paragraph_format.line_spacing=1.5
 if size or color or bold:
  for r in p.runs:
   if size:r.font.size=Pt(size)
   if color:r.font.color.rgb=RGBColor.from_string(color)
   if bold:r.bold=True
 return p

def title(text,eyebrow=None,toc_entry=True):
 global bk,pending_break
 if eyebrow:
  ep=para(eyebrow,size=8.5,color=GREEN,bold=True)
  if pending_break:ep.paragraph_format.page_break_before=True;pending_break=False
 p=DOC.add_paragraph(text,'Heading 1')
 if toc_entry:
  bk+=1;name='chapter'+str(bk);start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(bk));start.set(qn('w:name'),name);end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(bk));p._p.insert(0,start);p._p.append(end);entry_text='地点深度介绍' if text=='香港国际机场' else text
  toc.append((entry_text,page,name))
  tc=OxmlElement('w:fldSimple');tc.set(qn('w:instr'),' TC "'+entry_text+'" \\f C \\l 1 ');rp=OxmlElement('w:r');props=OxmlElement('w:rPr');props.append(OxmlElement('w:vanish'));rp.append(props);tc.append(rp);p._p.append(tc)
 return p

def newpage():
 global page,pending_break
 pending_break=True;page+=1

def link(p,text,url):
 h=OxmlElement('w:hyperlink');h.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True));r=OxmlElement('w:r');pr=OxmlElement('w:rPr');c=OxmlElement('w:color');c.set(qn('w:val'),GREEN);pr.append(c);sz=OxmlElement('w:sz');sz.set(qn('w:val'),'17');pr.append(sz);r.append(pr);t=OxmlElement('w:t');t.text=text;r.append(t);h.append(r);p._p.append(h)

def picture(path,width=17,height=None,caption=None):
 path=Path(path)
 if height:
  im=Image.open(path).convert('RGB');im=ImageOps.fit(im,(1700,int(1700*height/width)),centering=(.5,.43));path=TMP/(path.stem+f'-{height}.jpg');im.save(path,quality=94)
 p=DOC.add_paragraph();p.paragraph_format.space_after=Pt(3);p.paragraph_format.line_spacing=1;p.add_run().add_picture(str(path),width=Cm(width));p.paragraph_format.keep_with_next=bool(caption)
 if caption:para(caption,'Caption')

def table(headers,rows,widths):
 t=DOC.add_table(rows=1, cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
 for c,w in zip(t.columns,widths):c.width=Cm(w)
 for i,h in enumerate(headers):t.rows[0].cells[i].text=h
 trpr=t.rows[0]._tr.get_or_add_trPr();rep=OxmlElement('w:tblHeader');rep.set(qn('w:val'),'true');trpr.append(rep)
 for row in rows:
  cs=t.add_row().cells
  for c,x in zip(cs,row):c.text=str(x)
 for ri,row in enumerate(t.rows):
  cant=OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(cant)
  for ci,c in enumerate(row.cells):
   c.width=Cm(widths[ci]);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;pr=c._tc.get_or_add_tcPr();sh=OxmlElement('w:shd');sh.set(qn('w:fill'),INK if ri==0 else ('F0F4F2' if ri%2 else WHITE));pr.append(sh)
   margins=OxmlElement('w:tcMar')
   for x,v in [('top','90'),('bottom','90'),('left','100'),('right','100')]:z=OxmlElement('w:'+x);z.set(qn('w:w'),v);z.set(qn('w:type'),'dxa');margins.append(z)
   pr.append(margins)
   for p in c.paragraphs:
    p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.5
    for r in p.runs:r.font.size=Pt(9);r.font.color.rgb=RGBColor.from_string(WHITE if ri==0 else INK);r.bold=ri==0
 borders=OxmlElement('w:tblBorders')
 for x in ['top','left','bottom','right','insideH','insideV']:
  v=OxmlElement('w:'+x);v.set(qn('w:val'),'single');v.set(qn('w:sz'),'4');v.set(qn('w:color'),LINE);borders.append(v)
 t._tbl.tblPr.append(borders)
 return t

def daytimes(p):return '；'.join('10月'+str(day['day'])+'日 '+e['time']+('—'+e['end'] if e['end'] else '') for day in D['days'] for e in day['events'] if e['place']==p['id'])
def refline(ids):
 p=para('资料核验  ','Caption')
 for i,s in enumerate(ids):
  if s=='schedule':continue
  source=next(x for x in D['sources'] if x['id']==s);link(p,source['title'],source['url']);p.add_run('  ')

# 1 Cover
para('PEKING UNIVERSITY  /  FIELD TRIP 2026',size=9,color=GREEN,bold=True)
para('香江翱翔','Title');para('香港访学与交流路书','Subtitle',size=18)
para('第五届北京大学“香江翱翔”访港研学团',size=13)
para('2026年10月11日至17日  ·  中国香港',size=11,color=MUTED)
picture(R/'assets/images/hero.jpg',height=10.8)
para('走访香港，结识师友\n2026年秋',size=17,color=INK)
para(D['organizers'],size=10)
para('依据香港校友会2026年10月7日V4正式行程编制\n信息核查日期 2026年10月8日',size=9,color=MUTED)
para('照片为香港城市实景，拍摄时间早于本次访问；版权与署名见文末。','Caption')
# 2 Contents
newpage();title('阅读导航','CONTENTS',False);toc_holder=DOC.add_paragraph()
para('使用说明','Heading 2');para('本手册按香港校友会10月7日V4行程编制。标注“待确认”的细节，请留意领队后续通知。所列时间均为香港当地时间，与北京没有时差。')
para('地图可以帮助大家辨认楼宇和园区；具体从哪里进入、在哪里报到，请按接待通知。图中编号对应访问地点，未画行车路线。')
para('手机上也能查看每日安排、地点介绍和导航。出发前建议下载一份PDF，网络不稳定时仍能查阅；需要打印的同学可直接使用本手册。')
link(para('数字路书  '),'krkrstudy.github.io/hong-kong-field-trip-2026/','https://krkrstudy.github.io/hong-kong-field-trip-2026/')
# 3 framing
newpage();title('行前寄语','01 / BEFORE THE JOURNEY')
para('2026年是“十五五”开局之年。习近平主席在2025年12月听取李家超述职报告时，要求香港主动对接国家“十五五”规划，推动高质量发展，更好参与粤港澳大湾区建设。香港2026年施政报告也对科研成果转化、产业发展和创新合作作出部署。此时走进香港，我们有机会从实际工作中认识这些发展议题。')
para('在北京大学香港校友会和香港北大助学基金会的支持下，来自北大多个院系的15位同学将一同赴港。七天里，我们会走访政府部门、大学、科研机构和金融企业，也会与在港校友相聚。不同专业的同学同行，正好可以互相补充：你熟悉的知识，也许能帮助身边的人听懂一场交流。')
para('希望大家珍惜这次见面的机会。听到感兴趣的地方，不妨多问一句；遇到不熟悉的概念，也不必急着下结论。认识一座城市，需要留意制度和产业，也需要听见在其中学习、工作和生活的人。愿这几天的所见所闻，能让我们对香港多一份理解，对自己的学习与未来多一份思考。')
para('此行可以多留意','Heading 2')
for text in ['公共事务：法案怎样讨论，人才来到香港后有哪些支援，企业落地会遇到什么问题。','科研与创业：一项研究怎样验证，一个产品怎样找到用户，团队又怎样寻求合作。','学习与选择：向师生和校友请教他们的经历，也和同行伙伴聊聊各自的困惑与打算。']:para(text)
refline(['policy','policy2026'])
picture(R/D['overviewMap']['image'],height=6.3,caption='区域概览：1 机场；2 港岛参访区域（以上环酒店为参考）；3 科学园。底图 © OpenStreetMap contributors。')
#4 overview
newpage();title('七日行程总览','02 / THE WEEK AT A GLANCE')
rows=[]
for day in D['days']:
 events='\n'.join(e['time']+('—'+e['end'] if e['end'] else '')+' '+e['title'] for e in day['events'])
 rows.append([f"10月{day['day']}日\n{day['weekday']}",day['theme'],events])
table(['日期','每日主题','V4正式安排'],rows,[2.2,3.3,11.5])
para('港大两项活动时间重叠；爱诗地址、晚宴与校友座谈地点、中诚信房号等详见各日提醒及待确认清单。','Caption')
#5-11 daily
for idx,day in enumerate(D['days'],1):
 newpage();title(f"10月{day['day']}日  {day['theme']}",f"DAY {idx:02d} / {day['weekday']} / {day['region']}")
 rows=[]
 for e in day['events']:
  info=e['title'];
  if e['note']:info+='\n'+e['note']
  rows.append([e['time']+('—'+e['end'] if e['end'] else ''),info,e['status']])
 table(['时间','活动与地点','状态'],rows,[3.2,11,2.8])
 para('集合与交通','Heading 2');para(day['transport'],size=9.5)
 # image kept small enough for dense park day
 picture(R/day['map']['image'],height=4.8 if day['day']==15 else 5.6)
 legend='  ·  '.join(str(a['number'])+' '+a['name'] for a in day['map']['legend'])
 para(legend,'Caption');para(day['map'].get('note','编号对应地点；不连线、不推测道路路线。底图 © OpenStreetMap contributors。'),'Caption')
 para('今天可以聊什么','Heading 3');para(day['task'],size=9.5)
 para('出门前记得  '+day['notice'],size=9.5,color=GREEN)
#12..25 profiles
for i,p in enumerate(D['places'],1):
 newpage();title(p['name'],f"PLACE {i:02d} / {p['category']}",toc_entry=i==1)
 para(p['en'],size=9.5,color=MUTED)
 para(daytimes(p),size=9,color=GREEN,bold=True)
 if p['image']:
  picture(R/p['image'],height=7.0,caption=p['photo']['caption'])
 else:
  para('场地照片待确认','Heading 2');para('本次接待地址尚未明确，场地照片也待确认。出发前请留意领队补充通知。')
 para('地址与入口','Heading 2');para(p['address'],size=10)
 para(p['intro'])
 para(p['topic'],'Heading 2')
 for q in p['questions']:para('• '+q,size=10)
 para('参访注意事项','Heading 2');para(p['entry'],size=10)
 if p['coordinates']:
  url='https://www.google.com/maps/search/?api=1&query='+quote(p['address']+' '+p['name']) if p['id']=='dymon' else 'https://www.google.com/maps/search/?api=1&query='+','.join(map(str,p['coordinates']))
  link(para('',style='Caption'),'打开位置导航',url)
 refline(p['sources'])
# practical
newpage();title('出发前，记得这些','PRACTICAL NOTES')
for h,t in [
('集合与出发','行程已安排接送机和10月15日科学园中巴，其他日期的交通与发车时间待通知。建议为点名、安检和找房间多留15—20分钟；这是准备时间的建议，实际集合时刻请听领队安排。'),
('公共交通备用','如需查询公共交通，可使用港铁、香港出行易和科技园官网。机场快线香港站至机场的列车行程约24分钟，酒店接驳、候车和机场手续还要另留时间；科学园可参考大学站转乘272K的方式，出发前再查班次。团队车辆有变动时，请先联系领队。'),
('随身带好','出发前核对证件、适用的入境许可、机票和保险，准备好在港通讯与支付方式。每天出门带上常用药、雨具和饮水用品，穿适合步行的鞋。证件随身保管，领队联系方式也请存进手机。'),
('天气与步行','出门前查看香港天文台天气警告，遇到恶劣天气听从团队和当地部门安排。过街时留意来车，使用行人设施；拍照请选安全位置，也给行人和出入口留出通道。'),
('见面与交流','有问题可以提前想一想，现场交流时也请给对方留出回答的时间。拍照、录音或转发资料前，先问一句是否方便。参观实验室与企业时，按工作人员指引活动，设备和屏幕请勿自行操作。'),
('需要帮助时','团队联络表由领队另行发放，请提前保存。途中遇到困难，及时联系领队或同行老师；如有紧急危险，可拨打香港999求助，说清所在位置和发生的情况。')]:
 para(h,'Heading 2');para(t)
refline(['mtr','park','mobility','weather','emergency'])
newpage();title('行前待确认清单','CHECK BEFORE DEPARTURE')
for i,issue in enumerate(D['issues'],1):para(f'{i}. {issue}')
para('请领队提前核对','Heading 2')
para('建议先向港大确认14日是否分组，再核对13日爱诗科技接待地址和交通。15日请与接待方商定17W至15W的步行路线、登记方式，以及下午访问结束后的上车安排。')
para('11日预计落地至晚宴有95分钟；17日酒店出发至预计起飞有175分钟。两段都还要计入机场手续和当天交通，请结合航班与车况复核。')
para('收到后续通知后','Heading 2');para('表中的“V4已列明”表示这项活动及时间已写入正式行程。具体报到入口、房间和集合方式，请以领队与接待方的后续通知为准；收到更新后，记得一并核对手机路书和自己的记录。')
# Sources: requested anew supersedes old removal
newpage();title('资料来源与核验说明','SOURCES')
para('活动时间按香港校友会2026年10月7日V4正式行程编排。想进一步了解机构或查询出行信息，可点击下列官方资料。',size=9)
for s in D['sources']:
 p=para('',style='Caption')
 if s['url']:link(p,s['title'],s['url'])
 else:p.add_run(s['title'])
 p.paragraph_format.space_after=Pt(0)
para('地图坐标来自OpenStreetMap对象，使用Nominatim查询，并以官方地址核对；公爵大厦参考Commons大堂实景定位。完整坐标证据与照片署名见配套项目data.json及sources.json。地政总署接口本次未能取得可用响应，未将其写作核验成功的依据。',size=9)
# image credits 2 pages
photo_records=[]
seen=set()
for p in D['places']:
 if p['photo'] and p['image'] not in seen:seen.add(p['image']);photo_records.append((p['name'],p['photo']))
h=D['heroPhoto'];photo_records.insert(0,('封面 香港城市景观',dict(author=h['author'],source=h['source_page_url'],license=h['license'],licenseUrl=h['license_url'],caption=h['description'])))
for batch in range(math.ceil(len(photo_records)/7)):
 newpage();title('图片版权与署名'+(' 续' if batch else ''),'IMAGE CREDITS',toc_entry=batch==0)
 para('所有图片为真实地点的历史照片。适配中进行了尺寸缩放与局部显示裁切；照片继续适用原许可，下载及再利用请遵守来源页要求。',size=9)
 for name,a in photo_records[batch*7:batch*7+7]:
  para(name,'Heading 3');para(a['author']+' / Wikimedia Commons / '+a['license'],size=9)
  p=para('',style='Caption');link(p,'原始图片与文件信息',a['source']);p.add_run('  ·  ');link(p,'许可条款',a['licenseUrl'])
 para('地图底图与数据 © OpenStreetMap contributors。所有静态地图均为真实地理底图上的参考点，不绘制道路路线；在线地图底图使用OpenStreetMap服务。',size=9)
# Seed a real TOC field with cached contents. Page numbers updated from verified PDF if needed.
start=OxmlElement('w:r');fc=OxmlElement('w:fldChar');fc.set(qn('w:fldCharType'),'begin');start.append(fc);toc_holder._p.append(start)
r=OxmlElement('w:r');instr=OxmlElement('w:instrText');instr.set(qn('xml:space'),'preserve');instr.text=' TOC \\f C \\h \\z ';r.append(instr);toc_holder._p.append(r)
r=OxmlElement('w:r');f=OxmlElement('w:fldChar');f.set(qn('w:fldCharType'),'separate');r.append(f);toc_holder._p.append(r)
prev=toc_holder._p
for text,pg,name in toc:
 p=OxmlElement('w:p');pr=OxmlElement('w:pPr');sp=OxmlElement('w:spacing');sp.set(qn('w:after'),'60');sp.set(qn('w:line'),'360');sp.set(qn('w:lineRule'),'auto');pr.append(sp);tabs=OxmlElement('w:tabs');tab=OxmlElement('w:tab');tab.set(qn('w:val'),'right');tab.set(qn('w:leader'),'dot');tab.set(qn('w:pos'),'9500');tabs.append(tab);pr.append(tabs);p.append(pr)
 h=OxmlElement('w:hyperlink');h.set(qn('w:anchor'),name);r=OxmlElement('w:r');tx=OxmlElement('w:t');tx.text=text;r.append(tx);h.append(r);p.append(h);r=OxmlElement('w:r');tab=OxmlElement('w:tab');r.append(tab);p.append(r);f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),' PAGEREF '+name+' \\h ');r=OxmlElement('w:r');tx=OxmlElement('w:t');tx.text=str(pg);r.append(tx);f.append(r);p.append(f);prev.addnext(p);prev=p
r=OxmlElement('w:r');fc=OxmlElement('w:fldChar');fc.set(qn('w:fldCharType'),'end');r.append(fc);prev.append(r)
(R/'downloads').mkdir(exist_ok=True);out=R/'downloads/香港访学精美路书.docx';
for el in DOC.element.xpath('.//w:pBdr'):
 el.getparent().remove(el)
for el in DOC.styles.element.xpath('.//w:pBdr'):
 el.getparent().remove(el)
DOC.save(out);print('Saved',out,'planned pages',page)
(R/'.doc-build/page_plan.json').write_text(json.dumps(toc,ensure_ascii=False,indent=2))
