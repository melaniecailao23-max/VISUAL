"""
Modify Cailao_Exp22_PPT_Ch02_ML1_Design.pptx per assignment instructions.
Works directly with PPTX ZIP/XML structure.
"""
import zipfile, shutil, os
import xml.etree.ElementTree as ET

# Register namespaces
for prefix, uri in [
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('p', 'http://schemas.openxmlformats.org/presentationml/2006/main'),
    ('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main'),
    ('a14', 'http://schemas.microsoft.com/office/drawing/2010/main'),
    ('a16', 'http://schemas.microsoft.com/office/drawing/2014/main'),
]:
    ET.register_namespace(prefix, uri)

def emu(inches): return int(inches * 914400)

INPUT = '/projects/sandbox/VISUAL/Cailao_Exp22_PPT_Ch02_ML1_Design.pptx'
TEMP = '/tmp/pptx_work'
AUDIO = '/projects/sandbox/VISUAL/Welcome.m4a'

if os.path.exists(TEMP): shutil.rmtree(TEMP)
os.makedirs(TEMP)
with zipfile.ZipFile(INPUT, 'r') as z: z.extractall(TEMP)
print("Extracted PPTX")


A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
P14 = '{http://schemas.microsoft.com/office/powerpoint/2010/main}'

def read_xml(path):
    with open(os.path.join(TEMP, path), 'r', encoding='utf-8-sig') as f:
        return ET.fromstring(f.read())

def write_xml(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write('\ufeff' + s)

def write_xml_no_bom(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write(s)


# ===== STEP 2: Title formatting on slides 2, 4, 6, 8 =====
print("Step 2: Formatting titles...")
for sn in [2, 4, 6, 8]:
    root = read_xml(f'ppt/slides/slide{sn}.xml')
    for sp in root.iter(f'{P}sp'):
        nvPr = sp.find(f'.//{P}nvPr')
        if nvPr is not None:
            ph = nvPr.find(f'{P}ph')
            if ph is not None and ph.get('type') == 'title':
                for rPr in sp.iter(f'{A}rPr'):
                    rPr.set('b', '1')
                    rPr.set('sz', '6600')
                    sf = rPr.find(f'{A}solidFill')
                    if sf is not None: rPr.remove(sf)
                    nf = ET.SubElement(rPr, f'{A}solidFill')
                    sc = ET.SubElement(nf, f'{A}schemeClr')
                    sc.set('val', 'tx1')
                break
    write_xml(root, f'ppt/slides/slide{sn}.xml')
print("  Done")


# ===== STEP 3: Text box "Make it" on Slide 3 =====
print("Step 3: Adding text box...")
root = read_xml('ppt/slides/slide3.xml')
spTree = root.find(f'.//{P}spTree')
tb = ET.fromstring(f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvSpPr><p:cNvPr id="100" name="TextBox 100"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{emu(0.73)}" y="{emu(0.73)}"/><a:ext cx="3200400" cy="1371600"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
<a:p><a:r><a:rPr lang="en-US" sz="8000" dirty="0"/><a:t>Make it</a:t></a:r></a:p></p:txBody></p:sp>''')
spTree.append(tb)
write_xml(root, 'ppt/slides/slide3.xml')
print("  Done")


# ===== STEP 4: Rectangle on Slide 3 =====
print("Step 4: Adding rectangle...")
root = read_xml('ppt/slides/slide3.xml')
spTree = root.find(f'.//{P}spTree')
rect = ET.fromstring(f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvSpPr><p:cNvPr id="101" name="Rectangle 101"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{emu(4.7)}" y="{emu(0.4)}"/><a:ext cx="{emu(3.9)}" cy="{emu(3.9)}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:gradFill rotWithShape="1"><a:gsLst>
<a:gs pos="0"><a:schemeClr val="accent1"><a:satMod val="103000"/><a:lumMod val="118000"/><a:tint val="94000"/></a:schemeClr></a:gs>
<a:gs pos="50000"><a:schemeClr val="accent1"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
<a:gs pos="100000"><a:schemeClr val="accent1"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
<a:ln><a:noFill/></a:ln>
<a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst>
<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>
</p:spPr>
<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>
<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="16000" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Big</a:t></a:r></a:p></p:txBody></p:sp>''')
spTree.append(rect)
write_xml(root, 'ppt/slides/slide3.xml')
print("  Done")


# ===== STEP 5: Picture style on Slide 5 =====
print("Step 5: Applying picture style...")
root = read_xml('ppt/slides/slide5.xml')
for pic in root.iter(f'{P}pic'):
    xfrm = pic.find(f'.//{A}xfrm')
    if xfrm is None: continue
    ext = xfrm.find(f'{A}ext')
    if ext is None: continue
    cx, cy = int(ext.get('cx','0')), int(ext.get('cy','0'))
    if cy < 500000 or cx < 500000: continue  # skip thin decorative bars
    
    spPr = pic.find(f'{P}spPr')
    if spPr is None: continue
    
    # Remove existing style elements
    for tag in ['prstGeom','ln','effectLst','scene3d','sp3d','solidFill']:
        el = spPr.find(f'{A}{tag}')
        if el is not None: spPr.remove(el)
    
    # Add Reflected Perspective Right style
    geom = ET.SubElement(spPr, f'{A}prstGeom')
    geom.set('prst', 'rect')
    ET.SubElement(geom, f'{A}avLst')
    
    ln = ET.SubElement(spPr, f'{A}ln')
    ET.SubElement(ln, f'{A}noFill')
    
    eff = ET.SubElement(spPr, f'{A}effectLst')
    refl = ET.SubElement(eff, f'{A}reflection')
    for k, v in {'blurRad':'12700','stA':'48000','endA':'300','endPos':'55000',
                 'dist':'50800','dir':'5400000','sy':'-100000','algn':'bl','rotWithShape':'0'}.items():
        refl.set(k, v)
    
    sc = ET.SubElement(spPr, f'{A}scene3d')
    cam = ET.SubElement(sc, f'{A}camera')
    cam.set('prst', 'perspectiveRight')
    lr = ET.SubElement(sc, f'{A}lightRig')
    lr.set('rig', 'balanced')
    lr.set('dir', 't')
    
    sp3d = ET.SubElement(spPr, f'{A}sp3d')
    bev = ET.SubElement(sp3d, f'{A}bevelT')
    bev.set('w', '190500')
    bev.set('h', '38100')

write_xml(root, 'ppt/slides/slide5.xml')
print("  Done")


# ===== STEP 6: Numbered circle icons on Slide 10 =====
print("Step 6: Adding icons on Slide 10...")
root = read_xml('ppt/slides/slide10.xml')
spTree = root.find(f'.//{P}spTree')
icons = [(1, 3.2, 3.4), (2, 3.2, 5.5), (3, 9.1, 3.4), (4, 9.1, 5.5)]
for num, hx, vy in icons:
    iid = 200 + num
    sz = 609600
    icon = ET.fromstring(f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvSpPr><p:cNvPr id="{iid}" name="Graphic {iid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{emu(hx)}" y="{emu(vy)}"/><a:ext cx="{sz}" cy="{sz}"/></a:xfrm>
<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="000000"/></a:solidFill>
<a:ln><a:noFill/></a:ln>
<a:effectLst><a:outerShdw blurRad="50800" dist="50800" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>
</p:spPr>
<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>{num}</a:t></a:r></a:p></p:txBody></p:sp>''')
    spTree.append(icon)
write_xml(root, 'ppt/slides/slide10.xml')
print("  Done")


# ===== ANIMATION HELPER =====
def make_fly_in_entry(ctn_base, spid, para_idx):
    """Create a Fly In animation p:par block for one paragraph."""
    return f'''<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cTn id="{ctn_base}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="{ctn_base+1}" fill="hold"><p:stCondLst><p:cond delay="500"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="{ctn_base+2}" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="withEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst>
<p:set><p:cBhvr><p:cTn id="{ctn_base+3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>
<p:to><p:strVal val="visible"/></p:to></p:set>
<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{ctn_base+4}" dur="750" fill="hold"/>
<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>
<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{ctn_base+5}" dur="750" fill="hold"/>
<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="1+#ppt_h/2"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'''

def make_grow_shrink_entry(ctn_base, spid, para_idx):
    """Create a Grow/Shrink emphasis animation entry."""
    return f'''<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cTn id="{ctn_base}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="{ctn_base+1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="{ctn_base+2}" presetID="6" presetClass="emph" presetSubtype="32" fill="hold" grpId="1" nodeType="withEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:animScale><p:cBhvr><p:cTn id="{ctn_base+3}" dur="2000" fill="hold"/>
<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
</p:cBhvr><p:by x="150000" y="150000"/></p:animScale></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'''

def wrap_timing(inner_pars_xml):
    """Wrap animation par entries in full timing structure."""
    return f'''<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
<p:childTnLst><p:seq concurrent="1" nextAc="seek">
<p:cTn id="2" dur="indefinite" nodeType="mainSeq">
<p:childTnLst>{inner_pars_xml}</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''

def get_content_id(root):
    """Find the content placeholder (idx=1) shape id."""
    for sp in root.iter(f'{P}sp'):
        nvPr = sp.find(f'.//{P}nvPr')
        if nvPr is not None:
            ph = nvPr.find(f'{P}ph')
            if ph is not None and ph.get('idx') == '1':
                nvSpPr = sp.find(f'{P}nvSpPr')
                if nvSpPr is not None:
                    cNvPr = nvSpPr.find(f'{P}cNvPr')
                    if cNvPr is not None:
                        return cNvPr.get('id')
    return None


# ===== STEPS 7-10: Animations on Slide 2 =====
print("Steps 7-10: Slide 2 animations...")
root = read_xml('ppt/slides/slide2.xml')
spid = get_content_id(root)
print(f"  Content id: {spid}")

# Build animation entries (order per Step 10):
# 1. FlyIn para 0, 2. FlyIn para 1, 3. FlyIn para 2, 4. GrowShrink para 2, 5. FlyIn para 3
ctn = 3
entries = []
entries.append(make_fly_in_entry(ctn, spid, 0)); ctn += 6
entries.append(make_fly_in_entry(ctn, spid, 1)); ctn += 6
entries.append(make_fly_in_entry(ctn, spid, 2)); ctn += 6
entries.append(make_grow_shrink_entry(ctn, spid, 2)); ctn += 4
entries.append(make_fly_in_entry(ctn, spid, 3)); ctn += 6

inner = ''.join(entries)
timing_xml = wrap_timing(inner)
timing_elem = ET.fromstring(timing_xml)

# Remove old timing, add new
old = root.find(f'{P}timing')
if old is not None: root.remove(old)
root.append(timing_elem)
write_xml(root, 'ppt/slides/slide2.xml')
print("  Done")

# ===== STEP 8: Same fly-in on slides 4, 6, 8 =====
print("Step 8: Animations on slides 4, 6, 8...")
for sn in [4, 6, 8]:
    root = read_xml(f'ppt/slides/slide{sn}.xml')
    spid = get_content_id(root)
    if not spid:
        print(f"  Warning: no content placeholder on slide {sn}")
        continue
    ctn = 3
    entries = []
    entries.append(make_fly_in_entry(ctn, spid, 0)); ctn += 6
    entries.append(make_fly_in_entry(ctn, spid, 1)); ctn += 6
    entries.append(make_fly_in_entry(ctn, spid, 2)); ctn += 6
    entries.append(make_fly_in_entry(ctn, spid, 3)); ctn += 6
    inner = ''.join(entries)
    timing_xml = wrap_timing(inner)
    timing_elem = ET.fromstring(timing_xml)
    old = root.find(f'{P}timing')
    if old is not None: root.remove(old)
    root.append(timing_elem)
    write_xml(root, f'ppt/slides/slide{sn}.xml')
print("  Done")


# ===== STEP 11: Cube transition (From Bottom) on all slides =====
print("Step 11: Cube transition...")
ET.register_namespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main')
for sn in range(1, 11):
    root = read_xml(f'ppt/slides/slide{sn}.xml')
    old = root.find(f'{P}transition')
    if old is not None: root.remove(old)
    
    trans = ET.fromstring('''<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:cube dir="d"/></p:transition>''')
    
    # Insert after clrMapOvr, before timing
    timing = root.find(f'{P}timing')
    if timing is not None:
        idx = list(root).index(timing)
        root.insert(idx, trans)
    else:
        root.append(trans)
    
    write_xml(root, f'ppt/slides/slide{sn}.xml')
print("  Done")


# ===== STEP 12: Audio on Slide 1 =====
print("Step 12: Adding audio...")
# Copy audio file
shutil.copy2(AUDIO, os.path.join(TEMP, 'ppt/media/Welcome.m4a'))

# Update slide1 rels
ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/relationships')
rels_root = read_xml('ppt/slides/_rels/slide1.xml.rels')
ids = [r.get('Id','') for r in rels_root]
mx = max((int(i[3:]) for i in ids if i.startswith('rId')), default=0)
arid = f'rId{mx+1}'
mrid = f'rId{mx+2}'

ar = ET.SubElement(rels_root, 'Relationship')
ar.set('Id', arid)
ar.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio')
ar.set('Target', '../media/Welcome.m4a')

mr = ET.SubElement(rels_root, 'Relationship')
mr.set('Id', mrid)
mr.set('Type', 'http://schemas.microsoft.com/office/2007/relationships/media')
mr.set('Target', '../media/Welcome.m4a')

write_xml_no_bom(rels_root, 'ppt/slides/_rels/slide1.xml.rels')

# Add audio pic to slide 1
root = read_xml('ppt/slides/slide1.xml')
spTree = root.find(f'.//{P}spTree')

audio_pic = ET.fromstring(f'''<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
<p:nvPicPr><p:cNvPr id="300" name="Welcome.m4a"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
<p:nvPr><a:audioFile r:link="{arid}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media r:embed="{mrid}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>
<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="{emu(0.24)}" y="{emu(0.24)}"/><a:ext cx="{emu(0.3)}" cy="{emu(0.3)}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>''')
spTree.append(audio_pic)

# Add auto-play timing (start automatically, hide during show)
old = root.find(f'{P}timing')
if old is not None: root.remove(old)

audio_timing = ET.fromstring(f'''<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
<p:childTnLst><p:seq concurrent="1" nextAc="seek">
<p:cTn id="2" dur="indefinite" nodeType="mainSeq">
<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:cmd type="call" cmd="playFrom(0)"><p:cBhvr><p:cTn id="6" dur="1" fill="hold"/>
<p:tgtEl><p:spTgt spid="300"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst>
</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>''')
root.append(audio_timing)

write_xml(root, 'ppt/slides/slide1.xml')
print("  Done")


# ===== Update Content_Types.xml =====
print("Updating Content_Types...")
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('', CT_NS)
ct_root = read_xml('[Content_Types].xml')
has_m4a = any(d.get('Extension') == 'm4a' for d in ct_root.findall(f'{{{CT_NS}}}Default'))
if not has_m4a:
    d = ET.SubElement(ct_root, f'{{{CT_NS}}}Default')
    d.set('Extension', 'm4a')
    d.set('ContentType', 'audio/mp4')
write_xml_no_bom(ct_root, '[Content_Types].xml')
print("  Done")

# ===== FINAL: Repackage PPTX =====
print("Repackaging...")
if os.path.exists(INPUT): os.remove(INPUT)
with zipfile.ZipFile(INPUT, 'w', zipfile.ZIP_DEFLATED) as zf:
    for dirpath, dirs, files in os.walk(TEMP):
        for f in files:
            fp = os.path.join(dirpath, f)
            arcname = os.path.relpath(fp, TEMP)
            zf.write(fp, arcname)
print(f"Saved: {INPUT} ({os.path.getsize(INPUT)} bytes)")
print("\nALL DONE!")
