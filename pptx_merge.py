import os
import re
import shutil
import zipfile
from lxml import etree

NS = {
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

REL_NS = NS['rel']
R_NS = NS['r']
CT_NS = NS['ct']
P_NS = NS['p']

SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
LAYOUT_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
MASTER_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml'
SLIDE_REL = REL_NS.replace('package', 'officeDocument') if False else 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
NOTES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'


def _load_rels(path):
    if not os.path.exists(path):
        return None, None
    tree = etree.parse(path)
    return tree, tree.getroot()


def make_blank_base(source_pptx_path, out_path, work_dir):
    """Build a minimal empty .pptx (correct slide size, no slides/layouts/
    masters/themes/media) derived from source_pptx_path, to use as a clean
    starting point for Target (avoids dragging along the source's other,
    unrelated slides as dead weight)."""
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(source_pptx_path) as z:
        z.extractall(work_dir)

    # wipe sldIdLst and sldMasterIdLst in presentation.xml
    pres_path = os.path.join(work_dir, 'ppt/presentation.xml')
    tree = etree.parse(pres_path)
    root = tree.getroot()
    for tag in ('sldIdLst', 'sldMasterIdLst', 'notesMasterIdLst'):
        el = root.find(f'{{{P_NS}}}{tag}')
        if el is not None:
            root.remove(el)
    tree.write(pres_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # remove all relationships from presentation.xml.rels
    pres_rels_path = os.path.join(work_dir, 'ppt/_rels/presentation.xml.rels')
    rtree = etree.parse(pres_rels_path)
    rroot = rtree.getroot()
    for rel in list(rroot.findall(f'{{{REL_NS}}}Relationship')):
        rroot.remove(rel)
    rtree.write(pres_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # delete the physical parts + prune their Content_Types overrides
    ct_path = os.path.join(work_dir, '[Content_Types].xml')
    ct_tree = etree.parse(ct_path)
    ct_root = ct_tree.getroot()
    for sub in ('ppt/slides', 'ppt/slideLayouts', 'ppt/slideMasters',
                'ppt/theme', 'ppt/media', 'ppt/notesSlides', 'ppt/notesMasters',
                'ppt/embeddings'):
        d = os.path.join(work_dir, sub)
        if os.path.isdir(d):
            shutil.rmtree(d)
        for el in list(ct_root.findall(f'{{{CT_NS}}}Override')):
            if el.get('PartName', '').startswith('/' + sub + '/'):
                ct_root.remove(el)
    ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r, dirs, files in os.walk(work_dir):
            for f in files:
                full = os.path.join(r, f)
                rel = os.path.relpath(full, work_dir)
                zf.write(full, rel)


class Target:
    """A working directory holding an unzipped pptx that we grow by copying
    slides (and their layout/master/theme/media chains) from source decks."""

    def __init__(self, base_pptx_path, work_dir):
        self.dir = work_dir
        if os.path.exists(self.dir):
            shutil.rmtree(self.dir)
        os.makedirs(self.dir)
        with zipfile.ZipFile(base_pptx_path) as z:
            z.extractall(self.dir)
        self._source_master_cache = {}  # (source_path, source_layout_partname) -> target_layout_partname
        self._source_theme_cache = {}   # (source_path, source_theme_partname) -> target_theme_partname

    # ---------- helpers over target's own package parts ----------
    def _ct_tree(self):
        return etree.parse(os.path.join(self.dir, '[Content_Types].xml'))

    def _next_index(self, subdir, prefix, ext='.xml'):
        full_dir = os.path.join(self.dir, subdir)
        if not os.path.isdir(full_dir):
            os.makedirs(full_dir, exist_ok=True)
        existing = [f for f in os.listdir(full_dir)
                    if re.match(rf'^{prefix}\d+{re.escape(ext)}$', f)]
        nums = [int(re.match(rf'^{prefix}(\d+){re.escape(ext)}$', f).group(1)) for f in existing]
        return (max(nums) + 1) if nums else 1

    def _add_content_type_override(self, partname, content_type):
        path = os.path.join(self.dir, '[Content_Types].xml')
        tree = etree.parse(path)
        root = tree.getroot()
        for el in root.findall(f'{{{CT_NS}}}Override'):
            if el.get('PartName') == partname:
                return
        el = etree.SubElement(root, f'{{{CT_NS}}}Override')
        el.set('PartName', partname)
        el.set('ContentType', content_type)
        tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _rels_path_for(self, part_relpath):
        d, f = os.path.split(part_relpath)
        return os.path.join(self.dir, d, '_rels', f + '.rels')

    def _next_rid(self, rels_path):
        if not os.path.exists(rels_path):
            return 'rId1'
        tree = etree.parse(rels_path)
        ids = [int(re.match(r'rId(\d+)', r.get('Id')).group(1))
               for r in tree.getroot().findall(f'{{{REL_NS}}}Relationship')]
        return f'rId{(max(ids) + 1) if ids else 1}'

    def _add_relationship(self, owner_part_relpath, rel_type, target_relpath_from_owner, target_mode=None):
        rels_path = self._rels_path_for(owner_part_relpath)
        os.makedirs(os.path.dirname(rels_path), exist_ok=True)
        if os.path.exists(rels_path):
            tree = etree.parse(rels_path)
            root = tree.getroot()
        else:
            root = etree.Element(f'{{{REL_NS}}}Relationships')
            tree = etree.ElementTree(root)
        rid = self._next_rid(rels_path)
        el = etree.SubElement(root, f'{{{REL_NS}}}Relationship')
        el.set('Id', rid)
        el.set('Type', rel_type)
        el.set('Target', target_relpath_from_owner)
        if target_mode:
            el.set('TargetMode', target_mode)
        tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        return rid

    def _copy_media(self, src_pptx_dir, src_media_relpath, unique_tag):
        # src_media_relpath like 'ppt/media/image3.png'
        ext = os.path.splitext(src_media_relpath)[1]
        base = os.path.splitext(os.path.basename(src_media_relpath))[0]
        new_name = f'{base}_{unique_tag}{ext}'
        dst_relpath = f'ppt/media/{new_name}'
        os.makedirs(os.path.join(self.dir, 'ppt/media'), exist_ok=True)
        shutil.copyfile(os.path.join(src_pptx_dir, src_media_relpath),
                         os.path.join(self.dir, dst_relpath))
        ext_low = ext.lstrip('.').lower()
        self._ensure_default_extension(ext_low)
        return dst_relpath

    def _ensure_default_extension(self, ext, content_type=None):
        path = os.path.join(self.dir, '[Content_Types].xml')
        tree = etree.parse(path)
        root = tree.getroot()
        for el in root.findall(f'{{{CT_NS}}}Default'):
            if el.get('Extension').lower() == ext.lower():
                return
        guess = {
            'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
            'gif': 'image/gif', 'emf': 'image/x-emf', 'wmf': 'image/x-wmf',
            'svg': 'image/svg+xml', 'mp4': 'video/mp4', 'mp3': 'audio/mpeg',
            'wav': 'audio/wav', 'm4a': 'audio/mp4', 'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }
        content_type = content_type or guess.get(ext.lower(), 'application/octet-stream')
        el = etree.SubElement(root, f'{{{CT_NS}}}Default')
        el.set('Extension', ext)
        el.set('ContentType', content_type)
        tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # ---------- copying theme / master / layout chains ----------
    def _copy_theme(self, src_dir, src_theme_relpath, cache_key):
        if cache_key in self._source_theme_cache:
            return self._source_theme_cache[cache_key]
        idx = self._next_index('ppt/theme', 'theme')
        dst_relpath = f'ppt/theme/theme{idx}.xml'
        shutil.copyfile(os.path.join(src_dir, src_theme_relpath), os.path.join(self.dir, dst_relpath))
        self._add_content_type_override('/' + dst_relpath, THEME_CT)
        self._source_theme_cache[cache_key] = dst_relpath
        return dst_relpath

    def _copy_master_chain(self, src_dir, src_master_relpath, unique_tag):
        """Copy a slideMaster, its theme, and ALL its slideLayouts. Returns
        dict mapping src layout relpath -> dst layout relpath."""
        cache_key = (src_dir, src_master_relpath)
        if cache_key in self._source_master_cache:
            return self._source_master_cache[cache_key]

        master_idx = self._next_index('ppt/slideMasters', 'slideMaster')
        dst_master_relpath = f'ppt/slideMasters/slideMaster{master_idx}.xml'
        shutil.copyfile(os.path.join(src_dir, src_master_relpath), os.path.join(self.dir, dst_master_relpath))
        self._add_content_type_override('/' + dst_master_relpath, MASTER_CT)
        self._register_master(dst_master_relpath)

        src_master_rels_path = os.path.join(src_dir, 'ppt/slideMasters/_rels', os.path.basename(src_master_relpath) + '.rels')
        src_tree = etree.parse(src_master_rels_path)
        layout_map = {}
        master_rid_map = {}  # old rid (in source master) -> new rid (in dst master)
        for rel in src_tree.getroot().findall(f'{{{REL_NS}}}Relationship'):
            rtype = rel.get('Type')
            target = rel.get('Target')  # relative to ppt/slideMasters/
            abs_src_relpath = os.path.normpath(os.path.join('ppt/slideMasters', target)).replace('\\', '/')
            if rtype == THEME_REL:
                theme_dst = self._copy_theme(src_dir, abs_src_relpath, (src_dir, abs_src_relpath))
                new_rid = self._add_relationship(dst_master_relpath, THEME_REL,
                                                  '../' + theme_dst.split('/', 1)[1])
                master_rid_map[rel.get('Id')] = new_rid
            elif rtype == LAYOUT_REL:
                layout_idx = self._next_index('ppt/slideLayouts', 'slideLayout')
                dst_layout_relpath = f'ppt/slideLayouts/slideLayout{layout_idx}.xml'
                shutil.copyfile(os.path.join(src_dir, abs_src_relpath), os.path.join(self.dir, dst_layout_relpath))
                self._add_content_type_override('/' + dst_layout_relpath, LAYOUT_CT)
                # layout -> master relationship
                new_rid_layout_to_master = self._add_relationship(
                    dst_layout_relpath, MASTER_REL,
                    '../' + dst_master_relpath.split('/', 1)[1])
                self._apply_rid_map(dst_layout_relpath,
                                     {self._get_layout_master_rid(src_dir, abs_src_relpath): new_rid_layout_to_master})
                # master -> layout relationship
                new_rid_master_to_layout = self._add_relationship(
                    dst_master_relpath, LAYOUT_REL,
                    '../' + dst_layout_relpath.split('/', 1)[1])
                master_rid_map[rel.get('Id')] = new_rid_master_to_layout
                layout_map[abs_src_relpath] = dst_layout_relpath

        self._apply_rid_map(dst_master_relpath, master_rid_map)
        self._source_master_cache[cache_key] = (dst_master_relpath, layout_map)
        return dst_master_relpath, layout_map

    def _get_layout_master_rid(self, src_dir, src_layout_relpath):
        rels_path = os.path.join(src_dir, 'ppt/slideLayouts/_rels', os.path.basename(src_layout_relpath) + '.rels')
        tree = etree.parse(rels_path)
        for rel in tree.getroot().findall(f'{{{REL_NS}}}Relationship'):
            if rel.get('Type') == MASTER_REL:
                return rel.get('Id')
        return None

    def _apply_rid_map(self, part_relpath, rid_map):
        """Rewrite every quoted rId in `part_relpath` according to rid_map, in a
        SINGLE simultaneous pass (never sequential — sequential replacement can
        cascade when a new id happens to equal another entry's old id)."""
        rid_map = {k: v for k, v in rid_map.items() if v is not None and k != v}
        if not rid_map:
            return
        path = os.path.join(self.dir, part_relpath)
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        pattern = re.compile(r'(["\'])(' + '|'.join(re.escape(k) for k in rid_map) + r')(["\'])')

        def _repl(m):
            return m.group(1) + rid_map[m.group(2)] + m.group(3)

        content = pattern.sub(_repl, content)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)

    # ---------- public: copy one slide ----------
    def copy_slide(self, src_pptx_dir, src_slide_relpath, source_tag):
        """src_pptx_dir: unzipped source dir. src_slide_relpath: e.g. 'ppt/slides/slide13.xml'.
        source_tag: short unique string identifying the source deck (for media naming)."""
        src_rels_path = os.path.join(src_pptx_dir, 'ppt/slides/_rels', os.path.basename(src_slide_relpath) + '.rels')
        src_rels_tree = etree.parse(src_rels_path) if os.path.exists(src_rels_path) else None

        slide_idx = self._next_index('ppt/slides', 'slide')
        dst_slide_relpath = f'ppt/slides/slide{slide_idx}.xml'
        shutil.copyfile(os.path.join(src_pptx_dir, src_slide_relpath), os.path.join(self.dir, dst_slide_relpath))
        self._add_content_type_override('/' + dst_slide_relpath, SLIDE_CT)

        rid_rewrites = {}
        if src_rels_tree is not None:
            for rel in src_rels_tree.getroot().findall(f'{{{REL_NS}}}Relationship'):
                rtype = rel.get('Type')
                old_rid = rel.get('Id')
                target = rel.get('Target')
                target_mode = rel.get('TargetMode')

                if rtype == NOTES_REL:
                    # drop notes; remove reference from slide xml later if present
                    rid_rewrites[old_rid] = None
                    continue
                if target_mode == 'External':
                    new_rid = self._add_relationship(dst_slide_relpath, rtype, target, target_mode='External')
                    rid_rewrites[old_rid] = new_rid
                    continue
                if rtype == LAYOUT_REL:
                    abs_src_layout = os.path.normpath(os.path.join('ppt/slides', target)).replace('\\', '/')
                    src_layout_rels_path = os.path.join(src_pptx_dir, 'ppt/slideLayouts/_rels',
                                                          os.path.basename(abs_src_layout) + '.rels')
                    master_rid = None
                    with open(src_layout_rels_path, 'rb') as f:
                        lt = etree.parse(f)
                    src_master_relpath = None
                    for lrel in lt.getroot().findall(f'{{{REL_NS}}}Relationship'):
                        if lrel.get('Type') == MASTER_REL:
                            src_master_relpath = os.path.normpath(
                                os.path.join('ppt/slideLayouts', lrel.get('Target'))).replace('\\', '/')
                    dst_master_relpath, layout_map = self._copy_master_chain(src_pptx_dir, src_master_relpath, source_tag)
                    dst_layout_relpath = layout_map[abs_src_layout]
                    new_rid = self._add_relationship(dst_slide_relpath, LAYOUT_REL,
                                                       '../' + dst_layout_relpath.split('/', 1)[1])
                    rid_rewrites[old_rid] = new_rid
                    continue
                # media / embeddings / other package-relative parts
                abs_src_target = os.path.normpath(os.path.join('ppt/slides', target)).replace('\\', '/')
                if abs_src_target.startswith('ppt/media/'):
                    dst_media_relpath = self._copy_media(src_pptx_dir, abs_src_target, source_tag + f'_{slide_idx}')
                    new_rid = self._add_relationship(dst_slide_relpath, rtype,
                                                       '../' + dst_media_relpath.split('/', 1)[1])
                    rid_rewrites[old_rid] = new_rid
                else:
                    # unexpected part type (e.g. embedded chart/ole) - copy generically
                    ext = os.path.splitext(abs_src_target)[1]
                    idx2 = self._next_index('ppt/embeddings', 'oleObject', ext) if os.path.isdir(os.path.join(self.dir, 'ppt/embeddings')) else 1
                    os.makedirs(os.path.join(self.dir, 'ppt/embeddings'), exist_ok=True)
                    dst_relpath = f'ppt/embeddings/oleObject{idx2}{ext}'
                    shutil.copyfile(os.path.join(src_pptx_dir, abs_src_target), os.path.join(self.dir, dst_relpath))
                    self._ensure_default_extension(ext.lstrip('.'))
                    new_rid = self._add_relationship(dst_slide_relpath, rtype, '../' + dst_relpath.split('/', 1)[1])
                    rid_rewrites[old_rid] = new_rid

        # apply rid rewrites in ONE simultaneous pass (dropped/notes entries map to None and are skipped)
        self._apply_rid_map(dst_slide_relpath, rid_rewrites)

        # register the slide in presentation.xml + presentation.xml.rels
        new_rid_pres = self._add_relationship('ppt/presentation.xml', SLIDE_REL,
                                               dst_slide_relpath.split('/', 1)[1])
        self._append_slide_to_sldidlst(new_rid_pres)
        return dst_slide_relpath

    def _ensure_toplevel_before_sldSz(self, tag):
        """Ensure <p:TAG/> exists directly under <p:presentation>, inserted at
        the schema-correct position (sldMasterIdLst, notesMasterIdLst,
        handoutMasterIdLst, sldIdLst, sldSz, ... in that order)."""
        pres_path = os.path.join(self.dir, 'ppt/presentation.xml')
        tree = etree.parse(pres_path)
        root = tree.getroot()
        if root.find(f'{{{P_NS}}}{tag}') is not None:
            return
        order = ['sldMasterIdLst', 'notesMasterIdLst', 'handoutMasterIdLst', 'sldIdLst']
        my_rank = order.index(tag)
        idx = 0
        for i, child in enumerate(root):
            cname = etree.QName(child).localname
            if cname in order and order.index(cname) < my_rank:
                idx = i + 1
        el = etree.Element(f'{{{P_NS}}}{tag}')
        root.insert(idx, el)
        tree.write(pres_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _append_slide_to_sldidlst(self, rid):
        self._ensure_toplevel_before_sldSz('sldIdLst')
        pres_path = os.path.join(self.dir, 'ppt/presentation.xml')
        tree = etree.parse(pres_path)
        root = tree.getroot()
        sldidlst = root.find(f'{{{P_NS}}}sldIdLst')
        existing_ids = [int(s.get('id')) for s in sldidlst.findall(f'{{{P_NS}}}sldId')]
        new_id = (max(existing_ids) + 1) if existing_ids else 256
        el = etree.SubElement(sldidlst, f'{{{P_NS}}}sldId')
        el.set('id', str(new_id))
        el.set(f'{{{R_NS}}}id', rid)
        tree.write(pres_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _register_master(self, dst_master_relpath):
        new_rid = self._add_relationship('ppt/presentation.xml', MASTER_REL,
                                          dst_master_relpath.split('/', 1)[1])
        self._ensure_toplevel_before_sldSz('sldMasterIdLst')
        pres_path = os.path.join(self.dir, 'ppt/presentation.xml')
        tree = etree.parse(pres_path)
        root = tree.getroot()
        lst = root.find(f'{{{P_NS}}}sldMasterIdLst')
        existing_ids = [int(s.get('id')) for s in lst.findall(f'{{{P_NS}}}sldMasterId')]
        new_id = (max(existing_ids) + 1) if existing_ids else 2147483649
        el = etree.SubElement(lst, f'{{{P_NS}}}sldMasterId')
        el.set('id', str(new_id))
        el.set(f'{{{R_NS}}}id', new_rid)
        tree.write(pres_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def strip_shape_by_text(self, dst_slide_relpath, exact_text):
        """Remove the <p:sp> shape whose text frame equals exact_text (after
        stripping whitespace) from an already-copied slide."""
        path = os.path.join(self.dir, dst_slide_relpath)
        tree = etree.parse(path)
        root = tree.getroot()
        A_NS = NS['a']
        for sp in root.iter(f'{{{P_NS}}}sp'):
            texts = [t.text or '' for t in sp.iter(f'{{{A_NS}}}t')]
            joined = ''.join(texts).strip()
            if joined == exact_text:
                sp.getparent().remove(sp)
                break
        tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def apply_text_transform(self, dst_slide_relpath, transform_fn):
        """Apply transform_fn(text) -> text to every <a:t> run in the slide."""
        path = os.path.join(self.dir, dst_slide_relpath)
        tree = etree.parse(path)
        root = tree.getroot()
        A_NS = NS['a']
        changed = False
        for t in root.iter(f'{{{A_NS}}}t'):
            if t.text:
                new_text = transform_fn(t.text)
                if new_text != t.text:
                    t.text = new_text
                    changed = True
        if changed:
            tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def save(self, out_path):
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(self.dir):
                for f in files:
                    full = os.path.join(root, f)
                    rel = os.path.relpath(full, self.dir)
                    zf.write(full, rel)
