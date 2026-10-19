#!/usr/bin/env python3
"""Preserve native PPTX parts while selecting slides and editing named objects."""
import argparse
from copy import deepcopy
import json
import io
import math
from pathlib import Path
import posixpath
import sys
from xml.sax.saxutils import quoteattr
import xml.etree.ElementTree as E
import zipfile

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'p': P, 'a': A, 'r': R}
for prefix, uri in NS.items():
    E.register_namespace(prefix, uri)
ROOT = Path(__file__).resolve().parent.parent
DECLARATIONS = {}


def parse(b):
    namespaces = dict(pair for _, pair in E.iterparse(io.BytesIO(b), events=('start-ns',)))
    for prefix, uri in namespaces.items():
        if prefix and not (prefix.startswith('ns') and prefix[2:].isdigit()):
            E.register_namespace(prefix, uri)
    root = E.fromstring(b)
    DECLARATIONS[root] = namespaces
    return root


def encode(e):
    result = E.tostring(e, encoding='unicode')
    # Preserve namespace declarations used only in mc:Ignorable / QName values.
    end = result.index('>')
    head = result[:end]
    additions = []
    for prefix, uri in DECLARATIONS.get(e, {}).items():
        key = 'xmlns:'+prefix if prefix else 'xmlns'
        if key+'=' not in head:
            additions.append(' '+key+'='+quoteattr(uri))
    if additions:
        insert = end-1 if head.endswith('/') else end
        result = result[:insert]+''.join(additions)+result[insert:]
    return b'<?xml version="1.0" encoding="utf-8"?>\n'+result.encode('utf-8')


def load(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def relpath(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels') if part else '_rels/.rels'


def target(part, path):
    return path.lstrip('/') if path.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), path))


def ordered_slides(files):
    rels = {e.get('Id'): target('ppt/presentation.xml', e.get('Target'))
            for e in parse(files['ppt/_rels/presentation.xml.rels'])}
    return [rels[e.get('{'+R+'}id')] for e in parse(files['ppt/presentation.xml']).find('p:sldIdLst', NS)]


def shapes(root):
    return [(e, e.find('.//p:cNvPr', NS)) for e in root.find('p:cSld/p:spTree', NS)
            if e.tag not in ('{'+P+'}nvGrpSpPr', '{'+P+'}grpSpPr')]


def find_shape(root, key):
    found = [e for e, nv in shapes(root) if nv is not None and str(key) in (nv.get('id'), nv.get('name'))]
    if len(found) != 1:
        raise ValueError(f'Expected one top-level object {key!r}; found {len(found)}')
    return found[0]


def set_text(shape, value):
    body = shape.find('p:txBody', NS)
    if body is None:
        raise ValueError('Selected object has no native text body')
    oldp = body.find('a:p', NS)
    pp = oldp.find('a:pPr', NS) if oldp is not None else None
    oldpr = body.find('.//a:rPr', NS)
    if oldpr is None and pp is not None:
        oldpr = pp.find('a:defRPr', NS)
    for e in list(body):
        if e.tag == '{'+A+'}p':
            body.remove(e)
    for line in value.split('\n'):
        para = E.SubElement(body, '{'+A+'}p')
        if pp is not None:
            para.append(deepcopy(pp))
        run = E.SubElement(para, '{'+A+'}r')
        pr = deepcopy(oldpr) if oldpr is not None else E.Element('{'+A+'}rPr', sz='2400')
        pr.tag = '{'+A+'}rPr'
        for kind in ('latin', 'ea', 'cs'):
            font = pr.find('a:'+kind, NS)
            if font is None:
                font = E.SubElement(pr, '{'+A+'}'+kind)
            font.set('typeface', 'Microsoft YaHei')
        run.append(pr)
        E.SubElement(run, '{'+A+'}t').text = line


def transform(shape, values):
    if not isinstance(values, dict) or not values or set(values)-{'x', 'y', 'width', 'height'}:
        raise ValueError('Transform requires x/y/width/height values')
    if any(isinstance(v, bool) or not isinstance(v, (float, int)) or not math.isfinite(v) for v in values.values()):
        raise ValueError('Transform values must be finite numbers')
    xf = shape.find('p:grpSpPr/a:xfrm', NS) if shape.tag == '{'+P+'}grpSp' else shape.find('p:spPr/a:xfrm', NS)
    if xf is None:
        raise ValueError('Object has inherited/no transform; edit it with a presentation tool')
    off, ext = xf.find('a:off', NS), xf.find('a:ext', NS)
    ow, oh = int(ext.get('cx')), int(ext.get('cy'))
    if ow <= 0 or oh <= 0:
        raise ValueError('Object has invalid dimensions')
    for key in ('x', 'y'):
        if key in values:
            off.set(key, str(round(values[key]*9525)))
    w, h = values.get('width'), values.get('height')
    if w is not None and w <= 0 or h is not None and h <= 0:
        raise ValueError('Dimensions must be positive')
    if w is not None or h is not None:
        if w is None:
            w = h*ow/oh
        if h is None:
            h = w*oh/ow
        ext.set('cx', str(round(w*9525)))
        ext.set('cy', str(round(h*9525)))


def reachable(files):
    keep, pending = {'[Content_Types].xml'}, ['']
    while pending:
        part = pending.pop()
        if part:
            keep.add(part)
        rp = relpath(part)
        if rp not in files or rp in keep:
            continue
        keep.add(rp)
        for rel in parse(files[rp]):
            if rel.get('TargetMode') == 'External':
                continue
            dest = target(part, rel.get('Target'))
            if dest not in files:
                raise ValueError(f'Missing relationship target: {dest}')
            if dest not in keep:
                pending.append(dest)
    return keep


def validate(path):
    files = load(path)
    reachable(files)
    pages = ordered_slides(files)
    for p in pages:
        r = parse(files[p])
        ids = [e.get('id') for e in r.findall('.//p:cNvPr', NS)]
        if len(ids) != len(set(ids)):
            raise ValueError(f'Duplicate shape IDs in {p}')
    return {'status': 'pass', 'slide_count': len(pages), 'scope': 'Package relationships and native object IDs; WPS visual check still required'}


def catalog(path):
    files = load(path)
    result = []
    for i, p in enumerate(ordered_slides(files), 1):
        objects = []
        for e, nv in shapes(parse(files[p])):
            if nv is None:
                continue
            objects.append({'id': int(nv.get('id')), 'name': nv.get('name'),
                            'kind': e.tag.split('}')[-1], 'text': ''.join(e.itertext())})
        result.append({'slide': i, 'part': p, 'objects': objects})
    return result


def build(library, specpath, output):
    output = Path(output)
    if output.exists():
        raise ValueError('Output exists; use a new filename')
    files = load(library)
    pages = ordered_slides(files)
    spec = json.loads(Path(specpath).read_text(encoding='utf-8-sig'))
    if not isinstance(spec, dict) or set(spec) != {'slides'} or not isinstance(spec['slides'], list) or not spec['slides']:
        raise ValueError('Spec requires a nonempty slides array')
    pres = parse(files['ppt/presentation.xml'])
    slideids = pres.find('p:sldIdLst', NS)
    for e in list(slideids):
        slideids.remove(e)
    # Source section/custom-show lists may refer to the old slide order.
    for tag in ('extLst', 'custShowLst'):
        for e in pres.findall('p:'+tag, NS):
            pres.remove(e)
    rels = parse(files['ppt/_rels/presentation.xml.rels'])
    for e in list(rels):
        if e.get('Type').endswith('/slide'):
            rels.remove(e)
    types = parse(files['[Content_Types].xml'])
    for i, entry in enumerate(spec['slides'], 1):
        if not isinstance(entry, dict) or set(entry)-{'slide', 'texts', 'transforms', 'keep_shapes'}:
            raise ValueError('Unknown slide specification fields')
        num = entry.get('slide')
        if isinstance(num, bool) or not isinstance(num, int) or not 1 <= num <= len(pages):
            raise ValueError('Source slide number is outside library')
        src = pages[num-1]
        root = parse(files[src])
        tree = root.find('p:cSld/p:spTree', NS)
        if 'keep_shapes' in entry:
            chosen = [find_shape(root, k) for k in entry['keep_shapes']]
            for e, _ in shapes(root):
                if e not in chosen:
                    tree.remove(e)
        for name, value in entry.get('texts', {}).items():
            if not isinstance(value, str):
                raise ValueError('Replacement texts must be strings')
            set_text(find_shape(root, name), value)
        for name, values in entry.get('transforms', {}).items():
            transform(find_shape(root, name), values)
        for e in root.findall('p:timing', NS)+root.findall('p:transition', NS):
            root.remove(e)
        dest = f'ppt/slides/reuse-{i}.xml'
        files[dest] = encode(root)
        rr = parse(files[relpath(src)])
        for e in list(rr):
            if e.get('Type').endswith('/notesSlide'):
                rr.remove(e)
        files[relpath(dest)] = encode(rr)
        rid = f'rIdItoReuse{i}'
        if any(e.get('Id') == rid for e in rels):
            raise ValueError('Relationship ID collision')
        E.SubElement(rels, '{'+PKG+'}Relationship', Id=rid, Type=R+'/slide', Target='slides/reuse-'+str(i)+'.xml')
        E.SubElement(slideids, '{'+P+'}sldId', {'id': str(255+i), '{'+R+'}id': rid})
        E.SubElement(types, '{'+CT+'}Override', PartName='/'+dest, ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    files['ppt/presentation.xml'] = encode(pres)
    files['ppt/_rels/presentation.xml.rels'] = encode(rels)
    keep = reachable(files)
    for e in list(types):
        if e.get('PartName') and e.get('PartName').lstrip('/') not in keep:
            types.remove(e)
    files['[Content_Types].xml'] = encode(types)
    # Remove inherited personal/document-specific metadata from the generated copy.
    if 'docProps/core.xml' in keep:
        core = parse(files['docProps/core.xml'])
        for e in list(core):
            if e.tag.split('}')[-1] in ('creator', 'lastModifiedBy', 'title', 'subject', 'description', 'keywords'):
                core.remove(e)
        files['docProps/core.xml'] = encode(core)
    if 'docProps/app.xml' in keep:
        app = parse(files['docProps/app.xml'])
        for e in list(app):
            if e.tag.split('}')[-1] == 'Slides':
                e.text = str(len(spec['slides']))
            elif e.tag.split('}')[-1] in ('HeadingPairs', 'TitlesOfParts'):
                app.remove(e)
        files['docProps/app.xml'] = encode(app)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('xb') as stream:
        with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as z:
            for part in sorted(keep):
                z.writestr(part, files[part])
    return validate(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subs = parser.add_subparsers(dest='command', required=True)
    for command in ('catalog', 'build'):
        sub = subs.add_parser(command)
        sub.add_argument('--library', choices=('source', 'validated'), default='validated')
        if command == 'build':
            sub.add_argument('--spec', required=True)
            sub.add_argument('--output', required=True)
    sub = subs.add_parser('validate')
    sub.add_argument('pptx')
    args = parser.parse_args()
    try:
        if args.command == 'validate':
            result = validate(args.pptx)
        else:
            lib = ROOT/'assets'/('reference.pptx' if args.library == 'source' else 'validated.pptx')
            result = catalog(lib) if args.command == 'catalog' else build(lib, args.spec, args.output)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (ValueError, OSError, KeyError, E.ParseError, zipfile.BadZipFile) as exc:
        print(f'Error: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
