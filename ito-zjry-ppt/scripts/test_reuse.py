"""Behavior checks for native template reuse; run with Python standard library."""
import json
from pathlib import Path
import tempfile
import unittest
import reuse


class ReuseTests(unittest.TestCase):
    def test_copy_text_and_native_objects(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = Path(tmp)/'spec.json'
            spec.write_text(json.dumps({'slides': [
                {'slide': 1, 'texts': {'标题_可编辑': '新标题', '正文_长文本': '第一段\n第二段'}},
                {'slide': 2, 'transforms': {'复用_医师': {'x': 200, 'height': 160}}},
                {'slide': 2, 'texts': {'标题_可编辑': '独立副本'}}
            ]}, ensure_ascii=False), encoding='utf-8')
            out = Path(tmp)/'result.pptx'
            result = reuse.build(reuse.ROOT/'assets/validated.pptx', spec, out)
            self.assertEqual(result['slide_count'], 3)
            files = reuse.load(out)
            pages = reuse.ordered_slides(files)
            self.assertEqual(len(set(pages)), 3)
            first = reuse.parse(files[pages[0]])
            self.assertIn('新标题', ''.join(first.itertext()))
            paragraphs = reuse.find_shape(first, '正文_长文本').findall('p:txBody/a:p', reuse.NS)
            self.assertEqual(len(paragraphs), 2)
            for name in ('标题_可编辑', '正文_长文本'):
                for pr in reuse.find_shape(first, name).findall('.//a:rPr', reuse.NS):
                    self.assertEqual(pr.find('a:ea', reuse.NS).get('typeface'), 'Microsoft YaHei')
            second = reuse.parse(files[pages[1]])
            self.assertEqual(len(second.findall('p:cSld/p:spTree/p:grpSp', reuse.NS)), 3)
            third = reuse.parse(files[pages[2]])
            self.assertIn('独立副本', ''.join(third.itertext()))
            self.assertNotIn('独立副本', ''.join(second.itertext()))
            with self.assertRaises(ValueError):
                reuse.build(reuse.ROOT/'assets/validated.pptx', spec, out)

    def test_extract_original_icons(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = Path(tmp)/'spec.json'
            spec.write_text(json.dumps({'slides': [{'slide': 34, 'keep_shapes': [99, 249],
                'transforms': {'99': {'x': 200, 'height': 160}}}]}), encoding='utf-8')
            out = Path(tmp)/'icons.pptx'
            reuse.build(reuse.ROOT/'assets/reference.pptx', spec, out)
            files = reuse.load(out)
            self.assertEqual(len(reuse.shapes(reuse.parse(files[reuse.ordered_slides(files)[0]]))), 2)

    def test_unknown_shape_stops_before_writing(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = Path(tmp)/'spec.json'
            spec.write_text(json.dumps({'slides': [{'slide': 1, 'texts': {'missing': 'text'}}]}), encoding='utf-8')
            out = Path(tmp)/'bad.pptx'
            with self.assertRaises(ValueError):
                reuse.build(reuse.ROOT/'assets/validated.pptx', spec, out)
            self.assertFalse(out.exists())


if __name__ == '__main__':
    unittest.main()
