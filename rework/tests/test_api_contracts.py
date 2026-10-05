import json,sys,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from generate_api import generate
from generate_examples import generate as examples
class ApiContracts(unittest.TestCase):
 def test_generated_sources_match(self):
  for p,content in generate().items():self.assertEqual(p.read_text(),content,str(p))
 def test_generated_examples_match(self):
  self.assertEqual((ROOT/'workbook/api-examples.json').read_text(),examples())
 def test_every_interface_has_example(self):
  contracts=json.loads((ROOT/'api/contracts.json').read_text())['functions']
  examples_=json.loads(examples())['sheets']
  functions={s['name'] for s in contracts if not s['command']}
  self.assertEqual(len(functions),95)
  self.assertEqual(functions,{e['name'] for s in examples_ for e in s['examples']})
  self.assertEqual(sum(s['command'] for s in contracts),11)
 def test_vba_line_and_ascii_limits(self):
  for p in (ROOT/'src/vba').glob('*.bas'):
   p.read_text().encode('ascii')
   self.assertLessEqual(max(map(len,p.read_text().splitlines())),1023,str(p))
 def test_reserved_buffers_are_not_shrunk(self):
  c={s['name']:s for s in json.loads((ROOT/'api/contracts.json').read_text())['functions']}
  for name,arg,size in [('swe_heliacal_ut','dret',10),('swe_get_orbital_elements','dret',50),('swe_heliacal_pheno_ut','darr',50),('swe_houses_ex2','cusps',37)]:
   self.assertEqual(next(p['capacity'] for p in c[name]['parameters'] if p['name']==arg),size)
if __name__=='__main__':unittest.main()
