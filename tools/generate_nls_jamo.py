"""Capture Jamo composition states from Windows NLS API outputs.

Run on the same NLS version as generate_nls_weights.py. Indices and old-Hangul
flags are inferred from synthetic native sort keys, not copied weight tables.
Composition structure follows MS-UCODEREF MapOldHangulSortKey; weights are
independently captured. Validate with tests/test_portable_collation.py.
"""
from excel_model_lab.collation import _native_text_key as key
import unicodedata
P=lambda s:key(s).rsplit(b'\1',4)[0]
modern={P(chr(c)):c for c in range(0xac00,0xd7a4)}; modern[bytes.fromhex("c002")]=0xabff
points=[c for a,b in [(0x1100,0x11ff),(0xa960,0xa97f),(0xd7b0,0xd7ff)] for c in range(a,b+1) if unicodedata.category(chr(c))!='Cn']
def cls(c):return 0 if c<=0x115f or 0xa960<=c<=0xa97f else 1 if c<=0x11a7 or 0xd7b0<=c<=0xd7c6 else 2
states={}
for c in points:
 cl=cls(c)
 solo=P(chr(c));weights=(solo[3]&127,solo[5],solo[7])
 probe=P(chr(c)+chr(0x1161)) if cl==0 else P(chr(0x1100)+chr(c)) if cl==1 else P(chr(0x1100)+chr(0x1161)+chr(c))
 old=len(probe)==8
 if cl==0:index=modern[solo[:2]]-0xac00+1
 elif cl==1:
  probe=P(chr(0x1100)+chr(c)) if old else P(chr(0x1100)+chr(c)+chr(0x11c3))
  index=modern[probe[:2]]-0xac00
  if not old:index-=index%28
 else:index=modern[P(chr(0x1100)+chr(c))[:2]]-0xac00+1
 if c==0x115f:index=11171
 if c==0x1160:index=0
 states[c]=(old,(index//588,index//28%21,index%28),weights[cl])

import json
from pathlib import Path
for c,state in list(states.items()):
 states[c]=[*state,(P(chr(c)+"\u309d")[len(P(chr(c))):]).hex()]
Path(__file__).resolve().parents[1].joinpath("src/excel_model_lab/data/nls-jamo-v1.json").write_text(json.dumps(states,separators=(",",":")))
