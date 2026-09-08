from pathlib import Path

p = Path("tmp_bs_xlsm_build/build_xlsm_pure.py")
s = p.read_text(encoding="utf-8")
s = s.replace('"Baby boomer (60+)"', '"Baby boomer (60 - )"')
p.write_text(s, encoding="utf-8")
