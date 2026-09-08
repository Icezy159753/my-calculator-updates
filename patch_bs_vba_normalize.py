from pathlib import Path


root = Path(__file__).resolve().parent
module = root / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"
build = root / "tmp_bs_xlsm_build" / "build_xlsm.py"

module_text = module.read_text(encoding="utf-8-sig")
module.write_text(module_text, encoding="ascii")

text = build.read_text(encoding="utf-8")
old = '''        component = workbook.VBProject.VBComponents.Import(str(MODULE))
        del component

        workbook.SaveAs(str(OUTPUT), FileFormat=52)'''
new = '''        component = workbook.VBProject.VBComponents.Import(str(MODULE))
        del component

        if workbook.Worksheets(1).Name != "Control":
            control.Move(workbook.Worksheets(1))
        if workbook.Worksheets(2).Name != "Rawdata":
            raw.Move(workbook.Worksheets(2))

        workbook.SaveAs(str(OUTPUT), FileFormat=52)'''
assert text.count(old) == 1
build.write_text(text.replace(old, new, 1), encoding="utf-8")
print("normalized VBA module and workbook sheet order")
