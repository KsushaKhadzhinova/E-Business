import openpyxl
import os

base = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
errors = []
ok = []

for root, dirs, files in os.walk(base):
    for f in files:
        if f.endswith('.xlsx'):
            path = os.path.join(root, f)
            try:
                wb = openpyxl.load_workbook(path, data_only=True)
                sheets = wb.sheetnames
                wb.close()
                ok.append(os.path.relpath(path, base))
            except Exception as e:
                errors.append((os.path.relpath(path, base), str(e)))

print(f"OK: {len(ok)}")
print(f"ERRORS: {len(errors)}")
for p, e in errors:
    print(f"  BROKEN: {p}")
    print(f"    {e}")

if not errors:
    print("\nВсе файлы читаются нормально.")
