import openpyxl,warnings,os
warnings.filterwarnings("ignore")
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
def show(f,sh,rows):
    p=[os.path.join(r,f) for r,_,fs in os.walk(base) if f in fs][0]
    wv=openpyxl.load_workbook(p,data_only=True)[sh]; wf=openpyxl.load_workbook(p)[sh]
    for r in rows:
        print(sh,r,[(c.coordinate,repr(wv[c.coordinate].value)[:25],"F" if str(c.value).startswith("=") else "C") for c in wf[r] if c.value is not None][:14])
show("analiz_konkurentov_Levitt_Kano_NotaCode.xlsx","01_Конкуренты",[14,15,30])
show("analiz_konkurentov_Levitt_Kano_NotaCode.xlsx","06_Сравнение",[31,32,40])
show("klassifikaciya_urovni_konkurencii_NotaCode.xlsx","07_Карта",[25,26,60])
