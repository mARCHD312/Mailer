import codecs
import re

with codecs.open(r"E:\POSAO\Skripta 2.0 - mejl\mail_engine.py", "r", "utf-8") as f:
    lines = f.readlines()

new_lines = []
in_sheet = False

for i, line in enumerate(lines):
    if "ws = wb.active" in line:
        new_lines.append(line.replace("ws = wb.active", "for ws in wb.worksheets:\n                if not self.is_running or sent_count >= self.limit:\n                    break"))
        in_sheet = True
        continue
        
    if "Kampanja" in line and "poslato" in line.lower():
        in_sheet = False
        
    if in_sheet:
        if "except Exception as e:" in line and "GRESKA pri" in lines[i+1]:
            # This is the except block for wb.load_workbook
            in_sheet = False
            new_lines.append(line)
        elif line.strip() == "":
            new_lines.append(line)
        else:
            new_lines.append("    " + line)
    else:
        new_lines.append(line)

with codecs.open(r"E:\POSAO\Skripta 2.0 - mejl\mail_engine.py", "w", "utf-8") as f:
    f.writelines(new_lines)
