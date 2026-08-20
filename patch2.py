import codecs

with codecs.open(r'E:\POSAO\Skripta 2.0 - mejl\mail_engine.py', 'r', 'utf-8') as f:
    text = f.read()

parts = text.split('            header = {}', 1)

if len(parts) == 2:
    part1, part2 = parts
    part1 = part1.replace('                ws = wb.active\n', '')
    
    end_marker = '        self.log(f"\\n🎯 Kampanja završena'
    sub_parts = part2.split(end_marker, 1)
    
    if len(sub_parts) == 2:
        inner_block, end_block = sub_parts
        
        indented_inner = []
        for line in inner_block.split('\n'):
            if line.strip() == '':
                indented_inner.append('')
            else:
                indented_inner.append('    ' + line)
        
        new_text = part1 + '            for ws in wb.worksheets:\n' + '                header = {}\n' + '\n'.join(indented_inner) + end_marker + end_block
        
        with codecs.open(r'E:\POSAO\Skripta 2.0 - mejl\mail_engine.py', 'w', 'utf-8') as f:
            f.write(new_text)
        print('SUCCESS')
    else:
        print('FAIL to find end marker')
else:
    print('FAIL to find target_find')
