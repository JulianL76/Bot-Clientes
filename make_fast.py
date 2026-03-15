import codecs

with codecs.open('bot_test_debug_fast.py', 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    # Reduce globals
    if 'pyautogui.PAUSE = 1.0' in line:
        lines[i] = line.replace('1.0', '0.2')
    
    # We want to keep original times for Asesor, Canal, Segmento.
    if 195 <= i <= 240:
        continue
        
    # Replace sleeps
    if 'smart_sleep(0.5)' in line:
        lines[i] = line.replace('smart_sleep(0.5)', 'smart_sleep(0.1)')
    elif 'time.sleep(0.5)' in line:
        lines[i] = line.replace('time.sleep(0.5)', 'time.sleep(0.1)')
    elif 'time.sleep(1.0)' in line:
        lines[i] = line.replace('time.sleep(1.0)', 'time.sleep(0.2)')
    elif 'time.sleep(1.5)' in line:
        lines[i] = line.replace('time.sleep(1.5)', 'time.sleep(0.3)')

with codecs.open('bot_test_debug_fast.py', 'w', 'utf-8') as f:
    f.writelines(lines)
print("Transformacion completada")
