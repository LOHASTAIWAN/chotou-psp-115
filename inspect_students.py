# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('generate_deck_dual_page_v4.py', 'r', encoding='utf-8') as f:
    text = f.read()

pos_start = text.find('STUDENTS_DATA = [')
pos_end = text.find('def create_cover_slide')
code_str = text[pos_start:pos_end]

# create dummy color objects
class DummyColor:
    pass

C_ROSE_ACCENT = C_ROSE_LIGHT = C_AMBER_ACCENT = C_AMBER_LIGHT = C_BLUE_ACCENT = C_BLUE_LIGHT = C_EMERALD_ACC = C_EMERALD_LIGHT = 'color'
env = {
    'C_ROSE_ACCENT': C_ROSE_ACCENT,
    'C_ROSE_LIGHT': C_ROSE_LIGHT,
    'C_AMBER_ACCENT': C_AMBER_ACCENT,
    'C_AMBER_LIGHT': C_AMBER_LIGHT,
    'C_BLUE_ACCENT': C_BLUE_ACCENT,
    'C_BLUE_LIGHT': C_BLUE_LIGHT,
    'C_EMERALD_ACC': C_EMERALD_ACC,
    'C_EMERALD_LIGHT': C_EMERALD_LIGHT,
}
exec(code_str, env)
students = env['STUDENTS_DATA']

print(f"Total students: {len(students)}")
for s in students:
    print(f"=== Student #{s['id']:02d} ===")
    print(f"Title: {s['title']}")
    print(f"Risk: {s['risk']}")
    print(f"Choices count: {len(s['choices'])}")
    for i, c in enumerate(s['choices'][:5]):
        print(f"   {i+1}. {c[0]} - {c[1]} ({c[2]})")
    if len(s['choices']) > 5:
        print(f"   ... and {len(s['choices'])-5} more")
