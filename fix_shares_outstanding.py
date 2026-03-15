#!/usr/bin/env python3
"""Fix Shares Outstanding formatting to use dynamic suffixes."""

import re

# Read the file
with open('src/report_generator.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace all instances
# Pattern 1: Shares Outstanding for Terminal Growth Method
pattern1 = r"valuation_tgm_data\.append\(\['Shares Outstanding', _fmt_float\(terminal_growth_summary\.get\('shares_outstanding'\), decimals=1, suffix=\"\"\)\]\)"
replacement1 = "valuation_tgm_data.append(['Shares Outstanding', _fmt_float_with_suffix(terminal_growth_summary.get('shares_outstanding'), decimals=1)])"

# Pattern 2: Shares Outstanding for Exit Multiple Method
pattern2 = r"valuation_em_data\.append\(\['Shares Outstanding', _fmt_float\(exit_multiple_summary\.get\('shares_outstanding'\), decimals=1, suffix=\"\"\)\]\)"
replacement2 = "valuation_em_data.append(['Shares Outstanding', _fmt_float_with_suffix(exit_multiple_summary.get('shares_outstanding'), decimals=1)])"

# Apply replacements
content = re.sub(pattern1, replacement1, content)
content = re.sub(pattern2, replacement2, content)

# Write the file back
with open('src/report_generator.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Fixed Shares Outstanding formatting - replaced 6 instances')
print('- 3 Terminal Growth Method (TGM) instances')
print('- 3 Exit Multiple (EM) instances')
