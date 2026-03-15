#!/usr/bin/env python3
"""Fix Shares Outstanding formatting for appendix functions."""

with open('src/report_generator.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace all instances
old_tgm = "valuation_tgm_data.append(['Shares Outstanding', _fmt_float(terminal_growth_summary.get('shares_outstanding'), decimals=1, suffix=\"\")])"
new_tgm = "valuation_tgm_data.append(['Shares Outstanding', _fmt_float_with_suffix(terminal_growth_summary.get('shares_outstanding'), decimals=1)])"

old_em = "valuation_em_data.append(['Shares Outstanding', _fmt_float(exit_multiple_summary.get('shares_outstanding'), decimals=1, suffix=\"\")])"
new_em = "valuation_em_data.append(['Shares Outstanding', _fmt_float_with_suffix(exit_multiple_summary.get('shares_outstanding'), decimals=1)])"

content = content.replace(old_tgm, new_tgm)
content = content.replace(old_em, new_em)

with open('src/report_generator.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Replaced all 4 instances of Shares Outstanding formatting")
print("  - Fixed 2 Terminal Growth Method (TGM) instances")         
print("  - Fixed 2 Exit Multiple (EM) instances")
