# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based Excel data processing project that merges invoice data (9月销项.xlsx) into sales data (农品达8.26-9.25销售.xlsx) using both exact and fuzzy matching algorithms.

## Environment Setup

**Python Environment:**
```bash
# Activate the virtual environment
source venv/bin/activate  # macOS/Linux
```

**Dependencies:**
- pandas: Excel file reading/writing
- openpyxl: Excel formatting and styling
- difflib: String similarity matching (built-in)

**Dependencies:**
- streamlit: Web GUI framework

## Running the App

**Recommended: Streamlit GUI（面向非技术用户）**
```bash
streamlit run app.py
```

**Legacy CLI scripts (deprecated):**
1. **merge_invoices.py** - Basic version (single sheet)
2. **merge_invoices_final.py** - Enhanced version with red highlighting (single sheet)
3. **merge_all_sheets.py** - Full version processing all sheets

## File Structure

- `merge.py` - Core algorithm module (parameterized, importable)
- `app.py` - Streamlit GUI interface
- `merge_all_sheets.py` - Legacy CLI script
- `requirements.txt` - Python dependencies

## Core Architecture

### Data Processing Flow

1. **Invoice Data Aggregation**: Reads 9月销项.xlsx and aggregates quantities/amounts by product name
2. **Product Name Cleaning**: Removes category prefixes (format: `*category*productname`) from invoice data
3. **Two-Phase Matching Strategy**:
   - First collects all possible matches across all sheets
   - Deduplicates to keep only the best match per invoice product
   - Then applies matches to update sales data
4. **Multi-Sheet Processing**: Processes 4 sheets with different column configurations:
   - 蔬菜 (vegetables): uses "开票数量", "开票金额"
   - 肉蛋 (meat/eggs): uses "开票数量", "开票金额"
   - 9%: uses "已开票数量", "已开票金额"
   - 13%: uses "已开票数量", "已开票金额"

### Key Algorithms

**String Similarity Matching** (`similarity_ratio`, `find_best_match`):
- Uses `SequenceMatcher` from difflib for fuzzy matching
- Threshold: 75% similarity required
- Returns best match with highest similarity ratio

**Deduplication Strategy** (`deduplicate_matches`):
- Critical for handling cases where one invoice product could match multiple sales products
- Groups matches by invoice product
- Keeps only the match with highest similarity score
- Prevents double-counting of invoice data

**Product Name Cleaning** (`remove_category_prefix`):
- Removes `*category*` prefix from invoice product names
- Handles format: `*分类*商品名称` → `商品名称`
- Essential for matching since sales data lacks category prefixes

### Output Files

Scripts generate these files:
- `农品达8.26-9.25销售_已更新.xlsx`: Updated sales data with merged quantities/amounts
- `模糊匹配详情_所有sheet.xlsx`: Details of all fuzzy matches with similarity scores
- `9月销项_未匹配记录.xlsx`: Invoice records that weren't matched to any sales data (red highlighted)

### Visual Formatting

- Fuzzy match rows are highlighted in red font using openpyxl
- Unmatched invoice records are highlighted in red
- Original Excel structure and non-processed sheets preserved
