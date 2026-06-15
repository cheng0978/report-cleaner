# Report Cleaner

A small Python tool that turns a messy Excel or CSV file into a clean,
summarized report in seconds — built with pandas and Streamlit.

## What it does

1. Upload a `.csv` or `.xlsx` file
2. Automatically cleans the data — trims whitespace, removes blank and
   duplicate rows, and converts text-formatted numbers back to numbers
3. Pick a column to group by and a numeric column to summarize
4. Get totals, averages, and counts with a bar chart
5. Download a clean two-sheet Excel report (cleaned data + summary)

## Why I built it

Manually tidying spreadsheets is one of the most common, repetitive tasks
small businesses deal with. This tool automates the whole flow and shows the
user exactly what was cleaned, so nothing happens in a black box.

## Tech

- Python
- pandas (data cleaning & aggregation)
- Streamlit (interface)
- openpyxl (Excel export)

## Run it locally
