# Excel VBA Email Copy and Clean Macro

This repository contains an Excel VBA macro that copies, cleans, and normalizes email addresses while preserving the original source data. It is designed to handle real-world, messy email input commonly found in CRM exports, marketing lists, and legacy spreadsheets.

## Overview

The macro prompts the user to select up to four source columns. For each source column, a new column is inserted immediately to the right. The original values are copied into the new column and cleaned using extensive normalization logic. The original column is never modified.

Source cells are visually highlighted when changes occur, allowing quick auditing and verification.

## Features

- Supports up to four source columns
- Inserts cleaned columns to the right of each source column
- Non-destructive processing
- Visual change tracking
- Optimized for large datasets

## Cleaning Logic

The macro applies the following rules:

- Trims whitespace and removes internal spaces
- Removes placeholder values such as "no email" or "not available"
- Extracts email addresses from angle brackets
- Merges multiple @ symbols into a valid structure
- Collapses repeated punctuation
- Converts commas to periods in domains when appropriate
- Removes trailing invalid characters
- Forces lowercase output

## Domain and Provider Fixes

The macro corrects common domain errors, including:

- Invalid .com endings such as .cmo, .xom, .con, .coom
- Duplicate .com.com sequences
- Concatenated domains like gmailcom or hotmailcom

It also normalizes common email providers by correcting misspellings and forcing canonical names for:

- Gmail
- Yahoo
- Outlook
- Hotmail
- Live
- iCloud

Missing .com extensions are automatically appended for recognized providers.

## Visual Change Indicators

- Blue highlight: only letter casing changed
- Yellow highlight: content changed

Highlights are applied to the original source column for easy review.

## How to Use

1. Open Excel and enable macros
2. Press ALT + F11 to open the VBA editor
3. Insert a new module and paste the macro code
4. Save the workbook as a macro-enabled file (.xlsm)
5. Activate the worksheet containing email data
6. Run the macro `CopyAndCleanEmails_OriginalLogic`
7. Enter source column letters separated by commas (example: A,C,D)

## Requirements

- Microsoft Excel (desktop version)
- VBA enabled

## Notes

This macro prioritizes data safety, auditability, and transparency. The logic reflects real-world data correction needs rather than strict RFC email validation rules.


