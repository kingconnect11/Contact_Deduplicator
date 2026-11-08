# vCard Contact Merger v4 - Quick Guide

## What's Fixed in v4

### 🔧 Main Bug Fix: Name Order Detection
The previous version was treating these as **different people**:
- "John Smith" vs "Smith, John"
- "Mary Jane Otte" vs "Otte, Mary Jane"

**Now they correctly match as the same person!**

### ✨ New Features

1. **Nickname Matching**
   - "Bob Smith" = "Robert Smith" ✓
   - "Bill Johnson" = "William Johnson" ✓
   - "Liz Taylor" = "Elizabeth Taylor" ✓
   - 100+ nickname mappings included

2. **Phonetic Matching (Soundex)**
   - "Smith" matches "Smyth" ✓
   - "Meyer" matches "Mayer" ✓
   - "Johnson" matches "Johnsen" ✓

3. **Smart Email Normalization**
   - Gmail dots ignored: "john.doe@gmail.com" = "johndoe@gmail.com"
   - Plus addressing: "user+tag@gmail.com" = "user@gmail.com"
   - Case insensitive

4. **Phone Normalization**
   - "(650) 555-1234" = "+1-650-555-1234" = "650.555.1234"

5. **Match Explanations**
   - See WHY contacts were matched (email match, name match, etc.)

6. **Deduplication Report**
   - Statistics on what was merged

## How to Use

```bash
python3 vcard_merger-4.py
```

1. **Load Files** - Select your two vCard (.vcf) files
2. **Review Batches** - Groups are organized by confidence level
3. **Merge or Skip** - Check/uncheck groups, then merge
4. **Export** - Save your clean, deduplicated contacts

## Confidence Levels

| Score | Meaning | Recommendation |
|-------|---------|----------------|
| 95-100% | Very high confidence | Auto-merge safely |
| 90-94% | High confidence | Quick review |
| 80-89% | Medium confidence | Review warnings |
| 70-79% | Fair confidence | Review carefully |
| <70% | Low confidence | Manual review |

## Keyboard Tips

- All interactions are point-and-click
- No keyboard needed except when editing contact details

## File Requirements

- Input: vCard format (.vcf files)
- Output: vCard 3.0 format (.vcf file)
- Works with 2 files at a time (can be the same file twice)

Enjoy your clean contact list! ✨
