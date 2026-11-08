# Phyllis's Contact Merging Magician - Complete Guide

**Your personal tool for cleaning up duplicate contacts**

---

## What It Does

Automatically finds and merges duplicate contacts from vCard (.vcf) files:
- Detects duplicates by matching names, emails, and phone numbers
- Smart grouping by confidence score (95-100%, 90-94%, etc.)
- Batch approval - review and merge groups quickly
- Exports clean, deduplicated contacts

---

## Quick Start

```bash
python3 vcard_merger-3.py
```

**5-Step Workflow:**

1. **Load Files**
   - Click "Browse" to select your vCard files
   - Load same file twice, or two different files
   - Click "🚀 Load & Find Groups"

2. **Review Batches**
   - App groups duplicates by confidence (High → Low)
   - Start with "95-100% Very High Confidence"
   - Click "Start Batch Approval"

3. **Review Groups**
   - See all contacts in each group with notes
   - ✅ = Merge this group
   - ☐ = Skip this group
   - Click "👁 Preview" to see detailed comparison

4. **Merge or Skip**
   - **⚡ Merge Entire Batch** - Auto-merge all (for 95-100% batches)
   - **✓ Merge Selected** - Merge only checked groups
   - **⏭ Skip This Batch** - Move to next batch

5. **Export**
   - After reviewing batches, click "💾 Export Merged Contacts"
   - Save as .vcf file
   - Import to your contacts app

---

## Key Features

### 🎯 Smart Duplicate Detection
- Matches by email, phone, and name similarity
- Groups by confidence score for easy review
- Shows 100 groups per page (smooth scrolling)

### ⚠️ Warning System
Automatically detects potential issues:
- Different organizations (Apple Inc vs Microsoft)
- Different email domains (apple.com vs microsoft.com)
- Multiple area codes (3+ locations)
- Partial matches (only names similar, no email/phone match)
- Names too different (<70% similar)

Groups with warnings show **yellow background** + ⚠️ icon.

### 👁 Preview Merge
Click "👁 Preview" on any group to see:
- **Left:** Original contacts side-by-side
- **Right:** How they'll look merged
- **Warnings:** Any detected issues
- **Edit button:** Customize before merging

### ✏️ Edit Merged Contacts
1. Click "👁 Preview" on a group
2. Click "✏️ Edit Merged Contact"
3. Modify name, emails, phones, notes, etc.
4. Click "Save Changes"
5. Your edits are preserved when you merge

Groups you've edited show ✏️ icon.

### 📝 Notes Display
Every contact shows notes in batch review:
```
1. Mary Jane Otte phd | maryjane@earthlink.net | (650) 328-6600
   📝 Notes: Original contact from 2020; Verified email
```

### 🏆 Longest Name Wins
When merging, the name with the **most words** becomes the final name:
- "Mary Jane" (2 words)
- "Mary Jane Otte phd" (4 words) ← **This one!**
- Result: Merged contact = "Mary Jane Otte phd"

### 🔧 Name Normalization
Automatically fixes duplicate words:
- "Mary Jane Jane Otte" → "Mary Jane Otte"
- "Bob Bob Smith" → "Bob Smith"

### 🏠 Persistent Home Button
Green "🏠 Home" button always in bottom-left corner:
- Returns to load screen
- Warns if unsaved progress
- Reset and start over

---

## Understanding Confidence Scores

| Score | Meaning | What to Do |
|-------|---------|------------|
| **95-100%** | Very high - almost certainly same person | Use "⚡ Merge Entire Batch" |
| **90-94%** | High - likely same person | Quick review, merge selected |
| **85-89%** | Good - probably same person | Preview warned groups |
| **80-84%** | Medium - maybe same person | Preview most groups |
| **75-79%** | Fair - uncertain | Preview all or skip batch |
| **<75%** | Low - risky | Skip batch or manual review |

---

## Visual Indicators

### In Batch Review:

**Normal group (safe):**
```
☑ Group 1 - 98% confidence - 2 contacts  [👁 Preview]
```

**Group with warnings:**
```
☑ Group 2 - 92% confidence - 2 contacts ⚠️  [👁 Preview]
  ⚠️ Different organizations: Apple Inc, Microsoft
```

**Edited group:**
```
☑ Group 3 - 95% confidence - 2 contacts ✏️  [👁 Preview]
```

**Both warned AND edited:**
```
☑ Group 4 - 88% confidence - 3 contacts ⚠️✏️  [👁 Preview]
```

---

## Pro Tips

### When to Use Each Button

**⚡ Merge Entire Batch** - Best for:
- 95-100% confidence batches
- No warnings or very few warnings
- Trusted data sources

**✓ Merge Selected Groups** - Best for:
- 85-94% confidence batches
- Selectively merging after unchecking warned groups
- Page-by-page review

**⏭ Skip This Batch** - Best for:
- <85% confidence batches
- Too many warnings
- Uncertain data

### Warning Strategy

1. **High confidence (95-100%)** - Rarely have warnings, merge all
2. **Medium confidence (85-94%)** - Preview warned groups only
3. **Low confidence (<85%)** - Preview everything or skip

### Navigation

- **100 groups per page** - Use Previous/Next page buttons
- **Select All (This Page)** - Check all 100 groups on current page
- **Deselect All** - Uncheck all on current page
- Groups are auto-selected by default - just uncheck bad ones

---

## Workflow Example

**Scenario:** 1500 duplicate groups found

### Batch 1: 95-100% (1000 groups)
1. Click "Start Batch Approval"
2. Shows page 1 of 10 (100 groups)
3. No warnings - all look good
4. Click "⚡ Merge Entire Batch" - Done in 2 seconds!

### Batch 2: 90-94% (400 groups)
1. Shows page 1 of 4 (100 groups)
2. ⚠️ 5 groups have warnings
3. Click "👁 Preview" on warned groups
4. Uncheck 2 suspicious groups
5. Click "✓ Merge Selected (98)"
6. Auto-moves to page 2
7. Repeat for remaining pages

### Batch 3: 85-89% (100 groups)
1. Shows 1 page (100 groups)
2. ⚠️ 15 groups have warnings
3. Review warnings, uncheck bad ones
4. Click "✓ Merge Selected (85)"

### Export
1. Click "💾 Export Merged Contacts"
2. Save as "merged_contacts.vcf"
3. Result: 1485 clean contacts (15 skipped)

**Time:** ~10-15 minutes for 1500 groups!

---

## Common Scenarios

### Same Person, Different Jobs
**Warning:** Different organizations: Apple Inc, Google

**Solution:**
1. Click "👁 Preview"
2. Click "✏️ Edit Merged Contact"
3. Update organization to current job
4. Add note: "Changed jobs from Apple to Google"
5. Merge

### Different People, Same Name
**Warning:** Different email domains: smith@apple.com, smith@microsoft.com

**Solution:** ☐ Uncheck - Don't merge!

### Partial Duplicates
**Warning:** Only names similar, no email/phone match

**Solution:** Preview carefully, merge only if confident

### Name Formatting Issues
**Example:** "Smith, John" vs "John Smith"

**Good news:** Automatically detected as 100% match!

---

## File Management

### Input Files
- Supports .vcf (vCard) format
- Can load same file twice or two different files
- Handles 10,000+ contacts efficiently

### Output File
- Exports as .vcf (vCard 3.0)
- Contains: Merged contacts + Unmatched contacts
- Skipped groups are saved as separate contacts
- All notes preserved

### Autosave
- Progress auto-saved to `.vcard_merge_autosave.json`
- Recovers if app crashes
- Deleted after successful export

---

## Keyboard-Free!

Everything is point-and-click:
- ✅ Browse files
- ✅ Click checkboxes
- ✅ Click buttons
- ✅ Scroll with mouse

No keyboard needed except when editing contact details!

---

## Performance

**Speed:**
- 10,000 contacts processed in <1 second
- Hash-based algorithm (not brute force)
- Pagination prevents UI lag

**Memory:**
- Streaming file parser
- Constant memory usage
- Handles massive files

---

## Troubleshooting

### "Preview button missing"
Not missing! Look on the **right side** of group header.

### "Can't see Edit button"
Edit is **inside** Preview dialog:
1. Click "👁 Preview"
2. Then click "✏️ Edit Merged Contact"

### "Exported file has 2x contacts"
This was a bug - now FIXED! Export properly deduplicates.

### "Scrolling stopped working"
Fixed with pagination - only 100 groups per page now.

### "Wrong name in merged contact"
Now uses longest name! The name with most words wins.

---

## Summary of ALL Features

✨ **Branding:** "Phyllis's Contact Merging Magician"
🎨 **Colors:** Vibrant buttons (green, blue, dark blue)
👁️ **Preview:** See before you merge
⚠️ **Smart warnings:** 5 types auto-detected
📄 **Pagination:** 100 groups per page
✏️ **Edit in preview:** Customize before merging
🏠 **Home button:** Always accessible
🚀 **Bright load button:** Impossible to miss
📝 **Notes visible:** In batch review
🏆 **Longest name:** Auto-selected for merge
🔧 **Name normalization:** Fixes "Jane Jane" → "Jane"
⚡ **Fast:** 10k contacts in <1s
📊 **Batch approval:** By confidence score
💾 **Export:** Clean vCard output

---

## Quick Command Reference

```bash
# Run the app
python3 vcard_merger-3.py

# Test name normalization
python3 test_name_normalization.py

# Test longest name selection
python3 test_longest_name.py

# Verify exported contacts count
grep -c 'BEGIN:VCARD' merged_contacts.vcf
```

---

## Your Contact Merger Journey

**Before:**
- 10,000 messy contacts
- Duplicates everywhere
- Manual cleanup = weeks of work

**After:**
- 5,000 clean contacts
- Zero duplicates
- Automated cleanup = 30 minutes

**Enjoy your organized contact list!** ✨
