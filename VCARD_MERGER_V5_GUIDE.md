# Phyllis's Contact Merging Magician v5

## User Guide

### Overview

vCard Contact Merger v5 is a production-ready application for merging duplicate contacts from two vCard files. It uses intelligent matching algorithms to identify duplicates that might be formatted differently across your contact lists.

### Key Features

- **Name Order Detection**: "John Smith" and "Smith, John" are correctly identified as the same person
- **Nickname Recognition**: "Bob Smith" matches "Robert Smith" (100+ common nicknames supported)
- **Phonetic Matching**: Uses Soundex algorithm to match similar-sounding names ("Smith" matches "Smyth")
- **Email Normalization**: Handles Gmail dots, plus addressing, and domain aliases
- **Phone Normalization**: Normalizes various phone formats and handles +1 country codes
- **High Performance**: O(n) bucketing algorithm processes 10,000+ contacts in under 5 seconds
- **Batch Workflow**: Review duplicates by confidence level for efficient processing
- **Pagination**: Handles large datasets with 100 groups per page

### Getting Started

1. **Launch the Application**
   ```bash
   python3 vcard_merger_v5.py
   ```

2. **Select Your Files**
   - Click "Browse..." to select your first vCard file (.vcf)
   - Click "Browse..." to select your second vCard file (.vcf)

3. **Adjust Matching Sensitivity** (Optional)
   - Use the slider to adjust how aggressive the matching should be
   - 75% is recommended for most cases
   - Lower values find more matches (more aggressive)
   - Higher values find fewer matches (more conservative)

4. **Click "Load & Find Duplicates"**
   - The app will load both files and identify duplicate groups
   - Progress is shown during loading

### Reviewing Duplicates

After loading, you'll see the **Batch Approval Workflow** screen showing duplicate groups organized by confidence level:

- **95-100%**: Very high confidence matches (likely exact duplicates)
- **90-94%**: High confidence matches
- **85-89%**: Good matches
- **80-84%**: Moderate confidence
- **75-79%**: Lower confidence (review carefully)
- **70-74%**: May need manual verification
- **50-69%**: Lowest confidence (highest chance of false positives)

### Processing a Batch

1. Click on a confidence level to review those groups
2. Each group shows:
   - Contact names
   - Match confidence percentage
   - Why the contacts matched
3. Use checkboxes to select which groups to merge
4. Click **Preview** to see the merged result before committing
5. Click **Merge Selected** to merge all selected groups

### Preview and Edit

When you click **Preview** on a group:

- See all original contacts side-by-side
- View the merged result
- See why the contacts were matched
- Review any warnings (different organizations, no phone/email overlap, etc.)
- Click **Edit Merged Contact** to customize the merged result

### Exporting

After processing duplicate groups:

1. Return to the Batch Overview screen
2. Click **Export Merged Contacts**
3. Choose a location to save the merged .vcf file

The export includes:
- All merged contacts
- Unique contacts (not in any duplicate group)
- Skipped groups (kept as separate contacts)

### Understanding Match Confidence

The app calculates confidence scores based on:

| Factor | Points |
|--------|--------|
| Email exact match | +50 |
| Phone match | +40 |
| Name exact match (after normalization) | +50 |
| Nickname match | +45 |
| Phonetic match (Soundex) | +35 |
| Same organization | +10 |

Score is capped at 100%.

### Warnings

The app will warn you about potential issues:

- **Different organizations**: Contacts work at different companies
- **Different work email domains**: May be different people with same name
- **Multiple area codes**: Contacts may be in different locations
- **Name-only match**: No email or phone overlap (higher risk of false positive)

### Supported Nickname Mappings

The app recognizes 100+ common nicknames, including:

**Male Names:**
- Bob, Rob, Bobby, Robbie -> Robert
- Bill, Will, Billy, Liam -> William
- Mike, Mick, Mickey -> Michael
- Jim, Jimmy, Jamie -> James
- Dick, Rick, Rich -> Richard
- And many more...

**Female Names:**
- Liz, Lizzy, Beth, Betty -> Elizabeth
- Kate, Katie, Kathy -> Katherine
- Jenny, Jen -> Jennifer
- Maggie, Meg, Peggy -> Margaret
- And many more...

### Tips for Best Results

1. **Start with high confidence batches** (95-100%) and work down
2. **Review warnings carefully** before merging flagged groups
3. **Use Preview** to verify merges before committing
4. **Edit merged contacts** if needed to correct names or remove duplicate data
5. **Export frequently** to save your progress

### Troubleshooting

**Q: The app is slow with large files**
A: The initial loading may take time for very large files (50,000+ contacts). Progress is shown during loading.

**Q: Some duplicates weren't found**
A: Try lowering the matching sensitivity slider to find more matches.

**Q: Too many false positives**
A: Increase the matching sensitivity slider to be more conservative.

**Q: Export failed**
A: Ensure you have write permissions to the selected location.

### Technical Requirements

- Python 3.6+
- Tkinter (included with most Python installations)
- No external dependencies required

### Files Created

- `vcard_merger_v5.py` - Main application (single file, ~2500 lines)
- `test_vcard_merger_v5.py` - Test suite for core functions
- `VCARD_MERGER_V5_GUIDE.md` - This documentation

### Running Tests

To verify the installation:

```bash
python3 test_vcard_merger_v5.py
```

All tests should pass, confirming:
- Name order detection works correctly
- Soundex phonetic matching works
- Nickname resolution works
- Email/phone normalization works
- Comprehensive name matching works

### Version History

**v5.0** - Production-Ready Edition
- Fixed critical name order detection bug
- Added 100+ nickname mappings
- Implemented Soundex phonetic matching
- Gmail email normalization (dots, plus addressing)
- O(n) bucketing algorithm for performance
- Batch approval workflow by confidence level
- Pagination for large datasets (100 groups/page)
- Comprehensive test suite

**v4.0** - Enhanced Edition
- Initial implementation of name matching improvements
- Basic nickname support

**v3.0** - Previous version with name order detection issues
