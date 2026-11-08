# Ralph Wiggum Prompt: vCard Contact Merger v5

## TASK

Create a fully functional, production-ready vCard Contact Merger application (v5) that correctly identifies and merges duplicate contacts from two vCard files. The application must fix the critical name-order detection bug and implement intelligent matching algorithms.

---

## SUCCESS CRITERIA

All of the following must be true before outputting `<promise>COMPLETE</promise>`:

### 1. Core Bug Fixes (CRITICAL)
- [ ] "John Smith" and "Smith, John" are detected as the SAME person (100% match)
- [ ] "Mary Jane Otte" and "Otte, Mary Jane" are detected as the SAME person
- [ ] "Dr. John Smith" and "Smith, John" are detected as the SAME person
- [ ] "John Smith Jr." and "Smith, John" are detected as the SAME person
- [ ] Names with titles (Dr., Mr., Mrs., Prof.) and suffixes (Jr., Sr., PhD, III) are normalized before comparison

### 2. Name Matching Enhancements
- [ ] Nickname recognition: "Bob Smith" matches "Robert Smith" (100+ common nicknames mapped)
- [ ] Phonetic matching using Soundex: "Smith" matches "Smyth", "Meyer" matches "Mayer"
- [ ] Case-insensitive and punctuation-insensitive name comparison
- [ ] Handle multi-word first names: "Mary Jane" = "Mary Jane"
- [ ] Handle hyphenated last names: "Smith-Jones" matches appropriately

### 3. Email Normalization
- [ ] Case insensitive: "John.Doe@Gmail.com" = "john.doe@gmail.com"
- [ ] Gmail dot trick: "john.doe@gmail.com" = "johndoe@gmail.com"
- [ ] Plus addressing: "user+tag@gmail.com" = "user@gmail.com"
- [ ] Domain normalization: "googlemail.com" = "gmail.com"

### 4. Phone Normalization
- [ ] Strip all non-digits for comparison
- [ ] Handle +1 country code: "+1-650-555-1234" = "650-555-1234"
- [ ] Handle various formats: "(650) 555-1234" = "650.555.1234" = "6505551234"
- [ ] Compare last 10 digits for US numbers, last 7 as fallback

### 5. Confidence Scoring (0-100)
- [ ] Email exact match: +50 points
- [ ] Phone match: +40 points
- [ ] Name exact match (after normalization): +50 points
- [ ] Name reversed order match: +50 points
- [ ] Nickname match: +45 points
- [ ] Phonetic match (Soundex): +35 points
- [ ] Name similarity >80%: proportional points
- [ ] Score capped at 100

### 6. Bucketing Algorithm (Performance)
- [ ] Email buckets: exact normalized email → contact indices
- [ ] Phone buckets: last 10 digits → contact indices
- [ ] Canonical name buckets: "firstname lastname" sorted alphabetically → contact indices
- [ ] Soundex buckets: soundex(first) + soundex(last) → contact indices
- [ ] Nickname-expanded buckets: canonical name with nicknames resolved → contact indices
- [ ] O(n) bucketing, O(k) pair evaluation where k << n²

### 7. GUI Requirements (Tkinter)
- [ ] Application title: "Phyllis's Contact Merging Magician"
- [ ] Two-file selection interface with Browse buttons
- [ ] Similarity threshold slider (0.5 to 0.95)
- [ ] Batch approval workflow organized by confidence level (95-100%, 90-94%, 85-89%, etc.)
- [ ] Preview merge dialog showing original contacts side-by-side with merged result
- [ ] Edit merged contact before saving
- [ ] Warning system for potential false positives (different orgs, different email domains, etc.)
- [ ] Pagination: 100 groups per page for large datasets
- [ ] Home button always visible in status bar
- [ ] Progress indicators during file loading and group detection

### 8. Export Requirements
- [ ] Export to vCard 3.0 format (.vcf)
- [ ] Merged contacts include all emails, phones, addresses, notes from originals
- [ ] Skipped groups kept as separate individual contacts
- [ ] Unique (non-duplicate) contacts preserved
- [ ] No duplicate contacts in output
- [ ] Deduplication report available showing statistics

### 9. Code Quality
- [ ] All functions have docstrings
- [ ] No hardcoded paths
- [ ] Handles file encoding errors gracefully (UTF-8 with fallback)
- [ ] No unhandled exceptions in GUI
- [ ] Works on macOS, Windows, and Linux
- [ ] No external dependencies beyond Python standard library (tkinter, re, json, os, collections, difflib, threading, queue)

### 10. Test Verification
- [ ] Create a standalone test script that verifies all core matching functions
- [ ] Test: "John Smith" vs "Smith, John" → MATCH
- [ ] Test: "Bob Smith" vs "Robert Smith" → MATCH (nickname)
- [ ] Test: "Smith" vs "Smyth" → MATCH (phonetic)
- [ ] Test: "john.doe@gmail.com" vs "johndoe@gmail.com" → MATCH
- [ ] Test: "(650) 555-1234" vs "+1-650-555-1234" → MATCH
- [ ] All tests must pass

---

## PROCESS

1. **Analyze**: Read the existing `vcard_merger-3.py` to understand current architecture
2. **Design**: Plan the fix for name-order detection and enhanced matching
3. **Implement Core Functions First**:
   - `parse_name_parts(full_name)` - Extract first/last regardless of format
   - `get_canonical_name(contact)` - Returns normalized "firstname lastname"
   - `soundex(name)` - Phonetic encoding
   - `expand_nicknames(name)` - Map nicknames to canonical names
   - `normalize_email(email)` - Full email normalization
   - `normalize_phone(phone)` - Full phone normalization
   - `calculate_match_confidence(c1, c2)` - Enhanced scoring with explanations
4. **Implement Bucketing**: Create efficient O(n) bucketing with multiple strategies
5. **Implement GUI**: Build complete Tkinter interface with all features
6. **Test**: Run the standalone test script to verify all matching works
7. **Iterate**: If any test fails, fix and re-run until all pass

---

## DELIVERABLES

1. `vcard_merger-5.py` - Complete application (single file, ~2500+ lines)
2. `test_vcard_merger_5.py` - Standalone test script for core functions
3. `VCARD_MERGER_V5_GUIDE.md` - User documentation

---

## IMPLEMENTATION DETAILS

### Nickname Mapping (include at minimum)
```python
NICKNAMES = {
    # Male names
    'bob': 'robert', 'rob': 'robert', 'bobby': 'robert', 'robbie': 'robert',
    'bill': 'william', 'will': 'william', 'billy': 'william', 'liam': 'william',
    'mike': 'michael', 'mick': 'michael', 'mickey': 'michael',
    'jim': 'james', 'jimmy': 'james', 'jamie': 'james',
    'joe': 'joseph', 'joey': 'joseph',
    'tom': 'thomas', 'tommy': 'thomas',
    'dick': 'richard', 'rick': 'richard', 'rich': 'richard', 'ricky': 'richard',
    'dan': 'daniel', 'danny': 'daniel',
    'dave': 'david', 'davy': 'david',
    'steve': 'steven', 'stevie': 'steven',
    'chris': 'christopher', 'kit': 'christopher',
    'matt': 'matthew', 'matty': 'matthew',
    'tony': 'anthony', 'ant': 'anthony',
    'andy': 'andrew', 'drew': 'andrew',
    'nick': 'nicholas', 'nicky': 'nicholas',
    'ed': 'edward', 'eddie': 'edward', 'ted': 'edward', 'teddy': 'edward',
    'al': 'albert', 'bert': 'albert',
    'alex': 'alexander', 'xander': 'alexander',
    'ben': 'benjamin', 'benny': 'benjamin', 'benji': 'benjamin',
    'chuck': 'charles', 'charlie': 'charles',
    'frank': 'francis', 'frankie': 'francis',
    'fred': 'frederick', 'freddy': 'frederick',
    'greg': 'gregory',
    'harry': 'henry', 'hank': 'henry',
    'jack': 'john', 'johnny': 'john', 'jon': 'john',
    'jerry': 'gerald', 'gerry': 'gerald',
    'jeff': 'jeffrey', 'geoff': 'geoffrey',
    'ken': 'kenneth', 'kenny': 'kenneth',
    'larry': 'lawrence', 'laurie': 'lawrence',
    'lenny': 'leonard', 'leo': 'leonard',
    'lou': 'louis', 'louie': 'louis',
    'marty': 'martin',
    'max': 'maxwell', 'maxie': 'maxwell',
    'nate': 'nathan', 'nathaniel': 'nathan',
    'pat': 'patrick', 'paddy': 'patrick',
    'pete': 'peter',
    'phil': 'philip', 'pip': 'philip',
    'ray': 'raymond',
    'ron': 'ronald', 'ronnie': 'ronald',
    'sam': 'samuel', 'sammy': 'samuel',
    'stan': 'stanley',
    'tim': 'timothy', 'timmy': 'timothy',
    'vic': 'victor', 'vicky': 'victor',
    'walt': 'walter', 'wally': 'walter',
    'zach': 'zachary', 'zack': 'zachary',
    
    # Female names
    'liz': 'elizabeth', 'lizzy': 'elizabeth', 'beth': 'elizabeth', 'betty': 'elizabeth', 'eliza': 'elizabeth',
    'kate': 'katherine', 'kathy': 'katherine', 'katie': 'katherine', 'cathy': 'catherine',
    'sue': 'susan', 'susie': 'susan', 'suzy': 'susan',
    'jen': 'jennifer', 'jenny': 'jennifer',
    'meg': 'margaret', 'maggie': 'margaret', 'peggy': 'margaret', 'marge': 'margaret',
    'pat': 'patricia', 'patty': 'patricia', 'trish': 'patricia',
    'barb': 'barbara', 'barbie': 'barbara',
    'debbie': 'deborah', 'deb': 'deborah',
    'vicky': 'victoria', 'vic': 'victoria',
    'becky': 'rebecca', 'becca': 'rebecca',
    'mandy': 'amanda',
    'sandy': 'sandra',
    'cindy': 'cynthia',
    'nancy': 'ann', 'annie': 'ann', 'anna': 'ann',
    'alex': 'alexandra', 'lexi': 'alexandra',
    'abby': 'abigail', 'gail': 'abigail',
    'ally': 'allison', 'ali': 'allison',
    'chris': 'christine', 'tina': 'christine', 'christy': 'christine',
    'donna': 'madonna',
    'ellie': 'eleanor', 'ella': 'eleanor',
    'em': 'emily', 'emmy': 'emily',
    'gina': 'regina', 'reggie': 'regina',
    'jo': 'josephine', 'josie': 'josephine',
    'julie': 'julia', 'jules': 'julia',
    'kim': 'kimberly',
    'laurie': 'laura',
    'mel': 'melissa', 'missy': 'melissa',
    'mia': 'amelia',
    'nicky': 'nicole', 'niki': 'nicole',
    'pam': 'pamela',
    'penny': 'penelope',
    'sally': 'sarah', 'sadie': 'sarah',
    'steph': 'stephanie',
    'terri': 'theresa', 'terry': 'theresa',
    'val': 'valerie',
    'wendy': 'gwendolyn', 'gwen': 'gwendolyn',
}
```

### Soundex Algorithm
```python
def soundex(name):
    """Generate Soundex code for phonetic matching"""
    if not name:
        return ''
    
    name = name.upper()
    # Keep first letter
    first_letter = name[0]
    
    # Mapping
    mapping = {
        'B': '1', 'F': '1', 'P': '1', 'V': '1',
        'C': '2', 'G': '2', 'J': '2', 'K': '2', 'Q': '2', 'S': '2', 'X': '2', 'Z': '2',
        'D': '3', 'T': '3',
        'L': '4',
        'M': '5', 'N': '5',
        'R': '6',
    }
    
    # Convert to codes
    codes = [first_letter]
    prev_code = mapping.get(first_letter, '0')
    
    for char in name[1:]:
        code = mapping.get(char, '0')
        if code != '0' and code != prev_code:
            codes.append(code)
        prev_code = code if code != '0' else prev_code
    
    # Pad or truncate to 4 characters
    soundex_code = ''.join(codes)[:4].ljust(4, '0')
    return soundex_code
```

### Name Parsing Logic
```python
def parse_name_parts(full_name):
    """
    Parse a full name into (first_name, last_name, canonical_form) regardless of format.
    
    Handles:
    - "John Smith" → ("john", "smith", "john smith")
    - "Smith, John" → ("john", "smith", "john smith")
    - "Dr. John Smith Jr." → ("john", "smith", "john smith")
    - "Mary Jane Otte, PhD" → ("mary", "otte", "mary otte") # first word of first name
    """
    if not full_name:
        return ('', '', '')
    
    # Remove titles and suffixes
    titles = r'\b(dr|mr|mrs|ms|miss|prof|rev|hon|sir|dame)\.?\s*'
    suffixes = r',?\s*\b(jr|sr|ii|iii|iv|v|phd|md|esq|cpa|dds|dvm)\.?\s*$'
    
    name = re.sub(titles, '', full_name, flags=re.IGNORECASE)
    name = re.sub(suffixes, '', name, flags=re.IGNORECASE)
    name = ' '.join(name.split())  # Normalize whitespace
    
    if ',' in name:
        # Format: "Last, First" or "Last, First Middle"
        parts = name.split(',', 1)
        last_name = parts[0].strip().lower()
        first_parts = parts[1].strip().split() if len(parts) > 1 else []
        first_name = first_parts[0].lower() if first_parts else ''
    else:
        # Format: "First Last" or "First Middle Last"
        parts = name.split()
        if len(parts) == 0:
            return ('', '', '')
        elif len(parts) == 1:
            return (parts[0].lower(), '', parts[0].lower())
        else:
            first_name = parts[0].lower()
            last_name = parts[-1].lower()
    
    # Canonical form: always "firstname lastname"
    canonical = f"{first_name} {last_name}".strip()
    
    return (first_name, last_name, canonical)
```

---

## WHAT TO DO IF STUCK

After 15 iterations, if not complete:
1. Document what's blocking progress
2. List what was attempted
3. Identify which specific test is failing
4. Suggest alternative approaches
5. Output partial progress with clear notes on what remains

---

## IMPORTANT NOTES

- The app will be used to merge ~19,000 contacts across two files
- Performance matters: must handle 10,000+ contacts in under 5 seconds
- User is non-technical: GUI must be intuitive with clear visual feedback
- Preserve all contact data: never lose emails, phones, notes, etc. during merge
- The application is called "Phyllis's Contact Merging Magician" - keep this branding

---

## OUTPUT

When ALL success criteria are met and ALL tests pass:

```
<promise>COMPLETE</promise>
```

---

## COMMAND (for Claude Code)

```bash
/ralph-loop "$(cat RALPH_WIGGUM_PROMPT.md)" --max-iterations 30 --completion-promise "COMPLETE"
```
