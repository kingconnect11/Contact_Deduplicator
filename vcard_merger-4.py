#!/usr/bin/env python3
"""
vCard Contact Merger v4 - Enhanced Edition
Phyllis's Contact Merging Magician

Key improvements over v3:
- Fixed: "John Smith" vs "Smith, John" now correctly detected as same person
- Phonetic matching (Soundex) for similar-sounding names
- Nickname mapping (Bob→Robert, Bill→William, etc.)
- Email normalization (dots, case, plus addressing)
- Better phone normalization (international support)
- Enhanced confidence scoring with match explanations
- Bulk operations and keyboard shortcuts
- Deduplication report
- Progress tracking for large files
"""

import re
import json
import os
from collections import defaultdict
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext, simpledialog
import threading
import queue

# ============================================================================
# NICKNAME MAPPING - Common name variations
# ============================================================================

NICKNAME_MAP = {
    # Male names
    'bob': 'robert', 'rob': 'robert', 'bobby': 'robert', 'robbie': 'robert', 'bert': 'robert',
    'bill': 'william', 'will': 'william', 'billy': 'william', 'willy': 'william', 'liam': 'william',
    'mike': 'michael', 'mick': 'michael', 'mickey': 'michael', 'mikey': 'michael',
    'jim': 'james', 'jimmy': 'james', 'jamie': 'james',
    'joe': 'joseph', 'joey': 'joseph',
    'tom': 'thomas', 'tommy': 'thomas', 'thom': 'thomas',
    'dick': 'richard', 'rick': 'richard', 'ricky': 'richard', 'rich': 'richard',
    'dan': 'daniel', 'danny': 'daniel',
    'dave': 'david', 'davy': 'david',
    'steve': 'steven', 'stevie': 'steven',
    'chris': 'christopher', 'kit': 'christopher',
    'matt': 'matthew', 'matty': 'matthew',
    'tony': 'anthony', 'ant': 'anthony',
    'andy': 'andrew', 'drew': 'andrew',
    'nick': 'nicholas', 'nicky': 'nicholas',
    'ed': 'edward', 'eddie': 'edward', 'ted': 'edward', 'teddy': 'edward',
    'al': 'albert', 'bert': 'albert', 'albie': 'albert',
    'alex': 'alexander', 'xander': 'alexander', 'sandy': 'alexander',
    'ben': 'benjamin', 'benny': 'benjamin', 'benji': 'benjamin',
    'chuck': 'charles', 'charlie': 'charles', 'chas': 'charles',
    'frank': 'francis', 'frankie': 'francis', 'fran': 'francis',
    'fred': 'frederick', 'freddy': 'frederick', 'freddie': 'frederick',
    'greg': 'gregory',
    'hank': 'henry', 'harry': 'henry', 'hal': 'henry',
    'jack': 'john', 'johnny': 'john', 'jon': 'john',
    'jake': 'jacob',
    'jeff': 'jeffrey', 'geoff': 'geoffrey',
    'jerry': 'gerald', 'gerry': 'gerald',
    'josh': 'joshua',
    'larry': 'lawrence', 'lars': 'lawrence',
    'leo': 'leonard', 'lenny': 'leonard',
    'louie': 'louis', 'lou': 'louis',
    'mark': 'marcus',
    'marty': 'martin',
    'max': 'maxwell', 'maxie': 'maxwell',
    'nate': 'nathan', 'nathaniel': 'nathan',
    'pat': 'patrick', 'paddy': 'patrick',
    'pete': 'peter', 'petey': 'peter',
    'phil': 'philip', 'pip': 'philip',
    'ray': 'raymond',
    'ron': 'ronald', 'ronnie': 'ronald',
    'sam': 'samuel', 'sammy': 'samuel',
    'stan': 'stanley',
    'tim': 'timothy', 'timmy': 'timothy',
    'vic': 'victor', 'vicky': 'victor',
    'wally': 'walter', 'walt': 'walter',
    'zach': 'zachary', 'zack': 'zachary',
    
    # Female names
    'beth': 'elizabeth', 'liz': 'elizabeth', 'lizzy': 'elizabeth', 'betty': 'elizabeth',
    'libby': 'elizabeth', 'eliza': 'elizabeth', 'lisa': 'elizabeth', 'ellie': 'elizabeth',
    'kate': 'katherine', 'kathy': 'katherine', 'katie': 'katherine', 'cathy': 'catherine',
    'kat': 'katherine', 'kitty': 'katherine',
    'jenny': 'jennifer', 'jen': 'jennifer', 'jenn': 'jennifer',
    'sue': 'susan', 'susie': 'susan', 'suzy': 'susan',
    'maggie': 'margaret', 'meg': 'margaret', 'peggy': 'margaret', 'marge': 'margaret',
    'margie': 'margaret', 'madge': 'margaret', 'greta': 'margaret',
    'pam': 'pamela',
    'pat': 'patricia', 'patty': 'patricia', 'trish': 'patricia', 'tricia': 'patricia',
    'barb': 'barbara', 'barbie': 'barbara', 'babs': 'barbara',
    'deb': 'deborah', 'debbie': 'deborah', 'debby': 'deborah',
    'becky': 'rebecca', 'becca': 'rebecca',
    'vicky': 'victoria', 'vicki': 'victoria', 'tori': 'victoria',
    'chris': 'christine', 'chrissy': 'christine', 'tina': 'christine',
    'alex': 'alexandra', 'lexie': 'alexandra', 'sandy': 'alexandra',
    'mandy': 'amanda',
    'angie': 'angela', 'angel': 'angela',
    'annie': 'anne', 'ann': 'anne', 'anna': 'anne', 'nancy': 'anne',
    'bea': 'beatrice', 'trixie': 'beatrice',
    'carol': 'caroline', 'carrie': 'caroline',
    'cindy': 'cynthia',
    'connie': 'constance',
    'di': 'diana', 'diane': 'diana',
    'donna': 'madonna',
    'dot': 'dorothy', 'dottie': 'dorothy',
    'eleanor': 'ellen', 'ella': 'ellen', 'nell': 'ellen', 'nelly': 'ellen',
    'fran': 'frances', 'frannie': 'frances', 'francie': 'frances',
    'gail': 'abigail', 'abby': 'abigail',
    'ginny': 'virginia', 'ginger': 'virginia',
    'grace': 'gracie',
    'jan': 'janet', 'janice': 'janet',
    'jo': 'josephine', 'josie': 'josephine',
    'judy': 'judith', 'judi': 'judith',
    'jules': 'julia', 'julie': 'julia',
    'kay': 'katherine',
    'kim': 'kimberly', 'kimmy': 'kimberly',
    'laurie': 'laura', 'lori': 'laura',
    'linda': 'belinda', 'lindy': 'belinda',
    'lucy': 'lucille', 'lou': 'lucille',
    'lynn': 'linda',
    'maddie': 'madeline', 'maddy': 'madeline',
    'mary': 'marie', 'marie': 'mary', 'maria': 'mary', 'molly': 'mary', 'polly': 'mary',
    'mel': 'melanie', 'melinda': 'melanie',
    'mia': 'maria',
    'millie': 'millicent', 'mildred': 'millicent',
    'minnie': 'wilhelmina',
    'missy': 'melissa',
    'nan': 'nancy',
    'nat': 'natalie',
    'nicki': 'nicole', 'nikki': 'nicole',
    'penny': 'penelope',
    'polly': 'paula',
    'rose': 'rosemary', 'rosie': 'rosemary',
    'sally': 'sarah', 'sadie': 'sarah',
    'sam': 'samantha',
    'sandy': 'sandra',
    'shelly': 'michelle', 'shell': 'michelle',
    'stacy': 'anastasia',
    'terri': 'theresa', 'terry': 'theresa', 'tess': 'theresa',
    'tiff': 'tiffany',
    'val': 'valerie',
    'wendy': 'gwendolyn', 'gwen': 'gwendolyn',
}

# Create reverse mapping too (robert -> bob, etc.)
CANONICAL_NAMES = {}
for nickname, canonical in NICKNAME_MAP.items():
    CANONICAL_NAMES[nickname] = canonical
    CANONICAL_NAMES[canonical] = canonical  # canonical maps to itself


# ============================================================================
# SOUNDEX IMPLEMENTATION - Phonetic matching
# ============================================================================

def soundex(name):
    """
    Generate Soundex code for a name.
    This helps match names that sound similar but are spelled differently.
    Examples: "Smith" and "Smyth" both become "S530"
    """
    if not name:
        return ""
    
    # Convert to uppercase
    name = name.upper()
    
    # Keep first letter
    first_letter = name[0]
    
    # Map letters to numbers
    mapping = {
        'B': '1', 'F': '1', 'P': '1', 'V': '1',
        'C': '2', 'G': '2', 'J': '2', 'K': '2', 'Q': '2', 'S': '2', 'X': '2', 'Z': '2',
        'D': '3', 'T': '3',
        'L': '4',
        'M': '5', 'N': '5',
        'R': '6',
        'A': '0', 'E': '0', 'I': '0', 'O': '0', 'U': '0', 'H': '0', 'W': '0', 'Y': '0'
    }
    
    # Convert rest of name
    coded = first_letter
    prev_code = mapping.get(first_letter, '0')
    
    for char in name[1:]:
        code = mapping.get(char, '')
        if code and code != '0' and code != prev_code:
            coded += code
            prev_code = code
        elif code == '0':
            prev_code = '0'  # Vowels reset the previous code
    
    # Pad or truncate to 4 characters
    coded = (coded + '000')[:4]
    
    return coded


def metaphone_simple(name):
    """
    Simplified Metaphone algorithm for phonetic matching.
    More accurate than Soundex for English names.
    """
    if not name:
        return ""
    
    name = name.upper()
    
    # Common transformations
    replacements = [
        (r'GH', 'F'),
        (r'PH', 'F'),
        (r'CK', 'K'),
        (r'SCH', 'SK'),
        (r'SH', 'X'),
        (r'TH', '0'),
        (r'WR', 'R'),
        (r'WH', 'W'),
        (r'MB$', 'M'),
        (r'KN', 'N'),
        (r'GN', 'N'),
        (r'PN', 'N'),
        (r'X', 'KS'),
        (r'Q', 'K'),
        (r'V', 'F'),
        (r'Z', 'S'),
    ]
    
    for pattern, replacement in replacements:
        name = re.sub(pattern, replacement, name)
    
    # Remove vowels except at start
    if len(name) > 1:
        name = name[0] + re.sub(r'[AEIOU]', '', name[1:])
    
    # Remove double letters
    result = ''
    prev = ''
    for char in name:
        if char != prev:
            result += char
            prev = char
    
    return result[:6]  # Limit length


# ============================================================================
# NAME PARSING AND NORMALIZATION
# ============================================================================

def parse_name_parts(full_name):
    """
    Parse a full name into (first_name, last_name) regardless of format.
    Handles: "John Smith", "Smith, John", "Dr. John Smith Jr.", etc.
    
    Returns: (first_name, last_name, full_normalized)
    """
    if not full_name:
        return ('', '', '')
    
    # Remove common titles (with word boundaries, case insensitive)
    titles = r'\b(dr|mr|mrs|ms|miss|prof|rev)\.?\s*'
    suffixes = r',?\s*\b(jr|sr|ii|iii|iv|phd|md|esq|cpa)\.?\s*$'
    
    name = re.sub(titles, '', full_name, flags=re.IGNORECASE)
    name = re.sub(suffixes, '', name, flags=re.IGNORECASE)
    name = ' '.join(name.split())  # Clean up whitespace
    
    # Check for "Last, First" format
    if ',' in name:
        parts = name.split(',', 1)
        last_name = parts[0].strip()
        first_name = parts[1].strip() if len(parts) > 1 else ''
        # For multi-word first names, just take the first word for matching
        first_parts = first_name.split()
        first_only = first_parts[0] if first_parts else ''
    else:
        # Assume "First Last" or "First Middle Last"
        parts = name.split()
        if len(parts) == 0:
            return ('', '', '')
        elif len(parts) == 1:
            # Single name - could be first or last
            first_name = parts[0]
            first_only = parts[0]
            last_name = ''
        else:
            first_name = parts[0]
            first_only = parts[0]
            last_name = parts[-1]  # Last word is surname
    
    # Normalize to lowercase for comparison
    first_only = first_only.lower().strip()
    last_name = last_name.lower().strip()
    
    # Create normalized full name (always "first last" order)
    full_normalized = f"{first_only} {last_name}".strip()
    
    return (first_only, last_name, full_normalized)


def get_canonical_first_name(first_name):
    """Get the canonical form of a first name (expand nicknames)"""
    if not first_name:
        return ''
    return CANONICAL_NAMES.get(first_name.lower(), first_name.lower())


def names_match(name1, name2, threshold=0.75):
    """
    Comprehensive name matching that handles:
    - Different order: "John Smith" vs "Smith, John"
    - Nicknames: "Bob Smith" vs "Robert Smith"
    - Phonetic similarity: "Smith" vs "Smyth"
    - Partial matches: "John" vs "John Smith"
    
    Returns: (is_match, confidence, reasons)
    """
    if not name1 or not name2:
        return (False, 0, [])
    
    # Parse both names
    first1, last1, norm1 = parse_name_parts(name1)
    first2, last2, norm2 = parse_name_parts(name2)
    
    reasons = []
    confidence = 0
    
    # Exact normalized match
    if norm1 and norm2 and norm1 == norm2:
        return (True, 100, ["Exact name match"])
    
    # Check last name match (most important)
    last_match = False
    if last1 and last2:
        if last1 == last2:
            last_match = True
            confidence += 40
            reasons.append("Last name exact match")
        elif soundex(last1) == soundex(last2):
            last_match = True
            confidence += 30
            reasons.append(f"Last name phonetic match ({last1} ≈ {last2})")
        elif SequenceMatcher(None, last1, last2).ratio() > 0.85:
            last_match = True
            confidence += 25
            reasons.append(f"Last name similar ({last1} ~ {last2})")
    
    # Check first name match
    first_match = False
    if first1 and first2:
        # Direct match
        if first1 == first2:
            first_match = True
            confidence += 40
            reasons.append("First name exact match")
        # Nickname match
        elif get_canonical_first_name(first1) == get_canonical_first_name(first2):
            first_match = True
            confidence += 35
            reasons.append(f"Nickname match ({first1} = {first2})")
        # Phonetic match
        elif soundex(first1) == soundex(first2):
            first_match = True
            confidence += 25
            reasons.append(f"First name phonetic match ({first1} ≈ {first2})")
        # Initial match (J. vs John)
        elif len(first1) == 1 and first2.startswith(first1):
            first_match = True
            confidence += 20
            reasons.append(f"Initial match ({first1}. = {first2})")
        elif len(first2) == 1 and first1.startswith(first2):
            first_match = True
            confidence += 20
            reasons.append(f"Initial match ({first2}. = {first1})")
    
    # Both first and last must match for a name match
    is_match = first_match and last_match
    
    # Bonus for very high string similarity on full name
    full_sim = SequenceMatcher(None, norm1, norm2).ratio()
    if full_sim > 0.9:
        confidence += 10
        reasons.append(f"High overall similarity ({int(full_sim*100)}%)")
    
    # Ensure confidence doesn't exceed 100
    confidence = min(confidence, 100)
    
    return (is_match, confidence, reasons)


def normalize_name(name):
    """
    Remove duplicate consecutive words from name.
    Examples:
        'Mary Jane Jane Otte' -> 'Mary Jane Otte'
        'Bob Bob Smith' -> 'Bob Smith'
    """
    if not name:
        return name

    words = name.split()
    if len(words) <= 1:
        return name

    normalized = []
    prev_word = None

    for word in words:
        if prev_word is None or word.lower() != prev_word.lower():
            normalized.append(word)
            prev_word = word

    return ' '.join(normalized)


# ============================================================================
# EMAIL NORMALIZATION
# ============================================================================

def normalize_email(email):
    """
    Normalize email address for comparison.
    - Lowercase everything
    - Remove dots from Gmail addresses (Google ignores dots)
    - Remove plus addressing (user+tag@domain.com -> user@domain.com)
    """
    if not email or '@' not in email:
        return email.lower() if email else ''
    
    local, domain = email.lower().rsplit('@', 1)
    
    # Remove plus addressing
    if '+' in local:
        local = local.split('+')[0]
    
    # Remove dots for Gmail (Google ignores them)
    if domain in ('gmail.com', 'googlemail.com'):
        local = local.replace('.', '')
    
    return f"{local}@{domain}"


def get_email_domain(email):
    """Extract domain from email, or empty string if invalid"""
    if not email or '@' not in email:
        return ''
    return email.lower().split('@')[-1]


def is_generic_email_domain(domain):
    """Check if email domain is a generic provider (not company-specific)"""
    generic_domains = {
        'gmail.com', 'googlemail.com', 'yahoo.com', 'yahoo.co.uk',
        'hotmail.com', 'outlook.com', 'live.com', 'msn.com',
        'icloud.com', 'me.com', 'mac.com',
        'aol.com', 'protonmail.com', 'proton.me',
        'mail.com', 'email.com', 'inbox.com',
        'ymail.com', 'rocketmail.com',
        'comcast.net', 'verizon.net', 'att.net', 'sbcglobal.net',
        'earthlink.net', 'cox.net', 'charter.net',
    }
    return domain.lower() in generic_domains


# ============================================================================
# PHONE NORMALIZATION
# ============================================================================

def normalize_phone(phone):
    """
    Normalize phone number for comparison.
    Returns just the significant digits (usually last 10 for US, or all for international)
    """
    if not phone:
        return ''
    
    # Extract only digits
    digits = re.sub(r'\D', '', phone)
    
    if not digits:
        return ''
    
    # Handle US/Canada numbers
    if len(digits) == 11 and digits.startswith('1'):
        digits = digits[1:]  # Remove country code
    
    # For comparison, use last 10 digits (or all if shorter)
    if len(digits) >= 10:
        return digits[-10:]
    elif len(digits) >= 7:
        return digits[-7:]
    
    return digits


def phones_match(phone1, phone2):
    """
    Check if two phone numbers match.
    Returns: (is_match, confidence, reason)
    """
    norm1 = normalize_phone(phone1)
    norm2 = normalize_phone(phone2)
    
    if not norm1 or not norm2:
        return (False, 0, "")
    
    # Exact match on 10 digits
    if len(norm1) >= 10 and len(norm2) >= 10 and norm1 == norm2:
        return (True, 100, "Phone exact match (10 digits)")
    
    # Match on last 7 digits (area code might differ in formatting)
    if len(norm1) >= 7 and len(norm2) >= 7:
        if norm1[-7:] == norm2[-7:]:
            return (True, 90, "Phone match (last 7 digits)")
    
    return (False, 0, "")


# ============================================================================
# VCARD CONTACT CLASS
# ============================================================================

class VCardContact:
    """Contact with full field preservation and enhanced matching"""
    
    def __init__(self):
        self.fn = ''
        self.n_parts = []
        self.emails = []
        self.phones = []
        self.addresses = []
        self.notes = []
        self.org = ''
        self.title = ''
        self.birthday = ''
        self.url = ''
        self.raw_vcard = ''
        self.source_file = ''  # Track which file this came from
        
        # Cached normalized values for faster matching
        self._normalized_name = None
        self._parsed_name = None
        self._normalized_emails = None
        self._normalized_phones = None
    
    def parse_vcard(self, vcard_text):
        """Parse a vCard and extract all fields"""
        self.raw_vcard = vcard_text
        lines = vcard_text.split('\n')
        current_line = ''
        
        for line in lines:
            # Handle line continuation (folded lines)
            if line.startswith(' ') or line.startswith('\t'):
                current_line += line[1:]
                continue
            
            if current_line:
                self._process_line(current_line)
            current_line = line
        
        if current_line:
            self._process_line(current_line)
        
        # Clear caches after parsing
        self._normalized_name = None
        self._parsed_name = None
        self._normalized_emails = None
        self._normalized_phones = None
    
    def _process_line(self, line):
        if ':' not in line:
            return
        
        field, value = line.split(':', 1)
        field = field.upper()
        value = value.strip()
        
        if not value:
            return
        
        if field.startswith('FN'):
            self.fn = value
        elif field.startswith('N') and not field.startswith('NOTE'):
            self.n_parts = value.split(';')
        elif field.startswith('EMAIL'):
            if value.lower() not in [e.lower() for e in self.emails]:
                self.emails.append(value)
        elif field.startswith('TEL'):
            if value not in self.phones:
                self.phones.append(value)
        elif field.startswith('ADR'):
            if value not in self.addresses:
                self.addresses.append(value)
        elif field.startswith('NOTE'):
            if value not in self.notes:
                self.notes.append(value)
        elif field.startswith('ORG'):
            self.org = value
        elif field.startswith('TITLE'):
            self.title = value
        elif field.startswith('BDAY'):
            self.birthday = value
        elif field.startswith('URL'):
            self.url = value
    
    def get_normalized_name(self):
        """Get normalized name for matching (cached)"""
        if self._normalized_name is None:
            _, _, self._normalized_name = parse_name_parts(self.fn)
        return self._normalized_name
    
    def get_parsed_name(self):
        """Get parsed name parts (first, last, full) - cached"""
        if self._parsed_name is None:
            self._parsed_name = parse_name_parts(self.fn)
        return self._parsed_name
    
    def get_normalized_emails(self):
        """Get normalized emails for matching (cached)"""
        if self._normalized_emails is None:
            self._normalized_emails = [normalize_email(e) for e in self.emails]
        return self._normalized_emails
    
    def get_normalized_phones(self):
        """Get normalized phones for matching (cached)"""
        if self._normalized_phones is None:
            self._normalized_phones = [normalize_phone(p) for p in self.phones]
        return self._normalized_phones
    
    def get_display_name(self):
        """Get display name (normalized to remove duplicate words)"""
        if self.fn:
            return normalize_name(self.fn)
        return "Unnamed Contact"
    
    def get_summary(self):
        """Get one-line summary"""
        parts = [self.fn or "Unnamed"]
        if self.emails:
            parts.append(f"📧 {len(self.emails)}")
        if self.phones:
            parts.append(f"📞 {len(self.phones)}")
        return " | ".join(parts)
    
    def get_full_details(self):
        """Get full contact details as text"""
        details = []
        
        if self.fn:
            details.append(f"Name: {self.fn}")
        
        if self.source_file:
            details.append(f"Source: {self.source_file}")
        
        if self.emails:
            details.append(f"\nEmails ({len(self.emails)}):")
            for email in self.emails:
                details.append(f"  • {email}")
        
        if self.phones:
            details.append(f"\nPhones ({len(self.phones)}):")
            for phone in self.phones:
                details.append(f"  • {phone}")
        
        if self.addresses:
            details.append(f"\nAddresses ({len(self.addresses)}):")
            for addr in self.addresses:
                details.append(f"  • {addr}")
        
        if self.org:
            details.append(f"\nOrganization: {self.org}")
        
        if self.title:
            details.append(f"Title: {self.title}")
        
        if self.birthday:
            details.append(f"Birthday: {self.birthday}")
        
        if self.notes:
            details.append(f"\nNotes ({len(self.notes)}):")
            for note in self.notes:
                # Truncate long notes for display
                display_note = note[:200] + '...' if len(note) > 200 else note
                details.append(f"  • {display_note}")
        
        return "\n".join(details) if details else "No details available"
    
    def to_vcard(self):
        """Convert to vCard format"""
        lines = ['BEGIN:VCARD', 'VERSION:3.0']
        
        if self.fn:
            lines.append(f'FN:{self.fn}')
        
        if self.n_parts:
            lines.append(f'N:{";".join(self.n_parts)}')
        
        for email in self.emails:
            lines.append(f'EMAIL:{email}')
        
        for phone in self.phones:
            lines.append(f'TEL:{phone}')
        
        for addr in self.addresses:
            lines.append(f'ADR:{addr}')
        
        if self.org:
            lines.append(f'ORG:{self.org}')
        
        if self.title:
            lines.append(f'TITLE:{self.title}')
        
        if self.birthday:
            lines.append(f'BDAY:{self.birthday}')
        
        if self.url:
            lines.append(f'URL:{self.url}')
        
        for note in self.notes:
            lines.append(f'NOTE:{note}')
        
        lines.append('END:VCARD')
        return '\n'.join(lines)
    
    def copy(self):
        """Create a copy of this contact"""
        new_contact = VCardContact()
        new_contact.fn = self.fn
        new_contact.n_parts = self.n_parts.copy()
        new_contact.emails = self.emails.copy()
        new_contact.phones = self.phones.copy()
        new_contact.addresses = self.addresses.copy()
        new_contact.notes = self.notes.copy()
        new_contact.org = self.org
        new_contact.title = self.title
        new_contact.birthday = self.birthday
        new_contact.url = self.url
        new_contact.raw_vcard = self.raw_vcard
        new_contact.source_file = self.source_file
        return new_contact


# ============================================================================
# MERGING LOGIC
# ============================================================================

def merge_contacts(contacts):
    """Merge multiple contacts into one, preserving all data"""
    if not contacts:
        return None

    merged = VCardContact()

    # Select the best name (longest/most complete)
    best_name = ""
    for contact in contacts:
        if contact.fn:
            current_words = len(contact.fn.split())
            best_words = len(best_name.split()) if best_name else 0

            if current_words > best_words:
                best_name = contact.fn
            elif current_words == best_words and len(contact.fn) > len(best_name):
                best_name = contact.fn

    merged.fn = best_name if best_name else (contacts[0].fn if contacts else "")
    merged.n_parts = contacts[0].n_parts or []
    merged.org = contacts[0].org
    merged.title = contacts[0].title
    merged.birthday = contacts[0].birthday
    merged.url = contacts[0].url

    # Track sources
    sources = set()
    
    for contact in contacts:
        if contact.source_file:
            sources.add(contact.source_file)
        
        if not merged.org and contact.org:
            merged.org = contact.org
        if not merged.title and contact.title:
            merged.title = contact.title
        if not merged.birthday and contact.birthday:
            merged.birthday = contact.birthday
        if not merged.url and contact.url:
            merged.url = contact.url

        # Merge emails (case-insensitive dedup)
        for email in contact.emails:
            if email.lower() not in [e.lower() for e in merged.emails]:
                merged.emails.append(email)

        # Merge phones (normalize for comparison)
        for phone in contact.phones:
            norm_phone = normalize_phone(phone)
            existing_norms = [normalize_phone(p) for p in merged.phones]
            if norm_phone and norm_phone not in existing_norms:
                merged.phones.append(phone)

        # Merge addresses
        for addr in contact.addresses:
            if addr not in merged.addresses:
                merged.addresses.append(addr)

        # Merge notes
        for note in contact.notes:
            if note not in merged.notes:
                merged.notes.append(note)

    # Normalize the merged name
    if merged.fn:
        merged.fn = normalize_name(merged.fn)
    
    # Record source files
    if sources:
        merged.source_file = ', '.join(sorted(sources))

    return merged


# ============================================================================
# WARNING DETECTION
# ============================================================================

def detect_warnings(contacts):
    """
    Detect potential issues with merging these contacts.
    Returns: (has_warnings, list of warning messages)
    """
    warnings = []

    if len(contacts) < 2:
        return False, []

    # Extract data for analysis
    orgs = [c.org for c in contacts if c.org]
    emails = [e for c in contacts for e in c.emails]
    phones = [p for c in contacts for p in c.phones]
    names = [c.fn for c in contacts if c.fn]

    # Warning 1: Conflicting Organizations
    unique_orgs = set(o.lower().strip() for o in orgs)
    # Remove very similar orgs (e.g., "Acme" vs "Acme Inc")
    org_groups = []
    for org in unique_orgs:
        matched = False
        for group in org_groups:
            if SequenceMatcher(None, org, group[0]).ratio() > 0.8:
                group.append(org)
                matched = True
                break
        if not matched:
            org_groups.append([org])
    
    if len(org_groups) > 1:
        warnings.append(f"⚠️ Different organizations: {', '.join([g[0] for g in org_groups[:3]])}")

    # Warning 2: Email Domain Mismatch (different companies)
    if len(emails) > 1:
        domains = [get_email_domain(e) for e in emails]
        company_domains = set(d for d in domains if d and not is_generic_email_domain(d))
        if len(company_domains) > 1:
            warnings.append(f"⚠️ Different work email domains: {', '.join(list(company_domains)[:3])}")

    # Warning 3: Geographic Mismatch (very different area codes - US only)
    if len(phones) > 1:
        area_codes = []
        for phone in phones:
            norm = normalize_phone(phone)
            if len(norm) >= 10:
                area_codes.append(norm[:3])
        
        unique_area_codes = set(area_codes)
        if len(unique_area_codes) > 2:
            warnings.append(f"⚠️ Multiple area codes: {len(unique_area_codes)} different locations")

    # Warning 4: No email or phone match (name-only match)
    email_match = False
    norm_emails_by_contact = [set(c.get_normalized_emails()) for c in contacts]
    for i, emails1 in enumerate(norm_emails_by_contact):
        for emails2 in norm_emails_by_contact[i+1:]:
            if emails1 & emails2:
                email_match = True
                break
    
    phone_match = False
    for i, c1 in enumerate(contacts):
        for c2 in contacts[i+1:]:
            for p1 in c1.get_normalized_phones():
                for p2 in c2.get_normalized_phones():
                    if p1 and p2 and len(p1) >= 7 and len(p2) >= 7:
                        if p1[-7:] == p2[-7:]:
                            phone_match = True
                            break
    
    if not email_match and not phone_match:
        warnings.append("⚠️ Name-only match: No email or phone number overlap")

    # Warning 5: Name variation is large
    if len(names) >= 2:
        for i, name1 in enumerate(names):
            for name2 in names[i+1:]:
                _, _, norm1 = parse_name_parts(name1)
                _, _, norm2 = parse_name_parts(name2)
                similarity = SequenceMatcher(None, norm1, norm2).ratio()
                if similarity < 0.6:
                    warnings.append(f"⚠️ Names quite different: '{name1}' vs '{name2}'")
                    break

    return len(warnings) > 0, warnings


# ============================================================================
# FILE PARSING
# ============================================================================

def parse_vcard_file(filepath, source_name=None):
    """Parse vCard file using streaming approach for large files"""
    contacts = []
    current_vcard = []
    in_vcard = False
    
    if source_name is None:
        source_name = os.path.basename(filepath)

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            line = line.rstrip('\n\r')

            if line.upper().startswith('BEGIN:VCARD'):
                in_vcard = True
                current_vcard = [line]
                continue

            if line.upper().startswith('END:VCARD'):
                if in_vcard:
                    current_vcard.append(line)
                    vcard_text = '\n'.join(current_vcard)

                    contact = VCardContact()
                    contact.parse_vcard(vcard_text)
                    contact.source_file = source_name

                    if contact.fn or contact.emails or contact.phones:
                        contacts.append(contact)

                    current_vcard = []
                    in_vcard = False
                continue

            if in_vcard:
                current_vcard.append(line)

    return contacts


# ============================================================================
# CONFIDENCE CALCULATION
# ============================================================================

def calculate_match_confidence(contact1, contact2):
    """
    Calculate confidence score for a match (0-100) with detailed reasons.
    
    Factors:
    - Email match: +50 points
    - Phone match: +40 points  
    - Name match: up to +50 points (exact, nickname, phonetic)
    - Organization match: +10 points bonus
    """
    score = 0
    factors = []

    # Email match (highest confidence indicator)
    emails1 = set(contact1.get_normalized_emails())
    emails2 = set(contact2.get_normalized_emails())
    email_overlap = emails1 & emails2
    
    if email_overlap:
        score += 50
        factors.append(f"✓ Email match: {list(email_overlap)[0]}")

    # Phone match
    phones1 = contact1.get_normalized_phones()
    phones2 = contact2.get_normalized_phones()
    phone_matched = False
    
    for p1 in phones1:
        for p2 in phones2:
            match, conf, reason = phones_match(p1, p2)
            if match:
                score += 40
                factors.append(f"✓ {reason}")
                phone_matched = True
                break
        if phone_matched:
            break

    # Name match (comprehensive)
    name_match, name_conf, name_reasons = names_match(contact1.fn, contact2.fn)
    if name_match:
        # Scale name confidence to max 50 points
        name_score = int(name_conf * 0.5)
        score += name_score
        for reason in name_reasons:
            factors.append(f"✓ {reason}")
    else:
        # Partial name similarity still gets some points
        _, _, norm1 = parse_name_parts(contact1.fn)
        _, _, norm2 = parse_name_parts(contact2.fn)
        if norm1 and norm2:
            sim = SequenceMatcher(None, norm1, norm2).ratio()
            if sim > 0.5:
                partial_score = int(sim * 25)
                score += partial_score
                factors.append(f"~ Name {int(sim*100)}% similar")

    # Organization match bonus
    if contact1.org and contact2.org:
        org1 = contact1.org.lower().strip()
        org2 = contact2.org.lower().strip()
        if org1 == org2:
            score += 10
            factors.append(f"✓ Same organization: {contact1.org}")
        elif SequenceMatcher(None, org1, org2).ratio() > 0.8:
            score += 5
            factors.append(f"~ Similar organization")

    return min(score, 100), factors


# ============================================================================
# DUPLICATE GROUP FINDING
# ============================================================================

def find_similar_groups(contacts, threshold=0.75, progress_callback=None):
    """
    Group contacts by similarity with confidence scores.
    Uses hash-based bucketing for efficiency.
    
    Args:
        contacts: List of VCardContact objects
        threshold: Minimum name similarity threshold
        progress_callback: Optional function(current, total, message) for progress updates
    """
    total_contacts = len(contacts)
    
    if progress_callback:
        progress_callback(0, 100, "Building search indices...")

    # Phase 1: Create buckets for efficient matching
    email_buckets = defaultdict(list)
    phone_buckets = defaultdict(list)
    name_buckets = defaultdict(list)
    soundex_buckets = defaultdict(list)
    canonical_name_buckets = defaultdict(list)

    for i, contact in enumerate(contacts):
        # Email buckets (normalized)
        for email in contact.get_normalized_emails():
            if email:
                email_buckets[email].append(i)

        # Phone buckets (last 7-10 digits)
        for phone in contact.get_normalized_phones():
            if len(phone) >= 7:
                phone_buckets[phone[-7:]].append(i)
            if len(phone) >= 10:
                phone_buckets[phone[-10:]].append(i)

        # Name-based buckets
        first, last, norm_name = contact.get_parsed_name()
        
        if last:
            # Last name bucket
            name_buckets[last.lower()].append(i)
            
            # Soundex bucket for last name
            sx = soundex(last)
            if sx:
                soundex_buckets[sx].append(i)
        
        if first and last:
            # Canonical first + last name bucket (handles nicknames)
            canonical_first = get_canonical_first_name(first)
            canonical_key = f"{canonical_first}_{last.lower()}"
            canonical_name_buckets[canonical_key].append(i)
            
            # Also add reverse order bucket
            reverse_key = f"{last.lower()}_{canonical_first}"
            canonical_name_buckets[reverse_key].append(i)

    if progress_callback:
        progress_callback(20, 100, "Finding candidate pairs...")

    # Phase 2: Build candidate pairs from buckets
    candidate_pairs = set()

    # Add pairs from email buckets (very high confidence)
    for indices in email_buckets.values():
        if len(indices) > 1:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((min(indices[i], indices[j]), max(indices[i], indices[j])))

    # Add pairs from phone buckets (high confidence)
    for indices in phone_buckets.values():
        if len(indices) > 1:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((min(indices[i], indices[j]), max(indices[i], indices[j])))

    # Add pairs from canonical name buckets (handles nicknames + reversed names)
    for indices in canonical_name_buckets.values():
        if 1 < len(indices) <= 100:  # Limit large buckets
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((min(indices[i], indices[j]), max(indices[i], indices[j])))

    # Add pairs from soundex buckets (phonetic matching)
    for indices in soundex_buckets.values():
        if 1 < len(indices) <= 50:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((min(indices[i], indices[j]), max(indices[i], indices[j])))

    # Add pairs from exact last name buckets
    for indices in name_buckets.values():
        if 1 < len(indices) <= 100:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((min(indices[i], indices[j]), max(indices[i], indices[j])))

    if progress_callback:
        progress_callback(40, 100, f"Evaluating {len(candidate_pairs):,} candidate pairs...")

    # Phase 3: Evaluate candidate pairs
    match_graph = defaultdict(list)
    
    pairs_processed = 0
    total_pairs = len(candidate_pairs)

    for i, j in candidate_pairs:
        contact1 = contacts[i]
        contact2 = contacts[j]

        confidence, factors = calculate_match_confidence(contact1, contact2)

        # Keep if confidence is high enough
        if confidence >= 50:  # Minimum 50% confidence to consider
            match_graph[i].append((j, confidence, factors))
            match_graph[j].append((i, confidence, factors))
        
        pairs_processed += 1
        if progress_callback and pairs_processed % 1000 == 0:
            pct = 40 + int(40 * pairs_processed / total_pairs)
            progress_callback(pct, 100, f"Evaluated {pairs_processed:,} of {total_pairs:,} pairs...")

    if progress_callback:
        progress_callback(80, 100, "Building duplicate groups...")

    # Phase 4: Extract connected components as groups
    groups = []
    processed = set()

    for start_idx in range(len(contacts)):
        if start_idx in processed or start_idx not in match_graph:
            continue

        # BFS to find all connected contacts
        group_indices = []
        all_confidences = []
        all_factors = []
        queue_bfs = [start_idx]
        visited = {start_idx}

        while queue_bfs:
            current = queue_bfs.pop(0)
            group_indices.append(current)
            processed.add(current)

            for neighbor, conf, factors in match_graph.get(current, []):
                if neighbor not in visited:
                    visited.add(neighbor)
                    queue_bfs.append(neighbor)
                    all_confidences.append(conf)
                    all_factors.extend(factors)

        if len(group_indices) > 1:
            # Calculate average confidence for the group
            avg_confidence = sum(all_confidences) / len(all_confidences) if all_confidences else 50

            groups.append({
                'indices': group_indices,
                'confidence': int(avg_confidence),
                'match_factors': list(set(all_factors))  # Unique factors
            })

    # Sort by confidence (high to low)
    groups.sort(key=lambda g: g['confidence'], reverse=True)

    if progress_callback:
        progress_callback(100, 100, f"Found {len(groups):,} duplicate groups")

    return groups


# ============================================================================
# UI COMPONENTS
# ============================================================================

class PreviewMergeDialog:
    """Dialog for previewing merged contact result"""

    def __init__(self, parent, contacts, group_idx, app_ref, match_factors=None):
        self.result = None
        self.contacts = contacts
        self.group_idx = group_idx
        self.app_ref = app_ref
        self.match_factors = match_factors or []
        self.merged_contact = merge_contacts(contacts)

        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Preview Merge - Group {group_idx + 1}")
        self.dialog.geometry("1000x750")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.setup_ui()

        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')

    def setup_ui(self):
        main_frame = tk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)

        # Header
        header_frame = tk.Frame(main_frame, bg='#e3f2fd')
        header_frame.pack(fill='x', pady=(0, 15))

        tk.Label(header_frame, text=f"Preview Merge Result",
                font=('Arial', 14, 'bold'), bg='#e3f2fd').pack(pady=10)

        # Match explanation section
        if self.match_factors:
            match_frame = tk.Frame(main_frame, bg='#E8F5E9', relief='solid', borderwidth=1)
            match_frame.pack(fill='x', pady=(0, 10))
            
            tk.Label(match_frame, text="Why these contacts matched:",
                    font=('Arial', 10, 'bold'), bg='#E8F5E9').pack(anchor='w', padx=10, pady=5)
            
            for factor in self.match_factors[:5]:  # Show top 5 factors
                tk.Label(match_frame, text=f"  {factor}",
                        font=('Arial', 9), bg='#E8F5E9', fg='#2E7D32').pack(anchor='w', padx=15)
            
            if len(self.match_factors) > 5:
                tk.Label(match_frame, text=f"  ... and {len(self.match_factors) - 5} more factors",
                        font=('Arial', 9, 'italic'), bg='#E8F5E9', fg='#666').pack(anchor='w', padx=15, pady=(0, 5))

        # Content: Side by side view
        content_frame = tk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True)

        # Left: Original contacts
        left_frame = tk.LabelFrame(content_frame, text=f"Original Contacts ({len(self.contacts)})",
                                   font=('Arial', 11, 'bold'))
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        left_canvas = tk.Canvas(left_frame)
        left_scroll = tk.Scrollbar(left_frame, orient='vertical', command=left_canvas.yview)
        left_scrollable = tk.Frame(left_canvas)

        left_scrollable.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )

        left_canvas.create_window((0, 0), window=left_scrollable, anchor='nw')
        left_canvas.configure(yscrollcommand=left_scroll.set)

        for i, contact in enumerate(self.contacts):
            contact_frame = tk.Frame(left_scrollable, relief='solid', borderwidth=1, bg='#fafafa')
            contact_frame.pack(fill='x', padx=5, pady=5)

            # Show source file
            source_text = f" (from {contact.source_file})" if contact.source_file else ""
            tk.Label(contact_frame, text=f"Contact {i+1}{source_text}",
                    font=('Arial', 10, 'bold'), bg='#e3f2fd').pack(fill='x', padx=5, pady=3)

            details_text = scrolledtext.ScrolledText(contact_frame, height=10, width=35,
                                                     font=('Courier', 9), wrap='word')
            details_text.pack(fill='both', padx=5, pady=5)
            details_text.insert('1.0', contact.get_full_details())
            details_text.config(state='disabled')

        left_canvas.pack(side='left', fill='both', expand=True)
        left_scroll.pack(side='right', fill='y')

        # Right: Merged result
        right_frame = tk.LabelFrame(content_frame, text="Merged Result Preview",
                                    font=('Arial', 11, 'bold'))
        right_frame.pack(side='right', fill='both', expand=True, padx=(5, 0))

        # Check for warnings
        has_warnings, warning_list = detect_warnings(self.contacts)

        if has_warnings:
            warning_frame = tk.Frame(right_frame, bg='#FFF3CD', relief='solid', borderwidth=2)
            warning_frame.pack(fill='x', padx=10, pady=10)

            tk.Label(warning_frame, text="⚠️ Warnings Detected",
                    font=('Arial', 10, 'bold'), bg='#FFF3CD', fg='#856404').pack(pady=5)

            for warning in warning_list:
                tk.Label(warning_frame, text=warning,
                        font=('Arial', 9), bg='#FFF3CD', fg='#856404', anchor='w').pack(padx=10, anchor='w')

            tk.Label(warning_frame, text="Review carefully before merging",
                    font=('Arial', 9, 'italic'), bg='#FFF3CD', fg='#856404').pack(pady=5)

        # Merged text widget
        self.merged_text = scrolledtext.ScrolledText(right_frame, font=('Courier', 10),
                                               wrap='word', bg='#e8f5e9')
        self.merged_text.pack(fill='both', expand=True, padx=10, pady=10)
        self.update_merged_display()

        # Buttons
        btn_frame = tk.Frame(self.dialog, bg='white', relief='solid', borderwidth=1)
        btn_frame.pack(fill='x', padx=15, pady=15)

        tk.Button(btn_frame, text="✏️ Edit Merged Contact",
                 command=self.edit_merged,
                 font=('Arial', 12, 'bold'), bg='#2196F3', fg='white',
                 width=20, height=2).pack(side='left', padx=10, pady=10)

        tk.Button(btn_frame, text="Close",
                 command=self.dialog.destroy,
                 font=('Arial', 12), bg='#757575', fg='white',
                 width=15, height=2).pack(side='right', padx=10, pady=10)

    def update_merged_display(self):
        """Update the merged contact display"""
        self.merged_text.config(state='normal')
        self.merged_text.delete('1.0', 'end')
        self.merged_text.insert('1.0', self.merged_contact.get_full_details())
        self.merged_text.config(state='disabled')

    def edit_merged(self):
        """Open dialog to edit the merged contact"""
        edit_dialog = EditContactDialog(self.dialog, self.merged_contact)
        self.dialog.wait_window(edit_dialog.dialog)

        if edit_dialog.result:
            self.merged_contact = edit_dialog.result
            self.app_ref.edited_merges[self.group_idx] = self.merged_contact
            self.update_merged_display()
            messagebox.showinfo("Saved",
                              "Your edits have been saved!\n\n"
                              "This edited version will be used when you merge this group.",
                              parent=self.dialog)


class EditContactDialog:
    """Dialog for editing merged contact"""

    def __init__(self, parent, contact):
        self.result = None
        self.contact = contact.copy()
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit Merged Contact")
        self.dialog.geometry("600x750")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.setup_ui()
        
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def setup_ui(self):
        main_frame = tk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Name
        tk.Label(main_frame, text="Name:", font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        self.name_entry = tk.Entry(main_frame, font=('Arial', 11), width=50)
        self.name_entry.insert(0, self.contact.fn)
        self.name_entry.grid(row=0, column=1, pady=5, sticky='ew')
        
        # Organization
        tk.Label(main_frame, text="Organization:", font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='w', pady=5)
        self.org_entry = tk.Entry(main_frame, font=('Arial', 11), width=50)
        self.org_entry.insert(0, self.contact.org)
        self.org_entry.grid(row=1, column=1, pady=5, sticky='ew')
        
        # Title
        tk.Label(main_frame, text="Title:", font=('Arial', 11, 'bold')).grid(row=2, column=0, sticky='w', pady=5)
        self.title_entry = tk.Entry(main_frame, font=('Arial', 11), width=50)
        self.title_entry.insert(0, self.contact.title)
        self.title_entry.grid(row=2, column=1, pady=5, sticky='ew')
        
        # Emails
        tk.Label(main_frame, text="Emails:", font=('Arial', 11, 'bold')).grid(row=3, column=0, sticky='nw', pady=5)
        
        email_frame = tk.Frame(main_frame)
        email_frame.grid(row=3, column=1, pady=5, sticky='ew')
        
        self.email_listbox = tk.Listbox(email_frame, height=4, font=('Arial', 10))
        for email in self.contact.emails:
            self.email_listbox.insert('end', email)
        self.email_listbox.pack(side='left', fill='both', expand=True)
        
        email_btn_frame = tk.Frame(email_frame)
        email_btn_frame.pack(side='right', padx=5)
        tk.Button(email_btn_frame, text="Add", command=self.add_email, width=8).pack(pady=2)
        tk.Button(email_btn_frame, text="Remove", command=self.remove_email, width=8).pack(pady=2)
        
        # Phones
        tk.Label(main_frame, text="Phones:", font=('Arial', 11, 'bold')).grid(row=4, column=0, sticky='nw', pady=5)
        
        phone_frame = tk.Frame(main_frame)
        phone_frame.grid(row=4, column=1, pady=5, sticky='ew')
        
        self.phone_listbox = tk.Listbox(phone_frame, height=4, font=('Arial', 10))
        for phone in self.contact.phones:
            self.phone_listbox.insert('end', phone)
        self.phone_listbox.pack(side='left', fill='both', expand=True)
        
        phone_btn_frame = tk.Frame(phone_frame)
        phone_btn_frame.pack(side='right', padx=5)
        tk.Button(phone_btn_frame, text="Add", command=self.add_phone, width=8).pack(pady=2)
        tk.Button(phone_btn_frame, text="Remove", command=self.remove_phone, width=8).pack(pady=2)
        
        # Addresses
        tk.Label(main_frame, text="Addresses:", font=('Arial', 11, 'bold')).grid(row=5, column=0, sticky='nw', pady=5)
        
        addr_frame = tk.Frame(main_frame)
        addr_frame.grid(row=5, column=1, pady=5, sticky='ew')
        
        self.addr_listbox = tk.Listbox(addr_frame, height=3, font=('Arial', 10))
        for addr in self.contact.addresses:
            self.addr_listbox.insert('end', addr[:60] + '...' if len(addr) > 60 else addr)
        self.addr_listbox.pack(side='left', fill='both', expand=True)
        
        addr_btn_frame = tk.Frame(addr_frame)
        addr_btn_frame.pack(side='right', padx=5)
        tk.Button(addr_btn_frame, text="Remove", command=self.remove_address, width=8).pack(pady=2)
        
        # Notes
        tk.Label(main_frame, text="Notes:", font=('Arial', 11, 'bold')).grid(row=6, column=0, sticky='nw', pady=5)
        
        self.notes_text = scrolledtext.ScrolledText(main_frame, height=6, font=('Arial', 10))
        self.notes_text.grid(row=6, column=1, pady=5, sticky='ew')
        self.notes_text.insert('1.0', '\n---\n'.join(self.contact.notes))
        
        main_frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = tk.Frame(self.dialog)
        btn_frame.pack(fill='x', padx=20, pady=20)
        
        tk.Button(btn_frame, text="Cancel", command=self.cancel, 
                 font=('Arial', 11), width=15).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Save Changes", command=self.save, 
                 font=('Arial', 11, 'bold'), bg='#4CAF50', fg='white', width=15).pack(side='right', padx=5)
    
    def add_email(self):
        email = simpledialog.askstring("Add Email", "Enter email address:", parent=self.dialog)
        if email and email.strip():
            self.email_listbox.insert('end', email.strip())
    
    def add_phone(self):
        phone = simpledialog.askstring("Add Phone", "Enter phone number:", parent=self.dialog)
        if phone and phone.strip():
            self.phone_listbox.insert('end', phone.strip())
    
    def remove_email(self):
        selection = self.email_listbox.curselection()
        if selection:
            self.email_listbox.delete(selection[0])
    
    def remove_phone(self):
        selection = self.phone_listbox.curselection()
        if selection:
            self.phone_listbox.delete(selection[0])
    
    def remove_address(self):
        selection = self.addr_listbox.curselection()
        if selection:
            self.addr_listbox.delete(selection[0])
    
    def save(self):
        self.contact.fn = self.name_entry.get()
        self.contact.org = self.org_entry.get()
        self.contact.title = self.title_entry.get()
        
        self.contact.emails = list(self.email_listbox.get(0, 'end'))
        self.contact.phones = list(self.phone_listbox.get(0, 'end'))
        self.contact.addresses = list(self.addr_listbox.get(0, 'end'))
        
        notes_text = self.notes_text.get('1.0', 'end').strip()
        if notes_text:
            self.contact.notes = [n.strip() for n in notes_text.split('---') if n.strip()]
        else:
            self.contact.notes = []
        
        self.result = self.contact
        self.dialog.destroy()
    
    def cancel(self):
        self.result = None
        self.dialog.destroy()


class DeduplicationReportDialog:
    """Dialog showing deduplication statistics"""
    
    def __init__(self, parent, stats):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Deduplication Report")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        
        self.stats = stats
        self.setup_ui()
        
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def setup_ui(self):
        main_frame = tk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Header
        tk.Label(main_frame, text="📊 Deduplication Report",
                font=('Arial', 18, 'bold'), fg='#4CAF50').pack(pady=10)
        
        # Report text
        report_text = scrolledtext.ScrolledText(main_frame, font=('Courier', 11),
                                                wrap='word', height=25)
        report_text.pack(fill='both', expand=True, pady=10)
        
        report = self.generate_report()
        report_text.insert('1.0', report)
        report_text.config(state='disabled')
        
        # Buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        
        tk.Button(btn_frame, text="Copy Report",
                 command=lambda: self.copy_report(report),
                 font=('Arial', 11), bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        tk.Button(btn_frame, text="Close",
                 command=self.dialog.destroy,
                 font=('Arial', 11)).pack(side='right', padx=5)
    
    def generate_report(self):
        s = self.stats
        
        reduction = s['original'] - s['final']
        reduction_pct = (reduction / s['original'] * 100) if s['original'] > 0 else 0
        
        report = f"""
╔══════════════════════════════════════════════════════════════╗
║              CONTACT DEDUPLICATION REPORT                     ║
╚══════════════════════════════════════════════════════════════╝

INPUT FILES
═══════════════════════════════════════════════════════════════
  File 1: {s['file1_name']}
          {s['file1_count']:,} contacts
  
  File 2: {s['file2_name']}
          {s['file2_count']:,} contacts
  
  ─────────────────────────────────────────────────────────────
  TOTAL INPUT: {s['original']:,} contacts

DUPLICATE ANALYSIS
═══════════════════════════════════════════════════════════════
  Duplicate groups found: {s['groups_found']:,}
  
  By confidence level:
    • 95-100% (Very High): {s.get('groups_95_100', 0):,} groups
    • 90-94%  (High):      {s.get('groups_90_94', 0):,} groups
    • 80-89%  (Medium):    {s.get('groups_80_89', 0):,} groups
    • Below 80% (Low):     {s.get('groups_below_80', 0):,} groups

MERGE RESULTS
═══════════════════════════════════════════════════════════════
  Groups merged:  {s['groups_merged']:,}
  Groups skipped: {s['groups_skipped']:,}

OUTPUT
═══════════════════════════════════════════════════════════════
  Final contact count: {s['final']:,}
  
  ─────────────────────────────────────────────────────────────
  DUPLICATES REMOVED: {reduction:,} ({reduction_pct:.1f}% reduction)

═══════════════════════════════════════════════════════════════
Report generated by Phyllis's Contact Merging Magician v4
"""
        return report
    
    def copy_report(self, report):
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append(report)
        messagebox.showinfo("Copied", "Report copied to clipboard!", parent=self.dialog)


# ============================================================================
# MAIN APPLICATION
# ============================================================================

class MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phyllis's Contact Merging Magician v4")
        self.root.geometry("1300x850")

        self.contacts = []
        self.groups = []
        self.current_group_idx = 0
        self.merged_contacts = []
        self.skipped_groups = []
        self.history = []
        self.current_merged = None

        # File tracking
        self.file1_path = None
        self.file2_path = None
        self.file1_contacts = []
        self.file2_contacts = []

        # Pagination
        self.groups_per_page = 50
        self.current_overview_page = 0

        # Batch approval
        self.confidence_batches = []
        self.current_batch_idx = 0
        self.batch_selections = {}
        self.batch_review_page = 0
        self.groups_per_batch_page = 100
        self.edited_merges = {}
        self.merged_group_indices = set()

        self.task_queue = queue.Queue()

        self.setup_ui()
        self.show_load_screen()
        self.process_queue()

    def process_queue(self):
        """Process background thread messages"""
        try:
            while True:
                msg = self.task_queue.get_nowait()
                
                if msg['type'] == 'progress':
                    self.update_progress(msg['current'], msg['total'], msg['message'])
                elif msg['type'] == 'files_loaded':
                    self._handle_files_loaded(msg)
                elif msg['type'] == 'groups_found':
                    self._handle_groups_found(msg)
                elif msg['type'] == 'error':
                    messagebox.showerror("Error", msg['message'])
                    self.status_label.config(text="Error occurred", fg='red')
        except queue.Empty:
            pass
        
        self.root.after(100, self.process_queue)
    
    def update_progress(self, current, total, message):
        """Update progress bar and message"""
        if hasattr(self, 'progress_bar'):
            self.progress_bar['value'] = current
            self.progress_bar['maximum'] = total
        if hasattr(self, 'progress_label'):
            self.progress_label.config(text=message)
        self.root.update_idletasks()
    
    def setup_ui(self):
        """Setup main UI container"""
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill='both', expand=True)

        # Status bar at bottom
        status_frame = tk.Frame(self.root, bg='#f0f0f0', height=35)
        status_frame.pack(side='bottom', fill='x')

        # Home button
        self.home_btn = tk.Button(status_frame, text="🏠 Home",
                                  command=self.go_home,
                                  font=('Arial', 10, 'bold'),
                                  bg='#4CAF50', fg='white',
                                  relief='raised', padx=15, pady=2)
        self.home_btn.pack(side='left', padx=10, pady=3)

        self.status_label = tk.Label(status_frame, text="Ready",
                                     bg='#f0f0f0', anchor='w', padx=10)
        self.status_label.pack(side='left', fill='x', expand=True)
    
    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def go_home(self):
        """Go back to home/load screen"""
        if self.merged_contacts:
            confirm = messagebox.askyesno("Go Home",
                                         "Return to home screen?\n\n"
                                         "You have merged contacts that haven't been exported.\n"
                                         "Are you sure you want to start over?")
            if not confirm:
                return
        
        # Reset state
        self.contacts = []
        self.groups = []
        self.current_group_idx = 0
        self.merged_contacts = []
        self.skipped_groups = []
        self.history = []
        self.confidence_batches = []
        self.current_batch_idx = 0
        self.batch_selections = {}
        self.batch_review_page = 0
        self.edited_merges = {}
        self.merged_group_indices = set()
        self.file1_contacts = []
        self.file2_contacts = []

        self.show_load_screen()

    def show_load_screen(self):
        """Screen to load files"""
        self.clear_main_frame()

        # Title
        title_frame = tk.Frame(self.main_frame, bg='#4CAF50', height=120)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="✨ Phyllis's Contact Merging Magician ✨",
                font=('Arial', 26, 'bold'), bg='#4CAF50', fg='white').pack(pady=20)
        tk.Label(title_frame, text="Version 4.0 - Enhanced Edition",
                font=('Arial', 12), bg='#4CAF50', fg='white').pack()
        
        # Info section
        info_frame = tk.Frame(self.main_frame, bg='#E3F2FD')
        info_frame.pack(fill='x', padx=30, pady=20)
        
        tk.Label(info_frame, text="🔧 What's New in v4:",
                font=('Arial', 12, 'bold'), bg='#E3F2FD').pack(anchor='w', padx=15, pady=10)
        
        features = [
            "✓ Fixed: 'John Smith' vs 'Smith, John' now correctly detected as same person",
            "✓ Nickname matching: 'Bob Smith' matches 'Robert Smith'",
            "✓ Phonetic matching: 'Smith' matches 'Smyth'",
            "✓ Smarter email comparison (dots, plus addressing)",
            "✓ Match explanations show WHY contacts matched",
            "✓ Deduplication report with statistics",
        ]
        for feature in features:
            tk.Label(info_frame, text=f"  {feature}",
                    font=('Arial', 10), bg='#E3F2FD', fg='#1565C0').pack(anchor='w', padx=20)
        
        # File selection
        file_frame = tk.Frame(self.main_frame)
        file_frame.pack(pady=30, padx=50, fill='x')
        
        tk.Label(file_frame, text="First vCard file:", 
                font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='w', pady=8)
        
        self.file1_label = tk.Label(file_frame, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w', padx=10, width=60)
        self.file1_label.grid(row=0, column=1, sticky='ew', padx=10)
        
        tk.Button(file_frame, text="Browse...", 
                 command=self.select_file1, font=('Arial', 10)).grid(row=0, column=2)
        
        tk.Label(file_frame, text="Second vCard file:", 
                font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='w', pady=8)
        
        self.file2_label = tk.Label(file_frame, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w', padx=10, width=60)
        self.file2_label.grid(row=1, column=1, sticky='ew', padx=10)
        
        tk.Button(file_frame, text="Browse...", 
                 command=self.select_file2, font=('Arial', 10)).grid(row=1, column=2)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Threshold
        threshold_frame = tk.Frame(self.main_frame)
        threshold_frame.pack(pady=15, padx=50, fill='x')
        
        tk.Label(threshold_frame, text="Matching Sensitivity:", 
                font=('Arial', 11, 'bold')).pack(anchor='w')
        
        slider_frame = tk.Frame(threshold_frame)
        slider_frame.pack(fill='x', pady=5)
        
        self.threshold_var = tk.DoubleVar(value=0.75)
        self.threshold_slider = tk.Scale(slider_frame, from_=0.5, to=0.95, 
                                        resolution=0.05, orient='horizontal',
                                        variable=self.threshold_var,
                                        command=self.update_threshold,
                                        length=400)
        self.threshold_slider.pack(side='left', fill='x', expand=True)
        
        self.threshold_label = tk.Label(slider_frame, text="75% (Recommended)", 
                                        width=20, font=('Arial', 10, 'bold'))
        self.threshold_label.pack(side='right')
        
        tk.Label(threshold_frame, text="Lower = more matches found (more aggressive) | Higher = fewer matches (more conservative)", 
                font=('Arial', 9), fg='gray').pack(anchor='w')
        
        # Progress area (hidden initially)
        self.progress_frame = tk.Frame(self.main_frame)
        self.progress_frame.pack(pady=15, padx=50, fill='x')
        
        self.progress_label = tk.Label(self.progress_frame, text="", font=('Arial', 10))
        self.progress_label.pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=400, mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        
        # Load button
        self.load_btn = tk.Button(self.main_frame, text="🚀 Load & Find Duplicates",
                                  command=self.load_and_group,
                                  font=('Arial', 16, 'bold'),
                                  bg='#4CAF50', fg='white', height=2,
                                  disabledforeground='white',
                                  state='disabled')
        self.load_btn.pack(pady=20, padx=50, fill='x')
    
    def update_threshold(self, val):
        val = float(val)
        if val >= 0.85:
            desc = "(Conservative)"
        elif val >= 0.70:
            desc = "(Recommended)"
        else:
            desc = "(Aggressive)"
        self.threshold_label.config(text=f"{int(val*100)}% {desc}")
    
    def select_file1(self):
        path = filedialog.askopenfilename(
            title="Select first vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file1_path = path
            self.file1_label.config(text=os.path.basename(path))
            self.check_files_ready()
    
    def select_file2(self):
        path = filedialog.askopenfilename(
            title="Select second vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file2_path = path
            self.file2_label.config(text=os.path.basename(path))
            self.check_files_ready()
    
    def check_files_ready(self):
        if self.file1_path and self.file2_path:
            self.load_btn.config(state='normal')
    
    def load_and_group(self):
        """Load files and find groups in background thread"""
        self.load_btn.config(state='disabled', text="Processing...")
        self.status_label.config(text="Loading files...", fg='blue')
        
        thread = threading.Thread(target=self._load_and_group_thread)
        thread.daemon = True
        thread.start()
    
    def _load_and_group_thread(self):
        try:
            def progress_callback(current, total, message):
                self.task_queue.put({
                    'type': 'progress',
                    'current': current,
                    'total': total,
                    'message': message
                })
            
            progress_callback(0, 100, "Loading first file...")
            contacts1 = parse_vcard_file(self.file1_path, os.path.basename(self.file1_path))
            
            progress_callback(10, 100, "Loading second file...")
            contacts2 = parse_vcard_file(self.file2_path, os.path.basename(self.file2_path))
            
            all_contacts = contacts1 + contacts2
            
            self.task_queue.put({
                'type': 'files_loaded',
                'contacts': all_contacts,
                'contacts1': contacts1,
                'contacts2': contacts2,
                'file1_name': os.path.basename(self.file1_path),
                'file2_name': os.path.basename(self.file2_path)
            })
            
            progress_callback(20, 100, "Finding duplicates...")
            threshold = self.threshold_var.get()
            groups = find_similar_groups(all_contacts, threshold, progress_callback)
            
            self.task_queue.put({
                'type': 'groups_found',
                'groups': groups
            })
            
        except Exception as e:
            import traceback
            self.task_queue.put({
                'type': 'error',
                'message': f"Error: {str(e)}\n\n{traceback.format_exc()}"
            })
    
    def _handle_files_loaded(self, msg):
        self.contacts = msg['contacts']
        self.file1_contacts = msg['contacts1']
        self.file2_contacts = msg['contacts2']
        self.file1_name = msg['file1_name']
        self.file2_name = msg['file2_name']
        
        self.status_label.config(
            text=f"Loaded {len(msg['contacts1']):,} + {len(msg['contacts2']):,} = {len(self.contacts):,} contacts",
            fg='green'
        )
    
    def _handle_groups_found(self, msg):
        self.groups = msg['groups']

        if not self.groups:
            messagebox.showinfo("No Duplicates",
                              f"No duplicate groups found!\n\n"
                              f"This could mean:\n"
                              f"• Your contacts are already clean\n"
                              f"• The threshold is too high (try lowering it)\n\n"
                              f"Current threshold: {int(self.threshold_var.get()*100)}%")
            self.load_btn.config(state='normal', text="🚀 Load & Find Duplicates")
            return

        self.current_group_idx = 0
        self.merged_contacts = []
        self.skipped_groups = []
        self.history = []
        self.current_overview_page = 0

        self.organize_confidence_batches()
        self.show_batch_overview()

    def organize_confidence_batches(self):
        """Organize groups into confidence ranges"""
        ranges = [
            (95, 100, "95-100% Very High Confidence"),
            (90, 94, "90-94% High Confidence"),
            (80, 89, "80-89% Medium Confidence"),
            (70, 79, "70-79% Fair Confidence"),
            (50, 69, "50-69% Low Confidence"),
        ]

        self.confidence_batches = []

        for min_conf, max_conf, name in ranges:
            batch_groups = []
            for i, group_data in enumerate(self.groups):
                confidence = group_data['confidence']
                if min_conf <= confidence <= max_conf:
                    batch_groups.append(i)

            if batch_groups:
                self.confidence_batches.append({
                    'name': name,
                    'range': (min_conf, max_conf),
                    'group_indices': batch_groups,
                    'count': len(batch_groups)
                })

        self.current_batch_idx = 0

    def show_batch_overview(self):
        """Show overview of all confidence batches"""
        self.clear_main_frame()

        # Header
        header = tk.Frame(self.main_frame, bg='#4CAF50', height=100)
        header.pack(fill='x')
        header.pack_propagate(False)

        title_frame = tk.Frame(header, bg='#4CAF50')
        title_frame.pack(expand=True)

        tk.Label(title_frame, text="📊 Duplicate Groups Found",
                font=('Arial', 22, 'bold'), bg='#4CAF50', fg='white').pack()

        tk.Label(title_frame, text=f"Found {len(self.groups):,} duplicate groups in {len(self.contacts):,} contacts",
                font=('Arial', 12), bg='#4CAF50', fg='white').pack(pady=5)

        # File summary
        summary_frame = tk.Frame(self.main_frame, bg='#E3F2FD')
        summary_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(summary_frame, text=f"📁 {self.file1_name}: {len(self.file1_contacts):,} contacts",
                font=('Arial', 10), bg='#E3F2FD').pack(side='left', padx=20, pady=10)
        tk.Label(summary_frame, text=f"📁 {self.file2_name}: {len(self.file2_contacts):,} contacts",
                font=('Arial', 10), bg='#E3F2FD').pack(side='left', padx=20, pady=10)

        # Instructions
        info_frame = tk.Frame(self.main_frame, bg='#FFF8E1')
        info_frame.pack(fill='x', padx=20, pady=10)

        tk.Label(info_frame, text="💡 Tip: Start with highest confidence batches - they're most likely true duplicates",
                font=('Arial', 10), bg='#FFF8E1', fg='#F57C00').pack(pady=10)

        # Batches list
        list_frame = tk.Frame(self.main_frame)
        list_frame.pack(fill='both', expand=True, padx=40, pady=20)

        for i, batch in enumerate(self.confidence_batches):
            batch_frame = tk.Frame(list_frame, relief='solid', borderwidth=2, bg='#fafafa')
            batch_frame.pack(fill='x', pady=8)

            min_conf, max_conf = batch['range']
            if min_conf >= 95:
                color = '#4CAF50'
                icon = '✓✓✓'
            elif min_conf >= 90:
                color = '#8BC34A'
                icon = '✓✓'
            elif min_conf >= 80:
                color = '#FFC107'
                icon = '✓'
            elif min_conf >= 70:
                color = '#FF9800'
                icon = '~'
            else:
                color = '#F44336'
                icon = '⚠'

            header_frame = tk.Frame(batch_frame, bg=color, height=60)
            header_frame.pack(fill='x')
            header_frame.pack_propagate(False)

            left_frame = tk.Frame(header_frame, bg=color)
            left_frame.pack(side='left', padx=20, pady=10)

            tk.Label(left_frame, text=f"{icon} {batch['name']}",
                    font=('Arial', 14, 'bold'), bg=color, fg='white').pack(anchor='w')
            tk.Label(left_frame, text=f"{batch['count']:,} groups",
                    font=('Arial', 10), bg=color, fg='white').pack(anchor='w')

            tk.Button(header_frame, text=f"Review Batch →",
                     command=lambda idx=i: self.review_batch(idx),
                     font=('Arial', 11, 'bold'), bg='white', fg=color,
                     width=20, height=2).pack(side='right', padx=20, pady=10)

        # Bottom buttons
        button_frame = tk.Frame(self.main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', padx=20, pady=20)

        if self.confidence_batches:
            tk.Button(button_frame, text="▶ Start with Highest Confidence",
                     command=lambda: self.review_batch(0),
                     font=('Arial', 13, 'bold'), bg='#4CAF50', fg='white',
                     height=2).pack(side='left', fill='x', expand=True, padx=5)

    def review_batch(self, batch_idx, reset_page=True):
        """Show review screen for a specific confidence batch"""
        self.current_batch_idx = batch_idx

        if reset_page:
            self.batch_review_page = 0

        self.clear_main_frame()

        if batch_idx >= len(self.confidence_batches):
            self.show_batch_complete_screen()
            return

        batch = self.confidence_batches[batch_idx]
        all_group_indices = batch['group_indices']

        total_pages = (len(all_group_indices) + self.groups_per_batch_page - 1) // self.groups_per_batch_page
        start_idx = self.batch_review_page * self.groups_per_batch_page
        end_idx = min(start_idx + self.groups_per_batch_page, len(all_group_indices))
        group_indices = all_group_indices[start_idx:end_idx]

        min_conf, max_conf = batch['range']
        if min_conf >= 95:
            color = '#4CAF50'
        elif min_conf >= 90:
            color = '#8BC34A'
        elif min_conf >= 80:
            color = '#FFC107'
        elif min_conf >= 70:
            color = '#FF9800'
        else:
            color = '#F44336'

        # Header
        header = tk.Frame(self.main_frame, bg=color, height=80)
        header.pack(fill='x')
        header.pack_propagate(False)

        title_frame = tk.Frame(header, bg=color)
        title_frame.pack(side='left', padx=20, pady=15)

        tk.Label(title_frame, text=f"Batch {batch_idx + 1}/{len(self.confidence_batches)}: {batch['name']}",
                font=('Arial', 16, 'bold'), bg=color, fg='white').pack(anchor='w')

        page_info = f"Page {self.batch_review_page + 1}/{total_pages}" if total_pages > 1 else ""
        tk.Label(title_frame, text=f"Groups {start_idx + 1}-{end_idx} of {batch['count']} {page_info}",
                font=('Arial', 11), bg=color, fg='white').pack(anchor='w')

        progress_frame = tk.Frame(header, bg=color)
        progress_frame.pack(side='right', padx=20)
        tk.Label(progress_frame, text=f"✓ {len(self.merged_contacts)} merged | ⊘ {len(self.skipped_groups)} skipped",
                font=('Arial', 11), bg=color, fg='white').pack()

        # Toolbar
        toolbar = tk.Frame(self.main_frame, bg='#f5f5f5', height=50)
        toolbar.pack(fill='x')
        toolbar.pack_propagate(False)

        toolbar_left = tk.Frame(toolbar, bg='#f5f5f5')
        toolbar_left.pack(side='left', padx=15, pady=10)

        tk.Button(toolbar_left, text="← Back to Overview",
                 command=self.show_batch_overview,
                 font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=5)

        if total_pages > 1:
            pagination_frame = tk.Frame(toolbar, bg='#f5f5f5')
            pagination_frame.pack(side='left', padx=15, pady=10)

            if self.batch_review_page > 0:
                tk.Button(pagination_frame, text="← Prev",
                         command=self.prev_batch_page,
                         font=('Arial', 10)).pack(side='left', padx=3)

            tk.Label(pagination_frame, text=f"Page {self.batch_review_page + 1}/{total_pages}",
                    font=('Arial', 10, 'bold'), bg='#f5f5f5').pack(side='left', padx=10)

            if self.batch_review_page < total_pages - 1:
                tk.Button(pagination_frame, text="Next →",
                         command=self.next_batch_page,
                         font=('Arial', 10)).pack(side='left', padx=3)

        toolbar_right = tk.Frame(toolbar, bg='#f5f5f5')
        toolbar_right.pack(side='right', padx=15, pady=10)

        tk.Button(toolbar_right, text="Select All",
                 command=self.select_all_in_batch,
                 font=('Arial', 9)).pack(side='left', padx=3)
        tk.Button(toolbar_right, text="Deselect All",
                 command=self.deselect_all_in_batch,
                 font=('Arial', 9)).pack(side='left', padx=3)

        # Scrollable groups list
        list_frame = tk.Frame(self.main_frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)

        canvas = tk.Canvas(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        self.batch_selections = {}
        warning_count = 0

        for group_idx in group_indices:
            group_data = self.groups[group_idx]
            confidence = group_data['confidence']
            match_factors = group_data.get('match_factors', [])
            group_contacts = [self.contacts[idx] for idx in group_data['indices']]

            has_warnings, warning_list = detect_warnings(group_contacts)
            if has_warnings:
                warning_count += 1

            bg_color = '#FFF9E6' if has_warnings else '#fafafa'
            group_frame = tk.Frame(scrollable, relief='solid', borderwidth=1, bg=bg_color)
            group_frame.pack(fill='x', padx=5, pady=8)

            header_bg = '#FFF3CD' if has_warnings else '#e3f2fd'
            header_frame = tk.Frame(group_frame, bg=header_bg)
            header_frame.pack(fill='x', padx=8, pady=8)
            header_frame.grid_columnconfigure(0, weight=1)
            header_frame.grid_columnconfigure(1, weight=0)

            var = tk.BooleanVar(value=True)
            self.batch_selections[group_idx] = var

            warning_text = " ⚠️" if has_warnings else ""
            edited_text = " ✏️" if group_idx in self.edited_merges else ""

            checkbox = tk.Checkbutton(header_frame, 
                                     text=f"Group {group_idx + 1} - {confidence}% - {len(group_contacts)} contacts{warning_text}{edited_text}",
                                     variable=var, font=('Arial', 11, 'bold'),
                                     bg=header_bg, activebackground=header_bg, anchor='w')
            checkbox.grid(row=0, column=0, sticky='w', padx=(0, 10))

            preview_btn = tk.Button(header_frame, text="👁 Preview",
                                   command=lambda idx=group_idx, contacts=group_contacts, factors=match_factors: 
                                       self.preview_merge(idx, contacts, factors),
                                   font=('Arial', 10, 'bold'), bg='#2196F3', fg='white',
                                   relief='raised', padx=10, pady=3)
            preview_btn.grid(row=0, column=1, sticky='e', padx=5)

            # Show match reasons
            if match_factors:
                match_frame = tk.Frame(group_frame, bg=bg_color)
                match_frame.pack(fill='x', padx=15, pady=(0, 5))
                
                reasons_text = " | ".join(match_factors[:3])
                if len(match_factors) > 3:
                    reasons_text += f" (+{len(match_factors)-3} more)"
                
                tk.Label(match_frame, text=f"Match: {reasons_text}",
                        font=('Arial', 8), bg=bg_color, fg='#2E7D32', anchor='w').pack(fill='x')

            # Show warnings
            if has_warnings:
                warning_frame = tk.Frame(group_frame, bg=bg_color)
                warning_frame.pack(fill='x', padx=15, pady=(0, 5))
                for warning in warning_list[:2]:
                    tk.Label(warning_frame, text=warning,
                            font=('Arial', 8), bg=bg_color, fg='#856404', anchor='w').pack(fill='x')

            # Contact details
            contacts_frame = tk.Frame(group_frame, bg=bg_color)
            contacts_frame.pack(fill='x', padx=15, pady=8)

            for i, contact in enumerate(group_contacts):
                name = contact.get_display_name()
                source = f" [{contact.source_file}]" if contact.source_file else ""
                emails = ', '.join(contact.emails[:2]) if contact.emails else 'No email'
                phones = ', '.join(contact.phones[:1]) if contact.phones else 'No phone'

                contact_text = f"  {i+1}. {name}{source} | {emails} | {phones}"
                tk.Label(contacts_frame, text=contact_text,
                        font=('Arial', 9), bg=bg_color, anchor='w').pack(fill='x')

                if contact.notes:
                    notes_preview = contact.notes[0][:80] + "..." if len(contact.notes[0]) > 80 else contact.notes[0]
                    tk.Label(contacts_frame, text=f"     📝 {notes_preview}",
                            font=('Arial', 8), bg=bg_color, fg='#0066cc', anchor='w').pack(fill='x')

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Action buttons
        button_frame = tk.Frame(self.main_frame, bg='#E3F2FD', relief='solid', borderwidth=2)
        button_frame.pack(fill='x', padx=20, pady=20)

        tk.Button(button_frame, text="⊘ Skip Batch",
                 command=self.skip_current_batch,
                 font=('Arial', 12, 'bold'), bg='#FF9800', fg='white',
                 width=18, height=2).pack(side='left', padx=10, pady=10)

        tk.Button(button_frame, text="⚡ Merge Entire Batch",
                 command=self.merge_entire_batch,
                 font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white',
                 width=22, height=2).pack(side='left', padx=10, pady=10)

        tk.Button(button_frame, text=f"✓ Merge Selected ({len(self.batch_selections)})",
                 command=self.merge_selected_in_batch,
                 font=('Arial', 12, 'bold'), bg='#1976D2', fg='white',
                 width=22, height=2).pack(side='right', padx=10, pady=10)

        # Warning count
        if warning_count > 0:
            tk.Label(button_frame, text=f"⚠️ {warning_count} groups have warnings",
                    font=('Arial', 10), bg='#E3F2FD', fg='#856404').pack(side='right', padx=10)

    def select_all_in_batch(self):
        for var in self.batch_selections.values():
            var.set(True)

    def deselect_all_in_batch(self):
        for var in self.batch_selections.values():
            var.set(False)

    def preview_merge(self, group_idx, contacts, match_factors=None):
        dialog = PreviewMergeDialog(self.root, contacts, group_idx, self, match_factors)
        self.root.wait_window(dialog.dialog)

    def prev_batch_page(self):
        if self.batch_review_page > 0:
            self.batch_review_page -= 1
            self.review_batch(self.current_batch_idx, reset_page=False)

    def next_batch_page(self):
        batch = self.confidence_batches[self.current_batch_idx]
        total_pages = (len(batch['group_indices']) + self.groups_per_batch_page - 1) // self.groups_per_batch_page
        if self.batch_review_page < total_pages - 1:
            self.batch_review_page += 1
            self.review_batch(self.current_batch_idx, reset_page=False)

    def merge_entire_batch(self):
        batch = self.confidence_batches[self.current_batch_idx]

        confirm = messagebox.askyesno("Merge Entire Batch",
                                     f"Merge ALL {batch['count']} groups in this batch?\n\n"
                                     f"Batch: {batch['name']}\n\n"
                                     f"Any edits you've made will be preserved.")
        if not confirm:
            return

        merged_count = 0
        for group_idx in batch['group_indices']:
            if group_idx in self.edited_merges:
                merged_contact = self.edited_merges[group_idx]
            else:
                group_data = self.groups[group_idx]
                group_contacts = [self.contacts[idx] for idx in group_data['indices']]
                merged_contact = merge_contacts(group_contacts)

            if merged_contact:
                self.merged_contacts.append(merged_contact)
                self.merged_group_indices.add(group_idx)
                merged_count += 1

        messagebox.showinfo("Batch Merged",
                          f"Successfully merged {merged_count} groups!")

        self.current_batch_idx += 1
        if self.current_batch_idx < len(self.confidence_batches):
            self.review_batch(self.current_batch_idx)
        else:
            self.show_batch_complete_screen()

    def skip_current_batch(self):
        batch = self.confidence_batches[self.current_batch_idx]

        for group_idx in batch['group_indices']:
            group_indices = self.groups[group_idx]['indices']
            self.skipped_groups.append(group_indices)

        self.current_batch_idx += 1
        if self.current_batch_idx < len(self.confidence_batches):
            self.review_batch(self.current_batch_idx)
        else:
            self.show_batch_complete_screen()

    def merge_selected_in_batch(self):
        batch = self.confidence_batches[self.current_batch_idx]
        total_pages = (len(batch['group_indices']) + self.groups_per_batch_page - 1) // self.groups_per_batch_page

        selected_groups = [group_idx for group_idx, var in self.batch_selections.items() if var.get()]

        if not selected_groups:
            messagebox.showwarning("No Selection", "Please select at least one group to merge.")
            return

        merged_count = 0
        for group_idx in selected_groups:
            if group_idx in self.edited_merges:
                merged_contact = self.edited_merges[group_idx]
            else:
                group_data = self.groups[group_idx]
                group_contacts = [self.contacts[idx] for idx in group_data['indices']]
                merged_contact = merge_contacts(group_contacts)

            if merged_contact:
                self.merged_contacts.append(merged_contact)
                self.merged_group_indices.add(group_idx)
                merged_count += 1

        # Add unselected to skipped
        current_page_indices = batch['group_indices'][
            self.batch_review_page * self.groups_per_batch_page:
            min((self.batch_review_page + 1) * self.groups_per_batch_page, len(batch['group_indices']))
        ]
        for group_idx in current_page_indices:
            if group_idx not in selected_groups:
                self.skipped_groups.append(self.groups[group_idx]['indices'])

        if self.batch_review_page < total_pages - 1:
            messagebox.showinfo("Page Complete", f"Merged {merged_count} groups. Moving to next page...")
            self.batch_review_page += 1
            self.review_batch(self.current_batch_idx, reset_page=False)
        else:
            messagebox.showinfo("Batch Complete", f"Merged {merged_count} groups. Moving to next batch...")
            self.current_batch_idx += 1
            if self.current_batch_idx < len(self.confidence_batches):
                self.review_batch(self.current_batch_idx)
            else:
                self.show_batch_complete_screen()

    def show_batch_complete_screen(self):
        """Show completion screen"""
        self.clear_main_frame()

        tk.Label(self.main_frame, text="✅ Review Complete!",
                font=('Arial', 24, 'bold'), fg='#4CAF50').pack(pady=30)

        # Calculate statistics
        processed_indices = set()
        for group_idx in self.merged_group_indices:
            if group_idx < len(self.groups):
                processed_indices.update(self.groups[group_idx]['indices'])
        
        for skipped in self.skipped_groups:
            processed_indices.update(skipped)

        unprocessed_count = len(self.contacts) - len(processed_indices)
        final_count = len(self.merged_contacts) + unprocessed_count

        # Summary frame
        summary_frame = tk.Frame(self.main_frame, bg='#e8f5e9', relief='solid', borderwidth=2)
        summary_frame.pack(padx=100, pady=20, fill='x')

        tk.Label(summary_frame, text="📊 Summary",
                font=('Arial', 14, 'bold'), bg='#e8f5e9').pack(pady=10)

        stats = [
            f"Original contacts: {len(self.contacts):,}",
            f"  • {self.file1_name}: {len(self.file1_contacts):,}",
            f"  • {self.file2_name}: {len(self.file2_contacts):,}",
            f"",
            f"Groups merged: {len(self.merged_contacts):,}",
            f"Groups skipped: {len(self.skipped_groups):,}",
            f"",
            f"Final contact count: {final_count:,}",
            f"Duplicates removed: {len(self.contacts) - final_count:,}",
        ]
        
        for stat in stats:
            tk.Label(summary_frame, text=stat,
                    font=('Arial', 11), bg='#e8f5e9').pack()

        # Calculate confidence breakdown for report
        groups_95_100 = sum(1 for g in self.groups if g['confidence'] >= 95)
        groups_90_94 = sum(1 for g in self.groups if 90 <= g['confidence'] < 95)
        groups_80_89 = sum(1 for g in self.groups if 80 <= g['confidence'] < 90)
        groups_below_80 = sum(1 for g in self.groups if g['confidence'] < 80)

        # Store stats for report
        self.final_stats = {
            'original': len(self.contacts),
            'file1_name': self.file1_name,
            'file1_count': len(self.file1_contacts),
            'file2_name': self.file2_name,
            'file2_count': len(self.file2_contacts),
            'groups_found': len(self.groups),
            'groups_merged': len(self.merged_contacts),
            'groups_skipped': len(self.skipped_groups),
            'final': final_count,
            'groups_95_100': groups_95_100,
            'groups_90_94': groups_90_94,
            'groups_80_89': groups_80_89,
            'groups_below_80': groups_below_80,
        }

        # Buttons
        tk.Button(self.main_frame, text="💾 Export Merged Contacts",
                 command=self.export_merged,
                 font=('Arial', 14, 'bold'), bg='#4CAF50', fg='white',
                 height=2).pack(pady=15, padx=100, fill='x')

        tk.Button(self.main_frame, text="📊 View Deduplication Report",
                 command=self.show_report,
                 font=('Arial', 12), bg='#2196F3', fg='white',
                 height=2).pack(padx=100, fill='x', pady=5)

        tk.Button(self.main_frame, text="← Back to Batch Overview",
                 command=self.show_batch_overview,
                 font=('Arial', 12), bg='#757575', fg='white',
                 height=2).pack(padx=100, fill='x', pady=5)

    def show_report(self):
        """Show deduplication report"""
        if hasattr(self, 'final_stats'):
            DeduplicationReportDialog(self.root, self.final_stats)

    def export_merged(self):
        """Export the merged contacts"""
        output_path = filedialog.asksaveasfilename(
            title="Save merged contacts",
            defaultextension=".vcf",
            filetypes=[("vCard files", "*.vcf"), ("All files", "*.*")],
            initialfile="merged_contacts.vcf"
        )

        if not output_path:
            return

        try:
            # Collect all processed indices
            processed_indices = set()
            for group_idx in self.merged_group_indices:
                if group_idx < len(self.groups):
                    processed_indices.update(self.groups[group_idx]['indices'])

            for skipped in self.skipped_groups:
                processed_indices.update(skipped)

            # Build final contacts list
            final_contacts = self.merged_contacts.copy()

            for i, contact in enumerate(self.contacts):
                if i not in processed_indices:
                    final_contacts.append(contact)

            # Write to file
            with open(output_path, 'w', encoding='utf-8') as f:
                for contact in final_contacts:
                    f.write(contact.to_vcard() + '\n\n')

            messagebox.showinfo("Success",
                              f"Successfully exported {len(final_contacts):,} contacts to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Error exporting:\n{str(e)}")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    root = tk.Tk()
    app = MergerApp(root)
    root.mainloop()
