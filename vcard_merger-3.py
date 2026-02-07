#!/usr/bin/env python3
"""
vCard Contact Merger - Enhanced with Manual Controls
Full visibility and control over each merge
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

class VCardContact:
    """Contact with full field preservation"""
    
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
    
    def parse_vcard(self, vcard_text):
        """Parse a vCard and extract all fields"""
        self.raw_vcard = vcard_text
        lines = vcard_text.split('\n')
        current_line = ''
        
        for line in lines:
            if line.startswith(' ') or line.startswith('\t'):
                current_line += line[1:]
                continue
            
            if current_line:
                self._process_line(current_line)
            current_line = line
        
        if current_line:
            self._process_line(current_line)
    
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
        elif field.startswith('N'):
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
        """Get normalized name for matching"""
        return re.sub(r'[^\w\s]', '', self.fn.lower()).strip()
    
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
                details.append(f"  • {note[:100]}...")
        
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
        return new_contact


def normalize_name(name):
    """
    Remove duplicate consecutive words from name.
    Examples:
        'Mary Jane Jane Otte' -> 'Mary Jane Otte'
        'Bob Bob Smith' -> 'Bob Smith'
        'John Doe' -> 'John Doe' (no change)
    """
    if not name:
        return name

    words = name.split()
    if len(words) <= 1:
        return name

    normalized = []
    prev_word = None

    for word in words:
        # Compare case-insensitively but preserve original capitalization
        if prev_word is None or word.lower() != prev_word.lower():
            normalized.append(word)
            prev_word = word

    return ' '.join(normalized)


def merge_contacts(contacts):
    """Merge multiple contacts into one"""
    if not contacts:
        return None

    merged = VCardContact()

    # Select the longest/most complete name from all contacts
    best_name = ""
    for contact in contacts:
        if contact.fn:
            # Prefer name with more words, or if same word count, more characters
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

    for contact in contacts:
        if not merged.org and contact.org:
            merged.org = contact.org
        if not merged.title and contact.title:
            merged.title = contact.title

        for email in contact.emails:
            if email.lower() not in [e.lower() for e in merged.emails]:
                merged.emails.append(email)

        for phone in contact.phones:
            norm_phone = re.sub(r'\D', '', phone)
            existing_norms = [re.sub(r'\D', '', p) for p in merged.phones]
            if norm_phone not in existing_norms:
                merged.phones.append(phone)

        for addr in contact.addresses:
            if addr not in merged.addresses:
                merged.addresses.append(addr)

        for note in contact.notes:
            if note not in merged.notes:
                merged.notes.append(note)

    # Normalize the merged name to fix duplicate words
    if merged.fn:
        merged.fn = normalize_name(merged.fn)

    return merged


def detect_warnings(contacts):
    """Detect potential issues with merging these contacts
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
    unique_orgs = set(orgs)
    if len(unique_orgs) > 1:
        warnings.append(f"⚠️ Different organizations: {', '.join(unique_orgs)}")

    # Warning 2: Email Domain Mismatch (different companies)
    if len(emails) > 1:
        domains = [e.split('@')[1].lower() if '@' in e else '' for e in emails]
        unique_domains = set(d for d in domains if d and not d.endswith(('.gmail.com', '.yahoo.com', '.hotmail.com', '.outlook.com', '.icloud.com')))
        if len(unique_domains) > 1:
            warnings.append(f"⚠️ Different email domains: {', '.join(unique_domains)}")

    # Warning 3: Geographic Mismatch (different area codes)
    if len(phones) > 1:
        area_codes = []
        for phone in phones:
            digits = re.sub(r'\D', '', phone)
            if len(digits) >= 10:  # US/Canada format
                area_code = digits[:3]
                area_codes.append(area_code)

        unique_area_codes = set(area_codes)
        if len(unique_area_codes) > 2:  # More than 2 different area codes is suspicious
            warnings.append(f"⚠️ Multiple area codes: {len(unique_area_codes)} different locations")

    # Warning 4: Partial Match (only 1 field matches - risky)
    email_match = any(e1.lower() == e2.lower() for i, c1 in enumerate(contacts) for e1 in c1.emails
                     for c2 in contacts[i+1:] for e2 in c2.emails)

    phone_match = False
    for i, c1 in enumerate(contacts):
        for p1 in c1.phones:
            for c2 in contacts[i+1:]:
                for p2 in c2.phones:
                    norm1 = re.sub(r'\D', '', p1)
                    norm2 = re.sub(r'\D', '', p2)
                    if len(norm1) >= 7 and len(norm2) >= 7 and norm1[-7:] == norm2[-7:]:
                        phone_match = True
                        break

    if not email_match and not phone_match:
        warnings.append("⚠️ Partial match: Only names are similar (no email/phone match)")

    # Warning 5: Name variation too large
    if len(names) >= 2:
        from difflib import SequenceMatcher
        for i, name1 in enumerate(names):
            for name2 in names[i+1:]:
                similarity = SequenceMatcher(None, name1.lower(), name2.lower()).ratio()
                if similarity < 0.7:  # Less than 70% similar
                    warnings.append(f"⚠️ Names quite different: '{name1}' vs '{name2}'")
                    break

    return len(warnings) > 0, warnings


def parse_vcard_file(filepath):
    """Parse vCard file using streaming approach for large files"""
    contacts = []
    current_vcard = []
    in_vcard = False

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            line = line.rstrip('\n\r')

            # Check for BEGIN:VCARD
            if line.upper().startswith('BEGIN:VCARD'):
                in_vcard = True
                current_vcard = [line]
                continue

            # Check for END:VCARD
            if line.upper().startswith('END:VCARD'):
                if in_vcard:
                    current_vcard.append(line)
                    vcard_text = '\n'.join(current_vcard)

                    # Parse the complete vCard
                    contact = VCardContact()
                    contact.parse_vcard(vcard_text)

                    # Only keep contacts with at least name, email, or phone
                    if contact.fn or contact.emails or contact.phones:
                        contacts.append(contact)

                    # Reset for next vCard
                    current_vcard = []
                    in_vcard = False
                continue

            # Accumulate lines for current vCard
            if in_vcard:
                current_vcard.append(line)

    return contacts


def calculate_match_confidence(contact1, contact2):
    """Calculate confidence score for a match (0-100)"""
    score = 0
    factors = []

    # Email match = very high confidence
    email_match = any(e1.lower() == e2.lower()
                     for e1 in contact1.emails
                     for e2 in contact2.emails)
    if email_match:
        score += 50
        factors.append("Email match")

    # Phone match = high confidence
    phone_match = False
    for p1 in contact1.phones:
        for p2 in contact2.phones:
            norm1 = re.sub(r'\D', '', p1)
            norm2 = re.sub(r'\D', '', p2)
            if len(norm1) >= 7 and len(norm2) >= 7:
                if norm1[-7:] == norm2[-7:]:
                    phone_match = True
                    break
    if phone_match:
        score += 40
        factors.append("Phone match")

    # Name similarity with reversed name detection
    name1 = contact1.get_normalized_name()
    name2 = contact2.get_normalized_name()

    if name1 and name2:
        # Check for exact match
        if name1 == name2:
            name_score = 50
            factors.append("Name exact match")
        else:
            # Check for reversed name order (lastname, firstname vs firstname lastname)
            # Split by comma or space
            parts1 = [p.strip() for p in re.split(r'[,\s]+', name1) if p.strip()]
            parts2 = [p.strip() for p in re.split(r'[,\s]+', name2) if p.strip()]

            # Check if reversed order matches
            reversed_match = False
            if len(parts1) >= 2 and len(parts2) >= 2:
                # Check if parts1[0] parts1[1] == parts2[1] parts2[0]
                if (parts1[0] == parts2[1] and parts1[1] == parts2[0]) or \
                   (parts1[0] == parts2[-1] and parts1[-1] == parts2[0]):
                    reversed_match = True
                    name_score = 50  # Perfect match, just reversed
                    factors.append("Name match (reversed order)")

            if not reversed_match:
                # Calculate similarity normally
                name_sim = SequenceMatcher(None, name1, name2).ratio()
                name_score = int(name_sim * 50)  # Max 50 points
                if name_score > 0:
                    factors.append(f"Name {int(name_sim*100)}% similar")

        score += name_score

    return min(score, 100), factors


def find_similar_groups(contacts, threshold=0.75):
    """Group contacts by similarity with confidence scores - Optimized with hash-based bucketing"""

    # Phase 1: Create buckets for efficient matching
    # Instead of O(n²), we group contacts by potential match keys
    email_buckets = defaultdict(list)
    phone_buckets = defaultdict(list)
    name_buckets = defaultdict(list)

    for i, contact in enumerate(contacts):
        # Email buckets (exact match)
        for email in contact.emails:
            email_buckets[email.lower()].append(i)

        # Phone buckets (last 7 digits)
        for phone in contact.phones:
            norm_phone = re.sub(r'\D', '', phone)
            if len(norm_phone) >= 7:
                phone_buckets[norm_phone[-7:]].append(i)

        # Name buckets (first 3 chars + length)
        name = contact.get_normalized_name()
        if name:
            if len(name) >= 3:
                name_key = f"{name[:3]}_{len(name)}"
                name_buckets[name_key].append(i)

    # Phase 2: Build candidate pairs from buckets
    candidate_pairs = set()

    # Add pairs from email buckets
    for indices in email_buckets.values():
        if len(indices) > 1:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((indices[i], indices[j]))

    # Add pairs from phone buckets
    for indices in phone_buckets.values():
        if len(indices) > 1:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((indices[i], indices[j]))

    # Add pairs from name buckets (only if bucket not too large)
    for indices in name_buckets.values():
        if 1 < len(indices) <= 50:  # Limit bucket size to avoid explosion
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    candidate_pairs.add((indices[i], indices[j]))

    # Phase 3: Evaluate candidate pairs and group them
    match_graph = defaultdict(list)  # i -> [(j, confidence, factors), ...]

    for i, j in candidate_pairs:
        contact1 = contacts[i]
        contact2 = contacts[j]

        name1 = contact1.get_normalized_name()
        name2 = contact2.get_normalized_name()

        if not name1 or not name2:
            continue

        confidence, factors = calculate_match_confidence(contact1, contact2)
        similarity = SequenceMatcher(None, name1, name2).ratio()

        # Keep if either passes threshold or has high confidence
        if similarity >= threshold or confidence >= 70:
            match_graph[i].append((j, confidence, factors))
            match_graph[j].append((i, confidence, factors))

    # Phase 4: Extract connected components as groups
    groups = []
    processed = set()

    for start_idx in range(len(contacts)):
        if start_idx in processed or start_idx not in match_graph:
            continue

        # BFS to find all connected contacts
        group_indices = []
        confidences = {}
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
                    confidences[neighbor] = (conf, factors)

        if len(group_indices) > 1:
            # Calculate overall group confidence
            if confidences:
                avg_confidence = sum(c[0] for c in confidences.values()) / len(confidences)
            else:
                avg_confidence = 50

            groups.append({
                'indices': group_indices,
                'confidence': int(avg_confidence),
                'match_details': confidences
            })

    # Sort by confidence (high to low)
    groups.sort(key=lambda g: g['confidence'], reverse=True)

    return groups


class PreviewMergeDialog:
    """Dialog for previewing merged contact result"""

    def __init__(self, parent, contacts, group_idx, app_ref):
        self.result = None
        self.contacts = contacts
        self.group_idx = group_idx
        self.app_ref = app_ref  # Reference to main app for storing edited merges
        self.merged_contact = merge_contacts(contacts)

        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Preview Merge - Group {group_idx + 1}")
        self.dialog.geometry("900x700")
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

            tk.Label(contact_frame, text=f"Contact {i+1}",
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

        # Store merged text widget for updates
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
            # Update the merged contact with edited version
            self.merged_contact = edit_dialog.result

            # Store in app's edited merges dictionary
            self.app_ref.edited_merges[self.group_idx] = self.merged_contact

            # Update the display
            self.update_merged_display()

            # Show confirmation
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
        self.dialog.geometry("600x700")
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


class MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phyllis's Contact Merging Magician")
        self.root.geometry("1200x800")

        self.contacts = []
        self.groups = []
        self.current_group_idx = 0
        self.merged_contacts = []
        self.skipped_groups = []
        self.history = []  # For undo
        self.current_merged = None

        # Pagination for groups overview
        self.groups_per_page = 50
        self.current_overview_page = 0

        # Batch approval by confidence
        self.confidence_batches = []  # List of batches: {range, groups, name}
        self.current_batch_idx = 0
        self.batch_selections = {}  # group_idx -> BooleanVar
        self.batch_review_page = 0  # Current page in batch review
        self.groups_per_batch_page = 100  # Groups per page in batch review
        self.edited_merges = {}  # group_idx -> edited merged contact
        self.merged_group_indices = set()  # Track which groups have been merged in batch mode

        self.task_queue = queue.Queue()

        self.setup_ui()
        self.show_load_screen()
        self.process_queue()

        # Try to load autosave
        self.load_autosave()
    
    def process_queue(self):
        """Process background thread messages"""
        try:
            while True:
                msg = self.task_queue.get_nowait()
                
                if msg['type'] == 'files_loaded':
                    self._handle_files_loaded(msg)
                elif msg['type'] == 'groups_found':
                    self._handle_groups_found(msg)
                elif msg['type'] == 'error':
                    messagebox.showerror("Error", msg['message'])
                    self.status_label.config(text="Error occurred", fg='red')
        except queue.Empty:
            pass
        
        self.root.after(100, self.process_queue)
    
    def setup_ui(self):
        """Setup main UI container"""
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill='both', expand=True)

        # Status bar at bottom
        status_frame = tk.Frame(self.root, bg='#f0f0f0', height=35)
        status_frame.pack(side='bottom', fill='x')

        # Home button on left side of status bar
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
        confirm = messagebox.askyesno("Go Home",
                                     "Return to home screen?\n\n"
                                     "Any unsaved progress will be lost.\n"
                                     "Make sure you've exported your merged contacts if needed.")
        if confirm:
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

            # Show load screen
            self.show_load_screen()

    def show_load_screen(self):
        """Screen to load files"""
        self.clear_main_frame()

        tk.Label(self.main_frame, text="Phyllis's Contact Merging Magician",
                font=('Arial', 26, 'bold'), fg='#4CAF50').pack(pady=15)

        tk.Label(self.main_frame, text="✨ Transform your messy contacts into organized perfection ✨",
                font=('Arial', 12), fg='#666').pack(pady=5)
        
        # File selection
        file_frame = tk.Frame(self.main_frame)
        file_frame.pack(pady=30, padx=50, fill='x')
        
        tk.Label(file_frame, text="First vCard file:", 
                font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        
        self.file1_label = tk.Label(file_frame, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w', padx=10)
        self.file1_label.grid(row=0, column=1, sticky='ew', padx=10)
        
        tk.Button(file_frame, text="Browse...", 
                 command=self.select_file1).grid(row=0, column=2)
        
        tk.Label(file_frame, text="Second vCard file:", 
                font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='w', pady=5)
        
        self.file2_label = tk.Label(file_frame, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w', padx=10)
        self.file2_label.grid(row=1, column=1, sticky='ew', padx=10)
        
        tk.Button(file_frame, text="Browse...", 
                 command=self.select_file2).grid(row=1, column=2)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Threshold
        threshold_frame = tk.Frame(self.main_frame)
        threshold_frame.pack(pady=20, padx=50, fill='x')
        
        tk.Label(threshold_frame, text="Name Similarity Threshold:", 
                font=('Arial', 11, 'bold')).pack(anchor='w')
        
        slider_frame = tk.Frame(threshold_frame)
        slider_frame.pack(fill='x', pady=5)
        
        self.threshold_var = tk.DoubleVar(value=0.75)
        self.threshold_slider = tk.Scale(slider_frame, from_=0.5, to=0.95, 
                                        resolution=0.05, orient='horizontal',
                                        variable=self.threshold_var,
                                        command=self.update_threshold)
        self.threshold_slider.pack(side='left', fill='x', expand=True)
        
        self.threshold_label = tk.Label(slider_frame, text="75%", width=5)
        self.threshold_label.pack(side='right')
        
        tk.Label(threshold_frame, text="Lower = more groups found (more potential duplicates)", 
                font=('Arial', 9), fg='gray').pack(anchor='w')
        
        # Load button
        self.load_btn = tk.Button(self.main_frame, text="🚀 Load & Find Groups",
                                  command=self.load_and_group,
                                  font=('Arial', 16, 'bold'),
                                  bg='#4CAF50', fg='white', height=2,
                                  disabledforeground='white',
                                  state='disabled')
        self.load_btn.pack(pady=30, padx=50, fill='x')
    
    def update_threshold(self, val):
        self.threshold_label.config(text=f"{int(float(val)*100)}%")
    
    def select_file1(self):
        path = filedialog.askopenfilename(
            title="Select first vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file1_path = path
            self.file1_label.config(text=path.split('/')[-1])
            self.check_files_ready()
    
    def select_file2(self):
        path = filedialog.askopenfilename(
            title="Select second vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file2_path = path
            self.file2_label.config(text=path.split('/')[-1])
            self.check_files_ready()
    
    def check_files_ready(self):
        if hasattr(self, 'file1_path') and hasattr(self, 'file2_path'):
            self.load_btn.config(state='normal')
    
    def load_and_group(self):
        """Load files and find groups in background thread"""
        self.load_btn.config(state='disabled', text="Loading...")
        self.status_label.config(text="Loading files...", fg='blue')
        
        thread = threading.Thread(target=self._load_and_group_thread)
        thread.daemon = True
        thread.start()
    
    def _load_and_group_thread(self):
        try:
            contacts1 = parse_vcard_file(self.file1_path)
            contacts2 = parse_vcard_file(self.file2_path)
            all_contacts = contacts1 + contacts2
            
            threshold = self.threshold_var.get()
            groups = find_similar_groups(all_contacts, threshold)
            
            self.task_queue.put({
                'type': 'files_loaded',
                'contacts': all_contacts,
                'count1': len(contacts1),
                'count2': len(contacts2)
            })
            
            self.task_queue.put({
                'type': 'groups_found',
                'groups': groups
            })
            
        except Exception as e:
            self.task_queue.put({
                'type': 'error',
                'message': f"Error loading files: {str(e)}"
            })
    
    def _handle_files_loaded(self, msg):
        self.contacts = msg['contacts']
        self.status_label.config(
            text=f"Loaded {msg['count1']:,} + {msg['count2']:,} = {len(self.contacts):,} contacts",
            fg='green'
        )
    
    def _handle_groups_found(self, msg):
        self.groups = msg['groups']

        if not self.groups:
            messagebox.showinfo("No Duplicates",
                              "No duplicate groups found with current threshold. Try lowering the threshold.")
            self.load_btn.config(state='normal', text="Load & Find Groups")
            return

        self.current_group_idx = 0
        self.merged_contacts = []
        self.skipped_groups = []
        self.history = []
        self.current_overview_page = 0

        # Organize groups into confidence batches
        self.organize_confidence_batches()

        # Show batch overview for batch approval workflow
        self.show_batch_overview()

    def organize_confidence_batches(self):
        """Organize groups into confidence ranges"""
        # Define confidence ranges
        ranges = [
            (95, 100, "95-100% (Very High)"),
            (90, 94, "90-94% (High)"),
            (85, 89, "85-89% (Good)"),
            (80, 84, "80-84% (Medium-High)"),
            (70, 79, "70-79% (Medium)"),
            (60, 69, "60-69% (Medium-Low)"),
            (0, 59, "Below 60% (Low)")
        ]

        self.confidence_batches = []

        for min_conf, max_conf, name in ranges:
            # Find all groups in this range
            batch_groups = []
            for i, group_data in enumerate(self.groups):
                confidence = group_data['confidence']
                if min_conf <= confidence <= max_conf:
                    batch_groups.append(i)  # Store group index

            if batch_groups:
                self.confidence_batches.append({
                    'name': name,
                    'range': (min_conf, max_conf),
                    'group_indices': batch_groups,
                    'count': len(batch_groups)
                })

        # Set current batch to first one
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

        tk.Label(title_frame, text="Batch Approval by Confidence Score",
                font=('Arial', 22, 'bold'), bg='#4CAF50', fg='white').pack()

        tk.Label(title_frame, text=f"Found {len(self.groups)} duplicate groups organized into {len(self.confidence_batches)} confidence batches",
                font=('Arial', 12), bg='#4CAF50', fg='white').pack(pady=5)

        # Instructions
        info_frame = tk.Frame(self.main_frame, bg='#e8f5e9', height=80)
        info_frame.pack(fill='x', padx=20, pady=15)
        info_frame.pack_propagate(False)

        instructions = tk.Frame(info_frame, bg='#e8f5e9')
        instructions.pack(expand=True)

        tk.Label(instructions, text="Review and approve groups in batches based on confidence level.",
                font=('Arial', 11, 'bold'), bg='#e8f5e9').pack()
        tk.Label(instructions, text="Higher confidence batches are more likely to be true duplicates. Review each batch and select which groups to merge.",
                font=('Arial', 10), bg='#e8f5e9', fg='#555').pack()

        # Batches list
        list_frame = tk.Frame(self.main_frame)
        list_frame.pack(fill='both', expand=True, padx=40, pady=20)

        for i, batch in enumerate(self.confidence_batches):
            batch_frame = tk.Frame(list_frame, relief='solid', borderwidth=2, bg='#fafafa')
            batch_frame.pack(fill='x', pady=8)

            # Color based on confidence range
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

            # Header
            header_frame = tk.Frame(batch_frame, bg=color, height=60)
            header_frame.pack(fill='x')
            header_frame.pack_propagate(False)

            left_frame = tk.Frame(header_frame, bg=color)
            left_frame.pack(side='left', padx=20, pady=10)

            tk.Label(left_frame, text=f"{icon} {batch['name']}",
                    font=('Arial', 14, 'bold'), bg=color, fg='white').pack(anchor='w')
            tk.Label(left_frame, text=f"{batch['count']} groups in this batch",
                    font=('Arial', 10), bg=color, fg='white').pack(anchor='w')

            # Review button
            tk.Button(header_frame, text=f"Review & Approve Batch →",
                     command=lambda idx=i: self.review_batch(idx),
                     font=('Arial', 11, 'bold'), bg='white', fg=color,
                     width=25, height=2).pack(side='right', padx=20, pady=10)

        # Bottom buttons
        button_frame = tk.Frame(self.main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', padx=20, pady=20)

        tk.Button(button_frame, text="Start with Highest Confidence →",
                 command=lambda: self.review_batch(0),
                 font=('Arial', 13, 'bold'), bg='#4CAF50', fg='white',
                 height=2).pack(side='left', fill='x', expand=True, padx=5)

        tk.Button(button_frame, text="Switch to Manual Review Mode",
                 command=self.show_groups_overview,
                 font=('Arial', 11), bg='#9C27B0', fg='white',
                 height=2).pack(side='right', padx=5)

    def review_batch(self, batch_idx, reset_page=True):
        """Show review screen for a specific confidence batch"""
        self.current_batch_idx = batch_idx

        # Reset page when entering new batch
        if reset_page:
            self.batch_review_page = 0

        self.clear_main_frame()

        if batch_idx >= len(self.confidence_batches):
            # No more batches - show completion
            self.show_batch_complete_screen()
            return

        batch = self.confidence_batches[batch_idx]
        all_group_indices = batch['group_indices']

        # Calculate pagination
        total_pages = (len(all_group_indices) + self.groups_per_batch_page - 1) // self.groups_per_batch_page
        start_idx = self.batch_review_page * self.groups_per_batch_page
        end_idx = min(start_idx + self.groups_per_batch_page, len(all_group_indices))
        group_indices = all_group_indices[start_idx:end_idx]

        # Color based on confidence
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

        tk.Label(title_frame, text=f"Batch {batch_idx + 1} of {len(self.confidence_batches)}: {batch['name']}",
                font=('Arial', 16, 'bold'), bg=color, fg='white').pack(anchor='w')

        # Count warnings - will be updated in the loop below
        self.current_batch_warning_count = 0

        # Show pagination info
        page_info = f"Page {self.batch_review_page + 1} of {total_pages}" if total_pages > 1 else ""
        groups_info = f"Showing groups {start_idx + 1}-{end_idx} of {batch['count']} total" if total_pages > 1 else f"Review {batch['count']} groups"

        tk.Label(title_frame, text=f"{groups_info} {page_info}",
                font=('Arial', 11), bg=color, fg='white').pack(anchor='w')

        progress_frame = tk.Frame(header, bg=color)
        progress_frame.pack(side='right', padx=20)

        progress_text = f"Progress: {len(self.merged_contacts)} merged | {len(self.skipped_groups)} skipped"
        tk.Label(progress_frame, text=progress_text,
                font=('Arial', 11), bg=color, fg='white').pack()

        # Warning count label (will be updated after counting)
        self.warning_count_label = tk.Label(progress_frame, text="",
                                           font=('Arial', 10), bg=color, fg='white')
        self.warning_count_label.pack()

        # Toolbar
        toolbar = tk.Frame(self.main_frame, bg='#f5f5f5', height=50)
        toolbar.pack(fill='x')
        toolbar.pack_propagate(False)

        toolbar_left = tk.Frame(toolbar, bg='#f5f5f5')
        toolbar_left.pack(side='left', padx=15, pady=10)

        tk.Button(toolbar_left, text="← Back to Batch Overview",
                 command=self.show_batch_overview,
                 font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=5)

        # Pagination controls (if multiple pages)
        if total_pages > 1:
            pagination_frame = tk.Frame(toolbar, bg='#f5f5f5')
            pagination_frame.pack(side='left', padx=15, pady=10)

            if self.batch_review_page > 0:
                tk.Button(pagination_frame, text="← Previous Page",
                         command=self.prev_batch_page,
                         font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=3)

            tk.Label(pagination_frame, text=f"Page {self.batch_review_page + 1}/{total_pages}",
                    font=('Arial', 10, 'bold'), bg='#f5f5f5').pack(side='left', padx=10)

            if self.batch_review_page < total_pages - 1:
                tk.Button(pagination_frame, text="Next Page →",
                         command=self.next_batch_page,
                         font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=3)

        toolbar_right = tk.Frame(toolbar, bg='#f5f5f5')
        toolbar_right.pack(side='right', padx=15, pady=10)

        tk.Button(toolbar_right, text="Select All (This Page)",
                 command=self.select_all_in_batch,
                 font=('Arial', 9)).pack(side='left', padx=3)

        tk.Button(toolbar_right, text="Deselect All (This Page)",
                 command=self.deselect_all_in_batch,
                 font=('Arial', 9)).pack(side='left', padx=3)

        # Scrollable groups list
        list_frame = tk.Frame(self.main_frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)

        canvas = tk.Canvas(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas)

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        # Clear previous selections
        self.batch_selections = {}
        warning_count = 0

        # Display each group with checkbox
        for group_idx in group_indices:
            group_data = self.groups[group_idx]
            confidence = group_data['confidence']
            group_contacts = [self.contacts[idx] for idx in group_data['indices']]

            # Check for warnings
            has_warnings, warning_list = detect_warnings(group_contacts)
            if has_warnings:
                warning_count += 1

            # Group frame - highlight if warnings
            bg_color = '#FFF9E6' if has_warnings else '#fafafa'
            group_frame = tk.Frame(scrollable, relief='solid', borderwidth=1, bg=bg_color)
            group_frame.pack(fill='x', padx=5, pady=8)

            # Checkbox header - use grid layout to ensure button is always visible
            header_bg = '#FFF3CD' if has_warnings else '#e3f2fd'
            header_frame = tk.Frame(group_frame, bg=header_bg)
            header_frame.pack(fill='x', padx=8, pady=8)
            header_frame.grid_columnconfigure(0, weight=1)  # Left side expands
            header_frame.grid_columnconfigure(1, weight=0)  # Right side fixed

            var = tk.BooleanVar(value=True)  # Checked by default
            self.batch_selections[group_idx] = var

            # Warning icon and edited indicator
            warning_text = f" ⚠️" if has_warnings else ""
            edited_text = f" ✏️" if group_idx in self.edited_merges else ""

            # Checkbox on left (can expand)
            checkbox = tk.Checkbutton(header_frame, text=f"Group {group_idx + 1} - {confidence}% confidence - {len(group_contacts)} contacts{warning_text}{edited_text}",
                                     variable=var, font=('Arial', 11, 'bold'),
                                     bg=header_bg, activebackground=header_bg, anchor='w')
            checkbox.grid(row=0, column=0, sticky='w', padx=(0, 10))

            # Preview button on right (always visible, doesn't get pushed off)
            preview_btn = tk.Button(header_frame, text="👁 Preview",
                                   command=lambda idx=group_idx, contacts=group_contacts: self.preview_merge(idx, contacts),
                                   font=('Arial', 10, 'bold'), bg='#2196F3', fg='white',
                                   relief='raised', padx=10, pady=3)
            preview_btn.grid(row=0, column=1, sticky='e', padx=5)

            # Show warnings if present
            if has_warnings:
                warning_info = tk.Frame(group_frame, bg=bg_color)
                warning_info.pack(fill='x', padx=15, pady=(0, 5))

                for warning in warning_list[:2]:  # Show first 2 warnings
                    tk.Label(warning_info, text=warning,
                            font=('Arial', 8), bg=bg_color, fg='#856404', anchor='w').pack(fill='x')

                if len(warning_list) > 2:
                    tk.Label(warning_info, text=f"... and {len(warning_list) - 2} more warnings (click Preview to see all)",
                            font=('Arial', 8, 'italic'), bg=bg_color, fg='#856404', anchor='w').pack(fill='x')

            # Contact names preview
            contacts_frame = tk.Frame(group_frame, bg=bg_color)
            contacts_frame.pack(fill='x', padx=15, pady=8)

            for i, contact in enumerate(group_contacts):  # Show all contacts
                name = contact.get_display_name()
                emails = ', '.join(contact.emails[:2]) if contact.emails else 'No email'
                phones = ', '.join(contact.phones[:1]) if contact.phones else 'No phone'

                contact_text = f"  {i+1}. {name}"
                if emails != 'No email':
                    contact_text += f" | {emails}"
                if phones != 'No phone':
                    contact_text += f" | {phones}"

                tk.Label(contacts_frame, text=contact_text,
                        font=('Arial', 9), bg=bg_color, anchor='w').pack(fill='x')

                # Display notes if present
                if contact.notes:
                    notes_text = "     📝 Notes: " + "; ".join(contact.notes)
                    tk.Label(contacts_frame, text=notes_text,
                            font=('Arial', 8), bg=bg_color, fg='#0066cc', anchor='w').pack(fill='x')

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Update warning count label
        if warning_count > 0:
            self.warning_count_label.config(text=f"⚠️ {warning_count} groups have warnings")
        else:
            self.warning_count_label.config(text="✓ No warnings detected")

        # Action buttons
        button_frame = tk.Frame(self.main_frame, bg='#E3F2FD', relief='solid', borderwidth=2)
        button_frame.pack(fill='x', padx=20, pady=20)

        tk.Button(button_frame, text="⏭ Skip This Batch",
                 command=self.skip_current_batch,
                 font=('Arial', 14, 'bold'), bg='#1976D2', fg='white',
                 width=22, height=2).pack(side='left', padx=10, pady=10)

        tk.Button(button_frame, text="⚡ Merge Entire Batch (No Review)",
                 command=self.merge_entire_batch,
                 font=('Arial', 14, 'bold'), bg='#4CAF50', fg='white',
                 width=30, height=2).pack(side='left', padx=10, pady=10)

        selected_count = len([v for v in self.batch_selections.values() if v.get()])
        tk.Button(button_frame, text=f"✓ Merge Selected ({len(self.batch_selections)})",
                 command=self.merge_selected_in_batch,
                 font=('Arial', 14, 'bold'), bg='#0D47A1', fg='white',
                 width=25, height=2).pack(side='right', padx=10, pady=10)

    def select_all_in_batch(self):
        """Select all groups in current batch"""
        for var in self.batch_selections.values():
            var.set(True)

    def deselect_all_in_batch(self):
        """Deselect all groups in current batch"""
        for var in self.batch_selections.values():
            var.set(False)

    def preview_merge(self, group_idx, contacts):
        """Show preview dialog for merging contacts"""
        dialog = PreviewMergeDialog(self.root, contacts, group_idx, self)
        self.root.wait_window(dialog.dialog)

    def prev_batch_page(self):
        """Go to previous page in batch review"""
        if self.batch_review_page > 0:
            self.batch_review_page -= 1
            self.review_batch(self.current_batch_idx, reset_page=False)

    def next_batch_page(self):
        """Go to next page in batch review"""
        batch = self.confidence_batches[self.current_batch_idx]
        total_pages = (len(batch['group_indices']) + self.groups_per_batch_page - 1) // self.groups_per_batch_page
        if self.batch_review_page < total_pages - 1:
            self.batch_review_page += 1
            self.review_batch(self.current_batch_idx, reset_page=False)

    def merge_entire_batch(self):
        """Merge all groups in the entire batch without reviewing each page"""
        batch = self.confidence_batches[self.current_batch_idx]

        confirm = messagebox.askyesno("Merge Entire Batch",
                                     f"Merge ALL {batch['count']} groups in this batch?\n\n"
                                     f"Batch: {batch['name']}\n"
                                     f"Total groups: {batch['count']}\n\n"
                                     f"This will skip page-by-page review and merge everything.\n"
                                     f"Any edits you've made will be preserved.\n\n"
                                     f"Continue?")
        if not confirm:
            return

        # Merge all groups in batch
        merged_count = 0
        for group_idx in batch['group_indices']:
            # Check if this group has an edited merge
            if group_idx in self.edited_merges:
                merged_contact = self.edited_merges[group_idx]
            else:
                # Auto-merge the contacts
                group_data = self.groups[group_idx]
                group_contacts = [self.contacts[idx] for idx in group_data['indices']]
                merged_contact = merge_contacts(group_contacts)

            if merged_contact:
                self.merged_contacts.append(merged_contact)
                self.merged_group_indices.add(group_idx)  # Track merged group
                merged_count += 1

        # Show confirmation
        messagebox.showinfo("Batch Merged",
                          f"Successfully merged all {merged_count} groups in this batch!\n\n"
                          f"Moving to next batch...")

        # Auto-progress to next batch
        self.current_batch_idx += 1
        if self.current_batch_idx < len(self.confidence_batches):
            self.review_batch(self.current_batch_idx)
        else:
            self.show_batch_complete_screen()

    def skip_current_batch(self):
        """Skip the current batch and move to next"""
        batch = self.confidence_batches[self.current_batch_idx]

        # Add all groups in this batch to skipped
        for group_idx in batch['group_indices']:
            group_indices = self.groups[group_idx]['indices']
            self.skipped_groups.append(group_indices)

        # Move to next batch
        self.current_batch_idx += 1
        if self.current_batch_idx < len(self.confidence_batches):
            self.review_batch(self.current_batch_idx)
        else:
            self.show_batch_complete_screen()

    def merge_selected_in_batch(self):
        """Merge all selected groups in current batch"""
        batch = self.confidence_batches[self.current_batch_idx]
        total_pages = (len(batch['group_indices']) + self.groups_per_batch_page - 1) // self.groups_per_batch_page

        # Warn if multiple pages exist
        if total_pages > 1:
            selected_count = len([v for v in self.batch_selections.items() if v[1].get()])
            confirm = messagebox.askyesno("Multiple Pages",
                                         f"This batch has {total_pages} pages with {batch['count']} total groups.\n\n"
                                         f"You are currently on page {self.batch_review_page + 1}.\n"
                                         f"Selected on this page: {selected_count} groups.\n\n"
                                         f"Do you want to merge the selected groups on this page and continue?")
            if not confirm:
                return

        # Get selected group indices
        selected_groups = [group_idx for group_idx, var in self.batch_selections.items() if var.get()]

        if not selected_groups:
            messagebox.showwarning("No Selection", "Please select at least one group to merge.")
            return

        # Merge each selected group
        merged_count = 0
        for group_idx in selected_groups:
            # Check if this group has an edited merge
            if group_idx in self.edited_merges:
                # Use the edited version
                merged_contact = self.edited_merges[group_idx]
            else:
                # Auto-merge the contacts
                group_data = self.groups[group_idx]
                group_contacts = [self.contacts[idx] for idx in group_data['indices']]
                merged_contact = merge_contacts(group_contacts)

            if merged_contact:
                self.merged_contacts.append(merged_contact)
                self.merged_group_indices.add(group_idx)  # Track merged group
                merged_count += 1

        # Add unselected groups from current page to skipped
        current_page_indices = batch['group_indices'][
            self.batch_review_page * self.groups_per_batch_page:
            min((self.batch_review_page + 1) * self.groups_per_batch_page, len(batch['group_indices']))
        ]
        for group_idx in current_page_indices:
            if group_idx not in selected_groups:
                group_indices = self.groups[group_idx]['indices']
                self.skipped_groups.append(group_indices)

        # Check if there are more pages in this batch
        if self.batch_review_page < total_pages - 1:
            # Show confirmation and move to next page
            messagebox.showinfo("Page Complete",
                              f"Successfully merged {merged_count} groups from this page!\n"
                              f"Skipped {len(current_page_indices) - merged_count} groups.\n\n"
                              f"Moving to next page...")

            # Move to next page
            self.batch_review_page += 1
            self.review_batch(self.current_batch_idx, reset_page=False)
        else:
            # This was the last page - move to next batch
            messagebox.showinfo("Batch Complete",
                              f"Successfully merged {merged_count} groups from this page!\n"
                              f"Batch {self.current_batch_idx + 1} complete.\n\n"
                              f"Moving to next batch...")

            # Auto-progress to next batch
            self.current_batch_idx += 1
            if self.current_batch_idx < len(self.confidence_batches):
                self.review_batch(self.current_batch_idx)
            else:
                self.show_batch_complete_screen()

    def show_batch_complete_screen(self):
        """Show completion screen after all batches reviewed"""
        self.clear_main_frame()

        tk.Label(self.main_frame, text="✓ Batch Review Complete!",
                font=('Arial', 24, 'bold'), fg='#4CAF50').pack(pady=30)

        # Summary
        summary_frame = tk.Frame(self.main_frame, bg='#e8f5e9', relief='solid', borderwidth=2)
        summary_frame.pack(padx=100, pady=20, fill='x')

        total_original = len(self.contacts)

        processed_indices = set()
        for merged_contact in self.merged_contacts:
            # Groups that were merged
            pass
        for skipped_group in self.skipped_groups:
            processed_indices.update(skipped_group)

        # Calculate final count
        merged_group_indices = set()
        for batch in self.confidence_batches:
            for group_idx in batch['group_indices']:
                if group_idx < len(self.merged_contacts):
                    merged_group_indices.update(self.groups[group_idx]['indices'])

        unprocessed = len(self.contacts) - len(merged_group_indices) - len(processed_indices)
        final_count = len(self.merged_contacts) + len(self.contacts) - len(merged_group_indices)

        tk.Label(summary_frame, text="Summary",
                font=('Arial', 14, 'bold'), bg='#e8f5e9').pack(pady=10)

        tk.Label(summary_frame, text=f"Original contacts: {total_original:,}",
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Groups merged: {len(self.merged_contacts):,}",
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Groups skipped: {len(self.skipped_groups):,}",
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Estimated final contacts: ~{final_count:,}",
                font=('Arial', 14, 'bold'), bg='#e8f5e9', fg='#2e7d32').pack(pady=10)

        # Export button
        tk.Button(self.main_frame, text="Export Merged Contacts",
                 command=self.export_merged,
                 font=('Arial', 14, 'bold'), bg='#2196F3', fg='white',
                 height=2).pack(pady=20, padx=100, fill='x')

        tk.Button(self.main_frame, text="Back to Batch Overview",
                 command=self.show_batch_overview,
                 font=('Arial', 12), bg='#2196F3', fg='white',
                 height=2).pack(padx=100, fill='x', pady=5)

        tk.Button(self.main_frame, text="Start Over",
                 command=self.show_load_screen,
                 font=('Arial', 12), bg='#757575', fg='white',
                 height=2).pack(padx=100, fill='x')

    def show_groups_overview(self):
        """Show paginated overview of all duplicate groups"""
        self.clear_main_frame()

        # Header
        header = tk.Frame(self.main_frame, bg='#2196F3', height=80)
        header.pack(fill='x')
        header.pack_propagate(False)

        tk.Label(header, text=f"Found {len(self.groups)} Duplicate Groups",
                font=('Arial', 18, 'bold'), bg='#2196F3', fg='white').pack(pady=20)

        # Instructions
        info_frame = tk.Frame(self.main_frame, bg='#e3f2fd', height=60)
        info_frame.pack(fill='x', padx=20, pady=10)
        info_frame.pack_propagate(False)

        tk.Label(info_frame, text="Review groups one by one, or jump to a specific group. Changes are auto-saved.",
                font=('Arial', 11), bg='#e3f2fd').pack(pady=15)

        # Pagination controls
        total_pages = (len(self.groups) + self.groups_per_page - 1) // self.groups_per_page
        start_idx = self.current_overview_page * self.groups_per_page
        end_idx = min(start_idx + self.groups_per_page, len(self.groups))

        page_info_frame = tk.Frame(self.main_frame, bg='#f5f5f5')
        page_info_frame.pack(fill='x', padx=20, pady=10)

        tk.Label(page_info_frame, text=f"Page {self.current_overview_page + 1} of {total_pages} | Groups {start_idx + 1}-{end_idx} of {len(self.groups)}",
                font=('Arial', 12, 'bold'), bg='#f5f5f5').pack(side='left', padx=10)

        if self.current_overview_page > 0:
            tk.Button(page_info_frame, text="← Previous Page",
                     command=self.prev_overview_page,
                     font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=5)

        if self.current_overview_page < total_pages - 1:
            tk.Button(page_info_frame, text="Next Page →",
                     command=self.next_overview_page,
                     font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=5)

        # Groups list (scrollable)
        list_frame = tk.Frame(self.main_frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)

        canvas = tk.Canvas(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas)

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        # Show groups for current page
        for i in range(start_idx, end_idx):
            group_data = self.groups[i]
            confidence = group_data['confidence']
            group_contacts = [self.contacts[idx] for idx in group_data['indices']]

            # Group frame
            group_frame = tk.Frame(scrollable, relief='solid', borderwidth=1, bg='#fafafa')
            group_frame.pack(fill='x', padx=5, pady=5)

            # Confidence color
            if confidence >= 80:
                conf_color = '#4CAF50'
                conf_text = 'HIGH'
            elif confidence >= 60:
                conf_color = '#FF9800'
                conf_text = 'MEDIUM'
            else:
                conf_color = '#F44336'
                conf_text = 'LOW'

            # Header
            header_frame = tk.Frame(group_frame, bg=conf_color)
            header_frame.pack(fill='x')

            tk.Label(header_frame, text=f"Group {i+1}",
                    font=('Arial', 12, 'bold'), bg=conf_color, fg='white').pack(side='left', padx=10, pady=8)

            tk.Label(header_frame, text=f"{confidence}% {conf_text}",
                    font=('Arial', 10), bg=conf_color, fg='white').pack(side='left', padx=5)

            tk.Label(header_frame, text=f"{len(group_contacts)} contacts",
                    font=('Arial', 10), bg=conf_color, fg='white').pack(side='left', padx=5)

            tk.Button(header_frame, text="Review This Group →",
                     command=lambda idx=i: self.jump_to_group_from_overview(idx),
                     font=('Arial', 9), bg='white', fg=conf_color).pack(side='right', padx=10, pady=5)

            # Contact names preview
            names_frame = tk.Frame(group_frame, bg='#fafafa')
            names_frame.pack(fill='x', padx=10, pady=8)

            names_preview = ', '.join([c.get_display_name()[:30] for c in group_contacts])

            tk.Label(names_frame, text=names_preview,
                    font=('Arial', 10), bg='#fafafa', anchor='w').pack(fill='x')

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Action buttons
        button_frame = tk.Frame(self.main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', padx=20, pady=20)

        tk.Button(button_frame, text="Start Reviewing from Group 1",
                 command=lambda: self.jump_to_group_from_overview(0),
                 font=('Arial', 14, 'bold'), bg='#4CAF50', fg='white',
                 height=2).pack(fill='x')

    def prev_overview_page(self):
        """Go to previous page in overview"""
        if self.current_overview_page > 0:
            self.current_overview_page -= 1
            self.show_groups_overview()

    def next_overview_page(self):
        """Go to next page in overview"""
        total_pages = (len(self.groups) + self.groups_per_page - 1) // self.groups_per_page
        if self.current_overview_page < total_pages - 1:
            self.current_overview_page += 1
            self.show_groups_overview()

    def jump_to_group_from_overview(self, group_idx):
        """Jump to specific group from overview"""
        self.current_group_idx = group_idx
        self.show_review_screen()

    def show_review_screen(self):
        """Show detailed review screen for current group"""
        self.clear_main_frame()
        
        if self.current_group_idx >= len(self.groups):
            self.show_complete_screen()
            return
        
        group_data = self.groups[self.current_group_idx]
        group_indices = group_data['indices']
        confidence = group_data['confidence']
        
        # Store current group contacts for "remove from group" feature
        self.current_group_contacts = [self.contacts[i] for i in group_indices]
        self.current_group_indices = group_indices.copy()
        
        # Header with confidence
        header = tk.Frame(self.main_frame, height=80)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        # Confidence color
        if confidence >= 80:
            conf_color = '#4CAF50'
            conf_text = 'HIGH'
        elif confidence >= 60:
            conf_color = '#FF9800'
            conf_text = 'MEDIUM'
        else:
            conf_color = '#F44336'
            conf_text = 'LOW'
        
        header.config(bg=conf_color)
        
        title_frame = tk.Frame(header, bg=conf_color)
        title_frame.pack(side='left', padx=20, pady=15)
        
        tk.Label(title_frame, text=f"Group {self.current_group_idx + 1} of {len(self.groups)}", 
                font=('Arial', 16, 'bold'), bg=conf_color, fg='white').pack(anchor='w')
        tk.Label(title_frame, text=f"Confidence: {confidence}% ({conf_text})", 
                font=('Arial', 11), bg=conf_color, fg='white').pack(anchor='w')
        
        progress_text = f"{len(self.merged_contacts)} merged | {len(self.skipped_groups)} skipped"
        tk.Label(header, text=progress_text, 
                font=('Arial', 11), bg=conf_color, fg='white').pack(side='right', padx=20)
        
        # Search/Jump toolbar
        toolbar = tk.Frame(self.main_frame, bg='#f5f5f5', height=40)
        toolbar.pack(fill='x')
        toolbar.pack_propagate(False)

        # Back to overview button (if there are many groups)
        if len(self.groups) > 10:
            tk.Button(toolbar, text="← Back to Overview", command=self.show_groups_overview,
                     font=('Arial', 10), bg='#2196F3', fg='white').pack(side='left', padx=10)

        tk.Label(toolbar, text="Jump to group:", bg='#f5f5f5',
                font=('Arial', 10)).pack(side='left', padx=10)

        self.jump_entry = tk.Entry(toolbar, width=10, font=('Arial', 10))
        self.jump_entry.pack(side='left', padx=5)

        tk.Button(toolbar, text="Go", command=self.jump_to_group,
                 font=('Arial', 9)).pack(side='left', padx=5)

        if len(self.history) > 0:
            tk.Button(toolbar, text="← Undo", command=self.undo_last_action,
                     font=('Arial', 10), bg='#FFC107').pack(side='left', padx=20)

        tk.Button(toolbar, text="Save Progress", command=self.save_progress,
                 font=('Arial', 9), bg='#2196F3', fg='white').pack(side='right', padx=10)
        
        # Content area
        content = tk.Frame(self.main_frame)
        content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Left side - Individual contacts with remove option
        left_frame = tk.LabelFrame(content, text=f"Contacts in Group ({len(self.current_group_contacts)})", 
                                   font=('Arial', 12, 'bold'))
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        contacts_canvas = tk.Canvas(left_frame)
        scrollbar = tk.Scrollbar(left_frame, orient='vertical', command=contacts_canvas.yview)
        scrollable = tk.Frame(contacts_canvas)
        
        scrollable.bind(
            "<Configure>",
            lambda e: contacts_canvas.configure(scrollregion=contacts_canvas.bbox("all"))
        )
        
        contacts_canvas.create_window((0, 0), window=scrollable, anchor='nw')
        contacts_canvas.configure(yscrollcommand=scrollbar.set)
        
        for i, contact in enumerate(self.current_group_contacts):
            contact_frame = tk.Frame(scrollable, relief='solid', borderwidth=1, bg='#f9f9f9')
            contact_frame.pack(fill='x', padx=5, pady=5)
            
            header_frame = tk.Frame(contact_frame, bg='#e3f2fd')
            header_frame.pack(fill='x', padx=5, pady=5)
            
            tk.Label(header_frame, text=f"Contact {i+1}", 
                    font=('Arial', 10, 'bold'), bg='#e3f2fd', anchor='w').pack(side='left', padx=5)
            
            if len(self.current_group_contacts) > 2:
                tk.Button(header_frame, text="Remove from group", 
                         command=lambda idx=i: self.remove_from_group(idx),
                         font=('Arial', 8), bg='#F44336', fg='white').pack(side='right', padx=5)
            
            details_text = scrolledtext.ScrolledText(contact_frame, height=8, width=50, 
                                                     font=('Courier', 9), wrap='word')
            details_text.pack(fill='both', padx=5, pady=5)
            details_text.insert('1.0', contact.get_full_details())
            details_text.config(state='disabled')
        
        contacts_canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Right side - Merged preview
        right_frame = tk.LabelFrame(content, text="Merged Result Preview", 
                                    font=('Arial', 12, 'bold'))
        right_frame.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        self.current_merged = merge_contacts(self.current_group_contacts)
        
        merged_text = scrolledtext.ScrolledText(right_frame, font=('Courier', 10), 
                                               wrap='word', bg='#e8f5e9')
        merged_text.pack(fill='both', expand=True, padx=10, pady=10)
        merged_text.insert('1.0', self.current_merged.get_full_details())
        merged_text.config(state='disabled')
        
        # Action buttons
        button_frame = tk.Frame(self.main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', padx=20, pady=20)
        
        tk.Button(button_frame, text="← Skip This Group", 
                 command=self.skip_group,
                 font=('Arial', 12), bg='#ff9800', fg='white',
                 width=18, height=2).pack(side='left', padx=5)
        
        tk.Button(button_frame, text="✏️ Edit Merged", 
                 command=self.edit_merged,
                 font=('Arial', 12), bg='#2196F3', fg='white',
                 width=18, height=2).pack(side='left', padx=5)
        
        tk.Button(button_frame, text="✓ Merge This Group", 
                 command=self.merge_group,
                 font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white',
                 width=18, height=2).pack(side='right', padx=5)
        
        # Autosave
        self.autosave()
    
    def remove_from_group(self, contact_idx):
        """Remove a contact from the current group"""
        if len(self.current_group_contacts) <= 2:
            messagebox.showwarning("Cannot Remove", 
                                  "Cannot remove contact - group must have at least 2 contacts.")
            return
        
        contact = self.current_group_contacts[contact_idx]
        confirm = messagebox.askyesno("Remove Contact", 
                                      f"Remove '{contact.get_display_name()}' from this group?\n\n"
                                      "The contact will be kept as a separate, non-duplicate contact.")
        
        if confirm:
            # Remove from current group
            del self.current_group_contacts[contact_idx]
            del self.current_group_indices[contact_idx]
            
            # Update the group data
            self.groups[self.current_group_idx]['indices'] = self.current_group_indices.copy()
            
            # Refresh the screen
            self.show_review_screen()
    
    def edit_merged(self):
        """Open dialog to edit merged contact"""
        dialog = EditContactDialog(self.root, self.current_merged)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.current_merged = dialog.result
            # Refresh to show updated merge
            self.show_review_screen()
    
    def skip_group(self):
        """Skip this group - keep all contacts separate"""
        # Save to history for undo
        self.history.append({
            'action': 'skip',
            'group_idx': self.current_group_idx,
            'group_data': self.groups[self.current_group_idx].copy()
        })
        
        group_indices = self.groups[self.current_group_idx]['indices']
        self.skipped_groups.append(group_indices)
        self.current_group_idx += 1
        self.show_review_screen()
    
    def merge_group(self):
        """Merge this group"""
        # Save to history for undo
        self.history.append({
            'action': 'merge',
            'group_idx': self.current_group_idx,
            'merged_contact': self.current_merged.copy(),
            'group_data': self.groups[self.current_group_idx].copy()
        })
        
        self.merged_contacts.append(self.current_merged)
        self.current_group_idx += 1
        self.show_review_screen()
    
    def undo_last_action(self):
        """Undo the last merge or skip action"""
        if not self.history:
            return
        
        last_action = self.history.pop()
        
        if last_action['action'] == 'merge':
            # Remove the merged contact
            self.merged_contacts.pop()
        elif last_action['action'] == 'skip':
            # Remove from skipped groups
            self.skipped_groups.pop()
        
        # Go back to that group
        self.current_group_idx = last_action['group_idx']
        self.show_review_screen()
    
    def jump_to_group(self):
        """Jump to a specific group number"""
        try:
            group_num = int(self.jump_entry.get())
            if 1 <= group_num <= len(self.groups):
                self.current_group_idx = group_num - 1
                self.show_review_screen()
            else:
                messagebox.showwarning("Invalid Group", 
                                      f"Please enter a number between 1 and {len(self.groups)}")
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid group number")
    
    def save_progress(self):
        """Save progress to file"""
        save_data = {
            'current_group_idx': self.current_group_idx,
            'merged_indices': [g['indices'] for i, g in enumerate(self.groups[:self.current_group_idx]) 
                             if i < len(self.merged_contacts)],
            'skipped_indices': self.skipped_groups,
            'threshold': self.threshold_var.get(),
            'file1': self.file1_path if hasattr(self, 'file1_path') else None,
            'file2': self.file2_path if hasattr(self, 'file2_path') else None
        }
        
        save_path = filedialog.asksaveasfilename(
            title="Save Progress",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile="vcard_merge_progress.json"
        )
        
        if save_path:
            with open(save_path, 'w') as f:
                json.dump(save_data, f)
            messagebox.showinfo("Saved", f"Progress saved to {save_path}")
    
    def autosave(self):
        """Autosave progress"""
        if hasattr(self, 'file1_path') and hasattr(self, 'file2_path'):
            save_data = {
                'current_group_idx': self.current_group_idx,
                'merged_count': len(self.merged_contacts),
                'skipped_count': len(self.skipped_groups),
                'file1': self.file1_path,
                'file2': self.file2_path
            }
            
            try:
                with open('.vcard_merge_autosave.json', 'w') as f:
                    json.dump(save_data, f)
            except:
                pass
    
    def load_autosave(self):
        """Load autosave if exists"""
        try:
            if os.path.exists('.vcard_merge_autosave.json'):
                with open('.vcard_merge_autosave.json', 'r') as f:
                    data = json.load(f)
                
                # Ask user if they want to resume
                resume = messagebox.askyesno("Resume Previous Session", 
                                            f"Found autosaved progress:\n"
                                            f"Group {data['current_group_idx']} / {data.get('merged_count', 0)} merged\n"
                                            f"\nResume from where you left off?")
                
                if resume:
                    self.file1_path = data['file1']
                    self.file2_path = data['file2']
                    # Would need to re-load and continue...
                    messagebox.showinfo("Note", "Please reload the same files to resume")
        except:
            pass
    
    def show_complete_screen(self):
        """Show completion screen with export option"""
        self.clear_main_frame()
        
        tk.Label(self.main_frame, text="✓ Review Complete!", 
                font=('Arial', 24, 'bold'), fg='#4CAF50').pack(pady=30)
        
        # Summary
        summary_frame = tk.Frame(self.main_frame, bg='#e8f5e9', relief='solid', borderwidth=2)
        summary_frame.pack(padx=100, pady=20, fill='x')
        
        total_original = len(self.contacts)
        
        processed_indices = set()
        for i in range(self.current_group_idx):
            if i < len(self.merged_contacts):
                processed_indices.update(self.groups[i]['indices'])
        
        unprocessed = len(self.contacts) - len(processed_indices)
        final_count = len(self.merged_contacts) + unprocessed
        
        tk.Label(summary_frame, text="Summary", 
                font=('Arial', 14, 'bold'), bg='#e8f5e9').pack(pady=10)
        
        tk.Label(summary_frame, text=f"Original contacts: {total_original:,}", 
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Groups merged: {len(self.merged_contacts):,}", 
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Groups skipped: {len(self.skipped_groups):,}", 
                font=('Arial', 11), bg='#e8f5e9').pack()
        tk.Label(summary_frame, text=f"Final contact count: {final_count:,}", 
                font=('Arial', 14, 'bold'), bg='#e8f5e9', fg='#2e7d32').pack(pady=10)
        
        # Export button
        tk.Button(self.main_frame, text="Export Merged Contacts", 
                 command=self.export_merged,
                 font=('Arial', 14, 'bold'), bg='#2196F3', fg='white',
                 height=2).pack(pady=20, padx=100, fill='x')
        
        tk.Button(self.main_frame, text="Start Over",
                 command=self.show_load_screen,
                 font=('Arial', 12), bg='#757575', fg='white',
                 height=2).pack(padx=100, fill='x')
    
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
            processed_indices = set()

            # Sequential mode: use current_group_idx
            for i, group_data in enumerate(self.groups[:self.current_group_idx]):
                if i < len(self.merged_contacts):
                    processed_indices.update(group_data['indices'])

            # Batch mode: use merged_group_indices
            for group_idx in self.merged_group_indices:
                if group_idx < len(self.groups):
                    processed_indices.update(self.groups[group_idx]['indices'])

            final_contacts = self.merged_contacts.copy()
            
            for i, contact in enumerate(self.contacts):
                if i not in processed_indices:
                    final_contacts.append(contact)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                for contact in final_contacts:
                    f.write(contact.to_vcard() + '\n\n')
            
            # Clean up autosave
            try:
                os.remove('.vcard_merge_autosave.json')
            except:
                pass
            
            messagebox.showinfo("Success", 
                              f"Successfully exported {len(final_contacts):,} contacts to:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting:\n{str(e)}")


if __name__ == '__main__':
    root = tk.Tk()
    app = MergerApp(root)
    root.mainloop()
