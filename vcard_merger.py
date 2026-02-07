#!/usr/bin/env python3
"""
vCard Contact Merger - Efficiently merge large vCard files
Handles 10k+ contacts with smart duplicate detection
"""

import re
import hashlib
from collections import defaultdict
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class VCardContact:
    def __init__(self, vcard_text):
        self.raw = vcard_text
        self.fn = ''
        self.names = []
        self.emails = set()
        self.phones = set()
        self.addresses = []
        self.notes = []
        self.org = ''
        self.title = ''
        self.other_fields = []
        self._parse(vcard_text)
    
    def _parse(self, text):
        lines = text.split('\n')
        current_line = ''
        
        for line in lines:
            # Handle line continuation (lines starting with space/tab)
            if line.startswith(' ') or line.startswith('\t'):
                current_line += line[1:]
                continue
            
            if current_line:
                self._process_line(current_line)
            current_line = line
        
        if current_line:
            self._process_line(current_line)
    
    def _process_line(self, line):
        line = line.strip()
        if not line or ':' not in line:
            return
        
        field, value = line.split(':', 1)
        field = field.upper()
        value = value.strip()
        
        if not value:
            return
        
        if field.startswith('FN'):
            self.fn = value
            self.names.append(value.lower())
        elif field.startswith('N'):
            # N field format: Family;Given;Additional;Prefix;Suffix
            parts = value.split(';')
            if len(parts) >= 2 and parts[0] and parts[1]:
                full_name = f"{parts[1]} {parts[0]}".strip()
                self.names.append(full_name.lower())
        elif field.startswith('EMAIL'):
            self.emails.add(value.lower())
        elif field.startswith('TEL'):
            # Normalize phone: keep only digits and +
            normalized = re.sub(r'[^\d+]', '', value)
            if len(normalized) >= 7:  # Minimum valid phone length
                self.phones.add(normalized)
        elif field.startswith('ADR'):
            self.addresses.append(value)
        elif field.startswith('NOTE'):
            self.notes.append(value)
        elif field.startswith('ORG'):
            self.org = value
        elif field.startswith('TITLE'):
            self.title = value
        elif not field.startswith('BEGIN') and not field.startswith('END') and not field.startswith('VERSION'):
            self.other_fields.append(line)
    
    def get_hash_keys(self):
        """Return set of hash keys for quick duplicate lookup"""
        keys = set()
        
        # Email-based keys (most reliable)
        for email in self.emails:
            keys.add(f"email:{email}")
        
        # Phone-based keys (last 10 digits)
        for phone in self.phones:
            if len(phone) >= 10:
                keys.add(f"phone:{phone[-10:]}")
        
        # Name-based keys (first+last word)
        for name in self.names:
            words = name.split()
            if len(words) >= 2:
                keys.add(f"name:{words[0]}:{words[-1]}")
            elif len(words) == 1:
                keys.add(f"name:{words[0]}")
        
        return keys
    
    def similarity_score(self, other):
        """Calculate similarity score between contacts (0-1)"""
        score = 0.0
        checks = 0
        
        # Email match is definitive
        if self.emails and other.emails:
            if self.emails & other.emails:
                return 1.0
            checks += 1
        
        # Phone match (comparing last 7-10 digits)
        if self.phones and other.phones:
            for p1 in self.phones:
                for p2 in other.phones:
                    if len(p1) >= 7 and len(p2) >= 7:
                        if p1[-7:] == p2[-7:]:
                            score += 1.0
                            checks += 1
                            break
                if checks > score:
                    break
        
        # Name similarity
        if self.names and other.names:
            max_sim = 0
            for n1 in self.names:
                for n2 in other.names:
                    sim = SequenceMatcher(None, n1, n2).ratio()
                    max_sim = max(max_sim, sim)
            score += max_sim
            checks += 1
        
        return score / checks if checks > 0 else 0.0
    
    def merge_with(self, other):
        """Merge another contact into this one"""
        self.emails.update(other.emails)
        self.phones.update(other.phones)
        self.addresses.extend([a for a in other.addresses if a not in self.addresses])
        self.notes.extend([n for n in other.notes if n not in self.notes])
        self.other_fields.extend([f for f in other.other_fields if f not in self.other_fields])
        
        # Keep non-empty values
        if not self.fn and other.fn:
            self.fn = other.fn
        if not self.org and other.org:
            self.org = other.org
        if not self.title and other.title:
            self.title = other.title
        
        # Merge names
        for name in other.names:
            if name not in self.names:
                self.names.append(name)
    
    def to_vcard(self):
        """Convert back to vCard format"""
        lines = ['BEGIN:VCARD', 'VERSION:3.0']
        
        if self.fn:
            lines.append(f'FN:{self.fn}')
        
        for email in sorted(self.emails):
            lines.append(f'EMAIL:{email}')
        
        for phone in sorted(self.phones):
            lines.append(f'TEL:{phone}')
        
        for addr in self.addresses:
            lines.append(f'ADR:{addr}')
        
        if self.org:
            lines.append(f'ORG:{self.org}')
        
        if self.title:
            lines.append(f'TITLE:{self.title}')
        
        for note in self.notes:
            lines.append(f'NOTE:{note}')
        
        # Add other preserved fields
        for field in self.other_fields:
            lines.append(field)
        
        lines.append('END:VCARD')
        return '\n'.join(lines)


class VCardMerger:
    def __init__(self, threshold=0.85):
        self.threshold = threshold
        self.contacts = []
        self.duplicate_groups = []
    
    def parse_file(self, filepath):
        """Parse a vCard file and return list of contacts"""
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Split by BEGIN:VCARD
        vcards = re.split(r'BEGIN:VCARD', content, flags=re.IGNORECASE)
        contacts = []
        
        for vcard in vcards[1:]:  # Skip first empty split
            vcard_text = 'BEGIN:VCARD' + vcard
            if 'END:VCARD' in vcard_text.upper():
                try:
                    contact = VCardContact(vcard_text)
                    if contact.fn or contact.emails or contact.phones:
                        contacts.append(contact)
                except Exception as e:
                    print(f"Error parsing contact: {e}")
                    continue
        
        return contacts
    
    def find_duplicates(self, contacts1, contacts2):
        """Efficiently find duplicates using hash-based bucketing"""
        all_contacts = contacts1 + contacts2
        
        # Build hash index for fast lookup
        hash_index = defaultdict(list)
        for idx, contact in enumerate(all_contacts):
            for key in contact.get_hash_keys():
                hash_index[key].append(idx)
        
        # Find potential duplicate groups
        processed = set()
        groups = []
        
        for idx, contact in enumerate(all_contacts):
            if idx in processed:
                continue
            
            # Find candidates using hash index
            candidates = set()
            for key in contact.get_hash_keys():
                candidates.update(hash_index[key])
            
            # Remove self and already processed
            candidates.discard(idx)
            candidates -= processed
            
            # Check similarity scores
            group = [idx]
            for candidate_idx in candidates:
                score = contact.similarity_score(all_contacts[candidate_idx])
                if score >= self.threshold:
                    group.append(candidate_idx)
                    processed.add(candidate_idx)
            
            if len(group) > 1:
                groups.append(group)
            
            processed.add(idx)
        
        self.contacts = all_contacts
        self.duplicate_groups = groups
        return groups
    
    def merge_duplicates(self):
        """Merge all duplicate groups"""
        merged_contacts = []
        processed = set()
        
        for group in self.duplicate_groups:
            # Merge all contacts in group into first one
            base_contact = self.contacts[group[0]]
            for idx in group[1:]:
                base_contact.merge_with(self.contacts[idx])
                processed.add(idx)
            merged_contacts.append(base_contact)
            processed.add(group[0])
        
        # Add non-duplicate contacts
        for idx, contact in enumerate(self.contacts):
            if idx not in processed:
                merged_contacts.append(contact)
        
        return merged_contacts
    
    def export_to_file(self, contacts, filepath):
        """Export contacts to vCard file"""
        with open(filepath, 'w', encoding='utf-8') as f:
            for contact in contacts:
                f.write(contact.to_vcard() + '\n\n')


class MergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("vCard Contact Merger")
        self.root.geometry("700x550")
        
        self.file1_path = None
        self.file2_path = None
        self.merger = VCardMerger()
        
        self.setup_ui()
    
    def setup_ui(self):
        # Title
        title = tk.Label(self.root, text="vCard Contact Merger", 
                        font=('Arial', 20, 'bold'), pady=20)
        title.pack()
        
        # File 1
        frame1 = tk.Frame(self.root, pady=10)
        frame1.pack(fill='x', padx=20)
        
        tk.Label(frame1, text="First vCard file:", font=('Arial', 11)).pack(anchor='w')
        btn_frame1 = tk.Frame(frame1)
        btn_frame1.pack(fill='x', pady=5)
        
        self.file1_label = tk.Label(btn_frame1, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w')
        self.file1_label.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(btn_frame1, text="Browse...", command=self.select_file1,
                 width=12).pack(side='right')
        
        # File 2
        frame2 = tk.Frame(self.root, pady=10)
        frame2.pack(fill='x', padx=20)
        
        tk.Label(frame2, text="Second vCard file:", font=('Arial', 11)).pack(anchor='w')
        btn_frame2 = tk.Frame(frame2)
        btn_frame2.pack(fill='x', pady=5)
        
        self.file2_label = tk.Label(btn_frame2, text="No file selected", 
                                    bg='#f0f0f0', relief='sunken', anchor='w')
        self.file2_label.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(btn_frame2, text="Browse...", command=self.select_file2,
                 width=12).pack(side='right')
        
        # Threshold slider
        threshold_frame = tk.Frame(self.root, pady=15)
        threshold_frame.pack(fill='x', padx=20)
        
        tk.Label(threshold_frame, text="Name Matching Threshold:", 
                font=('Arial', 11)).pack(anchor='w')
        
        slider_frame = tk.Frame(threshold_frame)
        slider_frame.pack(fill='x', pady=5)
        
        self.threshold_var = tk.DoubleVar(value=0.85)
        self.threshold_slider = tk.Scale(slider_frame, from_=0.6, to=1.0, 
                                        resolution=0.05, orient='horizontal',
                                        variable=self.threshold_var,
                                        command=self.update_threshold_label)
        self.threshold_slider.pack(side='left', fill='x', expand=True)
        
        self.threshold_label = tk.Label(slider_frame, text="85%", width=5)
        self.threshold_label.pack(side='right', padx=(10, 0))
        
        tk.Label(threshold_frame, text="Higher = stricter matching", 
                font=('Arial', 9), fg='gray').pack(anchor='w')
        
        # Progress area
        self.progress_frame = tk.Frame(self.root, pady=10)
        self.progress_frame.pack(fill='x', padx=20)
        
        self.status_label = tk.Label(self.progress_frame, text="", 
                                     font=('Arial', 10))
        self.status_label.pack()
        
        self.progress = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=5)
        self.progress.pack_forget()
        
        # Results text
        self.results_text = tk.Text(self.root, height=8, width=70, 
                                   font=('Arial', 10), state='disabled',
                                   bg='#f9f9f9')
        self.results_text.pack(padx=20, pady=10, fill='both', expand=True)
        
        # Buttons
        btn_frame = tk.Frame(self.root, pady=10)
        btn_frame.pack(fill='x', padx=20)
        
        self.merge_btn = tk.Button(btn_frame, text="Merge Contacts", 
                                   command=self.start_merge,
                                   font=('Arial', 12, 'bold'),
                                   bg='#4CAF50', fg='white',
                                   height=2, state='disabled')
        self.merge_btn.pack(fill='x')
    
    def update_threshold_label(self, value):
        self.threshold_label.config(text=f"{int(float(value)*100)}%")
    
    def select_file1(self):
        path = filedialog.askopenfilename(
            title="Select first vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file1_path = path
            self.file1_label.config(text=path.split('/')[-1])
            self.check_ready()
    
    def select_file2(self):
        path = filedialog.askopenfilename(
            title="Select second vCard file",
            filetypes=[("vCard files", "*.vcf *.vcard"), ("All files", "*.*")]
        )
        if path:
            self.file2_path = path
            self.file2_label.config(text=path.split('/')[-1])
            self.check_ready()
    
    def check_ready(self):
        if self.file1_path and self.file2_path:
            self.merge_btn.config(state='normal')
    
    def update_results(self, text):
        self.results_text.config(state='normal')
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(1.0, text)
        self.results_text.config(state='disabled')
    
    def start_merge(self):
        self.merge_btn.config(state='disabled')
        self.progress.pack(fill='x', pady=5)
        self.progress.start(10)
        
        # Run merge in separate thread
        thread = threading.Thread(target=self.perform_merge)
        thread.start()
    
    def perform_merge(self):
        try:
            # Update status
            self.status_label.config(text="Loading files...")
            self.root.update()
            
            # Parse files
            contacts1 = self.merger.parse_file(self.file1_path)
            contacts2 = self.merger.parse_file(self.file2_path)
            
            self.status_label.config(
                text=f"Loaded {len(contacts1)} and {len(contacts2)} contacts. Finding duplicates..."
            )
            self.root.update()
            
            # Update threshold
            self.merger.threshold = self.threshold_var.get()
            
            # Find duplicates
            groups = self.merger.find_duplicates(contacts1, contacts2)
            
            self.status_label.config(text="Merging duplicates...")
            self.root.update()
            
            # Merge
            merged = self.merger.merge_duplicates()
            
            # Save file
            output_path = filedialog.asksaveasfilename(
                title="Save merged contacts",
                defaultextension=".vcf",
                filetypes=[("vCard files", "*.vcf"), ("All files", "*.*")]
            )
            
            if output_path:
                self.merger.export_to_file(merged, output_path)
                
                # Show results
                total = len(contacts1) + len(contacts2)
                duplicates = sum(len(g) - 1 for g in groups)
                
                results = f"""✓ Merge Complete!

Total contacts processed: {total:,}
Duplicate groups found: {len(groups):,}
Contacts merged: {duplicates:,}
Final unique contacts: {len(merged):,}

Output saved to:
{output_path}

Ready to import into your contacts app!"""
                
                self.update_results(results)
                self.status_label.config(text="Success!", fg='green')
                messagebox.showinfo("Success", "Contacts merged successfully!")
            else:
                self.status_label.config(text="Cancelled")
        
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg='red')
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        finally:
            self.progress.stop()
            self.progress.pack_forget()
            self.merge_btn.config(state='normal')


if __name__ == '__main__':
    root = tk.Tk()
    app = MergerGUI(root)
    root.mainloop()
