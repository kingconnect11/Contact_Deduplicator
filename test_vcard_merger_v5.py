#!/usr/bin/env python3
"""
Test Script for vCard Contact Merger v5

This script verifies all core matching functions work correctly.
Run with: python3 test_vcard_merger_v5.py
"""

import sys

# Import functions from the main module
from vcard_merger_v5 import (
    parse_name_parts,
    soundex,
    resolve_nickname,
    normalize_email,
    normalize_phone,
    names_match,
    phones_match,
    get_canonical_first_name,
)


def test_parse_name_parts():
    """Test name parsing and order detection - CRITICAL"""
    print("\n" + "="*60)
    print("TEST: parse_name_parts() - Name Order Detection")
    print("="*60)

    tests = [
        # (input, expected_canonical)
        ("John Smith", "john smith"),
        ("Smith, John", "john smith"),
        ("Dr. John Smith", "john smith"),
        ("John Smith Jr.", "john smith"),
        ("Dr. John Smith Jr.", "john smith"),
        ("Mary Jane Otte", "mary otte"),
        ("Otte, Mary Jane", "mary otte"),
        ("Prof. Robert Johnson PhD", "robert johnson"),
        ("Mrs. Elizabeth Taylor", "elizabeth taylor"),
        ("Smith, Dr. John", "john smith"),  # Edge case
    ]

    passed = 0
    failed = 0

    for name, expected_canonical in tests:
        first, last, canonical = parse_name_parts(name)
        if canonical == expected_canonical:
            print(f"  PASS: '{name}' -> '{canonical}'")
            passed += 1
        else:
            print(f"  FAIL: '{name}' -> '{canonical}' (expected '{expected_canonical}')")
            failed += 1

    # Critical test: Same person different order
    r1 = parse_name_parts("John Smith")
    r2 = parse_name_parts("Smith, John")
    if r1[2] == r2[2]:
        print(f"  PASS: 'John Smith' == 'Smith, John' (both -> '{r1[2]}')")
        passed += 1
    else:
        print(f"  FAIL: 'John Smith' != 'Smith, John' ('{r1[2]}' != '{r2[2]}')")
        failed += 1

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_soundex():
    """Test phonetic matching with Soundex"""
    print("\n" + "="*60)
    print("TEST: soundex() - Phonetic Matching")
    print("="*60)

    tests = [
        # (name1, name2, should_match)
        ("Smith", "Smyth", True),
        ("Meyer", "Mayer", True),
        ("Peterson", "Pederson", True),
        ("Johnson", "Jonson", True),
        ("Brown", "Braun", True),
        ("Smith", "Jones", False),
        ("Wilson", "Williams", False),
    ]

    passed = 0
    failed = 0

    for name1, name2, should_match in tests:
        s1 = soundex(name1)
        s2 = soundex(name2)
        matches = (s1 == s2)

        if matches == should_match:
            status = "PASS"
            passed += 1
        else:
            status = "FAIL"
            failed += 1

        match_text = "MATCH" if matches else "NO MATCH"
        expected_text = "expected match" if should_match else "expected no match"
        print(f"  {status}: '{name1}' ({s1}) vs '{name2}' ({s2}) -> {match_text} ({expected_text})")

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_resolve_nickname():
    """Test nickname resolution"""
    print("\n" + "="*60)
    print("TEST: resolve_nickname() - Nickname Matching")
    print("="*60)

    tests = [
        # (nickname, expected_canonical)
        ("bob", "robert"),
        ("robert", "robert"),
        ("bill", "william"),
        ("liz", "elizabeth"),
        ("mike", "michael"),
        ("jim", "james"),
        ("kate", "katherine"),
        ("peggy", "margaret"),
        ("dick", "richard"),
        ("ted", "edward"),
        ("unknown_name", "unknown_name"),  # Unknown should return itself
    ]

    passed = 0
    failed = 0

    for nickname, expected in tests:
        result = resolve_nickname(nickname)
        if result == expected:
            print(f"  PASS: '{nickname}' -> '{result}'")
            passed += 1
        else:
            print(f"  FAIL: '{nickname}' -> '{result}' (expected '{expected}')")
            failed += 1

    # Test that "Bob Smith" matches "Robert Smith"
    bob_canonical = get_canonical_first_name("bob")
    robert_canonical = get_canonical_first_name("robert")
    if bob_canonical == robert_canonical:
        print(f"  PASS: get_canonical_first_name('bob') == get_canonical_first_name('robert') ('{bob_canonical}')")
        passed += 1
    else:
        print(f"  FAIL: get_canonical_first_name('bob') != get_canonical_first_name('robert')")
        failed += 1

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_normalize_email():
    """Test email normalization"""
    print("\n" + "="*60)
    print("TEST: normalize_email() - Email Normalization")
    print("="*60)

    tests = [
        # (email1, email2, should_match_after_normalization)
        ("John.Doe@Gmail.com", "johndoe@gmail.com", True),  # Case + Gmail dots
        ("john.doe@gmail.com", "johndoe@gmail.com", True),  # Gmail dots
        ("user+tag@gmail.com", "user@gmail.com", True),     # Plus addressing
        ("user@googlemail.com", "user@gmail.com", True),    # Domain alias
        ("JOHN@EXAMPLE.COM", "john@example.com", True),     # Case only
        ("john@gmail.com", "john@yahoo.com", False),        # Different domain
    ]

    passed = 0
    failed = 0

    for email1, email2, should_match in tests:
        norm1 = normalize_email(email1)
        norm2 = normalize_email(email2)
        matches = (norm1 == norm2)

        if matches == should_match:
            status = "PASS"
            passed += 1
        else:
            status = "FAIL"
            failed += 1

        match_text = "MATCH" if matches else "NO MATCH"
        expected_text = "expected match" if should_match else "expected no match"
        print(f"  {status}: '{email1}' vs '{email2}' -> {match_text}")
        print(f"          (normalized: '{norm1}' vs '{norm2}')")

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_normalize_phone():
    """Test phone normalization"""
    print("\n" + "="*60)
    print("TEST: normalize_phone() - Phone Normalization")
    print("="*60)

    tests = [
        # (phone1, phone2, should_match)
        ("(650) 555-1234", "650.555.1234", True),
        ("650-555-1234", "6505551234", True),
        ("+1-650-555-1234", "650-555-1234", True),  # +1 country code
        ("1-650-555-1234", "(650) 555-1234", True),
        ("555-1234", "555.1234", True),  # 7 digit match
        ("650-555-1234", "415-555-1234", False),  # Different area code
    ]

    passed = 0
    failed = 0

    for phone1, phone2, should_match in tests:
        norm1 = normalize_phone(phone1)
        norm2 = normalize_phone(phone2)

        # Check if they match (considering 7 or 10 digit comparison)
        if len(norm1) >= 10 and len(norm2) >= 10:
            matches = (norm1 == norm2)
        elif len(norm1) >= 7 and len(norm2) >= 7:
            matches = (norm1[-7:] == norm2[-7:])
        else:
            matches = (norm1 == norm2)

        if matches == should_match:
            status = "PASS"
            passed += 1
        else:
            status = "FAIL"
            failed += 1

        match_text = "MATCH" if matches else "NO MATCH"
        expected_text = "expected match" if should_match else "expected no match"
        print(f"  {status}: '{phone1}' vs '{phone2}' -> {match_text}")
        print(f"          (normalized: '{norm1}' vs '{norm2}')")

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_names_match():
    """Test comprehensive name matching"""
    print("\n" + "="*60)
    print("TEST: names_match() - Comprehensive Name Matching")
    print("="*60)

    tests = [
        # (name1, name2, should_match, min_confidence)
        ("John Smith", "Smith, John", True, 90),       # Name order
        ("Bob Smith", "Robert Smith", True, 70),       # Nickname
        ("Dr. John Smith Jr.", "Smith, John", True, 70),  # Titles/suffixes
        ("Mary Jane Otte", "Otte, Mary Jane", True, 90),  # Multi-word first name
        ("Michael Brown", "Mike Brown", True, 70),     # Nickname
        ("Elizabeth Taylor", "Liz Taylor", True, 70),  # Nickname
        ("John Smith", "John Doe", False, 0),          # Different last name
        ("John Smith", "Jane Smith", False, 0),        # Different first name
    ]

    passed = 0
    failed = 0

    for name1, name2, should_match, min_conf in tests:
        is_match, confidence, reasons = names_match(name1, name2)

        match_ok = (is_match == should_match)
        conf_ok = (confidence >= min_conf) if should_match else True

        if match_ok and conf_ok:
            status = "PASS"
            passed += 1
        else:
            status = "FAIL"
            failed += 1

        match_text = "MATCH" if is_match else "NO MATCH"
        expected_text = f"expected {'match' if should_match else 'no match'}"
        if should_match:
            expected_text += f" (conf >= {min_conf}%)"

        print(f"  {status}: '{name1}' vs '{name2}'")
        print(f"          -> {match_text} ({confidence}%) - {expected_text}")
        if reasons:
            print(f"          Reasons: {', '.join(reasons[:2])}")

    print(f"\nResults: {passed} passed, {failed} failed")
    return failed == 0


def test_critical_verification():
    """Critical verification test matching ralph.sh expectations"""
    print("\n" + "="*60)
    print("CRITICAL VERIFICATION (ralph.sh compatibility)")
    print("="*60)

    all_pass = True

    # Test 1: Name order detection
    r1 = parse_name_parts('John Smith')
    r2 = parse_name_parts('Smith, John')

    if r1[2] == r2[2] == 'john smith':
        print("  PASS: 'John Smith' and 'Smith, John' both produce 'john smith'")
    else:
        print(f"  FAIL: 'John Smith' -> '{r1[2]}', 'Smith, John' -> '{r2[2]}'")
        all_pass = False

    # Test 2: Soundex
    if soundex('Smith') == soundex('Smyth'):
        print(f"  PASS: soundex('Smith') == soundex('Smyth') ({soundex('Smith')})")
    else:
        print(f"  FAIL: soundex('Smith') ({soundex('Smith')}) != soundex('Smyth') ({soundex('Smyth')})")
        all_pass = False

    # Test 3: Nickname resolution
    if resolve_nickname('bob') == 'robert':
        print("  PASS: resolve_nickname('bob') == 'robert'")
    else:
        print(f"  FAIL: resolve_nickname('bob') == '{resolve_nickname('bob')}' (expected 'robert')")
        all_pass = False

    if all_pass:
        print("\n" + "*"*60)
        print("*  ALL CRITICAL TESTS PASSED!")
        print("*"*60)
    else:
        print("\n" + "*"*60)
        print("*  SOME CRITICAL TESTS FAILED!")
        print("*"*60)

    return all_pass


def main():
    """Run all tests"""
    print("="*60)
    print("vCard Contact Merger v5 - Test Suite")
    print("="*60)

    results = []

    results.append(("parse_name_parts", test_parse_name_parts()))
    results.append(("soundex", test_soundex()))
    results.append(("resolve_nickname", test_resolve_nickname()))
    results.append(("normalize_email", test_normalize_email()))
    results.append(("normalize_phone", test_normalize_phone()))
    results.append(("names_match", test_names_match()))
    results.append(("CRITICAL", test_critical_verification()))

    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)

    all_passed = True
    for name, passed in results:
        status = "PASS" if passed else "FAIL"
        print(f"  {status}: {name}")
        if not passed:
            all_passed = False

    print()
    if all_passed:
        print("ALL TESTS PASSED!")
        return 0
    else:
        print("SOME TESTS FAILED!")
        return 1


if __name__ == '__main__':
    sys.exit(main())
