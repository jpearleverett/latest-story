#!/usr/bin/env python3
"""
Continue improvements: vary sentence structure, remove repetitive patterns, tighten chapters 5-8
"""

import docx
import glob
import re

def vary_repetitive_patterns(text):
    """Replace repetitive sentence starters and patterns"""
    replacements = [
        # Vary "I walked into"
        (r"I walked into ([^.]+)\.", r"I entered \1."),
        (r"I walked into ([^.]+) alone\.", r"I stepped inside \1 alone."),
        
        # Vary "I drove to"
        (r"I drove to ([^.]+)\.", r"I headed to \1."),
        (r"I drove straight to ([^.]+)\.", r"I made for \1."),
        
        # Vary "I rushed back"
        (r"I rushed back to ([^.]+)\.", r"I returned to \1."),
        
        # Vary "I looked at"
        (r"I looked at ([^.]+)\.", r"I studied \1."),
        (r"I looked at ([^.]+)\. ([^.]+)\.", r"I studied \1. \2."),
        
        # Remove "I knew" repetition
        (r"I knew ([^.]+)\. But I also knew", r"\1. But"),
        
        # Vary "My phone"
        (r"My phone buzzed\. ([A-Z])", r"My phone rang. \1"),
    ]
    
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    
    return text

def tighten_exposition(text):
    """Cut unnecessary exposition dumps"""
    # Remove overly explanatory sentences
    patterns_to_remove = [
        r"This was the Methodical path\. Slow\. Difficult\. Legal\. But clean\. We were rebuilding the foundation of justice piece by painful piece\.",
        r"The aggressive path had delivered the truth\. But now I had the choice that would define whether I was still a cop or just a monster with a badge\.",
    ]
    
    for pattern in patterns_to_remove:
        text = re.sub(pattern, "", text)
    
    return text

def process_file(filepath):
    """Process file with continued improvements"""
    doc = docx.Document(filepath)
    changes = 0
    current_chapter = 0
    
    for para in doc.paragraphs:
        text = para.text
        
        # Track chapter
        chapter_match = re.search(r'CHAPTER (\d+)', text)
        if chapter_match:
            current_chapter = int(chapter_match.group(1))
        
        # Apply improvements
        original = text
        
        # Vary repetitive patterns
        text = vary_repetitive_patterns(text)
        
        # Tighten chapters 5-8
        if 5 <= current_chapter <= 8:
            text = tighten_exposition(text)
        
        if text != original:
            para.text = text
            changes += 1
    
    doc.save(filepath)
    return changes

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    print("Continuing improvements: varying patterns, tightening chapters 5-8...")
    total_changes = 0
    
    for file in files:
        print(f"Processing {file}...", end=" ")
        try:
            changes = process_file(file)
            total_changes += changes
            print(f"✓ ({changes} changes)")
        except Exception as e:
            print(f"✗ Error: {e}")
    
    print(f"\n✓ Complete. {total_changes} total changes.")

if __name__ == "__main__":
    main()
