#!/usr/bin/env python3
"""
Expand short third subchapters in chapters 7-12 and remove all em dashes
"""

import docx
import glob
import re
import os

def remove_em_dashes(text):
    """Replace all em dashes with regular dashes or appropriate punctuation"""
    # Replace em dash (—) with regular dash (-) or comma depending on context
    text = re.sub(r' —([A-Z])', r'. \1', text)  # End of sentence em dash
    text = re.sub(r'—([A-Z])', r'. \1', text)  # Start of sentence em dash
    text = re.sub(r' —', r'. ', text)  # Space em dash
    text = re.sub(r'—', r'-', text)  # Any remaining em dashes
    return text

def expand_short_subchapter(text, subchapter_num, chapter_num):
    """Expand short third subchapters with more content"""
    if subchapter_num != 3:
        return text
    
    # Check if it's a third subchapter and if it's too short
    word_count = len(text.split())
    
    # Third subchapters should be at least 150 words
    if word_count < 150 and chapter_num >= 7:
        # Find the decision point
        decision_match = re.search(r'\[DECISION POINT\]', text)
        if decision_match:
            # Get text before decision point
            before_decision = text[:decision_match.start()]
            
            # Check if ending is weak (needs expansion)
            weak_endings = [
                r'But now I had a choice\.',
                r'The convergence point was here\.',
                r'I had to choose\.',
                r'I was trapped\.',
            ]
            
            needs_expansion = any(re.search(pattern, before_decision[-200:]) for pattern in weak_endings)
            
            if needs_expansion:
                # Add more tension and concrete threat before decision
                expansion = "\n\nMy phone vibrated. Sarah. 'Jack, we have a problem. The FBI just got a warrant. They're moving in ten minutes. You need to decide now.'\n\nI looked at the evidence. At the choice. The clock was ticking. The system was closing in."
                
                # Insert before decision point
                text = before_decision + expansion + "\n\n" + text[decision_match.start():]
    
    return text

def process_file(filepath, output_path):
    """Process a single file"""
    doc = docx.Document(filepath)
    changes = 0
    current_chapter = 0
    current_subchapter = 0
    
    for para in doc.paragraphs:
        text = para.text
        
        # Track chapter and subchapter numbers
        chapter_match = re.search(r'CHAPTER (\d+)', text)
        if chapter_match:
            current_chapter = int(chapter_match.group(1))
        
        subchapter_match = re.search(r'Subchapter \d+\.(\d+)[A-Z]*', text)
        if subchapter_match:
            current_subchapter = int(subchapter_match.group(1))
        
        # Remove em dashes
        original = text
        text = remove_em_dashes(text)
        if text != original:
            para.text = text
            changes += 1
        
        # Expand short third subchapters in chapters 7-12
        if current_chapter >= 7 and current_chapter <= 12:
            if current_subchapter == 3:
                original = para.text
                expanded = expand_short_subchapter(para.text, 3, current_chapter)
                if expanded != original:
                    para.text = expanded
                    changes += 1
    
    doc.save(output_path)
    return changes

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    # Focus on batches 7-26 (chapters 7-12)
    target_batches = [i for i in range(7, 27)]
    
    print(f"Expanding chapters 7-12 and removing em dashes...")
    total_changes = 0
    
    for file in files:
        batch_num = int(file.split('Batch')[1].split('.')[0])
        if batch_num in target_batches:
            print(f"Processing {file}...", end=" ")
            try:
                changes = process_file(file, file)
                total_changes += changes
                print(f"✓ ({changes} changes)")
            except Exception as e:
                print(f"✗ Error: {e}")
    
    print(f"\n✓ Complete. {total_changes} total changes.")

if __name__ == "__main__":
    main()
