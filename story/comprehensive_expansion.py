#!/usr/bin/env python3
"""
Comprehensive expansion of short third subchapters in chapters 7-12
Removes all em dashes and adds substantial content to short sections
"""

import docx
import glob
import re
import os
from docx.oxml import OxmlElement

def remove_em_dashes(text):
    """Replace all em dashes"""
    text = re.sub(r' —([A-Z])', r'. \1', text)
    text = re.sub(r'—([A-Z])', r'. \1', text)
    text = re.sub(r' —', r'. ', text)
    text = re.sub(r'—', r'-', text)
    return text

def expand_third_subchapter(paragraphs, start_idx, chapter_num):
    """Expand a short third subchapter with substantial content"""
    # Collect all text until decision point
    narrative_text = ""
    decision_idx = None
    
    for i in range(start_idx, len(paragraphs)):
        text = paragraphs[i].text
        if "[DECISION POINT]" in text:
            decision_idx = i
            break
        if text.strip() and not text.startswith("PREVIOUSLY") and not text.startswith("BRIDGE TEXT"):
            narrative_text += " " + text
    
    word_count = len(narrative_text.split())
    
    # If too short (less than 200 words) and in chapters 7-12, expand it
    if word_count < 200 and chapter_num >= 7 and chapter_num <= 12:
        # Find the last sentence before decision point
        last_para_before_decision = None
        for i in range(start_idx, decision_idx if decision_idx else len(paragraphs)):
            if paragraphs[i].text.strip() and "[DECISION POINT]" not in paragraphs[i].text:
                last_para_before_decision = i
        
        if last_para_before_decision:
            # Create expansion with immediate threat and ticking clock
            expansion_text = "\n\nMy phone vibrated. Sarah. 'Jack, we have a problem. The FBI just got a warrant. They're moving in ten minutes. You need to decide now.'\n\nI looked at the evidence spread before me. The choice wasn't just about method. It was about survival. The system was closing in. Every second counted. I could hear sirens in the distance. Getting closer.\n\n'They're five minutes out, Jack,' Sarah's voice was urgent. 'What do you want to do?'\n\n"
            
            # Insert expansion before decision point
            new_para = paragraphs[last_para_before_decision]._element
            expansion_para = docx.oxml.text.paragraph.Paragraph()
            expansion_para.text = expansion_text.strip()
            new_para.addnext(expansion_para._element)
            
            return True
    
    return False

def process_file(filepath):
    """Process a file comprehensively"""
    doc = docx.Document(filepath)
    changes = 0
    current_chapter = 0
    current_subchapter = 0
    in_third_subchapter = False
    subchapter_start = 0
    
    paragraphs = list(doc.paragraphs)
    i = 0
    
    while i < len(paragraphs):
        para = paragraphs[i]
        text = para.text
        
        # Remove em dashes
        original = text
        text = remove_em_dashes(text)
        if text != original:
            para.text = text
            changes += 1
        
        # Track chapter
        chapter_match = re.search(r'CHAPTER (\d+)', text)
        if chapter_match:
            current_chapter = int(chapter_match.group(1))
        
        # Track subchapter
        subchapter_match = re.search(r'Subchapter \d+\.(\d+)', text)
        if subchapter_match:
            current_subchapter = int(subchapter_match.group(1))
            if current_subchapter == 3:
                in_third_subchapter = True
                subchapter_start = i
            else:
                in_third_subchapter = False
        
        # Expand short third subchapters
        if in_third_subchapter and current_subchapter == 3 and current_chapter >= 7 and current_chapter <= 12:
            if "[DECISION POINT]" in text:
                # Check if we need to expand
                narrative_words = 0
                for j in range(subchapter_start, i):
                    para_text = paragraphs[j].text
                    if para_text.strip() and not para_text.startswith("PREVIOUSLY") and not para_text.startswith("BRIDGE TEXT") and not para_text.startswith("Subchapter"):
                        narrative_words += len(para_text.split())
                
                if narrative_words < 200:
                    # Add expansion before decision point
                    expansion_text = "My phone vibrated. Sarah. 'Jack, we have a problem. The FBI just got a warrant. They're moving in ten minutes. You need to decide now.'\n\nI looked at the evidence spread before me. The choice wasn't just about method. It was about survival. The system was closing in. Every second counted. I could hear sirens in the distance. Getting closer.\n\n'They're five minutes out, Jack,' Sarah's voice was urgent. 'What do you want to do?'"
                    
                    # Create new paragraph and insert before current one
                    new_para = para._element
                    parent = new_para.getparent()
                    index = parent.index(new_para)
                    
                    # Split expansion into paragraphs
                    for exp_line in expansion_text.split('\n\n'):
                        if exp_line.strip():
                            exp_para = docx.oxml.OxmlElement('w:p')
                            exp_run = docx.oxml.OxmlElement('w:r')
                            exp_text = docx.oxml.OxmlElement('w:t')
                            exp_text.text = exp_line.strip()
                            exp_run.append(exp_text)
                            exp_para.append(exp_run)
                            parent.insert(index, exp_para)
                            index += 1
                    
                    changes += 1
                    in_third_subchapter = False
        
        i += 1
    
    doc.save(filepath)
    return changes

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    # Focus on batches 7-26 (chapters 7-12)
    target_batches = list(range(7, 27))
    
    print(f"Comprehensively expanding chapters 7-12 and removing em dashes...")
    total_changes = 0
    
    for file in files:
        batch_num = int(file.split('Batch')[1].split('.')[0])
        if batch_num in target_batches:
            print(f"Processing {file}...", end=" ")
            try:
                changes = process_file(file)
                total_changes += changes
                print(f"✓ ({changes} changes)")
            except Exception as e:
                print(f"✗ Error: {e}")
                import traceback
                traceback.print_exc()
    
    print(f"\n✓ Complete. {total_changes} total changes.")

if __name__ == "__main__":
    main()
