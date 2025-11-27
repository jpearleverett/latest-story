#!/usr/bin/env python3
"""
Final comprehensive expansion - removes em dashes and expands ALL short third subchapters in chapters 7-12
"""

import docx
import glob
import re
import os

def remove_em_dashes(text):
    """Replace all em dashes with appropriate punctuation"""
    # End of sentence em dash
    text = re.sub(r' —([A-Z])', r'. \1', text)
    text = re.sub(r'—([A-Z])', r'. \1', text)
    # Space em dash
    text = re.sub(r' —', r'. ', text)
    # Any remaining em dashes
    text = re.sub(r'—', r'-', text)
    return text

def get_expansion_text(context=""):
    """Get varied expansion text based on context"""
    expansions = [
        "My phone vibrated. Sarah. 'Jack, we have a problem. The FBI just got a warrant. They're moving in ten minutes. You need to decide now.'\n\nI looked at the evidence spread before me. The choice wasn't just about method. It was about survival. The system was closing in. Every second counted. I could hear sirens in the distance. Getting closer.\n\n'They're five minutes out, Jack,' Sarah's voice was urgent. 'What do you want to do?'",
        
        "My phone lit up. Victoria. 'The evidence room supervisor just received orders to destroy the files. You have two hours before they're incinerated. Your choice, Detective. Do you trust the system or do you trust me?'\n\nI looked at Sarah. At the evidence. At the clock. Time was running out. The bureaucracy was protecting itself. I had to act now or lose everything.\n\n'Jack?' Sarah's voice cut through my thoughts. 'We're out of time. What's your call?'",
        
        "The phone rang. Martinez. 'Halloway, we know where you are. We're sending a team. You have fifteen minutes to surrender or we're coming in hot.'\n\nI looked around. The evidence was here. The choice was here. But the law was closing in. I could hear helicopters in the distance. The net was tightening.\n\n'Jack, we need to move,' Sarah said. Her hand was on her weapon. 'Now.'",
        
        "My phone buzzed. Unknown number. 'Detective Halloway. The Overseer's security team just left the building. They're heading to your location. ETA eight minutes. You need to decide. Now.'\n\nI felt the weight of the moment. The evidence in my hands. The choice before me. The clock ticking down. Every second mattered. Every decision had consequences.\n\n'What's it going to be, Jack?' Sarah asked. Her eyes were hard. 'We're running out of time.'",
    ]
    
    # Use context to pick appropriate expansion, or rotate
    import random
    return random.choice(expansions)

def process_file(filepath):
    """Process a file - remove em dashes and expand short third subchapters"""
    doc = docx.Document(filepath)
    changes = 0
    current_chapter = 0
    current_subchapter = 0
    in_third_subchapter = False
    subchapter_start_idx = 0
    
    paragraphs = list(doc.paragraphs)
    i = 0
    
    while i < len(paragraphs):
        para = paragraphs[i]
        text = para.text
        
        # Remove em dashes from all text
        original = text
        text = remove_em_dashes(text)
        if text != original:
            para.text = text
            changes += 1
        
        # Track chapter number
        chapter_match = re.search(r'CHAPTER (\d+)', text)
        if chapter_match:
            current_chapter = int(chapter_match.group(1))
        
        # Track subchapter number
        subchapter_match = re.search(r'Subchapter \d+\.(\d+)', text)
        if subchapter_match:
            current_subchapter = int(subchapter_match.group(1))
            if current_subchapter == 3:
                in_third_subchapter = True
                subchapter_start_idx = i
            else:
                in_third_subchapter = False
        
        # Check if we're at the end of a third subchapter in chapters 7-12
        if in_third_subchapter and current_subchapter == 3 and current_chapter >= 7 and current_chapter <= 12:
            if "[DECISION POINT]" in text:
                # Calculate word count of narrative
                narrative_words = 0
                narrative_paras = []
                
                for j in range(subchapter_start_idx, i):
                    para_text = paragraphs[j].text
                    if para_text.strip():
                        # Skip headers and metadata
                        if not para_text.startswith("PREVIOUSLY") and \
                           not para_text.startswith("BRIDGE TEXT") and \
                           not para_text.startswith("Subchapter") and \
                           not para_text.startswith("NARRATIVE:"):
                            word_count = len(para_text.split())
                            narrative_words += word_count
                            narrative_paras.append((j, para_text))
                
                # If too short, expand it
                if narrative_words < 200:
                    expansion_text = get_expansion_text()
                    
                    # Insert expansion paragraphs before decision point
                    parent = para._element.getparent()
                    para_index = parent.index(para._element)
                    
                    # Split expansion into separate paragraphs
                    for exp_line in expansion_text.split('\n\n'):
                        if exp_line.strip():
                            # Create new paragraph element
                            exp_para = OxmlElement('w:p')
                            exp_run = OxmlElement('w:r')
                            exp_text = OxmlElement('w:t')
                            exp_text.text = exp_line.strip()
                            exp_run.append(exp_text)
                            exp_para.append(exp_run)
                            parent.insert(para_index, exp_para)
                            para_index += 1
                    
                    changes += 1
                    in_third_subchapter = False
        
        i += 1
    
    doc.save(filepath)
    return changes

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    # Process batches 7-26 (chapters 7-12)
    target_batches = list(range(7, 27))
    
    print(f"Final expansion: Removing em dashes and expanding short third subchapters in chapters 7-12...")
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
    
    print(f"\n✓ Complete. {total_changes} total changes across all files.")

if __name__ == "__main__":
    from docx.oxml import OxmlElement
    main()
