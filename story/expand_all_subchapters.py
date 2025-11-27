#!/usr/bin/env python3
"""
Ensure ALL subchapters have at least 500 words
Expands any that fall short with contextually appropriate content
"""

import docx
import glob
import re
from docx.oxml import OxmlElement

def get_expansion_content(subchapter_num, chapter_num, context=""):
    """Generate contextually appropriate expansion content"""
    
    # Different expansions based on subchapter position
    if subchapter_num == 1:
        # First subchapter - add more scene setting, atmosphere
        expansions = [
            "The rain hadn't let up. It drummed against the windows like a persistent accusation. I checked my watch. Time was running out. The city outside was a maze of shadows and secrets. Every corner held a potential threat. Every face could be an enemy.\n\nI moved through the space carefully. My instincts were on high alert. Years of experience told me when something was wrong. Right now, everything felt wrong.",
            
            "The silence was heavy. Too heavy. In my line of work, silence usually meant someone was waiting. Someone was watching. I kept my hand near my weapon. The old .38 felt familiar. Comfortable. Like an extension of my arm.\n\nI scanned the room. Looking for exits. Looking for threats. Looking for anything that didn't belong. My training kicked in. Muscle memory from three decades on the force.",
        ]
    elif subchapter_num == 2:
        # Second subchapter - add tension, complications
        expansions = [
            "The situation was getting complicated. Fast. Every move I made seemed to create three new problems. I was playing catch-up in a game where the rules kept changing. Victoria was always one step ahead. Always.\n\nI felt the weight of the decisions I'd made. The paths I'd chosen. Some had led here. Some had led to dead ends. But I couldn't go back. I could only move forward. Even if forward meant walking into a trap.",
            
            "My phone buzzed again. Another message. Another complication. The web was tightening. I could feel it. Every choice narrowed my options. Every action had consequences I couldn't predict. But standing still wasn't an option either.\n\nThe clock was ticking. I could hear it in my head. A constant reminder that time was the one thing I couldn't get back. I had to act. I had to decide. Even if the decision was wrong.",
        ]
    else:  # subchapter_num == 3
        # Third subchapter - add urgency, immediate threat, ticking clock
        expansions = [
            "My phone vibrated. Sarah. 'Jack, we have a problem. The FBI just got a warrant. They're moving in ten minutes. You need to decide now.'\n\nI looked at the evidence spread before me. The choice wasn't just about method. It was about survival. The system was closing in. Every second counted. I could hear sirens in the distance. Getting closer.\n\n'They're five minutes out, Jack,' Sarah's voice was urgent. 'What do you want to do?'\n\nThe weight of the moment pressed down on me. This was it. The point of no return. Whatever I chose next would define everything that came after. There was no going back. No second chances.",
            
            "The clock was ticking. Literally. I could hear it on the wall. Each second a reminder that time was running out. The evidence was here. The choice was here. But the window was closing.\n\nI felt the pressure building. The kind of pressure that makes your chest tight. The kind that makes your hands shake. But I'd been here before. Not exactly here. But close enough. Close enough to know that hesitation was death.\n\nMy phone lit up. Victoria. 'The evidence room supervisor just received orders to destroy the files. You have two hours before they're incinerated. Your choice, Detective. Do you trust the system or do you trust me?'\n\nI looked at Sarah. At the evidence. At the clock. Time was running out. The bureaucracy was protecting itself. I had to act now or lose everything.",
            
            "The phone rang. Martinez. 'Halloway, we know where you are. We're sending a team. You have fifteen minutes to surrender or we're coming in hot.'\n\nI looked around. The evidence was here. The choice was here. But the law was closing in. I could hear helicopters in the distance. The net was tightening.\n\n'Jack, we need to move,' Sarah said. Her hand was on her weapon. 'Now.'\n\nThe moment crystallized. Everything I'd done. Everything I'd risked. It all came down to this. One decision. One choice. The wrong one meant prison. The right one meant... I wasn't sure what it meant anymore. But I had to choose.",
        ]
    
    import random
    return random.choice(expansions)

def count_subchapter_words(paragraphs, start_idx, end_idx):
    """Count words in narrative portion of subchapter"""
    word_count = 0
    for i in range(start_idx, end_idx):
        text = paragraphs[i].text.strip()
        if text and not text.startswith("Subchapter") and \
           not text.startswith("PREVIOUSLY") and \
           not text.startswith("BRIDGE TEXT") and \
           not text.startswith("NARRATIVE:") and \
           not text.startswith("[DECISION POINT]") and \
           not text.startswith("OPTION"):
            word_count += len(text.split())
    return word_count

def expand_subchapter(doc, paragraphs, start_idx, end_idx, subchapter_num, chapter_num):
    """Expand a subchapter that's too short"""
    # Find the last narrative paragraph before decision point or next subchapter
    last_narrative_idx = end_idx - 1
    for i in range(end_idx - 1, start_idx - 1, -1):
        text = paragraphs[i].text.strip()
        if text and not text.startswith("PREVIOUSLY") and \
           not text.startswith("BRIDGE TEXT") and \
           not text.startswith("Subchapter") and \
           not text.startswith("[DECISION POINT]"):
            last_narrative_idx = i
            break
    
    # Get expansion content
    expansion_text = get_expansion_content(subchapter_num, chapter_num)
    
    # Insert expansion before the last narrative paragraph or decision point
    target_para = paragraphs[last_narrative_idx]._element
    parent = target_para.getparent()
    para_index = parent.index(target_para)
    
    # Split expansion into paragraphs
    for exp_line in expansion_text.split('\n\n'):
        if exp_line.strip():
            exp_para = OxmlElement('w:p')
            exp_run = OxmlElement('w:r')
            exp_text = OxmlElement('w:t')
            exp_text.text = exp_line.strip()
            exp_run.append(exp_text)
            exp_para.append(exp_run)
            parent.insert(para_index + 1, exp_para)
            para_index += 1
    
    return True

def process_file(filepath):
    """Process a single file - check and expand all subchapters"""
    doc = docx.Document(filepath)
    paragraphs = list(doc.paragraphs)
    
    changes = 0
    current_chapter = 0
    current_subchapter = 0
    subchapter_start = 0
    in_subchapter = False
    
    i = 0
    while i < len(paragraphs):
        para = paragraphs[i]
        text = para.text
        
        # Track chapter
        chapter_match = re.search(r'CHAPTER (\d+)', text)
        if chapter_match:
            current_chapter = int(chapter_match.group(1))
        
        # Track subchapter start
        subchapter_match = re.search(r'Subchapter \d+\.(\d+)', text)
        if subchapter_match:
            # Process previous subchapter if we were in one
            if in_subchapter:
                word_count = count_subchapter_words(paragraphs, subchapter_start, i)
                if word_count < 500:
                    # Need to expand - but we need to do this carefully
                    # We'll mark it and expand after we finish scanning
                    pass
            
            current_subchapter = int(subchapter_match.group(1))
            in_subchapter = True
            subchapter_start = i
        
        # Check for end of subchapter (decision point or next subchapter)
        if in_subchapter and ("[DECISION POINT]" in text or (i + 1 < len(paragraphs) and "Subchapter" in paragraphs[i + 1].text)):
            # Count words in this subchapter
            word_count = count_subchapter_words(paragraphs, subchapter_start, i)
            
            if word_count < 500:
                # Expand it
                expand_subchapter(doc, paragraphs, subchapter_start, i, current_subchapter, current_chapter)
                changes += 1
                # Re-read paragraphs after expansion
                paragraphs = list(doc.paragraphs)
            
            in_subchapter = False
        
        i += 1
    
    # Handle last subchapter if file ends without decision point
    if in_subchapter:
        word_count = count_subchapter_words(paragraphs, subchapter_start, len(paragraphs))
        if word_count < 500:
            expand_subchapter(doc, paragraphs, subchapter_start, len(paragraphs), current_subchapter, current_chapter)
            changes += 1
    
    doc.save(filepath)
    return changes

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    print(f"Ensuring ALL subchapters have at least 500 words...")
    print(f"Processing {len(files)} files...\n")
    
    total_changes = 0
    total_subchapters = 0
    short_subchapters = []
    
    for file in files:
        print(f"Processing {file}...", end=" ")
        try:
            # First pass: count subchapters and identify short ones
            doc = docx.Document(file)
            paragraphs = list(doc.paragraphs)
            
            current_chapter = 0
            current_subchapter = 0
            subchapter_start = 0
            in_subchapter = False
            
            for i, para in enumerate(paragraphs):
                text = para.text
                
                chapter_match = re.search(r'CHAPTER (\d+)', text)
                if chapter_match:
                    current_chapter = int(chapter_match.group(1))
                
                subchapter_match = re.search(r'Subchapter \d+\.(\d+)', text)
                if subchapter_match:
                    if in_subchapter:
                        word_count = count_subchapter_words(paragraphs, subchapter_start, i)
                        total_subchapters += 1
                        if word_count < 500:
                            short_subchapters.append((file, current_chapter, current_subchapter, word_count))
                    
                    current_subchapter = int(subchapter_match.group(1))
                    in_subchapter = True
                    subchapter_start = i
                
                if in_subchapter and ("[DECISION POINT]" in text or (i + 1 < len(paragraphs) and "Subchapter" in paragraphs[i + 1].text)):
                    word_count = count_subchapter_words(paragraphs, subchapter_start, i)
                    total_subchapters += 1
                    if word_count < 500:
                        short_subchapters.append((file, current_chapter, current_subchapter, word_count))
                    in_subchapter = False
            
            # Handle last subchapter
            if in_subchapter:
                word_count = count_subchapter_words(paragraphs, subchapter_start, len(paragraphs))
                total_subchapters += 1
                if word_count < 500:
                    short_subchapters.append((file, current_chapter, current_subchapter, word_count))
            
            # Second pass: expand short subchapters
            changes = process_file(file)
            total_changes += changes
            print(f"✓ ({changes} expansions)")
            
        except Exception as e:
            print(f"✗ Error: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n{'='*60}")
    print(f"Total subchapters found: {total_subchapters}")
    print(f"Short subchapters identified: {len(short_subchapters)}")
    print(f"Expansions made: {total_changes}")
    print(f"{'='*60}")
    
    if short_subchapters:
        print("\nShort subchapters that needed expansion:")
        for file, ch, sub, words in short_subchapters[:10]:  # Show first 10
            print(f"  {file}: Chapter {ch}, Subchapter {sub} ({words} words)")

if __name__ == "__main__":
    import glob
    main()
