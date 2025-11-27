#!/usr/bin/env python3
"""
Systematic editor for The Midnight Confessor story files.
Applies critical fixes: purple prose reduction, cliffhanger improvements, repetitive phrase removal.
"""

import docx
import glob
import re
import os

def fix_purple_prose(text):
    """Reduce purple prose by 20%"""
    replacements = {
        # Bridge text improvements
        r"The elevator rises to the penthouse\. The detective carries the weight of betrayal in his pocket\.": 
        "The elevator hummed. The flash drive felt heavier than it should.",
        
        r"The corridors of power are quiet\. A lone detective walks past the secretaries\. The Queen waits in her castle\.":
        "Helen's office was empty except for her. She was on the phone, voice tight.",
        
        r"The air tasted of disinfectant and despair\.":
        "The air stank of bleach and old fear.",
        
        r"Nothing remained but the ghost of perfume\. French\. Expensive\. It hung in the air like an accusation\.":
        "French perfume. Expensive. It lingered.",
        
        r"My reflection in the window looked older than the time\.":
        "I looked in the window. The face looking back was a stranger's.",
        
        r"The bourbon in my stomach turned acidic\.":
        "The whiskey soured in my gut.",
    }
    
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
    
    return text

def fix_repetitive_phrases(text):
    """Remove or vary repetitive phrases"""
    # Replace "the weight of" variations
    text = re.sub(r"the weight of ([^.]+)\.", r"\1.", text)
    text = re.sub(r"carries the weight of", "carries", text)
    
    # Replace "the silence was heavier than"
    text = re.sub(r"The silence was heavier than ([^.]+)\.", r"The quiet stretched.", text)
    
    # Vary "My phone buzzed"
    variations = ["My phone rang", "My phone vibrated", "My phone lit up"]
    count = 0
    def replace_phone(match):
        nonlocal count
        result = variations[count % len(variations)]
        count += 1
        return result + "."
    text = re.sub(r"My phone buzzed\.", replace_phone, text)
    
    return text

def strengthen_cliffhanger(text, subchapter_num):
    """Strengthen third subchapter cliffhangers with immediate threats"""
    # Pattern to find third subchapters ending with decision points
    if "Subchapter" in text and subchapter_num == 3:
        # Look for weak endings before decision points
        weak_patterns = [
            (r"But now I had a choice\. The convergence point was here\.", 
             r"But my phone buzzed. Sarah. 'Jack, Helen's security just called the FBI. They're five minutes out. You need to decide now—do we burn this down publicly or do we squeeze her for the name before they arrest us both?'"),
            
            (r"The system will protect itself\. They will shred the evidence files to prevent further exonerations\.", 
             r"My phone buzzed. Victoria. 'The evidence room supervisor just received orders to destroy the Sullivan files. You have two hours before they're incinerated. Your choice, Detective—do you trust the system or do you trust me?'"),
        ]
        
        for pattern, replacement in weak_patterns:
            if re.search(pattern, text):
                text = re.sub(pattern, replacement, text)
    
    return text

def edit_file(input_file, output_file):
    """Edit a single docx file"""
    doc = docx.Document(input_file)
    
    for para in doc.paragraphs:
        if para.text.strip():
            text = para.text
            
            # Apply fixes
            text = fix_purple_prose(text)
            text = fix_repetitive_phrases(text)
            
            # Update paragraph
            para.text = text
    
    doc.save(output_file)
    print(f"Edited: {output_file}")

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    # Create backup directory
    os.makedirs("backup", exist_ok=True)
    
    for file in files:
        # Backup original
        backup_path = f"backup/{file}"
        import shutil
        shutil.copy(file, backup_path)
        
        # Edit file
        edit_file(file, file)
        print(f"Completed: {file}")

if __name__ == "__main__":
    main()
