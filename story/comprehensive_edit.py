#!/usr/bin/env python3
"""
Comprehensive editor for The Midnight Confessor - applies all critical fixes
"""

import docx
import glob
import re
import os
from docx.shared import Pt

def edit_paragraph_text(text):
    """Apply all text edits to a paragraph"""
    if not text.strip():
        return text
    
    # Fix purple prose in bridge texts
    replacements = {
        # Bridge text improvements
        r"^BRIDGE TEXT: The elevator rises to the penthouse\. The detective carries the weight of betrayal in his pocket\.$":
        "BRIDGE TEXT: The elevator hummed. The flash drive felt heavier than it should.",
        
        r"^BRIDGE TEXT: The corridors of power are quiet\. A lone detective walks past the secretaries\. The Queen waits in her castle\.$":
        "BRIDGE TEXT: Helen's office was empty except for her. She was on the phone, voice tight.",
        
        r"^BRIDGE TEXT: The prison stands in the rain\. A monument to justice or a warehouse for mistakes\. Jack enters the belly of the beast\.$":
        "BRIDGE TEXT: Greystone Correctional. Concrete walls. Steel doors. I walked in.",
        
        # Purple prose fixes
        r"The air tasted of disinfectant and despair\.":
        "The air stank of bleach and old fear.",
        
        r"Nothing remained but the ghost of perfume\. French\. Expensive\. It hung in the air like an accusation\.":
        "French perfume. Expensive. It lingered.",
        
        r"My reflection in the window looked older than the time\.":
        "I looked in the window. The face looking back was a stranger's.",
        
        r"The bourbon in my stomach turned acidic\.":
        "The whiskey soured in my gut.",
        
        # Fix weak cliffhangers - Subchapter 3.3AA
        r"I had the complete file\. The Queen was ours\. But now I had a choice\. The convergence point was here\.":
        "I had the complete file. The Queen was ours. But my phone vibrated. Sarah. 'Jack, Helen's security just called the FBI. They're five minutes out. You need to decide now—do we burn this down publicly or do we squeeze her for the name before they arrest us both?'",
        
        # Fix weak cliffhanger - Subchapter 3.3AB  
        r"I had the Queen\. And the entire conspiracy\. But now I had the final decision before the convergence\.":
        "I had the Queen. And the entire conspiracy. But my phone lit up. Victoria. 'Helen's lawyer just filed a motion to seal all evidence. The judge signs it in one hour. Your choice, Detective—public confession now or lose everything.'",
        
        # Fix weak cliffhanger - Subchapter 3.3BA
        r"I had the Queen\. The black ledger\. The name of my best friend\. The reckless path had delivered the truth\. But now I had the final decision before the convergence\.":
        "I had the Queen. The black ledger. The name of my best friend. The reckless path had delivered the truth. But my phone rang. Sarah. 'Jack, Internal Affairs is raiding your office. They found Silas. You have ten minutes before they charge you with obstruction. What do you want to do?'",
        
        # Remove repetitive "weight of"
        r"the weight of ([^.]+)\.([^.]+)":
        r"\1. \2",
        
        # Vary "My phone buzzed"
    }
    
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text, flags=re.MULTILINE)
    
    # Vary phone notifications (simple rotation)
    phone_count = [0]
    def replace_phone(match):
        variations = ["My phone rang", "My phone vibrated", "My phone lit up"]
        result = variations[phone_count[0] % len(variations)]
        phone_count[0] += 1
        return result + "."
    
    text = re.sub(r"My phone buzzed\.", replace_phone, text)
    
    return text

def edit_file(input_path, output_path):
    """Edit a single docx file"""
    doc = docx.Document(input_path)
    
    for para in doc.paragraphs:
        if para.text.strip():
            original = para.text
            edited = edit_paragraph_text(original)
            if edited != original:
                para.text = edited
    
    doc.save(output_path)
    return True

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    # Create edited directory
    os.makedirs("edited", exist_ok=True)
    
    print(f"Editing {len(files)} files...")
    
    for i, file in enumerate(files, 1):
        print(f"[{i}/{len(files)}] Processing {file}...")
        output_path = f"edited/{file}"
        try:
            edit_file(file, output_path)
            print(f"  ✓ Saved to {output_path}")
        except Exception as e:
            print(f"  ✗ Error: {e}")
    
    print("\nAll files processed. Edited versions in 'edited/' directory.")

if __name__ == "__main__":
    main()
