#!/usr/bin/env python3
"""
Comprehensive editor - fixes all critical issues identified in analysis
"""

import docx
import glob
import re
import os

# Comprehensive replacement dictionary
REPLACEMENTS = {
    # Purple prose in bridge texts
    r"^BRIDGE TEXT: The city sleeps but guilt never does\. A detective waits for a ghost in a bottle\.$":
    "BRIDGE TEXT: 2:47 AM. The city sleeps. I don't.",
    
    r"^BRIDGE TEXT: The prison stands in the rain\. A monument to justice or a warehouse for mistakes\. Jack enters the belly of the beast\.$":
    "BRIDGE TEXT: Greystone Correctional. Concrete walls. Steel doors.",
    
    r"^BRIDGE TEXT: The estate door is open\. A crime scene waiting for a detective\. A choice lies on the desk\.$":
    "BRIDGE TEXT: The front door stood open. Inside, a choice.",
    
    r"^BRIDGE TEXT: A greasy spoon at dawn\. The smell of burnt coffee and resentment\. The daughter waits\.$":
    "BRIDGE TEXT: 6 AM. The Blueline Diner. Burnt coffee and old anger.",
    
    r"^BRIDGE TEXT: The elevator rises to the penthouse\. The detective carries the weight of betrayal in his pocket\.$":
    "BRIDGE TEXT: The elevator hummed. The flash drive felt heavier than it should.",
    
    r"^BRIDGE TEXT: A partner in cuffs\. A girl in danger\. The law demands arrest but survival demands action\.$":
    "BRIDGE TEXT: Silas held out his wrists. Maya was missing. The clock was ticking.",
    
    r"^BRIDGE TEXT: The diner is a powder keg\. The phone buzzes\. A choice between a trap and an assault\.$":
    "BRIDGE TEXT: The diner was tense. My phone kept ringing. Sarah. Victoria. Both wanted answers.",
    
    r"^BRIDGE TEXT: The corridors of power are quiet\. A lone detective walks past the secretaries\. The Queen waits in her castle\.$":
    "BRIDGE TEXT: Helen's office was empty except for her. She was on the phone, voice tight.",
    
    r"^BRIDGE TEXT: The District Attorney's office\. The phone rings unanswered\. The endgame begins\.$":
    "BRIDGE TEXT: Helen's office. She saw me coming. Too late to run.",
    
    r"^BRIDGE TEXT: A slammed door\. A receptionist ignored\. The Queen is interrupted\.$":
    "BRIDGE TEXT: I locked the door behind me. Helen's face went white.",
    
    r"^BRIDGE TEXT: The DA's office\. No more appointments\. The Vigilante arrives\.$":
    "BRIDGE TEXT: I shoved past security. Helen saw the blood on my hands.",
    
    # Purple prose in narrative
    r"The air tasted of disinfectant and despair\.":
    "The air stank of bleach and old fear.",
    
    r"Nothing remained but the ghost of perfume\. French\. Expensive\. It hung in the air like an accusation\.":
    "French perfume. Expensive. It lingered.",
    
    r"My reflection in the window looked older than the time\.":
    "I looked in the window. The face looking back was a stranger's.",
    
    r"The bourbon in my stomach turned acidic\.":
    "The whiskey soured in my gut.",
    
    r"carried the collective weight of our evidence":
    "carried our evidence",
    
    r"carried the weight of":
    "carried",
    
    # Weak cliffhangers - need immediate threats
    r"I had the complete file\. The Queen was ours\. But now I had a choice\. The convergence point was here\.":
    "I had the complete file. The Queen was ours. But my phone vibrated. Sarah. 'Jack, Helen's security just called the FBI. They're five minutes out. You need to decide now—do we burn this down publicly or do we squeeze her for the name before they arrest us both?'",
    
    r"I had the Queen\. And the entire conspiracy\. But now I had the final decision before the convergence\.":
    "I had the Queen. And the entire conspiracy. But my phone lit up. Victoria. 'Helen's lawyer just filed a motion to seal all evidence. The judge signs it in one hour. Your choice, Detective—public confession now or lose everything.'",
    
    r"I had the Queen\. The black ledger\. The name of my best friend\. The reckless path had delivered the truth\. But now I had the final decision before the convergence\.":
    "I had the Queen. The black ledger. The name of my best friend. The reckless path had delivered the truth. But my phone rang. Sarah. 'Jack, Internal Affairs is raiding your office. They found Silas. You have ten minutes before they charge you with obstruction. What do you want to do?'",
    
    r"I had the Queen\. The conspiracy\. The name of my best friend\. The reckless path had delivered the truth faster than any warrant\. But now I had the choice that would define whether I was still a cop or just a monster with a badge\.":
    "I had the Queen. The conspiracy. The name of my best friend. The reckless path had delivered the truth faster than any warrant. But my phone buzzed. Sarah. 'Jack, the FBI just got a warrant for your arrest. They're at your apartment. You have five minutes to decide—turn yourself in or become a fugitive.'",
    
    r"I knew she was right\. But I also knew the system was slow and Victoria was fast\. I was trapped\. I had to choose the slow legal path to contain the damage from my aggressive actions or risk everything for the speed that Victoria had cultivated\.":
    "I knew she was right. But my phone lit up. Victoria. 'The evidence room supervisor just received orders to destroy the Sullivan files. You have two hours before they're incinerated. Your choice, Detective—do you trust the system or do you trust me?'",
    
    r"I looked at Sarah\. The impatience in Victoria's warning was valid\. \"We need to move fast\. Victoria is right\. The bureaucracy will protect itself\. They will shred the evidence files to prevent further exonerations\.\"":
    "I looked at Sarah. My phone vibrated. Victoria. 'The records department just received orders to destroy the Sullivan and Wade files. They're being shredded in one hour. Your choice—trust the system or trust me.'",
    
    # Remove "weight of" variations
    r"the weight of ([^.]+)\.([^.]+)":
    r"\1. \2",
    
    r"the unstoppable weight of":
    "the",
    
    # Vary phone notifications
}

def edit_text(text):
    """Apply all edits to text"""
    if not text.strip():
        return text
    
    # Apply all replacements
    for pattern, replacement in REPLACEMENTS.items():
        text = re.sub(pattern, replacement, text, flags=re.MULTILINE)
    
    # Vary "My phone buzzed" - simple rotation
    phone_variations = ["My phone rang", "My phone vibrated", "My phone lit up"]
    phone_count = 0
    def replace_phone(match):
        nonlocal phone_count
        result = phone_variations[phone_count % len(phone_variations)]
        phone_count += 1
        return result + "."
    
    text = re.sub(r"My phone buzzed\.", replace_phone, text)
    
    return text

def edit_file(input_path, output_path):
    """Edit a docx file"""
    doc = docx.Document(input_path)
    
    for para in doc.paragraphs:
        if para.text.strip():
            original = para.text
            edited = edit_text(original)
            if edited != original:
                para.text = edited
    
    doc.save(output_path)
    return True

def main():
    files = sorted(glob.glob("The_Midnight_Confessor_Batch*.docx"), 
                   key=lambda x: int(x.split('Batch')[1].split('.')[0]))
    
    os.makedirs("edited_v2", exist_ok=True)
    
    print(f"Comprehensively editing {len(files)} files...")
    
    for i, file in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {file}...")
        output_path = f"edited_v2/{file}"
        try:
            edit_file(file, output_path)
            print(f"  ✓ Done")
        except Exception as e:
            print(f"  ✗ Error: {e}")
    
    print(f"\n✓ All files edited. Check 'edited_v2/' directory.")

if __name__ == "__main__":
    main()
