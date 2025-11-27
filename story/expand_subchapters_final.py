#!/usr/bin/env python3
"""
Final script to expand all subchapters in The Midnight Confessor docx files.
"""

from docx import Document
import shutil
import re

def expand_subchapter(doc, subchapter_num, old_narrative_start, new_narrative_text):
    """Find and expand a subchapter's narrative section"""
    
    found_subchapter = False
    narrative_start_idx = None
    decision_point_idx = None
    
    # Find the subchapter
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        if f'Subchapter {subchapter_num}' in text:
            found_subchapter = True
            continue
        
        if found_subchapter and 'NARRATIVE:' in text and old_narrative_start[:30] in text:
            narrative_start_idx = i
            # Get the full original narrative by reading until [DECISION POINT]
            continue
        
        if narrative_start_idx is not None and '[DECISION POINT]' in text:
            decision_point_idx = i
            break
    
    if narrative_start_idx is None:
        print(f"ERROR: Could not find Subchapter {subchapter_num}")
        return False
    
    if decision_point_idx is None:
        print(f"ERROR: Could not find [DECISION POINT] for Subchapter {subchapter_num}")
        return False
    
    print(f"Found Subchapter {subchapter_num}: narrative at {narrative_start_idx}, decision at {decision_point_idx}")
    
    # Split new narrative into paragraphs
    new_paras = [p.strip() for p in new_narrative_text.split('\n\n') if p.strip()]
    
    # Replace the first narrative paragraph
    doc.paragraphs[narrative_start_idx].text = new_paras[0]
    
    # Insert remaining paragraphs before the decision point
    insert_idx = narrative_start_idx + 1
    for para_text in new_paras[1:]:
        # Find the paragraph before decision point to insert before
        target_para = doc.paragraphs[decision_point_idx - 1]
        new_para = target_para.insert_paragraph_before(para_text)
        decision_point_idx += 1  # Decision point moved down
    
    print(f"Successfully expanded Subchapter {subchapter_num}")
    return True

# Expanded content for Subchapter 2.3A
expanded_2_3A = """NARRATIVE: "We need to find her. Now." I grabbed Silas by the arm. The time for deliberation was over.

Silas looked at the flash drive then at me. His eyes were red-rimmed and desperate. The weight of his confession hung between us like a shroud. "You have the evidence, Jack. You have me cold. You can arrest me right now. Be the hero again. It would be clean. It would be the right thing to do."

He held out his wrists. The gesture was theatrical but genuine. I looked at the cuffs on my belt. Cold steel. The tools of justice. Then I looked at the text about Maya. A girl I had never met. A girl whose mother I had put away. A girl who was now in the hands of a woman who had spent years planning this moment.

The balcony air was cold. It cut through my coat and found the old wounds. My phone buzzed again. Sarah. Another text. The FBI is mobilizing. They want Silas in custody. They want the flash drive. They want to do this the right way.

The right way. I had spent thirty years doing things the right way. Following procedure. Building cases. Trusting the system. And where had it gotten me? A partner who framed innocent men. A best friend who would turn out to be something worse. A city full of corruption I had been too blind to see.

Silas's hands were still extended. Waiting. The cuffs would click. The arrest would be clean. The conviction would be certain. And Maya Bellamy would be in the wind. Victoria would have her. The system would move too slow. The bureaucracy would grind. The paperwork would pile up. And somewhere in the city, a terrified girl would wait for a rescue that might come too late.

I looked at Silas. Really looked at him. The man I had trusted. The man who had betrayed that trust. The man who was now offering himself up as a sacrifice to the machine. "If I arrest you now, Silas, what happens to Maya?"

"She dies," he said. The words came out flat. No emotion. Just fact. "Victoria told me. She said if I was arrested before she was done with me, she would kill the girl. She said the system moves too slow. She said you would choose the law over the life. She was betting on it."

My hand moved to the cuffs. The metal was cold against my palm. The weight of thirty years of procedure. The weight of doing things the right way. The weight of being the good cop. The legend.

But legends were built on bodies. On cases closed. On convictions secured. And somewhere in those thirty years, I had stopped asking if the bodies were the right ones. I had stopped questioning if the convictions were true. I had become the machine. Efficient. Certain. Blind.

The Methodical Path demanded I secure the criminal and trust the system to find the girl. It was clean. It was legal. It was what Sarah would do. It was what the good cop would do. Secure the evidence. Build the case. Trust the process.

The Pragmatic path demanded I use the criminal to hunt the greater threat. It was dirty. It was dangerous. It was what a desperate man would do. Use the leverage. Save the girl. Ask forgiveness later.

I looked at Silas. At his extended wrists. At the flash drive on the glass table. At my phone with Sarah's text. At the city spread out below us, gray and cold and full of secrets.

The choice wasn't between right and wrong. It was between two kinds of wrong. The wrong of letting a girl die for the sake of procedure. The wrong of using a criminal to save her. Both would cost me something. Both would change who I was. But only one would let me sleep at night.

My hand moved away from the cuffs. I picked up the flash drive. I looked at Silas. "You're going to help me find her. You're going to use every contact you have. Every favor you're owed. Every dirty secret you know. And if you try to run, if you try to warn Victoria, I will make sure your family knows exactly what you did. I will make sure your sons know their father framed an innocent man. I will burn your life to the ground."

Silas's hands dropped. He nodded. "I understand."

"Do you? Because this isn't a negotiation. This is me using you. This is me becoming exactly what Victoria wants me to become. A man who breaks the rules. A man who uses criminals. A man who puts justice before the law."

"I understand," he repeated. "And I'm sorry, Jack. For everything."

The apology hung in the air between us. Empty. Meaningless. But it was all he had left to give. I turned away from him. I looked at the city. At the rain. At the gray expanse of water beyond the buildings. Somewhere out there, Maya Bellamy was waiting. And I was coming for her. With a criminal as my guide. With a flash drive as my weapon. With everything I had left to lose."""

# Expand Subchapter 2.3A
print("Expanding Subchapter 2.3A...")
doc = Document('The_Midnight_Confessor_Batch1.docx')
shutil.copy('The_Midnight_Confessor_Batch1.docx', 'The_Midnight_Confessor_Batch1_backup.docx')

expand_subchapter(doc, '2.3A', 'NARRATIVE: "We need to find her. Now."', expanded_2_3A)
doc.save('The_Midnight_Confessor_Batch1.docx')
print("Batch 1 updated successfully")
