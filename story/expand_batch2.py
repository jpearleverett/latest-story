#!/usr/bin/env python3
"""
Script to expand all Batch 2 subchapters.
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
        
        if found_subchapter and 'NARRATIVE:' in text and old_narrative_start[:40] in text:
            narrative_start_idx = i
            continue
        
        if narrative_start_idx is not None and '[DECISION POINT]' in text:
            decision_point_idx = i
            break
        elif narrative_start_idx is not None and 'Subchapter' in text and i > narrative_start_idx + 2:
            decision_point_idx = i
            break
    
    if narrative_start_idx is None:
        print(f"ERROR: Could not find Subchapter {subchapter_num}")
        return False
    
    if decision_point_idx is None:
        decision_point_idx = len(doc.paragraphs)
    
    print(f"Found Subchapter {subchapter_num}: narrative at {narrative_start_idx}, end at {decision_point_idx}")
    
    # Split new narrative into paragraphs
    new_paras = [p.strip() for p in new_narrative_text.split('\n\n') if p.strip()]
    
    if not new_paras:
        print(f"ERROR: No paragraphs in expanded text for {subchapter_num}")
        return False
    
    # Replace the first narrative paragraph
    doc.paragraphs[narrative_start_idx].text = new_paras[0]
    
    # Insert remaining paragraphs before the decision point
    for para_text in new_paras[1:]:
        if decision_point_idx < len(doc.paragraphs):
            target_para = doc.paragraphs[decision_point_idx - 1]
            new_para = target_para.insert_paragraph_before(para_text)
            decision_point_idx += 1
        else:
            doc.add_paragraph(para_text)
    
    print(f"Successfully expanded Subchapter {subchapter_num} ({len(new_paras)} paragraphs)")
    return True

# Expanded content for Batch 2 subchapters
# Due to length, I'll create key expansions and we can iterate

expansions_batch2 = {
    '3.1AA': {
        'start': 'NARRATIVE: I sat across from Sarah',
        'text': '''NARRATIVE: I sat across from Sarah in her small sterile apartment. The only noise was the relentless hiss of the old radiator fighting the damp Ashport chill. Silas Reed was in federal holding. He had traded his expensive silk robe for penitentiary orange the moment Maya Bellamy was safe.

The apartment felt like a war room. Files were spread across every surface. Coffee cups littered the table. The air smelled of stale coffee and exhaustion. We had been at this for hours. Days. It felt like years.

"The numbers don't lie Jack," Sarah said. She traced a line on the Thornhill Ledger printouts she had been analyzing for eight hours. Her eyes were red from lack of sleep. Her hands were steady. Methodical. "Silas's funnel account paid out two organizations consistently. Webb Curiosities for operational cover and P&A Consulting. Price and Associates."

I looked at the numbers. The patterns. The money trail. It was all there. Clear as day. Clear as corruption. "Helen Price," I muttered. The name tasted like copper. Like blood. "The District Attorney. Her family firm handles the defense for the elite. She prosecuted all five of our innocent victims. Fifty-three wins. Zero losses. She was the golden girl. The untouchable."

"If her family is bankrolling the corruption," Sarah said, "then her record is bought and paid for. The foundation of Ashport's justice system is built on a shell company. This isn't just one bad cop. It is the institutional architecture. The whole system is rotten."

I stood up and paced the small room. The walls felt like they were closing in. The truth was too big. Too ugly. "I need to verify this link to specific cases. The ledger shows payments but I need direct orders. I need proof that she was giving the commands. Not just receiving the money."

I looked at the old cracked leather jacket hanging on her coat rack. A reminder of who we used to be. Who we still were. "The condemned precinct on Harbor Street. You are the only one who can go in clean. I'm too compromised. Too visible. But you? You can still move in the shadows."

Sarah didn't break eye contact. Her gaze was steady. Determined. "I pulled files from the Harbor Street precinct an hour ago. I signed out the condemned box under a fake internal transfer order citing historical preservation. Clean hands Jack. Even in the dirtiest place."

She slid a water-damaged folder across the table. The paper was yellowed. The edges were frayed. It looked like evidence. Real evidence. The kind that couldn't be faked. Inside were photocopies of handwritten memos from Assistant D.A. Helen Price to Silas Reed. They weren't explicit orders but the context was damningly clear. The notes were terse arrogant commands signed H.P. The fading ink smelled faintly of perfume and malice.

Reed. Witness A in the Chen case is unreliable. Remove her from the docket immediately. Have Tom adjust his forensics on the transfer date. Urgent.

Reed. Bellamy file. The timeline still needs revision. Ensure the secondary purchase order links directly to the safe deposit. Eliminate all testimony regarding a woman in red.

"She wasn't just prosecuting the cases," I said. I shone my tactical flashlight on the fading ink. The words were clear. The intent was clear. The corruption was clear. "She was the conductor. She told Silas which evidence to plant. She told him which witnesses to bury. She told him what forensic results to request from the lab. She was running the whole operation."

The memos confirmed the full chain of command. Helen Price was the architect. The mastermind. The Queen. The revelation settled heavy in my gut. The rot went higher than the badge. Higher than I had ever imagined. It went all the way to the top.'''
    },
    
    '3.2AA': {
        'start': 'NARRATIVE: We executed the next step',
        'text': '''NARRATIVE: We executed the next step with surgical precision. We hit Marcus Webb's shop together. Two methodical investigators armed with irrefutable documents. The shop was hushed. It was filled with artifacts that outlived their owners. It held secrets the city had forgotten.

The bell on the door chimed as we entered. The sound was soft. Almost gentle. It didn't match the tension in the air. The shop smelled of old wood and polish. Of history. Of secrets kept too long.

Webb was behind the counter. He was polishing a brass telescope. His movements were slow. Deliberate. The movements of a man trying to keep his hands steady. Trying to keep his fear hidden. He looked up when we entered. His face went pale. He knew why we were there.

Webb paled when he saw the Price-Reed Memos. His name appeared twice in the margins confirming cash drop-offs and witness coercion. He looked like an expensive antique about to shatter under its own weight. The facade was cracking. The mask was slipping.

"Helen promised me silence," Webb whimpered. He ran his hands through his thin gray hair. The gesture was nervous. Desperate. "She promised my secret life with Richard would never surface in public. She guaranteed that the framing of Eleanor would be the end of it. She said I would be safe. Protected."

He looked desperately at the silent clock on the wall. As if time could save him. As if the past could protect him from the present.

"She is a blackmailer Marcus," Sarah cut in. Her voice was cold and professional. She offered no quarter. No mercy. "She collects debts. And you owe Eleanor Bellamy eight years of her life. Silas has already confessed to the framework. You are next on the indictment unless you cooperate. You help us, or you go down with her."

"I loved Richard," he protested. His face collapsed. The mask fell away completely. "It was the only truth in my life. The only real thing. Everything else was a lie. But that? That was real."

"Then help us avenge him," I insisted. I placed the affidavit on the counter. The paper was crisp. Official. Final. "Sign this affidavit. Detail Helen Price's explicit knowledge of the blackmail against Richard and her direct orders to cover up the Bellamy murder. You clear Eleanor's name or you go down with Helen. Your secret gets exposed either way. But this way, you do some good. You help us bring down the real monster."

Webb looked at the methodical stack of evidence. He saw the unstoppable weight of the legal machine we had assembled. He saw his options. His choices. His future. He signed the affidavit. His hand shook violently. The pen scratched his confession into history. Into evidence. Into truth.

Webb's Affidavit was airtight and notarized by a nervous clerk Sarah had brought. This wasn't chaos. It was procedure. It was the law. It was justice. Or as close as we could get to it.

"Victoria sent a final instruction," Webb whispered. He looked over my shoulder desperate for relevance. Desperate to matter. "She said if you caught the Queen using clean hands she would be amused. She also said to ask Helen about a case file she keeps locked away. It is called The Insurance Policy. A black ledger of everyone who ever cooperated. The list of all the true untouchables who supplied the evidence. The real architects of the corruption."'''
    },
    
    '3.3AA': {
        'start': 'NARRATIVE: I walked into Helen Price\'s office alone',
        'text': '''NARRATIVE: I walked into Helen Price's office alone. I carried the collective weight of our evidence. Sarah waited outside ready to call the FBI on a prearranged signal. It was a moment of professional theater where the clean hands waited for the dirty work to be done.

The office was opulent. Expensive. The kind of office that screamed power. The kind of office that belonged to someone who had never questioned their right to it. The walls were lined with awards. With photos. With the evidence of a career built on lies.

Helen was on the phone. Her voice was tight. Strained. She already sensed the inevitable. She knew why I was there. She knew what was coming. She ended the call and turned. She attempted to regain control. To put on the mask. "Detective Halloway. I suggest you make an appointment."

"I already have one." I placed the evidence folder on her desk. The Thornhill Ledger linking her firm to the crime. The Price-Reed Memos proving her direction. Webb's Affidavit confirming her intent. It was all there. Everything. The whole case. The whole truth.

Helen looked at the stack. She didn't touch it. Her expression was pure dread. The mask was gone. The facade was broken. She was just a woman facing the consequences of her choices. Of her corruption. Of her crimes.

"The Queen of Convictions," I said. I walked slowly toward the Lady Justice painting behind her. The symbol of everything she had betrayed. "Fifty-three wins. How many were built on this? How many innocent lives were sacrificed to keep your father's firm liquid? How many people did you send to prison knowing they were innocent?"

"I was maintaining order," she snapped. She stood abruptly. The movement was sharp. Defensive. "The city needs stability. We protected the integrity of the institution against petty criminals. This was a necessary evil. You don't understand the pressure. The expectations. The need to win."

"You protected your stock portfolio." I pointed to the wall. To the awards. To the evidence of her success. "You protected your reputation. You protected your family's money. And you did it by destroying innocent lives. Now where is The Insurance Policy?"

Helen broke. The fight drained out of her instantly. She looked fragile and terrified. Small. Defeated. She fumbled with the wall panel revealing a small safe. The mechanism clicked. The door opened. Inside was a single black ledger. The Insurance Policy. The truth. The whole truth.

I opened it. A record of her family's firm. The blackmailed clients. The engineered cases. The whole conspiracy. And the name appearing over and over again next to every fabrication request. T.W.

Tom Wade. Chief Forensic Examiner. My best friend.

The name hit me like a physical blow. Tom. The man I had trusted. The man I had called my friend. The man who had been there for me through everything. Through my divorce. Through my retirement. Through my darkest moments. He was the architect. He was the monster.

"He was the architect," Helen whispered. Defeated. Broken. "He did the work. I just told him what we needed. He loved making the evidence perfect. He loved being the god who decided the truth. He was the one who made it all possible."

I had the complete file. The Queen was ours. The conspiracy was exposed. The truth was out. But now I had a choice. The convergence point was here. The moment of decision. The moment that would define everything.'''
    }
}

# Continue with remaining subchapters - I'll add them in batches
print("Expanding Batch 2 subchapters...")
doc2 = Document('The_Midnight_Confessor_Batch2.docx')
shutil.copy('The_Midnight_Confessor_Batch2.docx', 'The_Midnight_Confessor_Batch2_backup.docx')

# Expand the ones we have defined
for subchapter, expansion in expansions_batch2.items():
    print(f"\nExpanding Subchapter {subchapter}...")
    expand_subchapter(doc2, subchapter, expansion['start'], expansion['text'])

doc2.save('The_Midnight_Confessor_Batch2.docx')
print("\nBatch 2 partial expansion complete. More expansions needed.")
