#!/usr/bin/env python3
"""
Comprehensive script to expand all subchapters in The Midnight Confessor.
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
            # Next subchapter found, narrative ends here
            decision_point_idx = i
            break
    
    if narrative_start_idx is None:
        print(f"ERROR: Could not find Subchapter {subchapter_num}")
        return False
    
    if decision_point_idx is None:
        print(f"WARNING: Could not find [DECISION POINT] for Subchapter {subchapter_num}, using end of document")
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
        # Insert before the decision point paragraph
        if decision_point_idx < len(doc.paragraphs):
            target_para = doc.paragraphs[decision_point_idx - 1]
            new_para = target_para.insert_paragraph_before(para_text)
            decision_point_idx += 1  # Decision point moved down
        else:
            # Add to end
            doc.add_paragraph(para_text)
    
    print(f"Successfully expanded Subchapter {subchapter_num} ({len(new_paras)} paragraphs)")
    return True

# ========== EXPANDED CONTENT FOR BATCH 1 ==========

# Subchapter 2.2A - The Confrontation (Armed with Truth) - needs 61+ words
expanded_2_2A = """NARRATIVE: I didn't kick down Silas's door. I didn't need to. I had the ledger.

The doorman let me up to the 23rd floor. "Mr. Reed is expecting you." The doorman vanished before I realized Silas was already warned. Victoria had eyes everywhere. She was orchestrating this. Making sure I saw everything she wanted me to see.

Silas stood on his balcony. He looked out at the grey expanse of the bay. A glass of bourbon in his hand. He looked defeated. A statue weathered by decades of invisible rain. The expensive penthouse felt like a prison. The view was beautiful but it was just another cage.

"Jack," he said. He didn't turn around. His voice was tired. Resigned. "I wondered when you would get here. Did you enjoy breakfast at the diner? Claire makes terrible coffee but she has stamina. Four years of investigating. Four years of waiting. I should have known she would never give up."

"She is smarter than us, Silas." I walked out onto the balcony. The cold air hit me like a slap. The city spread out below us. A million stories. A million secrets. And we were just two men standing on a ledge, about to fall.

I tossed the flash drive onto the glass table. It made a sharp clack. The sound of inevitability. The sound of a case closing. The sound of a friendship ending.

Silas looked at it. He didn't feign ignorance. He didn't try to lie. He just sighed. A long exhale that deflated his frame. "She kept everything. I knew she would. I knew this day would come. I just didn't think it would be you standing here."

"Why?" I asked. The word sounded hollow in the cold air. "Eight million dollars? Was that the price of our friendship? Of Marcus's life?"

"It wasn't the money." He turned. His eyes were red-rimmed and filled with a desperate honesty. The honesty of a man who had been lying for so long he had forgotten what truth felt like. "I was blackmailed. Seven years ago. They had photos of me. Me and a man. They were going to ruin me. Ruin my life. Expose my lie to my wife and sons."

He picked up the drive. His fingers left smudged prints on the casing. The evidence of his guilt. The evidence of my failure. "They told me what to do. Create the accounts. Sign the papers. Frame Marcus. They said if I did it the photos would disappear. I sacrificed a stranger to save my life. I chose my family over his. I chose my secret over his freedom."

"You destroyed Marcus to save your reputation," I corrected. My voice was cold. Colder than the air around us.

"I did. And you let me. You were so desperate for a neat case. You never looked twice at the signatures. Never questioned why the evidence was so perfect. I'm a monster sure. But you? You were the blind man driving the getaway car. Too arrogant to read the map. Too certain to ask questions."

The words hit like physical blows. Each one true. Each one a reminder of how I had failed. Not just Silas. Not just Marcus. But everyone. The system. The truth. Myself.

My phone buzzed. A text from Sarah. Jack. Maya Bellamy is missing. Security footage shows her getting into a black Mercedes. Victoria has her. FBI is already issuing alerts.

I looked at Silas. The man I had trusted. The man who had betrayed that trust. "Victoria has Maya."

Silas went pale. The color drained from his face. "She is escalating. She told me she wouldn't involve innocents if I cooperated. She said she was only interested in us. The guilty. The corrupt. The ones who deserved punishment."

"She lied," I said. "She's been lying from the beginning. This isn't about justice. This is about revenge. And revenge doesn't care about innocence."""

# Subchapter 2.3B - The Storm - needs 336+ words
expanded_2_3B = """NARRATIVE: I had the confession. The evidence. The rage. I had achieved maximum chaos. Now I had to use it to corner Victoria.

The diner felt like a pressure cooker. The air was thick with tension. Silas sat in the corner booth, broken and weeping. His expensive suit was rumpled. His face was a mess of tears and snot. The man who had been my partner. The man who had framed innocent people. The man who was now just a weapon in my hand.

Claire stood over him with the steak knife still in her grip. Her knuckles were white. Her whole body was shaking. Not from fear. From rage. Four years of waiting. Four years of planning. Four years of watching the man who destroyed her father live his comfortable life. And now he was here. In her diner. At her mercy.

"Put the knife down, Claire," I said. My voice was calm. Too calm. The calm that comes right before the storm.

"Why?" she demanded. Her eyes never left Silas. "You brought him here. You want me to what? Forgive him? Thank you?"

"I want you to live," I said. "If you kill him, you go to prison. Victoria wins. She gets to watch another innocent person get locked away. That's what she wants. That's the whole point of this game."

Claire's grip tightened on the knife. Silas flinched. He looked like he was about to vomit. The fear was real. The terror was genuine. This wasn't an act. This was a man facing the consequences of his choices.

My phone buzzed. Sarah. Again. The third call in as many minutes. I ignored it. I couldn't talk to her right now. Not after what I had done. Not after dragging Silas here like a trophy. She would try to stop me. She would try to make me do things the right way. And the right way would get Maya killed.

I looked at the flash drive in my hand. The Thornhill Ledger. Everything Claire had spent four years collecting. Every wire transfer. Every signature. Every piece of evidence that proved Silas had framed her father. It was all here. On a piece of plastic smaller than my thumb.

I had everything I needed to destroy Silas. To put him away for the rest of his life. But that wasn't enough. Not anymore. Victoria had Maya. And Victoria was the real enemy. Silas was just a pawn. A broken, weeping pawn who had made terrible choices for terrible reasons.

"Claire," I said. "I need you to trust me. Just for a little while longer. I'm going to use him to find Victoria. I'm going to use him to save Maya Bellamy. And then I'm going to make sure he pays for what he did to your father. But I need him alive. I need him functional. I need him to be useful."

Claire's eyes flickered to me. The rage didn't leave, but something else joined it. Understanding. She nodded. Slowly. The knife lowered. Not all the way. But enough.

Silas let out a breath. A sob. Relief and terror mixed together.

I looked at the flash drive again. Then at Silas. Then at Claire. Three broken people in a greasy diner at dawn. Three people whose lives had been destroyed by the same system. The same corruption. The same lies.

My phone buzzed again. This time it was a text. Unknown number. Victoria.

You have the evidence. You have the confession. You have everything you need to be the hero again. But heroes follow rules. Heroes wait for permission. Heroes let the system work. And while the system works, Maya Bellamy dies. The choice is yours, Detective. Be the hero. Or be effective. —M.C.

I looked at the text. Then at Silas. Then at the flash drive. I had two paths. Two ways to use what I had.

The smart way. The methodical way. Use Silas as bait. Call Victoria. Set up a meeting. Draw her out into the open. Use the system. Use the law. Use everything I had learned in thirty years of being a cop.

The hard way. The direct way. Storm her location. Use Silas to find her. Break down the door. Take her by force. End this tonight. No rules. No system. Just action.

Both paths would work. Both paths would save Maya. But only one would let me sleep at night. Only one would let me look in the mirror tomorrow and recognize the man staring back.

I made my choice."""

# Subchapter 2.2B - The Missing Piece - needs 182+ words
expanded_2_2B = """NARRATIVE: I didn't waste time on the safe. I needed to get out before the police showed up.

The penthouse was a crime scene now. My crime scene. I had broken in. I had assaulted Silas. I had contaminated everything. But I didn't care. The rules were for people who had time. I was running out of it.

"Where is the proof? The hard evidence?" I demanded. I grabbed Silas by the collar. The silk was expensive. It felt wrong in my hands. Too soft. Too clean. "I need something I can use. Something that will hold up in court. Something that will put you away."

"Claire," he whispered. His voice was broken. Defeated. "Marcus's daughter. Claire Thornhill. She works at the Blueline Diner. She has been investigating us for four years. She has the proof, Jack. The physical proof. The Thornhill Ledger. Everything. Every wire transfer. Every signature. Every piece of evidence that proves I framed her father."

My phone buzzed. Sarah. Jack. Maya Bellamy is missing. Black Mercedes. Victoria has her. You need to call me now.

The text hit me like a physical blow. Maya. Eleanor's daughter. Another innocent caught in the crossfire. Another life on the line because of my mistakes. Because of my arrogance. Because of my blindness.

"Victoria has Maya," I said. I shoved Silas to his feet. "Get dressed. We're going to the diner. We're getting that ledger. And then we're finding Victoria. And if anything happens to that girl, I will make sure you spend the rest of your life in a cell next to the people you framed."

"I can't," he whimpered. "If I leave... the photos... they'll release them. My family..."

"The photos don't matter anymore! You are done! The only thing left is whether you go to prison as a man or a coward! Now get dressed!"

I drove the Bentley at 80 mph across town. The rain came down in sheets. The streets were slick. Dangerous. I didn't care. Silas sobbed in the passenger seat. The sound was grating. The sound of a man realizing his life was over. The sound of a man facing consequences.

We arrived at the Blueline Diner. The neon sign flickered. The windows were fogged. Inside, the early morning crowd was sparse. A few truckers. A few night shift workers. People who lived in the margins. People who understood what it meant to be broken.

I didn't care about the risk. My aggression demanded the evidence and the victim. I needed both. I needed Claire's ledger. I needed Silas's confession. I needed everything I could get my hands on. Because Victoria had Maya. And I was running out of time.

I dragged Silas in. The bell on the door chimed. Every head turned. Every eye watched. This was a show. A performance. And I was the director.

Claire Thornhill looked up. She saw me. Then she saw Silas. The man who killed her father. The man who had destroyed her family. The man who had stolen four years of her life.

Her face went white. Then red. Rage. Pure, unadulterated rage. She grabbed a steak knife off a table. The blade was clean. Sharp. Ready.

"You," she hissed at Silas. The word was a weapon. A promise. A threat.

"He confessed," I said. I shoved Silas toward a booth. The man stumbled. Fell. Looked up at Claire with terror in his eyes. "He framed your father. We need the drive, Claire. Now. Victoria has another girl. We need your ledger."

Claire stared at me. The aggressive half-crazed detective who delivered her villain on a silver platter. It was a chaotic move. Messy. Dangerous. Reckless. But it was also effective.

She looked at Silas. At the knife in her hand. At me. Then she made a decision. She threw the flash drive at me. "Take it. And get him out of my sight before I kill him."

I caught the drive. It felt heavy. Like a weapon. Like justice. Like revenge. I had everything I needed. Now I just had to use it."""

# Subchapter 2.1B - The Penthouse Raid - needs 9+ words
expanded_2_1B = """NARRATIVE: I didn't wait for morning. I didn't wait for Sarah's advice. I didn't wait for coffee.

I drove to the Marina District at 3 AM. The rain came down in sheets. It washed the city clean but couldn't touch the dirt I felt on me. I was fueled by betrayal. I needed immediate brutal confrontation. I needed to look Silas in the eye and make him tell me the truth. No more games. No more lies. Just the raw, ugly truth.

The Marina District was quiet. The kind of quiet that comes before a storm. The expensive buildings loomed in the darkness. Their windows were dark. Their residents were sleeping. Comfortable. Safe. Unaware that their neighbor was a monster.

I bypassed the doorman. Smashed the call box. Vaulted the velvet rope with my .38 visible. I took the elevator to the 23rd floor. The ascent felt like a rising tide of cold rage. Each floor was a countdown. Each second was a heartbeat. Each breath was a promise.

I kicked the door. The lock screamed a protest and splintered the expensive jamb. The sound was satisfying. The sound of breaking rules. The sound of consequences.

Silas stood there in a silk robe. Glass of bourbon in hand. He looked terrified. The fear of a man being found not the fear of a man being robbed. He knew why I was here. He knew what I had discovered. He knew his life was over.

"Jack?" he stammered. Liquid spilled on the marble. The expensive floor. The expensive life. All of it about to be destroyed. "What the hell..."

I shoved him. Hard. He stumbled back into the living room. The expensive furniture. The expensive art. The expensive lies. "Marcus Thornhill," I growled. "Tell me about the wire transfers you son of a bitch."

"I don't know what you're talking about! You're drunk, Jack. Leave before I call security!"

"I'm sober for the first time in twenty years." I grabbed him by the lapels of his expensive robe. I slammed him against the polished steel wall. The air rushed out of his lungs. The sound was satisfying. The sound of a man breaking. "I saw the signature. I saw the shell company. You framed him. You let him die in my lockup."

I pulled my gun. I didn't point it. I just let him see the cold weight of it. The promise of violence. The threat of death. "Tell me the truth, Silas. Or so help me God I will forget I'm a cop and remember you are just the man who betrayed me."

Silas didn't fight. He crumbled instantly. He slid down the wall weeping hysterically. Snot ran down his face. Tears. Real tears. The tears of a man who had been holding it all in for too long. "They made me," he sobbed. "They had photos. Me with a man. They were going to ruin me. I did it to save my family!"

"Who?"

"I don't know! They blackmailed me! Seven years ago! I signed the papers. I framed Marcus. I did it to save my life! And you helped! You were so arrogant, Jack! You wanted the win so bad you never checked the dates! You made it easy for me!"

I let him go. I felt sick. Physically ill. The confession was a relief and a devastation all at once. The truth was out. But the truth was worse than I had imagined.

"Victoria sent me a letter. She said you would come. She said if you came with rage you were finally waking up." Silas pointed vaguely to a safe under his desk. "The files. The internal affairs documents. The ones that prove they were watching us. They are all there."

I looked at the safe. Then at Silas. Then at my gun. I had the truth. Now I needed the proof. And I needed to get out before the police showed up."""

# Fix em dash in Subchapter 1.2
def fix_em_dash(doc):
    """Fix the em dash in Subchapter 1.2"""
    for para in doc.paragraphs:
        if '—M.C.' in para.text:
            para.text = para.text.replace('—M.C.', 'M.C.')
            print("Fixed em dash in Subchapter 1.2")
            return True
    return False

# ========== MAIN EXECUTION ==========

print("=" * 80)
print("EXPANDING BATCH 1 SUBCHAPTERS")
print("=" * 80)

# Batch 1
doc1 = Document('The_Midnight_Confessor_Batch1.docx')
shutil.copy('The_Midnight_Confessor_Batch1.docx', 'The_Midnight_Confessor_Batch1_backup.docx')

print("\n1. Fixing em dash in Subchapter 1.2...")
fix_em_dash(doc1)

print("\n2. Expanding Subchapter 2.2A...")
expand_subchapter(doc1, '2.2A', 'NARRATIVE: I didn\'t kick down Silas\'s door', expanded_2_2A)

print("\n3. Expanding Subchapter 2.3A...")
# Already expanded in previous run, skipping
# expand_subchapter(doc1, '2.3A', 'NARRATIVE: "We need to find her. Now."', expanded_2_3A)
print("   (Already expanded, skipping)")

print("\n4. Expanding Subchapter 2.1B...")
expand_subchapter(doc1, '2.1B', 'NARRATIVE: I didn\'t wait for morning', expanded_2_1B)

print("\n5. Expanding Subchapter 2.2B...")
expand_subchapter(doc1, '2.2B', 'NARRATIVE: I didn\'t waste time on the safe', expanded_2_2B)

print("\n6. Expanding Subchapter 2.3B...")
expand_subchapter(doc1, '2.3B', 'NARRATIVE: I had the confession', expanded_2_3B)

doc1.save('The_Midnight_Confessor_Batch1.docx')
print("\n" + "=" * 80)
print("BATCH 1 COMPLETE")
print("=" * 80)
