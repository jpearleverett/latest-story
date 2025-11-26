#!/usr/bin/env python3
"""
Comprehensive script to expand all subchapters in The Midnight Confessor.
Uses direct text replacement in the extracted text, then rebuilds docx.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import shutil
from collections import OrderedDict

# Expanded content for each subchapter
EXPANSIONS = {
    '2.3A': {
        'old_start': 'NARRATIVE: "We need to find her. Now." I grabbed Silas by the arm. The time for deliberation was over.',
        'new_text': '''NARRATIVE: "We need to find her. Now." I grabbed Silas by the arm. The time for deliberation was over.

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

The apology hung in the air between us. Empty. Meaningless. But it was all he had left to give. I turned away from him. I looked at the city. At the rain. At the gray expanse of water beyond the buildings. Somewhere out there, Maya Bellamy was waiting. And I was coming for her. With a criminal as my guide. With a flash drive as my weapon. With everything I had left to lose.'''
    },
    
    '2.3B': {
        'old_start': 'NARRATIVE: I had the confession. The evidence. The rage. I had achieved maximum chaos. Now I had to use it to corner Victoria.',
        'new_text': '''NARRATIVE: I had the confession. The evidence. The rage. I had achieved maximum chaos. Now I had to use it to corner Victoria.

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

I made my choice.'''
    }
}

def expand_docx_file(docx_path, expansions):
    """Expand subchapters in a docx file"""
    
    # Create backup
    backup_path = docx_path.replace('.docx', '_backup.docx')
    shutil.copy(docx_path, backup_path)
    print(f"Created backup: {backup_path}")
    
    # Extract text
    z = zipfile.ZipFile(docx_path, 'r')
    xml_content = z.read('word/document.xml')
    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Get all paragraphs
    paragraphs = root.findall('.//w:p', ns)
    
    # Convert to text for searching
    full_text = []
    for para in paragraphs:
        text_elems = para.findall('.//w:t', ns)
        para_text = ''.join([t.text if t.text else '' for t in text_elems])
        full_text.append(para_text)
    
    text_content = '\n'.join(full_text)
    
    # Apply expansions
    for subchapter, expansion in expansions.items():
        old_start = expansion['old_start']
        new_text = expansion['new_text']
        
        # Find the narrative section
        pattern = f'(Subchapter [0-9.]+[A-Z]* - [^\n]+\\nPREVIOUSLY:[^\\n]+\\nBRIDGE TEXT:[^\\n]+\\n){re.escape(old_start)}(.*?)(\\n\\[DECISION POINT\\])'
        
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            # Replace the narrative
            before = match.group(1)
            after = match.group(3)
            replacement = before + new_text + after
            text_content = text_content[:match.start()] + replacement + text_content[match.end():]
            print(f"Expanded Subchapter {subchapter}")
        else:
            print(f"WARNING: Could not find Subchapter {subchapter}")
    
    # Now we need to rebuild the docx with the new text
    # This is complex - for now, let's write the modified text to a file
    output_path = docx_path.replace('.docx', '_expanded.txt')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(text_content)
    print(f"Wrote expanded text to {output_path}")
    print("NOTE: Manual insertion into docx required, or use python-docx library")

if __name__ == '__main__':
    # Expand Batch 1
    expand_docx_file('The_Midnight_Confessor_Batch1.docx', {
        '2.3A': EXPANSIONS['2.3A'],
        '2.3B': EXPANSIONS['2.3B']
    })
