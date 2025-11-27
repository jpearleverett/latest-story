#!/usr/bin/env python3
"""
Script to expand subchapters in The Midnight Confessor docx files.
Modifies the XML structure to replace short subchapters with expanded versions.
"""

import zipfile
import xml.etree.ElementTree as ET
import shutil
import re
from io import BytesIO

def create_text_element(ns, text):
    """Create a w:t (text) element with the given text"""
    t_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t_elem.text = text
    return t_elem

def create_paragraph_with_text(ns, text):
    """Create a paragraph element containing the text"""
    p_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    r_elem = ET.SubElement(p_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t_elem = create_text_element(ns, text)
    r_elem.append(t_elem)
    return p_elem

def find_and_replace_narrative(docx_path, subchapter_title, old_narrative_start, new_narrative_text):
    """Find a subchapter's narrative section and replace it with expanded text"""
    
    # Create backup
    backup_path = docx_path.replace('.docx', '_backup.docx')
    shutil.copy(docx_path, backup_path)
    print(f"Created backup: {backup_path}")
    
    # Read the docx
    z = zipfile.ZipFile(docx_path, 'r')
    xml_content = z.read('word/document.xml')
    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Find all paragraphs
    paragraphs = root.findall('.//w:p', ns)
    
    # Find the narrative section
    in_narrative = False
    narrative_start_idx = None
    narrative_end_idx = None
    
    for i, para in enumerate(paragraphs):
        # Get text from paragraph
        text_elements = para.findall('.//w:t', ns)
        para_text = ''.join([t.text if t.text else '' for t in text_elements])
        
        if 'NARRATIVE:' in para_text and subchapter_title in '\n'.join([
            ''.join([t.text if t.text else '' for t in p.findall('.//w:t', ns)]) 
            for p in paragraphs[max(0, i-5):i]
        ]):
            in_narrative = True
            narrative_start_idx = i
            # Check if this paragraph contains the old narrative start
            if old_narrative_start in para_text:
                break
    
    if narrative_start_idx is None:
        print(f"ERROR: Could not find narrative section for {subchapter_title}")
        return False
    
    # Find where the narrative ends (before [DECISION POINT] or next subchapter)
    for i in range(narrative_start_idx, len(paragraphs)):
        text_elements = paragraphs[i].findall('.//w:t', ns)
        para_text = ''.join([t.text if t.text else '' for t in text_elements])
        
        if '[DECISION POINT]' in para_text or (i > narrative_start_idx and 'Subchapter' in para_text):
            narrative_end_idx = i
            break
    
    if narrative_end_idx is None:
        narrative_end_idx = len(paragraphs)
    
    print(f"Found narrative section: paragraphs {narrative_start_idx} to {narrative_end_idx}")
    
    # Split new narrative into paragraphs
    new_paragraphs_text = new_narrative_text.split('\n\n')
    
    # Replace the narrative paragraphs
    body = root.find('.//w:body', ns)
    if body is None:
        print("ERROR: Could not find document body")
        return False
    
    # Get all body children
    body_children = list(body)
    
    # Remove old narrative paragraphs
    for i in range(narrative_end_idx - 1, narrative_start_idx - 1, -1):
        if i < len(body_children):
            body.remove(body_children[i])
    
    # Insert new paragraphs
    insert_pos = narrative_start_idx
    for para_text in new_paragraphs_text:
        if para_text.strip():
            new_para = create_paragraph_with_text(ns, para_text.strip())
            body.insert(insert_pos, new_para)
            insert_pos += 1
    
    # Write back to docx
    z.close()
    
    # Create new docx with modified content
    new_docx = zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED)
    
    # Copy all files except document.xml
    old_z = zipfile.ZipFile(backup_path, 'r')
    for item in old_z.infolist():
        if item.filename != 'word/document.xml':
            new_docx.writestr(item, old_z.read(item.filename))
    old_z.close()
    
    # Write modified document.xml
    xml_str = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
    new_docx.writestr('word/document.xml', xml_str)
    new_docx.close()
    
    print(f"Successfully updated {docx_path}")
    return True

if __name__ == '__main__':
    print("Subchapter expansion script")
    print("This script will be called with specific parameters for each subchapter")
