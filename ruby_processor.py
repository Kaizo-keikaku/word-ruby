import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

def create_ruby_element(text, ruby_text):
    """
    Create a w:ruby XML element.
    Ex:
    <w:ruby>
      <w:rubyPr>
        <w:rubyAlign w:val="distributeSpace"/>
        <w:hps w:val="10"/>
        <w:hpsRaise w:val="20"/>
        <w:hpsBaseText w:val="21"/>
        <w:lid w:val="ja-JP"/>
      </w:rubyPr>
      <w:rt>
        <w:r>
          <w:rPr>
            <w:sz w:val="10"/>
          </w:rPr>
          <w:t>ruby_text</w:t>
        </w:r>
      </w:rt>
      <w:rubyBase>
        <w:r>
           <w:rPr>
            <w:sz w:val="21"/>
           </w:rPr>
           <w:t>text</w:t>
        </w:r>
      </w:rubyBase>
    </w:ruby>
    """
    ruby = OxmlElement('w:ruby')
    
    # Ruby Properties
    rubyPr = OxmlElement('w:rubyPr')
    rubyAlign = OxmlElement('w:rubyAlign')
    rubyAlign.set(qn('w:val'), 'distributeSpace')
    rubyPr.append(rubyAlign)
    
    # Font sizes (approximate defaults, usually word handles this if omitted, but let's be safe)
    # 10.5pt = 21 half-points
    # Ruby text usually 50-60% of base. Let's try omitting specific sizes first to let Word default
    # or just simple standard.
    # Actually, minimal structure is often enough.
    
    ruby.append(rubyPr)

    # Ruby Text (Furigana)
    rt = OxmlElement('w:rt')
    r_rt = OxmlElement('w:r')
    t_rt = OxmlElement('w:t')
    t_rt.text = ruby_text
    r_rt.append(t_rt)
    rt.append(r_rt)
    ruby.append(rt)

    # Ruby Base (Original Text)
    rubyBase = OxmlElement('w:rubyBase')
    r_base = OxmlElement('w:r')
    t_base = OxmlElement('w:t')
    t_base.text = text
    r_base.append(t_base)
    rubyBase.append(r_base)
    ruby.append(rubyBase)
    
    return ruby

def apply_ruby_to_document(doc_path, output_path, ruby_settings, mode='all'):
    """
    ruby_settings: list of dict {'word': '...', 'ruby': '...'}
    mode: 'once', 'per_page', 'all'
    """
    doc = Document(doc_path)
    
    # Flatten settings for easier lookup?
    # actually we just need to iterate.
    # But for 'once' mode, we need to track global usage.
    # For 'per_page', we need to track per page.
    
    # Global state for 'once'
    global_applied = {item['word']: False for item in ruby_settings}
    
    # Sort settings by length of word descending to multiple matches correctly (longest match first)
    # But simple iteration might fail on overlapping. We assume non-overlapping for simplicity in MVP.
    sorted_settings = sorted(ruby_settings, key=lambda x: len(x['word']), reverse=True)
    
    current_page_index = 0
    # Track per-page usage. 
    # Since we don't know page numbers easily, we increment an index whenever we see a break.
    items_applied_on_page = {item['word']: False for item in sorted_settings}

    for paragraph in doc.paragraphs:
        # Check for page breaks in the paragraph or previous siblings?
        # A paragraph might contain a page break.
        # w:lastRenderedPageBreak appears inside runs usually.
        
        # We process runs.
        # Because replacing text inside a run is easier if we split the run.
        
        # NOTE: This implementation is complex because we need to match text across runs? 
        # For simplicity, we only match text WITHIN a single run. 
        # Matching across runs is much harder (requires re-merging or complex state).
        
        # New approach: Iterate runs, check text, split if necessary.
        
        # We need to use a while loop because we act modifying the list of runs (splitting)
        
        # To detect page breaks:
        # Check for <w:lastRenderedPageBreak/> or <w:br w:type="page"/> in the run BEFORE processing text.
        
        p_element = paragraph._element
        
        # We will traverse children of paragraph to find runs and breaks
        # But iterating paragraph.runs is easier for text access.
        
        # Let's iterate over paragraph.runs, but we might modifications.
        # It's safer to iterate indices or copy.
        
        i = 0
        while i < len(paragraph.runs):
            run = paragraph.runs[i]
            run_xml = run._element
            
            # Check page break markers in this run's XML (before the text node typically)
            # w:br type="page" or w:lastRenderedPageBreak
            xml_str = etree.tostring(run_xml).decode('utf-8')
            if 'w:lastRenderedPageBreak' in xml_str or 'w:type="page"' in xml_str:
                 # Reset per-page tracker
                 items_applied_on_page = {item['word']: False for item in sorted_settings}
            
            original_text = run.text
            if not original_text:
                i += 1
                continue
            
            # Find match
            match_found = False
            for setting in sorted_settings:
                word = setting['word']
                ruby = setting['ruby']
                
                # Check mode logic
                if mode == 'once' and global_applied[word]:
                    continue
                if mode == 'per_page' and items_applied_on_page[word]:
                    continue
                
                idx = original_text.find(word)
                if idx != -1:
                    # Found a match!
                    match_found = True
                    
                    # Log application
                    if mode == 'once':
                        global_applied[word] = True
                    if mode == 'per_page':
                        items_applied_on_page[word] = True
                        
                    # Split run:
                    # 1. Text before match
                    # 2. Ruby element (match)
                    # 3. Text after match
                    
                    text_before = original_text[:idx]
                    text_after = original_text[idx+len(word):]
                    
                    # Modify current run to contain text_before
                    run.text = text_before
                    
                    # Create ruby element
                    # We need to insert it after the current run
                    # We can use lxml to insert after
                    
                    ruby_elem = create_ruby_element(word, ruby)
                    
                    # We insert the ruby element after the current run's element
                    run_xml.addnext(ruby_elem)
                    
                    # If there is text after, we need a new run for it
                    if text_after:
                        new_run = OxmlElement('w:r')
                        # Copy properties from original run if possible? 
                        # Ideally yes, but simpler to just make a run.
                        # Copying rPr
                        if run_xml.rPr is not None:
                            import copy
                            new_run.append(copy.deepcopy(run_xml.rPr))
                        
                        t_u = OxmlElement('w:t')
                        if len(text_after) > 0 and (text_after[0] == ' ' or text_after[-1] == ' '):
                             t_u.set(qn('xml:space'), 'preserve')
                        t_u.text = text_after
                        new_run.append(t_u)
                        
                        ruby_elem.addnext(new_run)
                        
                        # Now we need to update paragraph.runs list or just continue?
                        # Since we are operating on XML, paragraph.runs might be stale?
                        # Actually docx caches runs. This approach might break iteration if not careful.
                        # But since we found *one* match and processed it, we can break and re-scan the paragraph?
                        # Or recursive?
                        
                        # Simplest: After one modification in a paragraph, stop processing that paragraph or run?
                        # Better: Process the "text_after" recursively or in next iteration.
                        # But we inserted XML directly. python-docx's paragraph.runs list structure is based on XML.
                        # If we re-access paragraph.runs, it should reflect changes? 
                        # No, it's often cached.
                        
                        # Safe bet: If we modify, break inner loop and restart scanning the text_after?
                        # Or restart scanning the whole paragraph?
                        # If we restart whole paragraph, we might re-encounter the word if we didn't track it?
                        # But we are replacing the word with a ruby object (which is not a run text).
                        # So it won't be found as text again.
                    
                    # Break finding loop to restart/continue
                    break 
            
            if match_found:
                 # Since we modified the XML structure, the 'runs' list is invalid.
                 # We should reload runs or restart paragraph processing?
                 # If we restart paragraph processing, we need to handle the fact that 'text_before' is now the run.
                 # And 'ruby' is not a run. 'text_after' is a new run.
                 # So paragraph.runs will have more elements.
                 # Let's break the 'i' loop and restart it for the current paragraph?
                 # But we need to make sure we don't re-process the same text.
                 # We handled 'word' -> ruby. It's gone from text.
                 # So safe to restart processing runs of this paragraph.
                 i = -1 # will be incremented to 0
                 # NOTE: This is inefficient (O(N^2)) but safe.
            
            i += 1

    doc.save(output_path)
    return output_path
