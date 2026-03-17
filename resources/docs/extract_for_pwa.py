"""
Extract content from the consolidated GCC Playbook DOCX into structured JSON for the PWA.
"""
import json
import os
from docx import Document

DOCS_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(DOCS_DIR, "GCC_Playbook_2026_2030_FINAL.docx")
OUTPUT_DIR = os.path.join(os.path.dirname(DOCS_DIR), "..", "pwa", "public", "data")

def extract_content():
    print("Loading consolidated document...")
    doc = Document(INPUT_FILE)
    
    chapters = []
    current_part = None
    current_chapter = None
    current_section = None
    current_subsection = None
    
    glossary_terms = []
    in_glossary = False
    in_index = False
    in_references = False
    
    references = []
    
    # Parts structure
    parts = {
        "part1": {"title": "Part I: India's GCC Landscape", "subtitle": "Landscape, Maturity, Operating Models, Challenges, KPIs, Community", "chapters": []},
        "part2": {"title": "Part II: GCC by Scale & Stage", "subtitle": "Sizing, Setup, Lifecycle, Case Studies, Leaders", "chapters": []},
        "part3": {"title": "Part III: The Frontier", "subtitle": "AI, Deep Tech, M&A, Finance, Legal, ESG, 2030 Scenarios", "chapters": []},
    }
    
    current_part_key = "part1"
    chapter_counter = 0
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        style_name = para.style.name if para.style else "Normal"
        
        # Detect glossary section
        if text == "Key Terms and Definitions":
            in_glossary = True
            in_index = False
            in_references = False
            continue
        
        # Detect index section
        if text == "Subject Index":
            in_glossary = False
            in_index = True
            in_references = False
            continue
        
        # Detect references
        if "VALIDATED REFERENCES" in text.upper() or "Primary Sources" in text:
            in_glossary = False
            in_index = False
            in_references = True
            continue
        
        # Process glossary
        if in_glossary and not in_index:
            if style_name.startswith('Heading'):
                continue  # Skip alphabet headers
            if ':' in text or '—' in text or '-' in text:
                sep = ':' if ':' in text else ('—' if '—' in text else '-')
                parts_split = text.split(sep, 1)
                if len(parts_split) == 2:
                    term = parts_split[0].strip()
                    definition = parts_split[1].strip()
                    if len(term) < 80 and len(definition) > 5:
                        glossary_terms.append({"term": term, "definition": definition})
            continue
        
        # Process references
        if in_references:
            if text and not style_name.startswith('Heading'):
                references.append(text)
            continue
        
        if in_index:
            continue

        # Detect Part transitions
        if "PART III" in text or "Part III" in text and "Frontier" in text.lower() if "frontier" in text.lower() else False:
            current_part_key = "part3"
        
        # Process headings for chapter structure
        if style_name == 'Heading 1':
            # Determine part from numbering
            if text.startswith(('1.', '2.', '3.', '4.', '5.', '6.')) and '.' in text[:3]:
                num = text.split('.')[0].strip()
                section_num = text.split(' ')[0].strip()
                if num.isdigit() and int(num) <= 6 and current_part_key == "part1":
                    pass  # Part I
                elif current_part_key == "part2":
                    pass
            
            # Check if this is a Part II chapter (size tiers, setup, etc.)
            if any(kw in text for kw in ['Official Size Taxonomy', 'Pre-Launch', 'Four Defining Characteristics', 
                                          'Changes at 1,000', 'Large GCC', 'Eight Operating Models', 
                                          'Seven Stages', 'Case Study']):
                current_part_key = "part2"
            
            # Check if Part III
            if any(kw in text for kw in ['AI as GCC Operating', 'Global Location Competition',
                                          'GCC M&A', 'Deep-Tech GCCs', 'GCC Finance and Treasury',
                                          'Talent Intelligence System', 'Legal, Regulatory',
                                          'Innovation Engine', 'ESG and Sustainability',
                                          "Leader's Career", 'GCC 2030', 'Women in GCC',
                                          'Mid-Market GCC Playbook', 'Mental Health']):
                current_part_key = "part3"
            
            chapter_counter += 1
            current_chapter = {
                "id": f"ch-{chapter_counter}",
                "number": text.split(' ')[0].strip() if text[0].isdigit() else str(chapter_counter),
                "title": text,
                "sections": [],
                "content": [],
                "keyTakeaways": []
            }
            parts[current_part_key]["chapters"].append(current_chapter)
            current_section = None
            current_subsection = None
            
        elif style_name == 'Heading 2':
            if current_chapter:
                # Check for chapter summaries / key takeaways
                if 'Chapter Summary' in text or 'Things Every' in text or 'Principles' in text:
                    current_section = {
                        "title": text,
                        "content": [],
                        "isKeyTakeaway": True
                    }
                else:
                    current_section = {
                        "title": text,
                        "content": [],
                        "isKeyTakeaway": False
                    }
                current_chapter["sections"].append(current_section)
                current_subsection = None
                
        elif style_name == 'Heading 3':
            if current_section:
                current_subsection = {
                    "title": text,
                    "content": []
                }
                current_section.setdefault("subsections", []).append(current_subsection)
        else:
            # Regular paragraph content
            if len(text) < 5:
                continue
                
            target = current_subsection if current_subsection else (current_section if current_section else (current_chapter if current_chapter else None))
            if target:
                # Limit content length per section for PWA performance
                if len(target.get("content", [])) < 20:
                    target.setdefault("content", []).append(text)
    
    # Build the final data structure
    data = {
        "title": "The Complete India GCC Reference 2026–2030",
        "subtitle": "Landscape | Maturity Models | Setup | Governance | Who's Who | AI | Deep Tech | M&A | Finance | Legal | ESG | 2030 Scenarios",
        "author": "Curated by Kalilur Rahman",
        "parts": parts,
        "glossary": glossary_terms[:200],  # Cap at 200 terms
        "references": references[:100],  # Cap at 100 references
        "stats": {
            "totalChapters": chapter_counter,
            "part1Chapters": len(parts["part1"]["chapters"]),
            "part2Chapters": len(parts["part2"]["chapters"]),
            "part3Chapters": len(parts["part3"]["chapters"]),
            "glossaryTerms": len(glossary_terms),
            "references": len(references)
        }
    }
    
    # Write JSON
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "content.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\nContent extracted to: {output_path}")
    print(f"Stats: {json.dumps(data['stats'], indent=2)}")
    print(f"Glossary terms: {len(glossary_terms)}")
    print(f"References: {len(references)}")
    
    return data

if __name__ == "__main__":
    extract_content()
