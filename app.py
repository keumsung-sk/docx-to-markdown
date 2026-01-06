import streamlit as st
import mammoth
import markdownify
import re
import zipfile
import io
import os
import yaml  # pip install PyYAML
import requests # pip install requests
from PIL import Image, UnidentifiedImageError # pip install Pillow
from docx import Document # pip install python-docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ==========================================
# 1. Configuration & Constants
# ==========================================

EXCLUDED_KEYWORDS = [
    "00_ignore", "mockup pages required", "global sections", "inside page components",
    "optional specialty pages", "header", "footer", "badges", "navigation", 
    "inside form", "financing box", "contact info", "variables", "meta description",
    "contact us today", "homepage",
    "promotions", "contact"
]

FIXED_DATE = "2001-01-01"
DEFAULT_PHONE_NUMBER = "555-555-5555"
START_MARKER = "Footer (All Pages)"

st.set_page_config(page_title="Jekyll Parser (Upload Only)", layout="wide")

# ==========================================
# 2. UI Layout (Sidebar & Main)
# ==========================================

# --- Sidebar: Features & Instructions ---
with st.sidebar:
    st.header("📝 Docx to Jekyll Parser")
    st.markdown("""
    ### ✨ Key Features
    
    1. **📄 Seamless Markdown Conversion**
       - Converts uploaded `.docx` pages directly into clean Markdown (`.md`) format.
    
    2. **🖼️ Smart Image Extraction**
       - Automatically detects public image URLs within the document.
       - Allows users to extract and download the actual image files.
    
    3. **🧩 Modular Component Separation**
       - Identifies **Reviews** and **Navigation** sections.
       - Exports them into separate YAML/Markdown files.
    
    4. **👀 Instant Preview**
       - Preview the converted content on the right side.
    
    5. **📦 One-Click Bulk Download**
       - Download Main content, Review/Nav modules, and Images in a single ZIP package.
    """)
    st.info("💡 **Note:** Drag & Drop your file on the right to start.")

# --- Main Area: Title & Upload ---
st.title("📂 Document Converter (Upload & Preview)")

target_file = st.file_uploader("Drag and drop your Word (.docx) file here", type=["docx"])

# ==========================================
# 3. Utility Functions (Kept as provided)
# ==========================================

def clean_markdown_link(text):
    text = re.sub(r'\[([^\]]+)\]\(.*?\)', r'\1', text)
    text = text.replace('\\', '').replace('<', '').replace('>', '')
    return text.strip()

def clean_nav_text(text):
    text = re.sub(r'\[([^\]]+)\]\(.*?\)', r'\1', text)
    text = text.replace('{', '').replace('}', '')
    text = re.sub(r'^[\s*_\-+\.]+', '', text)
    text = re.sub(r'[\s*_\-+\.]+$', '', text)
    text = text.replace('#', '') 
    text = text.replace('**', '') 
    return text.strip()

def to_kebab_case(text):
    text = clean_nav_text(text)
    text = text.replace(' - ', ' ')
    text = text.replace('|', '')
    text = re.sub(r'[^\w\s-]', '', text).strip().lower()
    return re.sub(r'[\s]+', '-', text)

def should_skip_page(title, content):
    clean_title = title.lower().strip()
    for keyword in EXCLUDED_KEYWORDS:
        if keyword in clean_title: return True
    if len(content.strip()) < 10: return True 
    return False

# ==========================================
# 4. Image Processing Functions
# ==========================================

def download_and_convert_image(url):
    clean_match = re.search(r'(https?://[^\s<>")\]]+)', url)
    if clean_match:
        url = clean_match.group(1)
    else:
        return f"Skipped (Invalid URL format: {url})"

    if "youtube.com" in url or "youtu.be" in url:
        return f"Skipped (Video Link)"

    if 'drive.google.com' in url and '/view' in url:
        file_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if file_id_match:
            url = f"https://drive.google.com/uc?export=download&id={file_id_match.group(1)}"

    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
    ]

    for i, agent in enumerate(user_agents):
        try:
            headers = {'User-Agent': agent}
            response = requests.get(url, headers=headers, timeout=20)
            
            if response.status_code == 403 and i < len(user_agents) - 1: continue
            if response.status_code != 200: 
                return f"Failed (HTTP {response.status_code})"

            ct = response.headers.get('Content-Type', '').lower()
            if 'text/html' in ct:
                return f"Skipped (Target is a Webpage, not Image)"

            img = Image.open(io.BytesIO(response.content))
            output_buffer = io.BytesIO()
            img.save(output_buffer, format="WEBP", quality=80)
            return output_buffer.getvalue()
        except UnidentifiedImageError: return "Failed (Not an image file)"
        except Exception as e:
            if i == len(user_agents) - 1: return str(e)
    return "Failed (Unknown)"

# ==========================================
# 5. Hyperlink Extractor
# ==========================================

def extract_hyperlinks_from_docx(file_obj):
    hyperlinks_map = {}
    try:
        file_obj.seek(0)
        doc = Document(file_obj)
        rels = doc.part.rels
        for paragraph in doc.paragraphs:
            para_text = paragraph.text.strip()
            xml = paragraph._element.xml
            target_tags = ['[hero image]', '[image]', '[promo']
            found_url = None
            
            if 'w:hyperlink' in xml:
                for item in paragraph._element.xpath('.//w:hyperlink'):
                    try:
                        rid = item.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if rid and rid in rels:
                            url = rels[rid].target_ref
                            text_nodes = item.findall('.//w:t', namespaces=item.nsmap)
                            txt = "".join([node.text for node in text_nodes if node.text]).strip()
                            if txt: 
                                hyperlinks_map[txt] = url
                                hyperlinks_map[txt.replace(' ', '')] = url
                            if not found_url: found_url = url
                    except: continue
            
            if found_url:
                for tag in target_tags:
                    if tag in para_text.lower():
                        clean_val = re.sub(r'\[.*?\]', '', para_text).strip()
                        if clean_val:
                            hyperlinks_map[clean_val] = found_url
                            hyperlinks_map[clean_val.replace(' ', '')] = found_url
    except: pass
    file_obj.seek(0)
    return hyperlinks_map

# ==========================================
# 6. Extractors & Parsers
# ==========================================

def extract_nav_items_from_lines(lines):
    items = []
    for line in lines:
        clean = clean_nav_text(line)
        if clean and "Navigation (" not in clean and "Dev Note" not in clean: items.append(clean)
    return items

def generate_nav_yaml(lines):
    nav_structure = {}
    current_parent = None
    list_pattern = re.compile(r'^[\s]*[\-\*]\s+') 
    for line in lines:
        stripped = line.strip()
        if not stripped or "Dev Note" in stripped: continue
        is_child = list_pattern.match(stripped)
        clean_text = clean_nav_text(stripped)
        if not clean_text: continue
        if is_child:
            if current_parent: nav_structure[current_parent].append(clean_text)
        else:
            current_parent = clean_text
            if current_parent not in nav_structure: nav_structure[current_parent] = []

    yaml_output = ""
    for parent in nav_structure:
        children = nav_structure[parent]
        parent_kebab = to_kebab_case(parent)
        parent_href = "/contact-us/" if "contact" in parent_kebab else "#"
        yaml_output += f"- text: {parent}\n  href: \"{parent_href}\"\n"
        if not children:
            yaml_output += "\n"
            continue
        yaml_output += "  dropdown:\n"
        chunks = [children[i:i + 8] for i in range(0, len(children), 8)]
        for chunk in chunks:
            yaml_output += "    - title:\n      links:\n"
            for child in chunk:
                kebab = to_kebab_case(child)
                href = f"/services/{kebab}/"
                if parent == "Promotions": href = "/promotions/"
                elif parent == "About Us": href = f"/{kebab}/"
                yaml_output += f"        - text: {child}\n          href: {href}\n"
        yaml_output += "\n"
    return yaml_output

def generate_reviews_yaml(raw_text):
    raw_text = re.sub(r'^#.*', '', raw_text, flags=re.MULTILINE)
    normalized_text = raw_text.replace('**', '\n')
    lines = normalized_text.split('\n')
    reviews_data = []
    buffer = [] 
    for line in lines:
        stripped = line.strip()
        if not stripped: continue
        clean_line = re.sub(r'\[([^\]]+)\]\(.*?\)', r'\1', stripped).strip() 
        match = re.search(r'(.*?)\s*\(([^)]+)\)$', clean_line)
        if match:
            pre_paren = match.group(1).strip()
            service = match.group(2).strip()
            review_text = ""
            source = ""
            split = re.search(r'(.*[.!?"])\s+(.*)', pre_paren)
            if split:
                source = split.group(2).strip()
                review_text = " ".join(buffer + [split.group(1).strip()])
            else:
                words = pre_paren.split()
                if len(words) < 5:
                    source = pre_paren
                    review_text = " ".join(buffer)
                else:
                    review_text = " ".join(buffer + [pre_paren])
            
            review_text = review_text.strip().strip('"').strip("'").replace("**", "")
            if review_text:
                reviews_data.append({
                    "service_type": None, "text": review_text, 
                    "source": source.replace("**", ""), "service": f"({service.replace('**', '')})"
                })
            buffer = [] 
        else: buffer.append(clean_line)
    return yaml.dump({"reviews": reviews_data}, allow_unicode=True, sort_keys=False, default_flow_style=False)

def generate_services_yaml(service_box_data):
    if not service_box_data: return ""
    yaml_obj = {
        'services': {
            'heading': service_box_data['heading'],
            'sub_heading': service_box_data['sub_heading'],
            'variant': 'image',
            'cards_data': []
        }
    }
    for card in service_box_data['cards']:
        yaml_obj['services']['cards_data'].append({
            'title': card['title'],
            'permalink': f"/services/{card['slug']}/",
            'image_position': '[center_10%]'
        })
    return yaml.dump(yaml_obj, allow_unicode=True, sort_keys=False, default_flow_style=False)

def clean_body_line(text):
    text = text.strip().replace('\\_', '_').replace('\\[', '[').replace('\\]', ']')
    return text

def extract_tag_value(line, tag_name, extract_url=False):
    pattern = re.compile(rf'\[{tag_name}\].*?(\S.*)', re.IGNORECASE)
    match = pattern.search(line.replace('\\', ''))
    if match:
        raw_val = match.group(1).strip()
        if extract_url:
            link_match = re.search(r'\[.*?\]\((http[^)]+)\)', raw_val)
            if link_match: return link_match.group(1)
            url_match = re.search(r'(https?://[^\s<>")\]]+)', raw_val)
            if url_match: return url_match.group(1)
        return clean_markdown_link(raw_val)
    return None

def parse_page_content(raw_text, page_title, image_queue, hyperlink_map):
    lines = raw_text.split('\n')
    clean_title = page_title.replace("#", "").replace(" Page", "").strip()
    page_slug = to_kebab_case(clean_title)
    
    data = { 
        'title': clean_title, 'hero_image': page_slug, 'subheader': '', 
        'ctas': [], 'promos': [], 'body_lines': [] 
    }
    service_box = None
    is_body_started = False
    is_service_box = False 
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if is_body_started and not is_service_box and data['body_lines'] and data['body_lines'][-1] != "": 
                data['body_lines'].append("") 
            continue
        line_lower = stripped.lower().replace('\\', '') 
        
        if 'how can we help' in line_lower and '##' in stripped:
            is_service_box = True
            clean_head = clean_body_line(stripped.replace('##', ''))
            service_box = { 'heading': clean_head, 'sub_heading': '', 'cards': [] }
            continue

        if is_service_box:
            if (stripped.startswith('# ') or stripped.startswith('## ')) and 'how can we help' not in line_lower:
                is_service_box = False
            else:
                if '[p]' in line_lower:
                    sub = re.sub(r'\[p\]', '', stripped, flags=re.IGNORECASE).strip()
                    service_box['sub_heading'] = sub
                    continue
                if stripped and '[' not in stripped:
                    card_title = clean_body_line(stripped)
                    card_slug = to_kebab_case(card_title)
                    service_box['cards'].append({'title': card_title, 'slug': card_slug})
                continue 

        if not is_body_started:
            if stripped.startswith('##'): is_body_started = True 
            else:
                is_h1 = stripped.startswith('# ')
                is_title = clean_nav_text(stripped).lower() == data['title'].lower()
                if is_h1 or is_title: continue

        if '[hero image]' in line_lower:
            text_val = extract_tag_value(stripped, 'hero image', extract_url=True)
            if text_val:
                final_url = text_val.strip()
                if not final_url.startswith(('http', 'https')):
                    if final_url in hyperlink_map: final_url = hyperlink_map[final_url]
                    else: 
                        for k, v in hyperlink_map.items():
                            if final_url in k or k in final_url: 
                                final_url = v
                                break
                image_queue.append({'url': final_url, 'filename': page_slug})
            continue

        if '[para_subheader]' in line_lower:
            val = extract_tag_value(stripped, 'para_subheader')
            if val: data['subheader'] = val.replace('**', '')
            continue
        
        if '[promo' in line_lower or '[hero_promo]' in line_lower:
            promo = re.sub(r'\[.*?promo.*?\]', '', stripped, flags=re.IGNORECASE)
            data['promos'].append(clean_nav_text(promo).replace('**', ''))
            continue

        if '[cta' in line_lower:
            val = extract_tag_value(stripped, 'cta') or extract_tag_value(stripped, 'cta_1') or extract_tag_value(stripped, 'cta_2')
            if val:
                cta_text = val.replace('{', '').replace('}', '').strip().replace('**', '') 
                if DEFAULT_PHONE_NUMBER in cta_text:
                    display, link, icon, scheme, rev = "{{ site.phone }}", "tel:{{ site.phone }}", "phone", "accent", "false"
                else:
                    display, link, icon, scheme, rev = cta_text, "{{ site.contact_page }}", "mark_email_unread", "primary1", "true"
                    if re.search(r'\d{3}[-.\s]?\d{3}', cta_text):
                        link, icon, scheme, rev = f"tel:{{{{ site.phone }}}}", "phone", "accent", "false"
                data['ctas'].append({'text': display, 'link': link, 'icon': icon, 'type': "button-1", 'scheme': scheme, 'reverse': rev})
            continue
        
        if '##' in stripped: is_body_started = True
        
        cleaned = clean_body_line(stripped)
        if re.match(r'^(#+)([^#\s])', cleaned): cleaned = re.sub(r'^(#+)([^#\s])', r'\1 \2', cleaned)
        if cleaned.startswith('#') and data['body_lines'] and data['body_lines'][-1] != "": data['body_lines'].append("")
        
        is_list = re.match(r'^[\-\*]\s+', cleaned) or re.match(r'^\d+\.\s+', cleaned)
        if data['body_lines'] and data['body_lines'][-1] != "":
            prev = data['body_lines'][-1]
            prev_is_list = re.match(r'^[\-\*]\s+', prev) or re.match(r'^\d+\.\s+', prev)
            if not (is_list and prev_is_list) and not cleaned.startswith('#'): data['body_lines'].append("")
            
        data['body_lines'].append(cleaned)

    buttons_str = ""
    for cta in data['ctas']:
        buttons_str += f"""        - cta-link: '{cta['link']}'\n          cta-text: '{cta['text']}'\n          cta-icon: {cta['icon']}\n          cta-type: {cta['type']}\n          cta-color-scheme: {cta['scheme']}"""
        if cta['reverse'] == 'true': buttons_str += "\n          cta-reverse: true"
        buttons_str += "\n"
    
    promos_str = ""
    if data['promos']:
        for p in data['promos']: promos_str += f"    - heading: {p}\n"
    else: promos_str = "    # - heading:\n    #   disclaimer: \n    #   link:"

    final_content = f"""---
layout: post-sidebar
title: {data['title']}
title_override:
category: services
body_class:
show_steps_banner:
# top_review_slider:
#     show_slider: true
#     custom_class: "block lg:hidden"
# bottom_review_slider:
#     hide_slider: 
#     custom_class: "hidden lg:block"

hero:
  variant: split
  image: {data['hero_image']}
  image_position:
  content:
    - type: 'heading'
    # - type: 'lists'
    #   lists: 
    #     - item: Lorem ipsum
    - type: 'paragraph'
      paragraph: {data['subheader']}
    - type: 'cta'
      buttons:
{buttons_str.rstrip()}
  promos:
{promos_str}

hide_promo_carousel:
hide_sidebar_promo: true
hide_sidebar_review: true
hide_sidebar_financing: true
---

"""
    return final_content + '\n'.join(data['body_lines']), service_box

def process_docx(file_obj):
    hyperlink_map = extract_hyperlinks_from_docx(file_obj)
    result = mammoth.convert_to_html(file_obj)
    raw_md = markdownify.markdownify(result.value, heading_style="atx")
    
    sections = {}
    lines = raw_md.split('\n')
    current_page = "00_Ignore" 
    current_content = []
    
    nav_lines = []
    is_capturing_nav = False
    nav_captured = False
    
    for line in lines:
        stripped = line.strip()
        clean = clean_nav_text(stripped)
        if not nav_captured and "Navigation (All Pages)" in clean:
            is_capturing_nav = True
            continue
        if is_capturing_nav:
            if stripped.startswith("# ") or (stripped.startswith("## ") and "Navigation" not in clean):
                is_capturing_nav = False
                nav_captured = True
                sections['Navigation'] = nav_lines
            else: nav_lines.append(stripped)

    target_pages = set()
    if 'Navigation' in sections:
        raw_nav_items = extract_nav_items_from_lines(sections['Navigation'])
        for item in raw_nav_items: target_pages.add(item.lower())

    has_passed_footer = False 
    for line in lines:
        stripped = line.strip()
        clean = clean_nav_text(stripped)
        if START_MARKER in stripped or "Footer (All Pages)" in stripped:
            has_passed_footer = True
            current_page = "00_Ignore" 
            current_content = []
            continue 
        if not has_passed_footer: continue

        is_h1 = stripped.startswith("# ") 
        is_page_keyword = clean.lower().endswith('page')
        is_known = clean.lower() in target_pages
        is_reviews = "customer reviews" in clean.lower()
        
        if (is_h1 or is_page_keyword or is_known or is_reviews) and len(clean) < 80 and '[' not in stripped:
            is_exc = any(exc in clean.lower() for exc in EXCLUDED_KEYWORDS)
            if is_reviews: is_exc = False 
            if not is_exc:
                if current_content and current_page != "00_Ignore": sections[current_page] = "\n".join(current_content)
                current_page = clean.replace('#', '').strip()
                current_content = []
                continue 
        if current_page != "00_Ignore": current_content.append(line)
            
    if current_content and current_page != "00_Ignore": sections[current_page] = "\n".join(current_content)
    return sections, hyperlink_map

# ==========================================
# 7. Main Logic & Display
# ==========================================

if target_file:
    with st.spinner('Parsing Document...'):
        # 1. Parse File
        raw_sections, hyperlink_map = process_docx(target_file)
        
        zip_buffer = io.BytesIO()
        valid_cnt = 0
        reviews_yaml = None
        services_yaml_content = None 
        nav_yaml = None
        image_queue = [] 
        preview_data = {} # For displaying in tabs

        # 2. Generate ZIP Content
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            # A. Navigation
            if 'Navigation' in raw_sections:
                nav_yaml = generate_nav_yaml(raw_sections['Navigation'])
                preview_data['Navigation'] = nav_yaml
            
            # B. Pages & Reviews
            for page_name, content in raw_sections.items():
                if page_name == 'Navigation' or page_name == '00_Ignore': continue 
                if should_skip_page(page_name, content): continue
                
                # Review Section
                if "customer reviews" in page_name.lower():
                    reviews_yaml = generate_reviews_yaml(content)
                    zf.writestr("_data/reviews.yml", reviews_yaml)
                    continue

                # Standard Pages
                final_md, service_box_data = parse_page_content(content, page_name, image_queue, hyperlink_map)
                
                if service_box_data:
                    services_yaml_content = generate_services_yaml(service_box_data)
                    zf.writestr("_data/services.yml", services_yaml_content)

                slug = to_kebab_case(page_name.lower().replace(" page", "").strip())
                fname = f"{FIXED_DATE}-{slug}.md"
                zf.writestr(f"_posts/{fname}", final_md)
                
                # Add to Preview Data
                preview_data[fname] = final_md
                valid_cnt += 1

            # C. Images (This takes time, show progress)
            image_log = []
            if image_queue:
                progress_text = "🖼 Extracting Images..."
                my_bar = st.progress(0, text=progress_text)
                
                for i, img_data in enumerate(image_queue):
                    url, fname = img_data['url'], img_data['filename']
                    res = download_and_convert_image(url)
                    if isinstance(res, bytes): 
                        zf.writestr(f"img/{fname}.webp", res)
                        image_log.append(f"✅ {fname}.webp")
                    else: 
                        image_log.append(f"❌ {fname} -> {res}")
                    my_bar.progress((i + 1) / len(image_queue), text=f"Processing {fname}...")
                
                my_bar.empty()

        zip_buffer.seek(0)
        
        # 3. Display Results (Preview Area)
        st.success("✅ Conversion Complete!")

        # Tabs for better organization
        tab_list = ["📄 Converted Pages", "📂 Data Files"]
        if image_log: tab_list.append("🖼 Image Log")
        
        tabs = st.tabs(tab_list)

        # Tab 1: Converted Markdown Pages
        with tabs[0]:
            if preview_data:
                for fname, content in preview_data.items():
                    if fname == "Navigation": continue
                    with st.expander(f"📝 {fname}"):
                        st.code(content, language='yaml')
            else:
                st.info("No standard pages found.")

        # Tab 2: Data Files (Nav, Reviews, Services)
        with tabs[1]:
            col1, col2 = st.columns(2)
            with col1:
                if nav_yaml:
                    with st.expander("🧭 Navigation (full.yml)", expanded=True):
                        st.code(nav_yaml, language='yaml')
                else: st.warning("No Navigation detected.")
            
            with col2:
                if reviews_yaml:
                    with st.expander("⭐ Reviews (_data/reviews.yml)", expanded=True):
                        st.code(reviews_yaml, language='yaml')
                if services_yaml_content:
                    with st.expander("🛠 Services (_data/services.yml)", expanded=True):
                        st.code(services_yaml_content, language='yaml')

        # Tab 3: Image Log (if exists)
        if image_log:
            with tabs[2]:
                st.write(f"Processed {len(image_queue)} images.")
                st.text("\n".join(image_log))

        # 4. Download Button (At the bottom, pre-loaded with zip_buffer)
        st.divider()
        st.download_button(
            label="📦 Download All Files (ZIP)",
            data=zip_buffer,
            file_name="converted_files.zip",
            mime="application/zip",
            use_container_width=True
        )

else:
    # Empty State Area (Right Side)
    st.write("👈 Please refer to the instructions on the sidebar and upload a file to start.")
    for _ in range(5): st.write("") # Spacer