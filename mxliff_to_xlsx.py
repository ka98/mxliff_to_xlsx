import xlsxwriter
from lxml import etree
import html
import re

workbook = xlsxwriter.Workbook("output.xlsx")
worksheet = workbook.add_worksheet()

# Define formats
fmt_bold_cursive_underlined = workbook.add_format({"bold": True, "italic": True, "underline": True})
fmt_bold_cursive= workbook.add_format({"bold": True, "italic": True})
fmt_bold_underlined = workbook.add_format({"bold": True, "underline": True})
fmt_italic_underlined = workbook.add_format({"italic": True, "underline": True})
fmt_bold = workbook.add_format({"bold": True})
fmt_italic = workbook.add_format({"italic": True})
fmt_underline = workbook.add_format({"underline": True})
fmt_normal = workbook.add_format()

def get_format(fmt_dict):
    if not fmt_dict:
        return fmt_normal
    if fmt_dict.get("bold") and fmt_dict.get("italic") and fmt_dict.get("underline"):
        return fmt_bold_cursive_underlined 
    if fmt_dict.get("bold") and fmt_dict.get("italic"):
        return fmt_bold_cursive
    if fmt_dict.get("bold") and fmt_dict.get("underline"):
        return fmt_bold_underlined
    if fmt_dict.get("italic") and fmt_dict.get("underline"):
        return fmt_italic_underlined
    if fmt_dict.get("bold"):
        return fmt_bold
    if fmt_dict.get("italic"):
        return fmt_italic
    if fmt_dict.get("underline"):
        return fmt_underline
    return fmt_normal

def write_rich(cell_row, cell_col, text):
    if text is None or text == "":
        worksheet.write(cell_row, cell_col, "")
        return
    text = normalize_tags(text)
    segments = parse_mxliff_format(text)

    rich_parts = []

    for seg_text, fmt in segments:
        rich_parts.append(get_format(fmt))
        rich_parts.append(seg_text)
        
    if len(rich_parts) <= 2:
        worksheet.write_string(cell_row, cell_col, rich_parts[1], rich_parts[0])
    else:  
        worksheet.write_rich_string(cell_row, cell_col, *rich_parts)

def normalize_tags(text):
    text = html.unescape(text)
    text = text.replace("{1}", "").replace("{2}", "")
    text = text.replace("{3}", "").replace("{4}", "")
    text = text.replace("{5}", "").replace("{6}", "")
    text = text.replace("{7}", "").replace("{8}", "")
    text = text.replace("{9}", "").replace("{0}", "")
    text = text.replace("{j}", "")
    text = text.replace("{1>", "").replace("{2>", "")
    text = text.replace("{3>", "").replace("{4>", "")
    text = text.replace("{5>", "").replace("{6>", "")
    text = text.replace("{7>", "").replace("{8>", "")
    text = text.replace("{9>", "").replace("{0>", "")
    text = text.replace("<1}", "").replace("<2}", "")
    text = text.replace("<3}", "").replace("<4}", "")
    text = text.replace("<5}", "").replace("<6}", "")
    text = text.replace("<7}", "").replace("<8}", "")
    text = text.replace("<9}", "").replace("<0}", "")
    return text

def parse_mxliff_format(text):
    html.unescape(text)
    
    tokens = re.split(r'(\{[^}]+\})', text)

    segments = []
    flags = {}
    
    innerText = text 
    
    for token in tokens:
        if token.startswith("{") and token.endswith("}"):
            tag = token.strip("{}")
            tagText = tag.split(">")
            tag = tagText[0]
            innerText= tagText[1].split("<")[0]
            flags = {
                "bold": "b" in tag,
                "italic": "i" in tag,
                "underline": "u" in tag,
            }
            segments.append((innerText, flags))
        elif token != "":
            segments.append((token, {}))
            

    return segments

def mxliff_to_excel(input_file, output_file):
    tree = etree.parse(input_file)
    root = tree.getroot()
    
    print("Namespaces:", root.nsmap)


    data = []

    # Namespace handling (common in mxliff)
    ns = {"x": root.nsmap.get(None)}
    ns_m = root.nsmap.get("m")

    for unit in root.xpath("//x:trans-unit", namespaces=ns):
        source = unit.find("x:source", namespaces=ns)
        target = unit.find("x:target", namespaces=ns)
        altTransList = []
        for altTrans in unit.findall("x:alt-trans", namespaces=ns):
            altTarget = altTrans.find("x:target", namespaces=ns)
            altTransList.append({
                "origin": altTrans.get("origin"),
                "match-quality": altTrans.get("match-quality"),
                "text": altTarget.text if altTarget.text is not None else "",
                "mt-id": altTrans.get(f"{{{ns_m}}}mt-id"),
                "tm-id": altTrans.get("tm-id"),
            })
            
        score = unit.get(f"{{{ns_m}}}score")
        gross_score = unit.get(f"{{{ns_m}}}gross-score")
        trans_origin = unit.get(f"{{{ns_m}}}trans-origin") 

        data.append({
            "ID": unit.get("id"),
            "Source": source.text if source is not None else "",
            "Target": target.text if target is not None else "",
            "score": score,
            "gross-score": gross_score,
            "trans-origin": trans_origin,
            "alt-trans1-origin": altTransList[0]["origin"] if len(altTransList) > 0 else "",
            "alt-trans1-match-quality": altTransList[0]["match-quality"] if len(altTransList) > 0 else "",
            "alt-trans1-text": altTransList[0]["text"] if len(altTransList) > 0 else "",
            "alt-trans1-mt-id": altTransList[0]["mt-id"] if len(altTransList) > 0 else "",
            "alt-trans2-origin": altTransList[1]["origin"] if len(altTransList) > 1 else "",
            "alt-trans2-match-quality": altTransList[1]["match-quality"] if len(altTransList) > 1 else "",
            "alt-trans2-text": altTransList[1]["text"] if len(altTransList) > 1 else "",
            "alt-trans2-tm-id": altTransList[1]["tm-id"] if len(altTransList) > 1 else "",
        })

    # Write headers
    headers = list(data[0].keys())
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Write data
    for row_num, row_data in enumerate(data, start=1):
        for col_num, value in enumerate(row_data.values()):
            write_rich(row_num, col_num, value)

    workbook.close()

# Example usage
mxliff_to_excel("input.mxliff", "output.xlsx")