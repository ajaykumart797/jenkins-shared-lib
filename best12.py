from docx import Document
import re
import json

def extract_table_of_contents(doc_path, start_keyword, end_keyword):
    doc = Document(doc_path)
    toc_found = False
    toc_content = []
    for paragraph in doc.paragraphs:
        if toc_found and end_keyword in paragraph.text:
            break
        if toc_found:
            toc_content.append(paragraph.text)
        if start_keyword in paragraph.text:
            toc_found = True

    return toc_content






def remove_unwanted_chars(data):
    if isinstance(data, dict):
        for key, value in data.items():
            data[key] = remove_unwanted_chars(value)
    elif isinstance(data, list):
        for i in range(len(data)):
            data[i] = remove_unwanted_chars(data[i])
    elif isinstance(data, str):
        # Remove leading and trailing whitespaces
        data = data.strip()
        # Remove dot (.) if it exists at the beginning of the name
        data = re.sub(r'^\.', '', data)
        # Remove \u202f character
        data = data.replace('\u202f', '')
        # Remove \u200b character
        data = data.replace('\u200b', '')
        # Remove \u2019 character
        data = data.replace('\u2019', "'")
        # Remove \t followed by numbers
        data = re.sub(r'\t\d+', '', data)
        # Replace multiple spaces with a single space
        data = re.sub(r'\s+', ' ', data)
        
        data = data.replace('\u2013', '-')  

    return data


def frame_toc_as_json(toc_content):
    toc_json = {"data": []}
    section_stack = []
    appendix_stack = []

    for line in toc_content:
        line = line.strip()

        if line.startswith("Appendix "):
            entry_parts = re.match(r'Appendix (\S+)(?:\s(.+)|(.+)|$)', line)

            if entry_parts:
                appendix_id = entry_parts.group(1)
                entry_name = entry_parts.group(2) or entry_parts.group(3) or ""

                # Apply conditions for id and name in the Appendix main section
                if appendix_id[0].isdigit() or appendix_id[-1].isdigit():
                    full_appendix_id = f"Appendix {appendix_id}"
                    entry_name = entry_name
                else:
                    full_appendix_id = ""
                    if not entry_name:
                        entry_name = "Appendix " + appendix_id

                entry = {"id": full_appendix_id, "name": entry_name, "subsections": []}

                if not appendix_stack:
                    toc_json["data"].append(entry)
                else:
                    appendix_stack[-0]["subsections"].append(entry)

                appendix_stack.append(entry)

        elif not (line.startswith("Document") or line.startswith("Table of contents")):
            match = re.match(r'(\S+)(?:\s(.+)|$)', line)

            if match:
                # entry_id = match.group(1)
                entry_id = match.group(1).rstrip('.')
                entry_name = match.group(2) or ""

                # Check if the ID starts with a number
                if entry_id[0].isdigit() or entry_id[-1].isdigit():
                    entry = {"id": entry_id, "name": entry_name, "subsections": []}
                else:
                    entry = {"id": "", "name": entry_id, "subsections": []}

                while section_stack:
                    parent_id = section_stack[-1]["id"]
                    if entry_id.startswith(parent_id):
                        if not entry_name:
                            entry_name = match.group(2) or entry_id
                        entry = {"id": entry_id, "name": entry_name, "subsections": []}
                        section_stack[-1]["subsections"].append(entry)
                        section_stack.append(entry)
                        break
                    else:
                        section_stack.pop()

                if not section_stack:
                    entry = {"id": entry_id, "name": entry_name, "subsections": []}
                    toc_json["data"].append(entry)
                    section_stack.append(entry)

    return toc_json


# Example usage
doc_path = "C:/Users/ajay.kumart/jenkins-shared-lib/vars/Technology Modernization - Campus Network Solution NTT_Proposal (VA89406683) V1.0 (1).docx"
start_keyword = "Table of contents"
end_keyword = "List of tables"

toc_content = extract_table_of_contents(doc_path, start_keyword, end_keyword)

if toc_content:
    # Process the JSON data to remove unwanted characters
    cleaned_toc_content = remove_unwanted_chars(toc_content)
    toc_json = frame_toc_as_json(cleaned_toc_content)
    print(json.dumps(toc_json, indent=4))
else:
    print("Table of Contents not found in the document.")












