from docx import Document
import re
import json
import io
from azure.storage.blob import BlobServiceClient
import sys

storage_connection_string = "endpoint=https://sdi-virtual-architect-communicationservice-uat.communication.azure.com/;accesskey=ENpzcKuinRIU+zrXemEIvxm9bLpQ06yWyp02qj0n5CYIwksoU+IwHWyWbOn1s6y9YkDQQOwU/f23rJA9DGEIlA=="
container_name = "proposal-draft-template"
filename = sys.argv[2]
start_keyword = "Table of contents"
end_keyword = "List of tables"

def extract_table_of_contents_from_blob(storage_connection_string, container_name, filename, start_keyword, end_keyword):
    """
    Extracts the table of contents from a Word document stored in Azure Blob Storage.

    Args:
        storage_connection_string (str): Azure Storage connection string.
        container_name (str): Azure Blob Storage container name.
        filename (str): The filename to search for in Azure Blob Storage.
        start_keyword (str): The keyword indicating the start of the table of contents.
        end_keyword (str): The keyword indicating the end of the table of contents.

    Returns:
        list: List of strings representing the table of contents.
    """
    blob_service_client = BlobServiceClient.from_connection_string(storage_connection_string)
    blob_client = blob_service_client.get_blob_client(container=container_name, blob=filename)

    # Download blob content
    blob_data = blob_client.download_blob()
    doc_content = blob_data.readall().decode("utf-8")

    doc = Document(io.StringIO(doc_content))
    
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
    """
    Removes unwanted characters from the data.

    Args:
        data (dict, list, str): The data to be processed.

    Returns:
        dict, list, str: Processed data with unwanted characters removed.
    """
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
        # Replace \u2013 with a hyphen
        data = data.replace('\u2013', '-')

    return data

def frame_toc_as_json(toc_content):
    toc_json = {"data": []}
    section_stack = []

    for line in toc_content:
        line = line.strip()

        if not (line.startswith("Document") or line.startswith("Table of contents")) and not line.startswith("Appendix"):
            match = re.match(r'(\S+)(?:\s(.+)|$)', line)
            if match:
                entry_id = match.group(1).rstrip('.')
                if entry_id[0].isdigit():
                    entry_name = match.group(2) or entry_id
                else:
                    entry_id = ""  # Set entry_id to an empty string or the parent section's ID
                    entry_name = line

                entry = {"id": entry_id, "name": entry_name, "subsections": []}

                if entry_id == "":
                    # If entry_id is empty, add the entry as a subsection of its parent section
                    if section_stack:
                        section_stack[-1]["subsections"].append(entry)
                else:
                    while section_stack:
                        parent_id = section_stack[-1]["id"]
                        if entry_id.startswith(parent_id):
                            section_stack[-1]["subsections"].append(entry)
                            section_stack.append(entry)
                            break
                        else:
                            section_stack.pop()

                    if not section_stack:
                        toc_json["data"].append(entry)
                        section_stack.append(entry)

    return toc_json



toc_content = extract_table_of_contents_from_blob(storage_connection_string, container_name, filename, start_keyword, end_keyword)

if toc_content:
    # Process the JSON data to remove unwanted characters
    cleaned_toc_content = remove_unwanted_chars(toc_content)
    toc_json = frame_toc_as_json(cleaned_toc_content)
    
    # Modify success response structure
    success_response = {
        "data": toc_json["data"],
        "status_info": {
            "status": "pass",
            "status_code": {
                "$numberLong": "200"
            }
        }
    }
    print(json.dumps(success_response, indent=4))
else:
    # Add failure response structure
    response = {"data": [], "status_info": {"status": "fail", "status_code": {"$numberLong": "400"}}}
    print(json.dumps(response, indent=4))
