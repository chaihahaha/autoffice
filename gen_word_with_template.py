import re
import docx
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

master_doc = Document('master_template.docx')
empty_table_doc = Document('empty_table.docx')
N = 1

# Function to append a template's content into the master document
def append_template_content(src_doc, dest_doc):
    for element in src_doc.element.body:
        # Import the element from the source doc into the destination doc
        dest_doc.element.body.append(element)
    return

def replace_text_in_docx(doc_path, output_doc_path, entry_list, pattern):
    # Load the template document
    doc = Document(doc_path)

    # Function to replace text in a given XML element
    def replace_text_in_element(element, entry):
        for child in element.iter():
            if child.tag == qn('w:t'):  # A text element
                # Apply regex replacement
                if pattern.search(child.text):
                    child.text = pattern.sub(f"{entry}", child.text)

    # Iterate over the entry pairs to create a modified document for each pair
    for entry in entry_list:
        # Create a copy of the loaded document
        new_doc = Document(doc_path)

        # Replace text in the main document body
        replace_text_in_element(new_doc.element.body, entry)

        global N
        # Save the modified document with a unique name
        new_file_name = f"{N}_{entry}.docx"
        N += 1
        new_doc.save(os.path.join(output_doc_path, new_file_name))
        append_template_content(new_doc, master_doc)
        append_template_content(empty_table_doc, master_doc)
        print(f"Saved: {new_file_name}")
