import xml.etree.ElementTree as ET
import pandas as pd

# Load and parse the .layout file
file_path = 'Opportunity-CM 商談 ページレイアウト%EF%BC%88フィード%EF%BC%89.layout'  # Replace with your actual .layout file path
tree = ET.parse(file_path)
root = tree.getroot()

# Get namespace
namespace = {'ns': root.tag.split('}')[0].strip('{')}

# Initialize lists to store data
sections = []

# Loop through layout sections and check if data is being extracted
for section in root.findall("ns:layoutSections", namespaces=namespace):
    section_label = section.find("ns:label", namespaces=namespace).text if section.find("ns:label", namespaces=namespace) is not None else "No Label"
    style = section.find("ns:style", namespaces=namespace).text if section.find("ns:style", namespaces=namespace) is not None else "No Style"
    
    # Debug: Print section details
    print(f"Processing Section: {section_label}, Style: {style}")
    
    for column in section.findall("ns:layoutColumns", namespaces=namespace):
        for item in column.findall("ns:layoutItems", namespaces=namespace):
            field_name = item.find("ns:field", namespaces=namespace).text if item.find("ns:field", namespaces=namespace) is not None else "No Field"
            behavior = item.find("ns:behavior", namespaces=namespace).text if item.find("ns:behavior", namespaces=namespace) is not None else "No Behavior"
            
            # Debug: Print field and behavior details
            print(f"Field: {field_name}, Behavior: {behavior}")
            
            # Append row data
            sections.append({
                "Section Label": section_label,
                "Style": style,
                "Field Name": field_name,
                "Behavior": behavior
            })

# Check if any data was added to sections
if not sections:
    print("No data was extracted. Check the XML structure.")

# Convert to DataFrame and write to Excel if data exists
if sections:
    df = pd.DataFrame(sections)
    output_file = "layout_data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data has been successfully written to {output_file}")
else:
    print("Data extraction failed; no rows to write.")
