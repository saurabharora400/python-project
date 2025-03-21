import pandas as pd  # Importing pandas for handling Excel files
import win32com.client as win32  # Importing win32com.client to interact with Outlook
import os
from docx import Document  # Importing Document from python-docx to handle Word documents
import base64  # Importing base64 to encode images for embedding in emails
import env  # Importing env (assuming it's needed for some environment variables, not used directly in this code)

# Load your Excel file containing the data
excel_file = "book.xlsx"  # Specify the path to your Excel file
df = pd.read_excel(excel_file)  # Read the Excel file into a pandas DataFrame

# Paths to the Word document and PDF document
word_doc_path = "Email1.docx"  # Specify the path to your Word document
pdf_doc_path = "attachment.pdf"  # Specify the path to your PDF attachment

# Function to extract text, images, and hyperlinks from a Word document
def extract_text_images_links(doc_path):
    doc = Document(doc_path)  # Load the Word document
    content_parts = []  # List to store the content extracted from the Word document
    rels = doc.part.rels  # Get relationships in the document (for handling images and hyperlinks)
    
    # Loop through each paragraph in the document
    for para in doc.paragraphs:
        paragraph_xml = para._element  # Get the XML element of the paragraph

        # Loop through each child element in the paragraph
        for child in paragraph_xml:
            if child.tag.endswith('}hyperlink'):
                # If the element is a hyperlink
                rId = child.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if rId and rId in rels:
                    href = rels[rId].target_ref  # Get the hyperlink target
                    link_text = ''.join([node.text for node in child.iter() if node.tag.endswith('}t')])
                    content_parts.append(f'<a href="{href}">{link_text}</a>')  # Add the hyperlink to the content
            elif child.tag.endswith('}r'):
                # If the element is a run (a segment of text with a common set of properties)
                run = child
                run_text = ''
                bold = False
                italic = False
                underline = False
                highlight = None
                font_size = None
                font_name = None

                # Iterate through the nodes in the run to extract text and formatting
                for node in run.iter():
                    if node.tag.endswith('}t'):
                        # Extract the text
                        run_text += node.text
                    elif node.tag.endswith('}blip'):
                        # Handle inline images
                        rId = node.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if rId and rId in rels:
                            image = rels[rId].target_part.blob  # Get the image data
                            img_format = rels[rId].target_part.content_type.split("/")[-1]  # Get the image format
                            img_b64 = base64.b64encode(image).decode('utf-8')  # Encode the image in base64

                            # Find the size of the image in the XML structure
                            ext_lst = node.getparent().findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
                            if ext_lst:
                                cx = ext_lst[0].get('cx')
                                cy = ext_lst[0].get('cy')

                                # Convert EMUs to pixels (fixed values in this example)
                                width_px = int(3147060)/ 9525
                                height_px = int(2098040) / 9525

                                # Add the image to the HTML content
                                img_html = f'<img src="data:image/{img_format};base64,{img_b64}" style="width:{width_px}px; height:{height_px}px;" />'
                                print(f"Image inserted with size: {width_px}px x {height_px}px")
                                content_parts.append(img_html)
                            else:
                                # If no size attributes found, add the image with default settings
                                print("Size attributes not found, using default size.")
                                img_html = f'<img src="data:image/{img_format};base64,{img_b64}" />'
                                content_parts.append(img_html)
                    elif node.tag.endswith('}b'):
                        # Handle bold text
                        bold = True
                    elif node.tag.endswith('}i'):
                        # Handle italic text
                        italic = True
                    elif node.tag.endswith('}u'):
                        # Handle underline text
                        underline = True
                    elif node.tag.endswith('}highlight'):
                        # Handle highlighted text
                        highlight = node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                    elif node.tag.endswith('}sz'):
                        # Handle font size
                        font_size = int(node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")) / 2
                    elif node.tag.endswith('}rFonts'):
                        # Handle font name
                        font_name = node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii")

                # Apply the collected styles to the text
                if run_text:
                    style = []
                    if bold:
                        style.append('font-weight: bold;')
                    if italic:
                        style.append('font-style: italic;')
                    if underline:
                        style.append('text-decoration: underline;')
                    if font_size:
                        style.append(f'font-size: {font_size}pt;')
                    if font_name:
                        style.append(f'font-family: {font_name};')
                    if highlight:
                        style.append(f'background-color: yellow;')

                    style_str = ''.join(style)
                    run_text = f'<span style="{style_str}">{run_text}</span>'
                    content_parts.append(run_text)

        content_parts.append('<br>')  # Add a line break after each paragraph

    return ''.join(content_parts)  # Join all parts into a single HTML string

# Iterate over the rows in the Excel file to generate and send emails
for index, row in df.iterrows():
    content_parts = []
    content_parts.append(f"Dear {row['Name']},<br><br>")
    content_parts.append(f"Thank you for indicating your interest in joining <b>NTU Heritage Club Recruitment Drive 2024!</b><br><br>")
    content_parts.append(f"The following is the details for your <b>allocated timeslot</b> for the interviews:<br>")
    content_parts.append(f"Date:<b>26th August 2024</b><br>")
    content_parts.append(f"Timeslot:<b>{row['Time']}</b><br>")
    content_parts.append(f"Venue: <b> NS TR+3 </b>for registration<br><br>")
    
    # Extract the body content with images and hyperlinks from the Word document
    content_parts.append(extract_text_images_links(word_doc_path))
    email_body_html = ''.join(content_parts)  # Combine all parts into a single HTML string

    recipient_name = row['Name']  # Get recipient's name from the Excel sheet
    recipient_email = row['Email']  # Get recipient's email address from the Excel sheet

    # Initialize Outlook and create an email item
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = "Test Code"  # Set the subject of the email
    mail.To = recipient_email  # Set the recipient email address
    mail.HTMLBody = email_body_html  # Set the HTML body of the email

    # Attach the PDF document
    attachment_path = os.path.abspath(pdf_doc_path)
    mail.Attachments.Add(attachment_path)

    # Send the email
    mail.Send()
    print(f"Email sent to {recipient_name} at {recipient_email}")

print("All emails have been sent.")
