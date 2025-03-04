import win32com.client
import os
import re

# Maximum length for folder names
MAX_FOLDER_NAME_LENGTH = 100

def sanitize_folder_name(name, max_length=MAX_FOLDER_NAME_LENGTH):
    """Sanitize and truncate folder names to be Windows-safe."""
    if not name or not name.strip():
        return "UnnamedDocument"
    # Remove forbidden characters for Windows
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Replace multiple spaces or underscores with a single underscore
    name = re.sub(r'[\s_]+', '_', name)
    # Truncate to max_length and strip leading/trailing underscores
    return name[:max_length].strip('_')

def get_document_subject(doc):
    """Retrieve the subject of a document, falling back to form name if missing."""
    subject = None
    for item in doc.Items:
        if item.Name.lower() == "subject":
            subject = item.Values[0] if item.Values else None
            break
    if not subject:
        for item in doc.Items:
            if item.Name.lower() == "form":
                subject = f"Form_{item.Values[0]}" if item.Values else None
                break
    return subject or "UnnamedDocument"

def get_document_folder_paths(doc):
    """
    Returns a list of folder paths (each as a list of folder parts) for a document,
    based on its hidden "$Folders" field. If none exist, returns a single path for "Uncategorized".
    """
    try:
        folder_names = doc.GetItemValue("$Folders")
    except Exception:
        folder_names = []
    if not folder_names:
        folder_names = ["Uncategorized"]

    folder_paths = []
    for folder in folder_names:
        # If the folder name includes hierarchy delimiters (e.g. backslash), split it
        parts = folder.split("\\")
        # Sanitize each part and ignore empty parts
        sanitized_parts = [sanitize_folder_name(part) for part in parts if part.strip()]
        if sanitized_parts:
            folder_paths.append(sanitized_parts)
        else:
            folder_paths.append(["Uncategorized"])
    return folder_paths

def extract_nsf_data_all_documents(password, nsf_path, output_dir="output"):
    """
    Extracts all documents from the NSF using db.AllDocuments.
    For each document, it uses the "$Folders" field to determine folder membership.
    Documents are placed in folder hierarchies based on their folder names.
    """
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(password)

    db = session.GetDatabase("", nsf_path)
    if not db.IsOpen:
        db.Open()

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    collection = db.AllDocuments
    doc = collection.GetFirstDocument()
    doc_counter = 0

    while doc:
        # Get the next document pointer before processing the current document
        next_doc = collection.GetNextDocument(doc)

        # Determine the folder paths for this document
        folder_paths = get_document_folder_paths(doc)
        # Get the subject to use as the document folder name
        subject = get_document_subject(doc)
        try:
            doc_id = doc.UniversalID[:8]  # Shortened for uniqueness
        except Exception:
            doc_id = "unknown"
        safe_subject = sanitize_folder_name(f"{subject}_{doc_id}")

        # For each folder path the document belongs to, create the full directory structure
        for folder_parts in folder_paths:
            # Build the full path (e.g., output/Folder/Subfolder/...)
            folder_path_full = os.path.join(output_dir, *folder_parts)
            os.makedirs(folder_path_full, exist_ok=True)
            # Create a folder for the document within that folder
            doc_folder = os.path.join(folder_path_full, safe_subject)
            os.makedirs(doc_folder, exist_ok=True)

            # Write document fields to a text file
            text_file_path = os.path.join(doc_folder, "document.txt")
            with open(text_file_path, "w", encoding="utf-8") as f:
                f.write(f"----- Document: {subject} ({doc_id}) -----\n")
                for item in doc.Items:
                    try:
                        # Convert the item's values to string safely
                        f.write(f"{item.Name}: {str(item.Values)}\n")
                    except Exception as e:
                        f.write(f"{item.Name}: <Error reading value: {e}>\n")
                f.write("--------------------\n")
            print(f"Saved document to: {text_file_path}")

            # Extract attachments, if any
            for item in doc.Items:
                if hasattr(item, "EmbeddedObjects"):
                    embedded_objects = item.EmbeddedObjects
                    if not embedded_objects:
                        continue
                    try:
                        if hasattr(embedded_objects, "Count"):
                            for i in range(1, embedded_objects.Count + 1):
                                embedded_obj = embedded_objects.Item(i)
                                attachment_name = sanitize_folder_name(embedded_obj.Name)
                                attachment_path = os.path.join(doc_folder, attachment_name)
                                embedded_obj.ExtractFile(attachment_path)
                                print(f"Extracted attachment to: {attachment_path}")
                        elif hasattr(embedded_objects, "__iter__"):
                            for embedded_obj in embedded_objects:
                                attachment_name = sanitize_folder_name(embedded_obj.Name)
                                attachment_path = os.path.join(doc_folder, attachment_name)
                                embedded_obj.ExtractFile(attachment_path)
                                print(f"Extracted attachment to: {attachment_path}")
                    except Exception as e:
                        print(f"Failed to extract attachment in document {doc_id}: {e}")

        doc_counter += 1
        doc = next_doc  # Move to the next document in the collection

    print(f"Extracted {doc_counter} documents.")

if __name__ == '__main__':
    # Replace with your actual password and NSF file path
    extract_nsf_data_all_documents("", "FND-CHHAD-Reference-Libraryl.nsf")
