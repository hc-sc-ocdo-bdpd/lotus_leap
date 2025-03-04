import win32com.client
import os
import re

NSF_PATH = "FND-CHHAD-Reference-Libraryl.nsf"
LOTUS_PASSWORD = ""  # If needed
OUTPUT_DIR = "output_all_views_categories"

# Which column index to parse for backslash-delimited categories?
# If your category is in the first column, use 0. If second column, use 1, etc.
CATEGORY_COLUMN_INDEX = 0

MAX_FOLDER_NAME_LENGTH = 100

def sanitize_folder_name(name, max_length=MAX_FOLDER_NAME_LENGTH):
    """Removes invalid characters and truncates for Windows-safe folder names."""
    if not name or not name.strip():
        return "Unnamed"
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'[\s_]+', '_', name)
    return name[:max_length].strip('_')

def get_document_subject(doc):
    """Return 'Subject' field, or fallback to 'Form' if needed."""
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

def extract_document(doc, folder_path):
    """
    Creates a subfolder named after doc subject + short UniversalID,
    writes fields to 'document.txt', and extracts attachments.
    """
    subject = get_document_subject(doc)
    try:
        doc_id = doc.UniversalID[:8]
    except Exception:
        doc_id = "unknown"

    doc_folder_name = sanitize_folder_name(f"{subject}_{doc_id}")
    doc_folder_path = os.path.join(folder_path, doc_folder_name)
    os.makedirs(doc_folder_path, exist_ok=True)

    # Write all fields to a text file
    text_file_path = os.path.join(doc_folder_path, "document.txt")
    with open(text_file_path, "w", encoding="utf-8") as f:
        f.write(f"----- Document: {subject} ({doc_id}) -----\n")
        for item in doc.Items:
            try:
                f.write(f"{item.Name}: {item.Values}\n")
            except Exception as e:
                f.write(f"{item.Name}: <Error reading value: {e}>\n")
        f.write("--------------------\n")

    # Extract attachments (if any)
    for item in doc.Items:
        if hasattr(item, "EmbeddedObjects"):
            embedded_objects = item.EmbeddedObjects
            if not embedded_objects:
                continue
            if hasattr(embedded_objects, "Count"):
                for i in range(1, embedded_objects.Count + 1):
                    embedded_obj = embedded_objects.Item(i)
                    attachment_name = sanitize_folder_name(embedded_obj.Name)
                    attachment_path = os.path.join(doc_folder_path, attachment_name)
                    embedded_obj.ExtractFile(attachment_path)

def extract_all_views_with_categories(password, nsf_path, output_dir="output_all_views_categories"):
    """
    1) Enumerate ALL views in the NSF.
    2) For each view:
       - Create a folder named after the view.
       - For each document entry, parse the first column for a backslash-delimited category path.
       - Extract the doc under that category path, with a subfolder named after the doc's subject + short UID.
    """
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(password)

    db = session.GetDatabase("", nsf_path)
    if not db.IsOpen:
        raise Exception(f"Unable to open NSF at '{nsf_path}'")

    os.makedirs(output_dir, exist_ok=True)

    views = db.Views
    print(f"[INFO] Found {len(views)} views in the database.\n")

    view_count = 0
    for view in views:
        view_name = view.Name
        # Optional: skip hidden/system views
        # if view_name.startswith("(") or view_name.startswith("$"):
        #     continue

        safe_view_name = sanitize_folder_name(view_name)
        view_folder = os.path.join(output_dir, safe_view_name)
        os.makedirs(view_folder, exist_ok=True)

        print(f"[INFO] Processing view '{view_name}' -> folder '{safe_view_name}'")

        all_entries = view.AllEntries
        entry = all_entries.GetFirstEntry()
        doc_count = 0

        while entry:
            next_entry = all_entries.GetNextEntry(entry)
            if entry.IsDocument:
                doc = entry.Document
                if doc:
                    # Get the category path from the specified column
                    col_vals = entry.ColumnValues
                    if len(col_vals) > CATEGORY_COLUMN_INDEX:
                        cat_string = str(col_vals[CATEGORY_COLUMN_INDEX])
                    else:
                        cat_string = ""

                    cat_string = cat_string.strip()
                    if not cat_string:
                        cat_string = "Uncategorized"

                    # Split on backslash for multi-level categories
                    parts = [p.strip() for p in cat_string.split("\\") if p.strip()]
                    if not parts:
                        parts = ["Uncategorized"]

                    # Sanitize each part
                    parts = [sanitize_folder_name(p) for p in parts]

                    # Build final folder path: e.g. output/viewName/CatA/SubCatB
                    final_folder_path = os.path.join(view_folder, *parts)
                    os.makedirs(final_folder_path, exist_ok=True)

                    # Extract the doc
                    extract_document(doc, final_folder_path)
                    doc_count += 1

            entry = next_entry

        print(f"[INFO] Extracted {doc_count} documents from view '{view_name}'\n")
        view_count += 1

    print(f"[DONE] Processed {view_count} views total.")

if __name__ == '__main__':
    extract_all_views_with_categories(LOTUS_PASSWORD, NSF_PATH, OUTPUT_DIR)
