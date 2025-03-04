import win32com.client
import os
import re

NSF_PATH = "FND-CHHAD-Reference-Libraryl.nsf"
LOTUS_PASSWORD = ""  # If needed
VIEW_NAME = "English\\Document\\By Category"  # Double-check exact name!
OUTPUT_DIR = "output_blended_debug"

# Max length for folder names on Windows
MAX_FOLDER_NAME_LENGTH = 100

def sanitize_folder_name(name, max_length=MAX_FOLDER_NAME_LENGTH):
    if not name or not name.strip():
        return "UnnamedDocument"
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'[\s_]+', '_', name)
    return name[:max_length].strip('_')

def get_document_subject(doc):
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
    subject = get_document_subject(doc)
    try:
        doc_id = doc.UniversalID[:8]
    except Exception:
        doc_id = "unknown"

    doc_folder_name = sanitize_folder_name(f"{subject}_{doc_id}")
    doc_folder_path = os.path.join(folder_path, doc_folder_name)
    os.makedirs(doc_folder_path, exist_ok=True)

    # Write fields
    text_file_path = os.path.join(doc_folder_path, "document.txt")
    with open(text_file_path, "w", encoding="utf-8") as f:
        f.write(f"----- Document: {subject} ({doc_id}) -----\n")
        for item in doc.Items:
            try:
                f.write(f"{item.Name}: {item.Values}\n")
            except Exception as e:
                f.write(f"{item.Name}: <Error reading value: {e}>\n")
        f.write("--------------------\n")

    # Attachments
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

def gather_view_categories(db, view_name):
    """
    Build a dict: doc_id -> [ [catPath1], [catPath2], ... ]
    Each catPath is a list of category strings.
    """
    doc_id_to_paths = {}

    print(f"\n[DEBUG] Attempting to open view: '{view_name}'")
    view = db.GetView(view_name)
    if not view:
        print(f"[DEBUG] View '{view_name}' not found. Returning empty mapping.")
        return doc_id_to_paths

    all_entries = view.AllEntries
    entry_count = 0

    entry = all_entries.GetFirstEntry()
    current_path = []
    while entry:
        entry_count += 1
        next_entry = all_entries.GetNextEntry(entry)

        if entry.IsCategory:
            cat_name = entry.ColumnValues[0] or "Uncategorized"
            # Adjust current_path based on the category level
            while len(current_path) >= entry.Level:
                current_path.pop()
            current_path.append(cat_name)

        elif entry.IsDocument:
            doc = entry.Document
            if doc is not None:
                uid = doc.UniversalID
                if uid:
                    if uid not in doc_id_to_paths:
                        doc_id_to_paths[uid] = []
                    doc_id_to_paths[uid].append(list(current_path))  # copy current category path

        entry = next_entry

    print(f"[DEBUG] View '{view_name}' opened. Entries found: {entry_count}")
    print(f"[DEBUG] Documents found in this view: {len(doc_id_to_paths)} unique doc IDs.\n")
    return doc_id_to_paths

def blended_export(password, nsf_path, view_name, output_dir="output"):
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(password)

    db = session.GetDatabase("", nsf_path)
    if not db.IsOpen:
        raise Exception(f"Unable to open NSF at '{nsf_path}'")

    print("[DEBUG] Database opened successfully.")
    os.makedirs(output_dir, exist_ok=True)

    # 1) Gather categories from the view
    doc_id_to_paths = gather_view_categories(db, view_name)

    # 2) Iterate all docs in the DB
    all_docs = db.AllDocuments
    doc = all_docs.GetFirstDocument()
    doc_count = 0
    fallback_count = 0
    view_count = 0

    while doc:
        doc_id_full = doc.UniversalID
        # Some docs might not have a valid UniversalID
        # but that's rare unless they are design docs.
        doc_id_full = doc_id_full or "UNKNOWN_UNID"

        next_doc = all_docs.GetNextDocument(doc)

        # If doc ID is in our dictionary, use the view-based paths
        if doc_id_full in doc_id_to_paths:
            cat_path_list = doc_id_to_paths[doc_id_full]
            # cat_path_list is a list of category paths, e.g. [ ["CatA"], ["CatB","SubB"] ]
            if cat_path_list:
                for cat_path in cat_path_list:
                    # Sanitize each category part
                    sanitized_parts = [sanitize_folder_name(x) for x in cat_path if x.strip()]
                    folder_path = os.path.join(output_dir, *sanitized_parts)
                    os.makedirs(folder_path, exist_ok=True)

                    extract_document(doc, folder_path)
                view_count += 1
            else:
                # If there's an empty category path from the view, fallback
                fallback_count += 1
                category_values = doc.GetItemValue("Category")
                if not category_values:
                    category_values = ["Uncategorized"]
                for cat in category_values:
                    cat = cat.strip() or "Uncategorized"
                    parts = cat.split("\\")
                    parts = [sanitize_folder_name(p) for p in parts if p.strip()]
                    if not parts:
                        parts = ["Uncategorized"]
                    folder_path = os.path.join(output_dir, *parts)
                    os.makedirs(folder_path, exist_ok=True)
                    extract_document(doc, folder_path)
        else:
            # 3) Fallback to doc's 'Category' field
            fallback_count += 1
            category_values = doc.GetItemValue("Category")
            if not category_values:
                category_values = ["Uncategorized"]
            for cat in category_values:
                cat = cat.strip() or "Uncategorized"
                parts = cat.split("\\")
                parts = [sanitize_folder_name(p) for p in parts if p.strip()]
                if not parts:
                    parts = ["Uncategorized"]
                folder_path = os.path.join(output_dir, *parts)
                os.makedirs(folder_path, exist_ok=True)
                extract_document(doc, folder_path)

        doc_count += 1
        doc = next_doc

    print("\n[DEBUG] Finished blended export.")
    print(f"[DEBUG] Total documents processed: {doc_count}")
    print(f"[DEBUG] Documents using view-based categories: {view_count}")
    print(f"[DEBUG] Documents using fallback category field: {fallback_count}\n")

if __name__ == '__main__':
    blended_export(LOTUS_PASSWORD, NSF_PATH, VIEW_NAME, OUTPUT_DIR)
