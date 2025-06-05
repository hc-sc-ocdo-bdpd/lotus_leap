import win32com.client
import os
import re

LOTUS_PASSWORD = ""  # If needed
OUTPUT_DIR = "output_all_dbs"
CATEGORY_COLUMN_INDEX = 0
MAX_FOLDER_NAME_LENGTH = 100

def sanitize_folder_name(name, max_length=MAX_FOLDER_NAME_LENGTH):
    if not name or not name.strip():
        return "Unnamed"
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
    
    text_file_path = os.path.join(doc_folder_path, "document.txt")
    with open(text_file_path, "w", encoding="utf-8") as f:
        f.write(f"----- Document: {subject} ({doc_id}) -----\n")
        for item in doc.Items:
            try:
                value = item.Values
                # Dump only if the value is text (either a string or list of strings)
                if isinstance(value, str):
                    f.write(f"{item.Name}: {value}\n")
                elif isinstance(value, list) and all(isinstance(v, str) for v in value):
                    f.write(f"{item.Name}: {'; '.join(value)}\n")
            except Exception as e:
                f.write(f"{item.Name}: <Error reading value: {e}>\n")
        f.write("--------------------\n")

def extract_document_old(doc, folder_path):
    subject = get_document_subject(doc)
    try:
        doc_id = doc.UniversalID[:8]
    except Exception:
        doc_id = "unknown"
    
    doc_folder_name = sanitize_folder_name(f"{subject}_{doc_id}")
    doc_folder_path = os.path.join(folder_path, doc_folder_name)
    os.makedirs(doc_folder_path, exist_ok=True)
    
    text_file_path = os.path.join(doc_folder_path, "document.txt")
    with open(text_file_path, "w", encoding="utf-8") as f:
        f.write(f"----- Document: {subject} ({doc_id}) -----\n")
        for item in doc.Items:
            try:
                f.write(f"{item.Name}: {item.Values}\n")
            except Exception as e:
                f.write(f"{item.Name}: <Error reading value: {e}>\n")
        f.write("--------------------\n")

def extract_all_objects(password, db, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    print(f"[INFO] Enumerating all objects in the database {db.Title}.")

    try:
        for design_element in db.DesignElements:
            element_name = sanitize_folder_name(design_element.Name)
            element_type = design_element.Type
            print(f"[INFO] Found design element: {element_name} ({element_type})")
    except Exception as e:
        print(f"[ERROR] Failed to enumerate design elements in {db.Title}: {e}")

def extract_all_views_with_categories(password, db, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    views = db.Views
    print(f"[INFO] Found {len(views)} views in the database {db.Title}.")
    
    for view in views:
        view_name = view.Name
        safe_view_name = sanitize_folder_name(view_name)
        view_folder = os.path.join(output_dir, safe_view_name)
        os.makedirs(view_folder, exist_ok=True)
        
        print(f"[INFO] Processing view '{view_name}' -> folder '{safe_view_name}'")
        
        all_entries = view.AllEntries
        entries_list = []
        try:
            entry = all_entries.GetFirstEntry()
        except Exception as e:
            print(f"[ERROR] Failed to get first entry for view '{view_name}': {e}")
            continue
        
        while entry:
            entries_list.append(entry)
            try:
                entry = all_entries.GetNextEntry(entry)
            except Exception as e:
                print(f"[ERROR] Failed to get next entry in view '{view_name}': {e}")
                break
        
        for entry in entries_list:
            if entry.IsDocument:
                doc = entry.Document
                if doc:
                    col_vals = entry.ColumnValues
                    cat_string = (str(col_vals[CATEGORY_COLUMN_INDEX])
                                  if len(col_vals) > CATEGORY_COLUMN_INDEX
                                  else "Uncategorized")
                    
                    parts = [sanitize_folder_name(p.strip()) for p in cat_string.split("\\") if p.strip()]
                    final_folder_path = os.path.join(view_folder, *parts) if parts else os.path.join(view_folder, "Uncategorized")
                    os.makedirs(final_folder_path, exist_ok=True)
                    extract_document(doc, final_folder_path)


def extract_all_views_with_categories_old(password, db, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    views = db.Views
    print(f"[INFO] Found {len(views)} views in the database {db.Title}.")
    
    for view in views:
        view_name = view.Name
        safe_view_name = sanitize_folder_name(view_name)
        view_folder = os.path.join(output_dir, safe_view_name)
        os.makedirs(view_folder, exist_ok=True)
        
        print(f"[INFO] Processing view '{view_name}' -> folder '{safe_view_name}'")
        
        all_entries = view.AllEntries
        entries_list = []
        entry = all_entries.GetFirstEntry()
        while entry:
            entries_list.append(entry)
            entry = all_entries.GetNextEntry(entry)
        
        for entry in entries_list:
            if entry.IsDocument:
                doc = entry.Document
                if doc:
                    col_vals = entry.ColumnValues
                    cat_string = str(col_vals[CATEGORY_COLUMN_INDEX]) if len(col_vals) > CATEGORY_COLUMN_INDEX else "Uncategorized"
                    
                    parts = [sanitize_folder_name(p.strip()) for p in cat_string.split("\\") if p.strip()]
                    final_folder_path = os.path.join(view_folder, *parts) if parts else os.path.join(view_folder, "Uncategorized")
                    os.makedirs(final_folder_path, exist_ok=True)
                    extract_document(doc, final_folder_path)

def enumerate_all_databases(password, output_dir):
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(password)

    # Get all address books (NSF databases)
    dbs = session.AddressBooks  # Includes local and remote NSF files

    print(f"[INFO] Found {len(dbs)} databases in the workspace.")

    for db in dbs:
        if not db.IsOpen:
            db.Open()  # This is an illustrative call; refer to your API documentation.
        db_output_dir = os.path.join(output_dir, sanitize_folder_name(db.Title))
        print(f"[INFO] Processing database: {db.Title}")
        extract_all_objects(password, db, db_output_dir)
        extract_all_views_with_categories(password, db, db_output_dir)

    print("[DONE] Processed all databases.")

if __name__ == '__main__':
    enumerate_all_databases(LOTUS_PASSWORD, OUTPUT_DIR)
