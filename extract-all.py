import win32com.client
import os

def extract_nsf_data_to_folders(password, nsf_path, output_dir="output"):
    # Create a Lotus Notes session
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(password)
    
    # Open the NSF database; if running locally, the server parameter can be empty
    db = session.GetDatabase("", nsf_path)
    if not db.IsOpen:
        db.Open()
    
    # Ensure the base output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Access all documents in the database
    collection = db.AllDocuments
    doc = collection.GetFirstDocument()
    
    doc_counter = 1
    while doc:
        # Create a folder for this document.
        try:
            doc_id = doc.UniversalID
        except Exception:
            doc_id = f"doc_{doc_counter}"
        doc_folder = os.path.join(output_dir, doc_id)
        if not os.path.exists(doc_folder):
            os.makedirs(doc_folder)
        
        # Write the document's fields to a text file
        text_file_path = os.path.join(doc_folder, "document.txt")
        with open(text_file_path, "w", encoding="utf-8") as f:
            f.write("----- Document -----\n")
            for item in doc.Items:
                f.write(f"{item.Name}: {item.Values}\n")
            f.write("--------------------\n")
        print(f"Saved text document: {text_file_path}")
        
        # Attempt to extract attachments (embedded objects) from items
        for item in doc.Items:
            # Check if the item has an EmbeddedObjects property
            if hasattr(item, "EmbeddedObjects"):
                embedded_objects = item.EmbeddedObjects
                if not embedded_objects:
                    # Sometimes EmbeddedObjects can be None or empty
                    continue
                try:
                    # If the embedded_objects object has a Count attribute, use it as a COM collection
                    if hasattr(embedded_objects, "Count"):
                        count = embedded_objects.Count
                        for i in range(1, count + 1):
                            embedded_obj = embedded_objects.Item(i)
                            attachment_name = embedded_obj.Name
                            attachment_path = os.path.join(doc_folder, attachment_name)
                            try:
                                embedded_obj.ExtractFile(attachment_path)
                                print(f"Extracted attachment '{attachment_name}' to {attachment_path}")
                            except Exception as e:
                                print(f"Failed to extract attachment '{attachment_name}': {e}")
                    # Otherwise, if it is iterable (like a tuple), iterate directly
                    elif hasattr(embedded_objects, "__iter__"):
                        for embedded_obj in embedded_objects:
                            attachment_name = embedded_obj.Name
                            attachment_path = os.path.join(doc_folder, attachment_name)
                            try:
                                embedded_obj.ExtractFile(attachment_path)
                                print(f"Extracted attachment '{attachment_name}' to {attachment_path}")
                            except Exception as e:
                                print(f"Failed to extract attachment '{attachment_name}': {e}")
                    else:
                        print(f"EmbeddedObjects in item '{item.Name}' is not iterable or a COM collection.")
                except Exception as e:
                    print(f"Error processing embedded objects in item '{item.Name}': {e}")
        
        doc_counter += 1
        doc = collection.GetNextDocument(doc)

# Example usage
if __name__ == '__main__':
    # Replace with your actual Lotus Notes password and NSF file path
    #
    extract_nsf_data_to_folders("", "FND-CHHAD-Reference-Libraryl.nsf")
