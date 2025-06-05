import win32com.client

# Configuration
LOTUS_PASSWORD = ""  # Provide your Lotus Notes password if required

def list_nsf_databases():
    """Enumerates all available NSF databases on the Lotus Notes server."""
    
    # Initialize Lotus Notes session
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(LOTUS_PASSWORD)  # Provide password if required

    # Get current server
    server = session.CurrentDatabase.Server

    # Get database directory and list all databases
    db_directory = session.GetDbDirectory(server)
    db_list = db_directory.ListDbs()

    # Print available databases
    print("\nAvailable NSF Databases on Server:")
    for db in db_list:
        print(f"- Title: {db.Title}\n  File Path: {db.FilePath}\n")

# Run the function
try:
    list_nsf_databases()
except Exception as e:
    print(f"Error: {e}")
