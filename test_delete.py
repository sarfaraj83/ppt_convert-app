import os

file_path = 'presentations/e72de8df-dd1a-4f73-af35-10578e1c3bcb.pptx'  # Ensure this file exists
try:
    os.remove(file_path)
    print(f"Deleted {file_path}")
except Exception as e:
    print(f"Error deleting file {file_path}: {e}")
