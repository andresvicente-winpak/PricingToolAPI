import zipfile
import os
import tempfile

def process_file(input_path):
    # Create a temporary ZIP file
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, 'output.zip')

    # Write the input file into a zip
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        zipf.write(input_path, arcname=os.path.basename(input_path))

    return zip_path
