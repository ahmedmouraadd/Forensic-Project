import os
import exifread
import hashlib
import pandas as pd
from PIL import Image
import pillow_heif
import subprocess

# Register HEIF opener with PIL
pillow_heif.register_heif_opener()

# Convert DMS (Degrees, Minutes, Seconds) to Decimal format for GPS coordinates
def dms_to_decimal(dms, direction=None):
    if isinstance(dms, tuple) and len(dms) == 3:
        d, m, s = dms
        decimal = d + m / 60 + s / 3600
    elif isinstance(dms, list) and len(dms) == 3:
        try:
            d = float(dms[0])
            m = float(dms[1])
            s = eval(str(dms[2]))
            decimal = d + m / 60 + s / 3600
        except Exception as e:
            print(f"Error converting list format DMS to decimal: {dms} - {e}")
            return None
    else:
        return None

    if direction in ['S', 'W']:
        decimal = -decimal
    return round(decimal, 6)

# Parse GPS coordinates in different formats, returning in decimal format
def parse_gps(coord_str, is_longitude=False):
    if coord_str:
        if isinstance(coord_str, str) and 'deg' in coord_str:
            parts = coord_str.replace('deg', '').replace("'", "").replace('"', "").split()
            try:
                degrees = float(parts[0])
                minutes = float(parts[1])
                seconds = float(parts[2])
                direction = parts[3] if len(parts) > 3 else None
                return dms_to_decimal((degrees, minutes, seconds), direction)
            except (ValueError, IndexError) as e:
                print(f"Error parsing ExifTool format GPS coordinates: {coord_str} - {e}")
                return None
        elif isinstance(coord_str, str):
            try:
                dms_list = eval(coord_str)
                direction = "W" if is_longitude else "N"
                return dms_to_decimal(dms_list, direction)
            except Exception as e:
                print(f"Error parsing ExifRead format GPS coordinates: {coord_str} - {e}")
                return None
    return None

# Calculate hash of an image file using specified hash type (default: MD5)
def calculate_image_hash(file_path, hash_type="md5"):
    hash_func = hashlib.md5() if hash_type == "md5" else hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_func.update(chunk)
    return hash_func.hexdigest()

# Extract metadata from common image formats (e.g., JPG, PNG)
def extract_metadata_image(file_path):
    metadata = {}
    try:
        with open(file_path, 'rb') as file:
            tags = exifread.process_file(file)
            metadata['File Name'] = os.path.basename(file_path)
            metadata['Date Taken'] = tags.get('EXIF DateTimeOriginal')
            metadata['Camera Model'] = tags.get('Image Model')

            gps_latitude = str(tags.get('GPS GPSLatitude')) if tags.get('GPS GPSLatitude') else None
            gps_longitude = str(tags.get('GPS GPSLongitude')) if tags.get('GPS GPSLongitude') else None
            latitude_direction = str(tags.get('GPS GPSLatitudeRef')).strip() if tags.get('GPS GPSLatitudeRef') else 'N'
            longitude_direction = str(tags.get('GPS GPSLongitudeRef')).strip() if tags.get('GPS GPSLongitudeRef') else 'E'
            
            metadata['Latitude (Decimal)'] = dms_to_decimal(eval(gps_latitude), latitude_direction) if gps_latitude else None
            metadata['Longitude (Decimal)'] = dms_to_decimal(eval(gps_longitude), longitude_direction) if gps_longitude else None
            
            if metadata['Latitude (Decimal)'] is not None and metadata['Longitude (Decimal)'] is not None:
                metadata['Google Maps Link'] = f"https://www.google.com/maps?q={metadata['Latitude (Decimal)']},{metadata['Longitude (Decimal)']}"

            metadata['File Size (Bytes)'] = os.path.getsize(file_path)

            with Image.open(file_path) as img:
                metadata['Width (px)'], metadata['Height (px)'] = img.size

            metadata['MD5 Hash'] = calculate_image_hash(file_path)

    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        metadata = None
    return metadata

# Extract metadata from HEIC image files using ExifTool
def extract_metadata_heic(file_path):
    metadata = {}
    try:
        result = subprocess.run(
            ["C:\\Users\\Dator\\Downloads\\exiftool-13.01_64\\exiftool-13.01_64\\exiftool.exe", file_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        for line in result.stdout.splitlines():
            key, _, value = line.partition(': ')
            key = key.strip()
            value = value.strip()
            
            if key == 'File Name':
                metadata['File Name'] = value
            elif key == 'Date/Time Original':
                metadata['Date Taken'] = value
            elif key in ['Model', 'Camera Model Name']:
                metadata['Camera Model'] = value
            elif key == 'GPS Latitude':
                metadata['Latitude (Decimal)'] = parse_gps(value)
            elif key == 'GPS Longitude':
                metadata['Longitude (Decimal)'] = parse_gps(value, is_longitude=True)

        if metadata.get('Latitude (Decimal)') and metadata.get('Longitude (Decimal)'):
            metadata['Google Maps Link'] = f"https://www.google.com/maps?q={metadata['Latitude (Decimal)']},{metadata['Longitude (Decimal)']}"

        metadata['File Size (Bytes)'] = os.path.getsize(file_path)

        with Image.open(file_path) as img:
            metadata['Width (px)'], metadata['Height (px)'] = img.size

        metadata['MD5 Hash'] = calculate_image_hash(file_path)

    except Exception as e:
        print(f"Error processing HEIC file with ExifTool: {file_path}: {e}")
        metadata = None
    return metadata

# Save extracted metadata to an Excel file
def save_metadata_to_excel(metadata_list, output_path):
    df = pd.DataFrame(metadata_list)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Metadata')

    workbook = writer.book
    worksheet = writer.sheets['Metadata']

    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'font_color': '#000000', 'border': 1})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    for i, col in enumerate(df.columns):
        max_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_width)

    writer.close()

# Main function to process images in a directory and save metadata to Excel
def main(directory_path, output_excel):
    metadata_list = []
    
    image_formats = ('.jpg', '.jpeg', '.png')
    heic_formats = ('.heic',)
    
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith(image_formats):
                metadata = extract_metadata_image(file_path)
            elif file.lower().endswith(heic_formats):
                metadata = extract_metadata_heic(file_path)
            else:
                continue
            if metadata:
                metadata_list.append(metadata)
                
    save_metadata_to_excel(metadata_list, output_excel)
    print(f"Metadata saved to {output_excel}")

# Define paths for input directory and output Excel file
directory_path = "images"
output_excel = "output/extracted_metadata.xlsx"

# Run main function
main(directory_path, output_excel)
