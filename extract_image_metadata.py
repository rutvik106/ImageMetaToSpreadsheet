import os
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
from datetime import datetime
import pytz


def get_exif_data(image):
    exif_data = {}
    gps_data = {}
    try:
        info = image._getexif()
        if info:
            for tag, value in info.items():
                tag_name = TAGS.get(tag, tag)
                if tag_name == 'GPSInfo':
                    for t in value:
                        sub_tag = GPSTAGS.get(t, t)
                        gps_data[sub_tag] = value[t]
                else:
                    exif_data[tag_name] = value
    except Exception as e:
        print(f"Error getting EXIF data: {e}")

    return exif_data, gps_data


def convert_to_degrees(value):
    d, m, s = value
    return d + (m / 60.0) + (s / 3600.0)


def get_lat_lon(gps_data):
    lat = lon = None
    try:
        if 'GPSLatitude' in gps_data and 'GPSLongitude' in gps_data:
            lat = convert_to_degrees(gps_data['GPSLatitude'])
            if gps_data['GPSLatitudeRef'] != 'N':
                lat = -lat
            lon = convert_to_degrees(gps_data['GPSLongitude'])
            if gps_data['GPSLongitudeRef'] != 'E':
                lon = -lon
    except KeyError:
        pass

    return lat, lon


def convert_to_iso8601(date_str):
    try:
        local_tz = pytz.timezone("Asia/Kolkata")  # Replace with the appropriate timezone if needed
        dt = datetime.strptime(date_str, "%Y:%m:%d %H:%M:%S")
        local_dt = local_tz.localize(dt, is_dst=None)
        return local_dt.isoformat()
    except Exception as e:
        print(f"Error converting date: {e}")
        return None


def process_images(directory):
    image_data = []

    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.jpg'):
                filepath = os.path.join(root, file)
                try:
                    image = Image.open(filepath)
                    exif_data, gps_data = get_exif_data(image)
                    lat, lon = get_lat_lon(gps_data)

                    date = exif_data.get("DateTime", None)
                    if date:
                        date = convert_to_iso8601(date)

                    image_details = {
                        "filename": file,
                        "path": filepath,
                        "latitude": lat,
                        "longitude": lon,
                        "date": date,
                        "device": exif_data.get("Model", None)
                    }
                    image_data.append(image_details)
                except Exception as e:
                    print(f"Error processing image {filepath}: {e}")

    return image_data


def create_spreadsheet(image_data, output_file):
    df = pd.DataFrame(image_data)
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    directory = input("Enter the directory containing images: ")
    output_file = input("Enter the output Excel file name (with .xlsx extension): ")
    image_data = process_images(directory)
    create_spreadsheet(image_data, output_file)
    print(f"Spreadsheet {output_file} created successfully.")
