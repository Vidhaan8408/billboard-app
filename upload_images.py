
import os
import re
import cloudinary
import cloudinary.uploader

# CONFIG — replace with your credentials
cloudinary.config(
    cloud_name="dub8ndson",
    api_key="532366637732586",
    api_secret="8ZV9QHoRaR_xYj6gqXU5MSTHe5E"
)

FOLDER = "images"  # your local folder

def clean_name(filename):
    name, ext = os.path.splitext(filename)

    # remove random suffix like _abc123
    cleaned = re.sub(r'_[a-zA-Z0-9]{5,}$', '', name)

    return cleaned, ext

for file in os.listdir(FOLDER):
    if not file.lower().endswith((".jpg", ".jpeg", ".png")):
        continue

    path = os.path.join(FOLDER, file)

    public_id, ext = clean_name(file)

    print(f"Uploading: {file} → {public_id}")

    cloudinary.uploader.upload(
        path,
        public_id=public_id,
        overwrite=True,
        resource_type="image"
    )
