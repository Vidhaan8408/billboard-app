import cloudinary
import cloudinary.uploader
import os

cloudinary.config(
    cloud_name="dub8ndson",
    api_key="532366637732586",
    api_secret="8ZV9QHoRaR_xYj6gqXU5MSTHe5E"
)

folder = "images"

for file in os.listdir(folder):
    if file.endswith(".jpg"):
        path = os.path.join(folder, file)
        print(f"Uploading {file}...")
        res = cloudinary.uploader.upload(path, public_id=file.split(".")[0])
        print(res["secure_url"])
