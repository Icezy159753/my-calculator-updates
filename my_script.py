from PIL import Image
import os

ICON_FOLDER = "Icon"
ICON_NAME = "Iconcat.ico" # หรือ Iconcat.png ถ้าลองวิธีที่ 1 แล้ว

# หากรันสคริปต์ทดสอบนี้จากโฟลเดอร์เดียวกับ Launcher หลัก
icon_path = os.path.join(ICON_FOLDER, ICON_NAME)

# หรือระบุ full path โดยตรงเพื่อความแน่นอนในการทดสอบ
# icon_path = r"C:\path\to\your\project\Icon\Iconcat.ico"

print(f"Attempting to load: {icon_path}")

if not os.path.exists(icon_path):
    print(f"ERROR: File not found at {icon_path}")
else:
    try:
        img = Image.open(icon_path)
        print(f"Successfully loaded: {ICON_NAME}")
        print(f"Image mode: {img.mode}") # ดูว่า image mode เป็นอะไร (เช่น RGBA, P, L)
        print(f"Image size: {img.size}")
        img.show() # ลองให้ PIL แสดงภาพโดยตรง
    except Exception as e:
        print(f"ERROR loading {ICON_NAME} with PIL: {e}")