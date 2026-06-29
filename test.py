import os

print("📁 Проверка файлов в текущей папке:")
print("=" * 40)

files_needed = ["Доверенность_шаблон.docx", "pechat_shablon.png", "bot.py"]

for f in files_needed:
    if os.path.exists(f):
        print(f"✅ {f} — найден")
    else:
        print(f"❌ {f} — ОТСУТСТВУЕТ!")

print("=" * 40)
print(f"📂 Текущая папка: {os.getcwd()}")