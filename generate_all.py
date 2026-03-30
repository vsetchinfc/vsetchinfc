"""
generate_all.py
Master script that runs all 4 document generators sequentially.
Usage: python generate_all.py
"""

import subprocess
import sys
import os


SCRIPTS = [
    "doc1_lesson_plan.py",
    "doc2_visual_content.py",
    "doc3_check_tasks.py",
    "doc4_homework.py",
]


def main():
    print("=" * 60)
    print("  Генерация документов — Части речи")
    print("  Воскресная Русская Школа, Мельбурн")
    print("=" * 60)
    print()

    os.makedirs("output", exist_ok=True)

    all_ok = True
    for script in SCRIPTS:
        print(f"▶  Запускаю {script} ...")
        result = subprocess.run(
            [sys.executable, script],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            print(result.stdout.strip())
        else:
            print(f"❌  Ошибка в {script}:")
            print(result.stderr)
            all_ok = False

    print()
    if all_ok:
        print("=" * 60)
        print("✅  Все 4 документа успешно созданы в папке output/")
        print()
        for f in sorted(os.listdir("output")):
            if f.endswith(".docx"):
                path = os.path.join("output", f)
                size_kb = os.path.getsize(path) // 1024
                print(f"   📄  {f}  ({size_kb} KB)")
        print("=" * 60)
    else:
        print("❌  Некоторые файлы не были созданы. Проверьте ошибки выше.")
        sys.exit(1)


if __name__ == "__main__":
    main()
