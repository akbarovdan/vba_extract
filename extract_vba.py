import os
import sys
import tempfile
import urllib.request
from typing import Tuple
from oletools.olevba import VBA_Parser

OUTPUT_BASE_DIR: str = "output"


def detect_module_type(filename: str, code: str) -> str:
    """Определяет тип модуля для раскладки по папкам"""
    code_lower = code.lower()
    fn_lower = filename.lower()

    if fn_lower.endswith(".frm") or "begin {c62a69f0-" in code_lower:
        return "Forms"
    elif "attribute vb_customizable = true" in code_lower or any(
        fn_lower.startswith(prefix)
        for prefix in ["лист", "sheet", "thisworkbook", "этакнига"]
    ):
        return "Sheets"
    elif (
        fn_lower.endswith(".cls")
        or "attribute vb_creatable = false" in code_lower
    ):
        return "Classes"
    else:
        return "Modules"


def download_file_if_url(source: str) -> Tuple[str, bool]:
    """Скачивает файл во временную директорию, если передана URL-ссылка"""
    if source.startswith(("http://", "https://")):
        print(f"[↓] Скачиваем файл по ссылке: {source} ...")
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsm")
        os.close(tmp_fd)

        req = urllib.request.Request(
            source, headers={"User-Agent": "Mozilla/5.0"}
        )
        with (
            urllib.request.urlopen(req) as response,
            open(tmp_path, "wb") as out_file,
        ):
            out_file.write(response.read())

        print(f"[✓] Временно сохранен в: {tmp_path}")
        return tmp_path, True
    return source, False


def main() -> None:
    # 1. Получаем путь или ссылку
    if len(sys.argv) > 1:
        source_input = sys.argv[1].strip()
    else:
        source_input = input(
            "Введите путь к файлу или URL-ссылку: "
        ).strip('"\' ')

    if not source_input:
        print("[-] Ошибка: путь или ссылка не указаны.")
        return

    # Инициализируем переменные заранее для чистоты анализатора
    file_path: str = ""
    is_temp: bool = False

    try:
        file_path, is_temp = download_file_if_url(source_input)

        if not os.path.exists(file_path):
            print(f"[-] Ошибка: файл '{file_path}' не найден!")
            return

        print(f"[+] Анализ файла: {source_input}")

        # 2. Создаем структуру папок
        categories = ["Modules", "Sheets", "Forms", "Classes", "Other"]
        for cat in categories:
            os.makedirs(os.path.join(OUTPUT_BASE_DIR, cat), exist_ok=True)

        vba_parser = VBA_Parser(file_path)
        if not vba_parser.detect_vba_macros():
            print("[-] В файле не найдено макросов VBA.")
            vba_parser.close()
            return

        summary = []
        count = 0

        # 3. Выгружаем каждый модуль
        for (
            filename,
            stream_path,
            vba_filename,
            vba_code,
        ) in vba_parser.extract_macros():
            # Принудительно гарантируем строковые типы для линтера
            clean_name: str = str(vba_filename or "module").replace("/", "_").replace("\\", "_")
            code_text: str = str(vba_code or "")
            
            mod_type: str = detect_module_type(clean_name, code_text)

            if not clean_name.endswith((".bas", ".cls", ".frm", ".txt")):
                ext = (
                    ".frm"
                    if mod_type == "Forms"
                    else (
                        ".cls"
                        if mod_type in ["Sheets", "Classes"]
                        else ".bas"
                    )
                )
                clean_name += ext

            target_path: str = os.path.join(OUTPUT_BASE_DIR, mod_type, clean_name)

            with open(target_path, "w", encoding="utf-8") as f:
                f.write(code_text)

            lines_count = len(code_text.splitlines())
            summary.append(
                f"[{mod_type:<7}] {clean_name:<35} | {lines_count:>4} строк | {target_path}"
            )
            print(f"  ✓ [{mod_type}] {clean_name} ({lines_count} строк)")
            count += 1

        vba_parser.close()

        # 4. Сохраняем сводку summary.txt
        summary_file = os.path.join(OUTPUT_BASE_DIR, "summary.txt")
        with open(summary_file, "w", encoding="utf-8") as f:
            f.write(f"Источник: {source_input}\n")
            f.write(f"Всего выгружено модулей: {count}\n\n")
            f.write("\n".join(summary))

        print(f"\n[✓] Успешно выгружено модулей: {count}")
        print(f"[✓] Файлы сохранены в: {os.path.abspath(OUTPUT_BASE_DIR)}/")
        print(f"[✓] Сводка: {os.path.abspath(summary_file)}")

    finally:
        if is_temp and file_path and os.path.exists(file_path):
            os.remove(file_path)


if __name__ == "__main__":
    main()