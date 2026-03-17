import sys
import os
import argparse
import chardet


# Определяет кодировку по первым sample_size байтам.
def detect_encoding(file_path, sample_size=10000):
    with open(file_path, 'rb') as f:
        raw = f.read(sample_size)
    result = chardet.detect(raw)
    return result['encoding']

# Конвертирует файл из from_enc в to_enc. При ошибке завершает программу.
def convert_text_file(input_file, output_file, from_enc, to_enc, overwrite=False):
    if input_file == output_file and not overwrite:
        sys.stderr.write("Ошибка: входной и выходной файлы совпадают. Используйте --overwrite.\n")
        sys.exit(1)

    if from_enc is None:
        from_enc = detect_encoding(input_file)
        if from_enc is None:
            sys.stderr.write("Не удалось определить кодировку. Укажите её явно через --from.\n")
            sys.exit(1)

    try:
        with open(input_file, 'r', encoding=from_enc) as f_in:
            with open(output_file, 'w', encoding=to_enc) as f_out:
                for line in f_in:
                    f_out.write(line)
    except UnicodeDecodeError:
        sys.stderr.write(f"Ошибка декодирования: файл не может быть прочитан в кодировке {from_enc}.\n")
        sys.exit(1)
    except Exception as e:
        sys.stderr.write(f"Ошибка при обработке файла: {e}\n")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="Конвертация кодировки текстового файла (без лишнего вывода).")
    parser.add_argument("input", help="Путь к входному файлу (или '-' для stdin)")
    parser.add_argument("output", nargs='?', help="Путь к выходному файлу (по умолчанию stdout или перезапись входа)")
    parser.add_argument("--from", dest="from_enc", help="Исходная кодировка (если не указана, определяется автоматически)")
    parser.add_argument("--to", dest="to_enc", default="utf-8", help="Целевая кодировка (по умолчанию utf-8)")
    parser.add_argument("--overwrite", action="store_true", help="Разрешить перезапись входного файла")

    args = parser.parse_args()

    # Обработка stdin/stdout
    if args.input == '-':
        if args.from_enc is None:
            sys.stderr.write("При чтении из stdin необходимо указать исходную кодировку через --from.\n")
            sys.exit(1)
        try:
            data = sys.stdin.buffer.read().decode(args.from_enc)
        except UnicodeDecodeError as e:
            sys.stderr.write(f"Ошибка декодирования stdin: {e}\n")
            sys.exit(1)

        if args.output:
            with open(args.output, 'w', encoding=args.to_enc) as f:
                f.write(data)
        else:
            sys.stdout.buffer.write(data.encode(args.to_enc))
        return

    # Проверка существования входного файла
    if not os.path.isfile(args.input):
        sys.stderr.write(f"Файл не найден: {args.input}\n")
        sys.exit(1)

    # Определение выходного файла
    out_file = args.output if args.output else args.input

    # Для Excel-файлов предупреждение (можно удалить, если не нужно)
    ext = os.path.splitext(args.input)[1].lower()
    if ext in ('.xlsx', '.xls'):
        sys.stderr.write("Предупреждение: Excel-файлы обычно содержат текст в Unicode и не требуют перекодировки. "
                         "Продолжение может привести к потере данных.\n")
        # При желании можно реализовать обработку Excel через openpyxl

    convert_text_file(args.input, out_file, args.from_enc, args.to_enc, args.overwrite)

if __name__ == "__main__":
    main()