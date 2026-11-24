import json
import re
import pandas as pd
import pdfplumber


def parse_belmarket_receipt(pdf_path: str) -> dict:
    full_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text.append(text)

    raw_text = "\n".join(full_text)
    lines = [line.strip() for line in raw_text.split("\n") if line.strip()]

    receipt_data = {
        "store": "Белмаркет",
        "date": None,
        "time": None,
        "total": None,
        "items": [],
    }

    datetime_pattern = re.compile(
        r"(\d{2}\.\d{2}\.\d{2,4})\s+(\d{2}:\d{2}(?::\d{2})?)"
    )
    total_pattern = re.compile(
        r"(?:ИТОГО|ВСЕГО|К ОПЛАТЕ)\s*:?\s*(\d+[.,]\d{2})", re.IGNORECASE
    )
    qty_price_pattern = re.compile(
        r"(\d+[.,]\d{1,3})\s*[xXхХ*]\s*(\d+[.,]\d{2})\s*=?\s*(\d+[.,]\d{2})"
    )
    single_line_item_pattern = re.compile(
        r"^(.+?)\s+(\d+[.,]\d{1,3})\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})$"
    )

    for i, line in enumerate(lines):
        if not receipt_data["date"]:
            dt_match = datetime_pattern.search(line)
            if dt_match:
                receipt_data["date"] = dt_match.group(1)
                receipt_data["time"] = dt_match.group(2)

        total_match = total_pattern.search(line)
        if total_match:
            receipt_data["total"] = float(
                total_match.group(1).replace(",", ".")
            )

        qp_match = qty_price_pattern.search(line)
        if qp_match and i > 0:
            item_name = lines[i - 1]
            if not any(
                k in item_name.upper() for k in ["КАССА", "ЧЕК", "УНП", "---"]
            ):
                receipt_data["items"].append({
                    "Дата": receipt_data["date"],
                    "Товар": item_name,
                    "Количество": float(qp_match.group(1).replace(",", ".")),
                    "Цена": float(qp_match.group(2).replace(",", ".")),
                    "Сумма": float(qp_match.group(3).replace(",", ".")),
                })
                continue

        sl_match = single_line_item_pattern.match(line)
        if sl_match:
            name_candidate = sl_match.group(1).strip()
            if not any(
                k in name_candidate.upper()
                for k in ["ИТОГО", "НДС", "СУММА", "НАЛИЧНЫМИ", "КАРТОЙ"]
            ):
                receipt_data["items"].append({
                    "Дата": receipt_data["date"],
                    "Товар": name_candidate,
                    "Количество": float(sl_match.group(2).replace(",", ".")),
                    "Цена": float(sl_match.group(3).replace(",", ".")),
                    "Сумма": float(sl_match.group(4).replace(",", ".")),
                })

    return receipt_data


def save_results(data: dict):
    # 1. Запись в JSON-файл
    with open("result.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    # 2. Запись списка товаров в Excel-таблицу
    if data["items"]:
        df = pd.DataFrame(data["items"])
        df.to_excel("result.xlsx", index=False)
        print("Готово! Данные записаны в файлы result.json и result.xlsx")
    else:
        print("Позиции товаров не найдены. Проверьте структуру PDF-файла.")


if __name__ == "__main__":
    parsed_data = parse_belmarket_receipt("Check/check.pdf")
    save_results(parsed_data)
