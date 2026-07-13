"""
Отладочный инструмент для сверки сырых данных Roistat API с тем, что видно
в дашборде Roistat вручную. НЕ отправляет ничего в Anthropic/Claude — только
Roistat API. Это не веб-сервер и не эндпоинт — обычный скрипт, запускается
из терминала и печатает JSON в stdout (можно перенаправить в файл для diff).

Использует ТОЛЬКО переменные окружения ROISTAT_API_KEY / ROISTAT_PROJECT_ID
(через .env, как и ksm_daily_agent.py) — ключ нигде не хардкодится и никогда
не печатается в вывод.

Запуск:
  python3 roistat_debug.py                          # за вчера (по Алматы)
  python3 roistat_debug.py --date 2026-07-12         # за конкретный день
  python3 roistat_debug.py --from 2026-07-06 --to 2026-07-12   # за диапазон

Дальше сравниваете date_from/date_to и totals из вывода с тем же отчётом,
открытым вручную в дашборде Roistat за ТЕ ЖЕ даты — числа должны совпасть
день в день (уже проверено на практике в этом проекте).
"""

import sys
import json
import argparse

import ksm_daily_agent as agent


def build_debug_report(date_from, date_to):
    raw = agent.fetch_roistat_data(date_from, date_to)

    totals = agent.extract_totals(raw)
    channels = agent.extract_channels(raw)

    # Что именно ушло в запрос — те же значения, что в fetch_roistat_data(),
    # продублированы явно здесь для читаемости отчёта (не пересобираются
    # заново из функции, чтобы не разойтись случайно, если она поменяется —
    # см. TODO ниже).
    request_body = {
        "dimensions": ["marker_level_1", "marker_level_2"],
        "metrics": ["visits", "leads", "marketing_cost", "impressions", "clicks"],
        "period": {
            "from": f"{date_from}T00:00:00{agent.ROISTAT_TZ_OFFSET}",
            "to": f"{date_to}T23:59:59{agent.ROISTAT_TZ_OFFSET}",
        },
    }

    # Roistat не даёт выбрать модель атрибуции в этом запросе — она приходит
    # в самом ответе (attribution_model_id на каждой метрике). Вытаскиваем
    # то, что реально использовал Roistat, а не то, что мы якобы попросили.
    attribution_models_seen = set()
    items = agent._iter_items(raw)
    for item in items:
        for metric in item.get("metrics", []):
            amid = metric.get("attribution_model_id")
            if amid:
                attribution_models_seen.add(amid)

    return {
        "projectId": agent.ROISTAT_PROJECT_ID,
        "dateFrom": date_from,
        "dateTo": date_to,
        "timezone_used": agent.ROISTAT_TZ_OFFSET + " (Алматы, фиксированный, без перехода на летнее время)",
        "metrics_requested": request_body["metrics"],
        "dimensions_requested": request_body["dimensions"],
        "filters_used": "нет — запрос идёт без дополнительных фильтров (весь трафик по проекту)",
        "attribution_model": {
            "requested": "не указывается явно в этом запросе (Roistat применяет свою модель по умолчанию)",
            "seen_in_response": sorted(attribution_models_seen) or None,
        },
        "request_body_sent": request_body,
        "raw_roistat_response": raw,
        "calculated_totals": {
            "leads": totals["leads"],
            "spend": totals["cost"],
            "cpl": totals["cpl"],
            "visits": totals["visits"],
            "impressions": totals["impressions"],
            "clicks": totals["clicks"],
            "ctr_pct": totals["ctr"],
            "cpc": totals["cpc"],
            "conversion_rate_pct": totals["cr"],
            "conversions": "не запрошено у Roistat в этом вызове (метрика 'leads' используется как конверсия в заявку)",
            "revenue": "не запрошено в этом вызове; ранее проверено отдельно — sales/revenue/profit/roi/romi всегда возвращают 0 для этого проекта (в Roistat не настроена передача выручки из CRM)",
        },
        "channels_breakdown": channels,
    }


def main():
    parser = argparse.ArgumentParser(description="Debug: сырые данные Roistat vs дашборд")
    parser.add_argument("--date", help="один день, например 2026-07-12")
    parser.add_argument("--from", dest="date_from", help="начало диапазона")
    parser.add_argument("--to", dest="date_to", help="конец диапазона")
    args = parser.parse_args()

    if args.date:
        date_from = date_to = args.date
    elif args.date_from and args.date_to:
        date_from, date_to = args.date_from, args.date_to
    elif args.date_from or args.date_to:
        parser.error("--from и --to нужно указывать вместе")
    else:
        yesterday = agent.almaty_today() - agent.timedelta(days=1)
        date_from = date_to = str(yesterday)

    report = build_debug_report(date_from, date_to)
    json.dump(report, sys.stdout, ensure_ascii=False, indent=2, default=str)
    print()


if __name__ == "__main__":
    main()
