"""
KSM Marketing Agent — ежедневная автоматизация
Пайплайн: Roistat -> Claude (Anthropic API) -> хранилище -> Telegram

Это НЕ веб-сайт и не то, что открывается в браузере — это скрипт,
который запускается ОДИН раз за вызов (например, каждое утро по
расписанию: cron, GitHub Actions, облачный планировщик — см. пояснение
в чате). Мини-сайт из демо должен читать результат его работы
(latest_snapshot.json), а не дёргать Roistat напрямую из браузера.

Секреты нигде не хардкожены — только через переменные окружения
(см. .env.example рядом). Перед запуском:
  pip install requests python-dotenv
"""

import os
import json
import datetime
import requests
from dotenv import load_dotenv

load_dotenv()  # подхватывает значения из файла .env, лежащего рядом со скриптом

# ---------- Конфигурация (из переменных окружения) ----------
ROISTAT_API_KEY = os.environ["ROISTAT_API_KEY"]
ROISTAT_PROJECT_ID = os.environ["ROISTAT_PROJECT_ID"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

# Для ежедневных отчётов достаточно Sonnet; Haiku дешевле и быстрее,
# если хочется сэкономить на частых запусках.
CLAUDE_MODEL = "claude-sonnet-5"

SNAPSHOT_PATH = "latest_snapshot.json"
CPL_ALERT_PCT = float(os.environ.get("CPL_ALERT_PCT", 15))
LEADS_DROP_ALERT_PCT = float(os.environ.get("LEADS_DROP_ALERT_PCT", 15))


# ---------- 1. Roistat: получение данных ----------
# Подтверждено по официальной документации (help-ru.roistat.com/API):
#   - метод: POST https://cloud.roistat.com/api/v1/project/analytics/data
#   - ключ передаётся заголовком Api-key, номер проекта — параметром в URL
#   - метрики: visits, leads, sales, profit, marketing_cost, roi (расход — marketing_cost, не cost)
#   - измерение marker_level_1 — верхний уровень источника (канал)
#   - period.from / period.to — с временем и часовым поясом, не просто дата
ROISTAT_TZ_OFFSET = "+06:00"  # часовой пояс проекта в Roistat; поменяйте, если у вас другой

def fetch_roistat_data(date_from, date_to):
    url = "https://cloud.roistat.com/api/v1/project/analytics/data"
    headers = {"Api-key": ROISTAT_API_KEY}
    params = {"project": ROISTAT_PROJECT_ID}
    body = {
        "dimensions": ["marker_level_1"],
        "metrics": ["visits", "leads", "marketing_cost"],
        "period": {
            "from": f"{date_from}T00:00:00{ROISTAT_TZ_OFFSET}",
            "to": f"{date_to}T23:59:59{ROISTAT_TZ_OFFSET}",
        },
    }
    resp = requests.post(url, headers=headers, params=params, json=body, timeout=30)
    if resp.status_code >= 400:
        print(f"Roistat вернул ошибку, статус {resp.status_code}.")
        print("Тело ответа:", resp.text[:1000])
        resp.raise_for_status()
    try:
        return resp.json()
    except requests.exceptions.JSONDecodeError:
        print("=" * 60)
        print(f"Roistat ответил статусом {resp.status_code}, но это не JSON.")
        print("Заголовок ответа (Content-Type):", resp.headers.get("Content-Type"))
        print("Первые 500 символов ответа:")
        print(repr(resp.text[:500]))
        print("=" * 60)
        raise


def extract_totals(roistat_response):
    """
    Реальная форма ответа Roistat (подтверждено на практике):
    data -> список из одного объекта -> в нём items (строки отчёта).
    Внутри каждой строки metrics — список объектов {metric_name, value},
    а не отдельные поля leads/marketing_cost напрямую.
    """
    data = roistat_response.get("data")
    if isinstance(data, list) and data:
        block = data[0]
    elif isinstance(data, dict):
        block = data
    else:
        block = {}
    items = block.get("items") or []

    leads = 0
    cost = 0
    for item in items:
        for metric in item.get("metrics", []):
            name = metric.get("metric_name")
            value = metric.get("value") or 0
            if name == "leads":
                leads += value
            elif name == "marketing_cost":
                cost += value

    print(f"[Roistat] строк обработано: {len(items)}, сумма leads={leads}, сумма marketing_cost={cost}")
    if items and leads == 0 and cost == 0:
        print("Суммы всё равно нулевые — вот пример одной строки целиком, чтобы найти точную причину:")
        print(json.dumps(items[0], ensure_ascii=False))

    return {"leads": leads, "cost": cost}


# ---------- 2. Claude: анализ данных (через Anthropic API) ----------
SYSTEM_PROMPT = """Ты — ИИ-агент «Маркетинг-аналитик KSM» для отдела маркетинга
компании, производящей строительные материалы. Тебе присылают сырые данные
по рекламе из Roistat: расход, лиды, CPL по каналам и кампаниям за вчера
и позавчера.

Твои задачи (из технического задания отдела):
1. Сравнить период с предыдущим: что выросло, что упало, на сколько процентов.
2. Объяснить вероятные причины роста или падения CPL.
3. Найти узкие места — кампании, которые тратят бюджет без заявок, аномальный CPC.
4. Дать конкретные рекомендации: какие кампании отключить или усилить, какие
   ключевые слова добавить или исключить, что проверить на посадочной странице.
5. Сформировать 2-4 вопроса рекламному агентству по сути отклонений.

Отвечай на русском, кратко и по делу, строго такой структурой:
Что произошло / Почему / Что делать / Вопросы агентству.
Не придумывай цифры, которых нет во входных данных."""


def analyze_with_claude(roistat_payload):
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": CLAUDE_MODEL,
            "max_tokens": 1500,
            "system": SYSTEM_PROMPT,
            "messages": [{
                "role": "user",
                "content": "Вот данные:\n" + json.dumps(roistat_payload, ensure_ascii=False)
                            + "\n\nПроанализируй их по инструкции.",
            }],
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["content"]
    return "".join(block["text"] for block in content if block["type"] == "text")


# ---------- 3. Хранилище: сохраняем снимок в формате, который читает мини-сайт ----------
def save_snapshot(totals_yesterday, totals_day_before, analysis_text, alerts):
    """
    Схема специально простая и совпадает с тем, что ждёт ksm-marketing-agent-demo.html
    (функция tryLoadLive). Если структуру поменяете — поправьте её и там тоже.
    """
    snapshot = {
        "updated_at": datetime.datetime.utcnow().isoformat() + "Z",
        "totals": {
            "yesterday": totals_yesterday,
            "day_before": totals_day_before,
        },
        "analysis": analysis_text,
        "alerts": alerts,
    }
    with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
    return snapshot


# ---------- 4. Telegram: уведомление при отклонениях ----------
def send_telegram_alert(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram не настроен (нет токена/chat_id) — пропускаю отправку.")
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=15)


# ---------- 5. Быстрая проверка порогов (без ИИ, для мгновенного сигнала) ----------
def check_thresholds(today_totals, yesterday_totals):
    alerts = []
    if yesterday_totals.get("leads", 0) > 0:
        change = (today_totals["leads"] - yesterday_totals["leads"]) / yesterday_totals["leads"] * 100
        if change <= -LEADS_DROP_ALERT_PCT:
            alerts.append(f"Лиды упали на {abs(change):.1f}% за сутки ({today_totals['leads']} против {yesterday_totals['leads']}).")

    cpl_today = today_totals["cost"] / today_totals["leads"] if today_totals.get("leads") else None
    cpl_yesterday = yesterday_totals["cost"] / yesterday_totals["leads"] if yesterday_totals.get("leads") else None
    if cpl_today and cpl_yesterday:
        change = (cpl_today - cpl_yesterday) / cpl_yesterday * 100
        if change >= CPL_ALERT_PCT:
            alerts.append(f"CPL вырос на {change:.1f}% (с {cpl_yesterday:.0f} до {cpl_today:.0f} ₸).")
    return alerts


# ---------- main: одна ежедневная итерация ----------
def main():
    today = datetime.date.today()
    yesterday_date = today - datetime.timedelta(days=1)
    day_before_date = today - datetime.timedelta(days=2)

    raw_yesterday = fetch_roistat_data(str(yesterday_date), str(yesterday_date))
    raw_day_before = fetch_roistat_data(str(day_before_date), str(day_before_date))

    totals_yesterday = extract_totals(raw_yesterday)
    totals_day_before = extract_totals(raw_day_before)

    analysis_text = analyze_with_claude({"вчера": raw_yesterday, "позавчера": raw_day_before})
    alerts = check_thresholds(totals_yesterday, totals_day_before)
    save_snapshot(totals_yesterday, totals_day_before, analysis_text, alerts)

    if alerts:
        send_telegram_alert(
            "Внимание, KSM-маркетинг:\n" + "\n".join(alerts)
            + f"\n\nПодробный разбор:\n{analysis_text[:500]}"
        )

    print("Готово:", datetime.datetime.now().isoformat())


if __name__ == "__main__":
    main()
