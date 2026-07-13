"""
Разовый бэкфилл истории по каналам — забирает последние N дней из Roistat и
наполняет data/channel_history.json, чтобы агент не начинал "обучение" с нуля,
а сразу использовал данные, которые уже накоплены в самом Roistat.

Запускается ВРУЧНУЮ ОДИН РАЗ (не часть ежедневного GitHub Actions workflow):
  python3 backfill_history.py [дней]

По умолчанию — 90 дней (совпадает с окном хранения в ksm_daily_agent.py).
Безопасно запускать повторно — append_daily_record идемпотентен по дате.
"""

import sys
import time
import datetime

import ksm_daily_agent as agent


def main():
    days = int(sys.argv[1]) if len(sys.argv) > 1 else agent.HISTORY_WINDOW_DAYS
    today = agent.almaty_today()
    history = agent.load_history()

    print(f"Бэкфилл за последние {days} дней (сегодня по Алматы: {today})")
    for i in range(days, 0, -1):
        d = today - datetime.timedelta(days=i)
        try:
            raw = agent.fetch_roistat_data(str(d), str(d))
            channels = agent.extract_channels(raw)
            agent.append_daily_record(history, str(d), channels)
            print(f"  {d}: {len(channels)} каналов")
        except Exception as e:
            print(f"  {d}: ОШИБКА, пропускаю — {e}")
        time.sleep(0.3)  # не бомбить API большим количеством запросов подряд

    agent.save_history(history)
    print("Готово. Файл:", agent.HISTORY_PATH)


if __name__ == "__main__":
    main()
