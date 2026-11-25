import json
import pandas as pd

EXCEL_FILE = "calendar.xlsx"
TZ_OFFSET = "+03:00"  # Московское время


def is_empty(val):
    return pd.isna(val) or str(val).strip() == ""


def clean_str(val):
    return "" if is_empty(val) else str(val).strip()


def clean_num(val):
    if is_empty(val):
        return None
    try:
        f = float(val)
        return int(f) if f.is_integer() else f
    except (ValueError, TypeError):
        return str(val).strip()


def is_truthy(val):
    return clean_str(val).lower() in ("1", "1.0", "true", "да", "yes", "+")


def format_dt(val):
    if is_empty(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime(f"%Y-%m-%dT%H:%M:%S{TZ_OFFSET}")
    s = str(val).strip()
    # Если уже в ISO с часовым поясом — оставляем как есть
    if "T" in s and (s.endswith("Z") or "+" in s[10:] or "-" in s[10:]):
        return s
    dt = pd.to_datetime(s, dayfirst="." in s)
    return dt.strftime(f"%Y-%m-%dT%H:%M:%S{TZ_OFFSET}")


def main():
    xl = pd.ExcelFile(EXCEL_FILE)
    sheets = xl.sheet_names

    # 1. Сессии (группируем по race_id)
    sessions_map = {}
    if "sessions" in sheets:
        for _, row in xl.parse("sessions").iterrows():
            rid = clean_str(row.get("race_id"))
            if not rid or is_empty(row.get("name")) or is_empty(row.get("start")):
                continue
            item = {
                "name": clean_str(row["name"]),
                "start": format_dt(row["start"]),
            }
            if is_truthy(row.get("main")):
                item["main"] = True
            sessions_map.setdefault(rid, []).append(item)

    # 2. Результаты (группируем по race_id)
    results_map = {}
    if "results" in sheets:
        for _, row in xl.parse("results").iterrows():
            rid = clean_str(row.get("race_id"))
            if not rid or is_empty(row.get("driver")):
                continue
            item = {"driver": clean_str(row["driver"])}
            if not is_empty(row.get("pos")):
                item["pos"] = clean_num(row["pos"])
            if not is_empty(row.get("team")):
                item["team"] = clean_str(row["team"])
            if not is_empty(row.get("time")):
                item["time"] = clean_str(row["time"])
            if not is_empty(row.get("points")):
                item["points"] = clean_num(row["points"])
            results_map.setdefault(rid, []).append(item)

    # 3. Основной список гонок -> races.json
    races = []
    for _, row in xl.parse("races").iterrows():
        if is_empty(row.get("series")) or is_empty(row.get("name")):
            continue
        rid = clean_str(row.get("id"))
        race = {
            "series": clean_str(row["series"]),
            "name": clean_str(row["name"]),
            "circuit": clean_str(row.get("circuit")),
            "country": clean_str(row.get("country")),
            "wheels": clean_str(row.get("wheels")) or "open",
            "start": format_dt(row.get("start")),
            "color": clean_str(row.get("color")) or "#e10600",
        }
        if not is_empty(row.get("watch")):
            race["watch"] = clean_str(row["watch"])
        if not is_empty(row.get("watchLabel")):
            race["watchLabel"] = clean_str(row["watchLabel"])
        if rid in sessions_map:
            race["sessions"] = sessions_map[rid]
        if rid in results_map:
            race["results"] = results_map[rid]
        races.append(race)

    with open("races.json", "w", encoding="utf-8") as f:
        json.dump(races, f, ensure_ascii=False, indent=2)

    # 4. Таблицы чемпионатов -> standings.json
    standings = {}
    if "standings" in sheets:
        type_aliases = {
            "drivers": "drivers", "driver": "drivers", "пилоты": "drivers", "личный": "drivers",
            "teams": "teams", "team": "teams", "команды": "teams", "конструкторы": "teams",
        }
        for _, row in xl.parse("standings").iterrows():
            series = clean_str(row.get("series"))
            st_type = type_aliases.get(clean_str(row.get("type")).lower())
            if not series or not st_type or is_empty(row.get("name")):
                continue

            entry = standings.setdefault(series, {})
            if not is_empty(row.get("updated")) and "updated" not in entry:
                entry["updated"] = clean_str(row["updated"])

            rec = {"name": clean_str(row["name"])}
            if not is_empty(row.get("pos")):
                rec["pos"] = clean_num(row["pos"])
            if st_type == "drivers" and not is_empty(row.get("team")):
                rec["team"] = clean_str(row["team"])
            if not is_empty(row.get("points")):
                rec["points"] = clean_num(row["points"])

            entry.setdefault(st_type, []).append(rec)

    with open("standings.json", "w", encoding="utf-8") as f:
        json.dump(standings, f, ensure_ascii=False, indent=2)

    print(f"Готово: {len(races)} гонок (races.json), {len(standings)} серий в таблицах (standings.json).")


if __name__ == "__main__":
    main()
