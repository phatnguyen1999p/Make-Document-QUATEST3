"""
Calculation & Excel Automation Engine for QUATEST 3 (Room N2).
Handles:
1. PMS Lasair III 110 raw CSV (database_all.csv)
2. TSI AeroTrak 9350 raw .xlsx / .xls (6ChannelFixed)
3. Excel workbook inspection (formulas, cached values, #VALUE!/#N/A/#REF! error scan)
4. Safe Excel COM filling & recalculation (TinhToanKhiNen, TinhToanPhongSach, TinhToanVSKK, KQ bui)
"""
import os
import sys
import csv
import json
import argparse
from datetime import datetime
import openpyxl
import xlrd

sys.stdout.reconfigure(encoding="utf-8")

ISO_14644_LIMITS = {
    "ISO 5": {0.5: 3520, 1.0: 832, 5.0: 29},
    "ISO 6": {0.5: 35200, 1.0: 8320, 5.0: 293},
    "ISO 7": {0.5: 352000, 1.0: 83200, 5.0: 2930},
    "ISO 8": {0.5: 3520000, 1.0: 832000, 5.0: 29300},
    "ISO 9": {0.5: 35200000, 1.0: 8320000, 5.0: 293000},
    "Grade A": {0.5: 3520, 5.0: 20},
    "Grade B (at rest)": {0.5: 3520, 5.0: 29},
    "Grade B (in operation)": {0.5: 352000, 5.0: 2900},
    "Grade C (at rest)": {0.5: 352000, 5.0: 2900},
    "Grade C (in operation)": {0.5: 3520000, 5.0: 29000},
    "Grade D (at rest)": {0.5: 3520000, 5.0: 29000},
}


def parse_lasair_csv(csv_path: str, loc_filter: str = None):
    """
    Parse PMS Lasair III database_all.csv (; delimiter).
    Channels 1..8 = 0.1, 0.15, 0.2, 0.25, 0.3, 0.5, 1.0, 5.0 um (differential counts).
    """
    rows_out = []
    with open(csv_path, "r", encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f, delimiter=";")
        for r in reader:
            loc = (r.get("Location") or "").strip()
            if loc_filter and loc_filter.lower() not in loc.lower():
                continue
            data_id = (r.get("Data ID") or "").strip()
            date_s = (r.get("Sample Date") or "").strip()
            time_s = (r.get("Sample Time") or "").strip()
            vol_unit = (r.get("Volume Units") or "").strip()
            try:
                vol_raw = float(r.get("Sample Volume") or 0)
            except ValueError:
                continue
            if vol_raw <= 0:
                continue
            vol_m3 = vol_raw / 1000.0 if vol_unit.upper() == "L" else vol_raw

            counts = []
            for i in range(1, 9):
                try:
                    c_val = float(r.get(f"Counts {i}") or 0)
                except ValueError:
                    c_val = 0.0
                counts.append(c_val)

            # Cumulative concentrations (particles / m3)
            c_ge_01 = sum(counts[0:8]) / vol_m3
            c_ge_03 = sum(counts[4:8]) / vol_m3
            c_ge_05 = sum(counts[5:8]) / vol_m3
            c_ge_10 = sum(counts[6:8]) / vol_m3
            c_ge_50 = counts[7] / vol_m3

            # Compressed air ISO 8573-1 bands (raw counts and /m3 before N*)
            band_01_05 = sum(counts[0:5]) / vol_m3
            band_05_10 = counts[5] / vol_m3
            band_10_50 = counts[6] / vol_m3

            rows_out.append({
                "data_id": data_id,
                "date": date_s,
                "time": time_s,
                "location": loc,
                "loc_changed": (r.get("Location Changed") or "").strip(),
                "vol_m3": round(vol_m3, 6),
                "counts": [int(x) for x in counts],
                "ge_0.1_m3": round(c_ge_01),
                "ge_0.3_m3": round(c_ge_03),
                "ge_0.5_m3": round(c_ge_05),
                "ge_1.0_m3": round(c_ge_10),
                "ge_5.0_m3": round(c_ge_50),
                "band_0.1_0.5_m3": round(band_01_05),
                "band_0.5_1.0_m3": round(band_05_10),
                "band_1.0_5.0_m3": round(band_10_50),
            })

    print(f"=== Lasair III CSV: {csv_path} (filter={loc_filter!r}, matched={len(rows_out)} rows) ===")
    for item in rows_out:
        print(
            f"  ID={item['data_id']:>4s} | {item['date']} {item['time']} | Loc='{item['location']}' ({item['loc_changed']}) | "
            f"Vol={item['vol_m3']:.4f}m3 | >=0.5={item['ge_0.5_m3']:>8,d} | >=1.0={item['ge_1.0_m3']:>7,d} | >=5.0={item['ge_5.0_m3']:>6,d}"
        )
    return rows_out


def parse_aerotrak_file(file_path: str, zone_filter: str = None):
    """
    Parse TSI AeroTrak 9350 export (.xlsx or .xls) with sheet '6ChannelFixed'.
    Automatically detects Count Mode ('Σ' vs 'Δ') and Units ('# / m³' vs '#').
    """
    ext = os.path.splitext(file_path)[1].lower()
    grid = []
    if ext == ".xls":
        wb = xlrd.open_workbook(file_path)
        sh = wb.sheet_by_index(0)
        for r in range(sh.nrows):
            grid.append([sh.cell_value(r, c) for c in range(sh.ncols)])
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sh = wb["6ChannelFixed"] if "6ChannelFixed" in wb.sheetnames else wb.active
        for r in range(1, sh.max_row + 1):
            grid.append([sh.cell(r, c).value for c in range(1, sh.max_column + 1)])
        wb.close()

    count_mode = "Σ"
    units = "# / m³"
    serial = ""
    header_row_idx = None
    col_map = {}

    for r_idx, row in enumerate(grid):
        row_strs = [str(x).strip() if x is not None else "" for x in row]
        for c_i, val in enumerate(row_strs):
            if val == "Count Mode:":
                for nxt in row_strs[c_i + 1 : c_i + 6]:
                    if nxt:
                        count_mode = nxt
                        break
            elif val == "Units:":
                for nxt in row_strs[c_i + 1 : c_i + 6]:
                    if nxt:
                        units = nxt
                        break
            elif val == "Serial:":
                for nxt in row_strs[c_i + 1 : c_i + 6]:
                    if nxt:
                        serial = nxt
                        break
        if "#" in row_strs and "Zone" in row_strs and "Location" in row_strs:
            header_row_idx = r_idx
            for c_i, val in enumerate(row_strs):
                val_clean = val.replace("\n", " ").strip()
                if val_clean == "#":
                    col_map["id"] = c_i
                elif val_clean == "Zone":
                    col_map["zone"] = c_i
                elif val_clean == "Location":
                    col_map["location"] = c_i
                elif "Date/Time" in val_clean or "Date &" in val_clean:
                    col_map["datetime"] = c_i
                elif "Sample Time" in val_clean:
                    col_map["sample_time"] = c_i
                elif "0.3" in val_clean or "0,3" in val_clean:
                    col_map["0.3"] = c_i
                elif "0.5" in val_clean or "0,5" in val_clean:
                    col_map["0.5"] = c_i
                elif "Ch3" in val_clean or val_clean.startswith("1 "):
                    col_map["1.0"] = c_i
                elif "Ch4" in val_clean or val_clean.startswith("2 "):
                    col_map["2.0"] = c_i
                elif "Ch5" in val_clean or val_clean.startswith("3 "):
                    col_map["3.0"] = c_i
                elif "Ch6" in val_clean or val_clean.startswith("5 "):
                    col_map["5.0"] = c_i
                elif "Vol." in val_clean:
                    col_map["vol_l"] = c_i
            break

    if header_row_idx is None:
        raise ValueError(f"Could not find '#' / 'Zone' header row in {file_path}")

    samples = []
    for row in grid[header_row_idx + 1 :]:
        if len(row) <= col_map.get("id", 0):
            continue
        raw_id = row[col_map["id"]]
        if raw_id is None or str(raw_id).strip() == "":
            continue
        try:
            sample_id = int(float(raw_id))
        except ValueError:
            continue
        zone = str(row[col_map["zone"]]).strip() if "zone" in col_map else ""
        if zone_filter and zone_filter.lower() not in zone.lower():
            continue
        loc = str(row[col_map["location"]]).strip() if "location" in col_map else ""
        dt_str = str(row[col_map["datetime"]]).strip() if "datetime" in col_map else ""
        st_str = str(row[col_map["sample_time"]]).strip() if "sample_time" in col_map else ""
        vol_l = float(row[col_map["vol_l"]] or 0) if "vol_l" in col_map and row[col_map["vol_l"]] not in (None, "") else 50.0

        ch = {}
        for k in ("0.3", "0.5", "1.0", "2.0", "3.0", "5.0"):
            c_idx = col_map.get(k)
            val = float(row[c_idx] or 0) if c_idx is not None and row[c_idx] not in (None, "") else 0.0
            ch[k] = val

        # Convert to cumulative #/m3 if needed
        if "Δ" in count_mode or "Delta" in count_mode:
            cum_50 = ch["5.0"]
            cum_10 = ch["1.0"] + ch["2.0"] + ch["3.0"] + ch["5.0"]
            cum_05 = ch["0.5"] + cum_10
            cum_03 = ch["0.3"] + cum_05
        else:
            cum_03, cum_05, cum_10, cum_50 = ch["0.3"], ch["0.5"], ch["1.0"], ch["5.0"]

        if "m³" not in units and "m3" not in units.lower() and vol_l > 0:
            factor = 1000.0 / vol_l
            cum_03 *= factor
            cum_05 *= factor
            cum_10 *= factor
            cum_50 *= factor

        samples.append({
            "id": sample_id,
            "zone": zone,
            "location": loc,
            "datetime": dt_str,
            "sample_time": st_str,
            "vol_L": vol_l,
            "ge_0.3_m3": round(cum_03),
            "ge_0.5_m3": round(cum_05),
            "ge_1.0_m3": round(cum_10),
            "ge_5.0_m3": round(cum_50),
            "raw_channels": {k: round(v) for k, v in ch.items()},
        })

    print(
        f"=== AeroTrak 9350 ({os.path.basename(file_path)}) | S/N={serial} | "
        f"Mode={count_mode} | Units={units} | Samples={len(samples)} ==="
    )
    for s in samples:
        print(
            f"  #{s['id']:<4d} | Zone='{s['zone']}' | {s['datetime']} ({s['sample_time']}, {s['vol_L']:.0f}L) | "
            f">=0.5={s['ge_0.5_m3']:>8,d} | >=1.0={s['ge_1.0_m3']:>7,d} | >=5.0={s['ge_5.0_m3']:>6,d}"
        )
    return samples


def inspect_excel(file_path: str, max_rows: int = 80):
    """Inspect sheets, formulas, cached values, and scan for formula errors (#VALUE!, #N/A, #REF!)."""
    wb_f = openpyxl.load_workbook(file_path, data_only=False)
    wb_v = openpyxl.load_workbook(file_path, data_only=True)
    print(f"=== EXCEL INSPECT: {file_path} ===")
    for sname in wb_f.sheetnames:
        ws_f = wb_f[sname]
        ws_v = wb_v[sname]
        print(f"\n--- Sheet '{sname}' (rows={ws_f.max_row}, cols={ws_f.max_column}) ---")
        errors = []
        for r in range(1, ws_f.max_row + 1):
            row_items = []
            for c in range(1, ws_f.max_column + 1):
                fv = ws_f.cell(r, c).value
                vv = ws_v.cell(r, c).value
                coord = ws_f.cell(r, c).coordinate
                if isinstance(vv, str) and vv.startswith("#"):
                    # Ignore Feller lookup table tail Z382:AA550 per Bẫy 43
                    if not (c >= 26 and r >= 380):
                        errors.append((coord, fv, vv))
                if r <= max_rows and (fv is not None or vv is not None):
                    if isinstance(fv, str) and fv.startswith("="):
                        row_items.append(f"{coord}:{fv} => {vv!r}")
                    else:
                        row_items.append(f"{coord}:{vv!r}")
            if row_items and r <= max_rows:
                print(f"  R{r:02d}: " + " | ".join(row_items[:10]))
        if errors:
            print(f"  [FORMULA ERRORS in '{sname}']: {errors}")
        else:
            print(f"  [OK] 0 formula errors in '{sname}'")
    wb_f.close()
    wb_v.close()


def fill_excel_via_com(xlsx_path: str, spec_json_path: str):
    """
    Safely fill an Excel workbook via Excel COM (Bẫy 17, 29, 54, 55),
    recalculate, and scan for formula errors.
    spec_json format:
    {
      "sheets": {
        "Tinh_Toan_Khi_Co_Ap_V6": {
          "C34": "=DiemSuong!E5",
          "C41": 2.0,
          "C44": "KPH"
        },
        "DiemSuong": { ... }
      }
    }
    """
    import win32com.client
    import win32process

    abs_xlsx = os.path.abspath(xlsx_path)
    lock_file = os.path.join(os.path.dirname(abs_xlsx), "~$" + os.path.basename(abs_xlsx))
    if os.path.exists(lock_file):
        raise RuntimeError(f"LOCKED (Bẫy 4/29): Lock file exists at {lock_file}. Close Excel first.")

    with open(spec_json_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    print(f"Started isolated Excel COM instance (PID={xl_pid})")

    try:
        wb = xl.Workbooks.Open(abs_xlsx)
        if wb.ReadOnly:
            raise RuntimeError(f"Excel opened {abs_xlsx} as ReadOnly (Bẫy 29)!")

        sheets_spec = spec.get("sheets", spec)
        for sname, cell_map in sheets_spec.items():
            ws = wb.Worksheets(sname)
            for coord, val in cell_map.items():
                rg = ws.Range(coord)
                if rg.MergeCells:
                    rg = rg.MergeArea.Cells(1, 1)
                if isinstance(val, (int, float)):
                    rg.Formula = f"{val}"
                elif isinstance(val, str) and val.startswith("="):
                    rg.Formula = val
                else:
                    rg.Value2 = str(val)
            print(f"  Filled {len(cell_map)} cell(s) in sheet '{sname}'")

        xl.CalculateFull()
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"Saved and closed {os.path.basename(abs_xlsx)}")
    finally:
        xl.Quit()


def main():
    parser = argparse.ArgumentParser(description="Calculation & Excel Engine for QUATEST 3")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_las = sub.add_parser("lasair", help="Parse Lasair III database_all.csv")
    p_las.add_argument("csv", help="Path to database_all.csv")
    p_las.add_argument("--filter", help="Location substring filter")

    p_aero = sub.add_parser("aerotrak", help="Parse AeroTrak 9350 .xlsx or .xls (6ChannelFixed)")
    p_aero.add_argument("file", help="Path to .xlsx or .xls file")
    p_aero.add_argument("--filter", help="Zone substring filter")

    p_ins = sub.add_parser("inspect-excel", help="Inspect Excel sheets, formulas, values, and # errors")
    p_ins.add_argument("file", help="Path to .xlsx file")
    p_ins.add_argument("--max-rows", type=int, default=80)

    p_fill = sub.add_parser("fill-excel", help="Safely fill Excel via COM and recalculate")
    p_fill.add_argument("xlsx", help="Path to .xlsx file")
    p_fill.add_argument("spec_json", help="Path to JSON input spec")

    p_iso = sub.add_parser("iso8573-audit", help="Audit & calculate ISO 8573-1:2010 from TinhToanKhiNen.xlsx")
    p_iso.add_argument("xlsx", help="Path to TinhToanKhiNen.xlsx")
    p_iso.add_argument("--fix", action="store_true", help="Fix formula bugs (K6, Bẫy 15) in-place")

    p_dp = sub.add_parser("iso8573-dewpoint", help="Run CS Humidity Calculator & ISO 8573-1 Table 2 classification")
    p_dp.add_argument("--t1", required=True, help="Start,End dewpoint T1 (°C) e.g. '-39.77,-44.09'")
    p_dp.add_argument("--tkhi", required=True, help="Start,End gas temp (°C) e.g. '22.02,21.57'")
    p_dp.add_argument("--p", required=True, help="Start,End gauge pressure (bar) e.g. '7.31,7.32'")

    args = parser.parse_args()
    if args.cmd == "lasair":
        parse_lasair_csv(args.csv, args.filter)
    elif args.cmd == "aerotrak":
        parse_aerotrak_file(args.file, args.filter)
    elif args.cmd == "inspect-excel":
        inspect_excel(args.file, args.max_rows)
    elif args.cmd == "fill-excel":
        fill_excel_via_com(args.xlsx, args.spec_json)
    elif args.cmd == "iso8573-audit":
        try:
            from .iso8573_cal import audit_and_calc_from_excel
        except ImportError:
            from iso8573_cal import audit_and_calc_from_excel
        rep = audit_and_calc_from_excel(args.xlsx, fix_bugs=args.fix)
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    elif args.cmd == "iso8573-dewpoint":
        try:
            from .iso8573_cal import calc_dewpoint_cs
        except ImportError:
            from iso8573_cal import calc_dewpoint_cs
        t1_s, t1_e = map(float, args.t1.split(","))
        tk_s, tk_e = map(float, args.tkhi.split(","))
        p_s, p_e = map(float, args.p.split(","))
        res = calc_dewpoint_cs(t1_s, t1_e, tk_s, tk_e, p_s, p_e)
        print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
