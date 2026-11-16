# -*- coding: utf-8 -*-
"""
Automated Intake Helper for QUATEST 3 Monitoring Dossiers.
Combines:
1. Cách 2: 100% Offline Local Windows OCR (0 credit) on raw scan PDFs/images.
2. Cách 3: Ultra-lightweight profile.json generation (~300 - 500 tokens).

Usage:
    python Core/Tools/intake_helper.py --scan "Input/131638"
    python Core/Tools/intake_helper.py --ocr "Input/131638/scan.pdf"
"""

import os
import sys
import re
import json
import argparse
from typing import Dict, Any

WORKSPACE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

from Core.Tools.md_ocr import ocr_scanned_pdf_pages, pages_to_markdown, convert_document
from Core.Services.checkpoint import CheckpointStore, combined_files_hash, sha256_text
from Core.Tools.tax_helper import detect_tax_codes, verify_customer_with_tax, format_verification_report
from Core.Tools.equipment_names import as_dict  # tên máy chuẩn (AGENTS mục 14)


OCR_DPI = 200  # DPI render ảnh trước khi OCR (phải khớp giữa các lần chạy thì cache mới dùng lại được)


def extract_metadata_from_text(text: str, default_code: str = "", mst_override: str = "") -> Dict[str, Any]:
    """Extracts common metadata from OCR text using pattern matching."""
    meta: Dict[str, Any] = {
        "code": default_code,
        "reg_date": "",
        "sample_date": "",
        "cert_date": "",
        "client_name_vn": "",
        "client_name_en": "",
        "address_vn": "",
        "address_en": "",
        "dossier_type": "bsc",
        "executor": "Phát",
        "equipments": []
    }

    # 1. Dossier Code
    m_code = re.search(r"(Q3\d{2}[A-Z]\d{6}|\b\d{6}\b)", text)
    if m_code:
        c = m_code.group(1)
        if len(c) == 6:
            c = f"Q326D{c}"
        meta["code"] = c

    # 2. Dates (exclude form effective date 01/08/2025 per Rule 5)
    raw_dates = re.findall(r"\b(\d{1,2}/\d{1,2}/202[4-9])\b", text)
    valid_dates = [d for d in raw_dates if d not in ("01/08/2025", "01/8/2025", "29/09/2025")]
    if valid_dates:
        meta["reg_date"] = valid_dates[0]
        meta["sample_date"] = valid_dates[-1] if len(valid_dates) > 1 else valid_dates[0]
        meta["cert_date"] = meta["sample_date"]
    else:
        meta["reg_date"] = "22/10/2026"
        meta["sample_date"] = "22/10/2026"
        meta["cert_date"] = "22/10/2026"

    # 3. Candidate Client Name & Address from text
    cand_client_vn = ""
    cand_address_vn = ""
    m_client = re.search(r"(?:Đơn vị yêu cầu|Don vi yeu cau)[\s:\.]*([^\n\r\d]{5,80})", text, re.IGNORECASE)
    if m_client:
        c_name = m_client.group(1).strip()
        c_name = re.sub(r"^(là\s*|la\s*)", "", c_name, flags=re.IGNORECASE)
        cand_client_vn = c_name
    m_addr = re.search(r"(?:Địa chỉ|Dia chi|Địa điểm thực hiện|Dia diem thuc hien)[\s:\.]*([^\n\r]{10,120})", text, re.IGNORECASE)
    if m_addr:
        cand_address_vn = m_addr.group(1).strip()

    # 4. Tax Code Detection & Verification
    detected_msts = [mst_override] if mst_override else detect_tax_codes(text)
    if detected_msts:
        target_mst = detected_msts[0]
        tax_res = verify_customer_with_tax(cand_client_vn, target_mst, cand_address_vn)
        meta["tax_verification"] = tax_res
        if tax_res.get("has_discrepancy"):
            print("\n" + format_verification_report(tax_res) + "\n")
        else:
            print(f"[Tax Lookup] Khách hàng: {tax_res.get('suggested_client_vn')} (MST: {target_mst}) -> Trùng khớp chuẩn xác 100%.")

        meta["client_name_vn"] = tax_res.get("suggested_client_vn") or cand_client_vn
        meta["client_name_en"] = tax_res.get("suggested_client_en") or ""
        meta["address_vn"] = cand_address_vn or tax_res.get("suggested_address_vn") or ""
    elif "CONG TY TNHH GIAI PHAP Y SINH ABT" in text.upper() or "ABT" in text.upper():
        meta["client_name_vn"] = "Công ty TNHH Giải pháp Y sinh ABT"
        meta["client_name_en"] = "ABT Biomedical Solutions Company Limited"
        meta["address_vn"] = "Nhà xưởng 5.02, Lô L2, đường Long Hậu - Hiệp Phước, Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Long An"
        meta["address_en"] = "Factory 5.02, Lot L2, Long Hau - Hiep Phuoc street, Long Hau Industrial Park, Can Giuoc commune, Long An province"
    elif "PHUONG HAI" in text.upper() or "PHƯỢNG HẢI" in text.upper():
        meta["client_name_vn"] = "Công ty Cổ phần Khoa học Kỹ thuật Phượng Hải"
        meta["client_name_en"] = "Phuong Hai Technology Science Joint Stock Company"
        meta["address_vn"] = "103/23 Nguyễn Hữu Dật, phường Tây Thạnh, Thành phố Hồ Chí Minh"
        meta["address_en"] = "103/23 Nguyen Huu Dat Street, Tay Thanh Ward, Ho Chi Minh City"
    else:
        m_client = re.search(r"(?:Đơn vị yêu cầu|Don vi yeu cau)[\s:\.]*([^\n\r\d]{5,80})", text, re.IGNORECASE)
        if m_client:
            c_name = m_client.group(1).strip()
            c_name = re.sub(r"^(là\s*|la\s*)", "", c_name, flags=re.IGNORECASE)
            meta["client_name_vn"] = c_name
        m_addr = re.search(r"(?:Địa chỉ|Dia chi|Địa điểm thực hiện|Dia diem thuc hien)[\s:\.]*([^\n\r]{10,120})", text, re.IGNORECASE)
        if m_addr:
            meta["address_vn"] = m_addr.group(1).strip()

    # 5. Dossier Type detection
    lower_text = text.lower()
    if "tu hut" in lower_text or "fume hood" in lower_text or "bl.gha" in lower_text or "pha mau" in lower_text:
        meta["dossier_type"] = "fumehood"
    elif "khi nen" in lower_text or "compressed air" in lower_text or "diem suong" in lower_text or "iso 8573" in lower_text:
        meta["dossier_type"] = "compressed_air"
    elif "phong sach" in lower_text or "clean room" in lower_text or "cleanroom" in lower_text:
        meta["dossier_type"] = "cleanroom"
    elif "laf" in lower_text:
        meta["dossier_type"] = "laf"
    elif "an toan sinh hoc" in lower_text or "biosafety" in lower_text or "bsc" in lower_text or "en 12469" in lower_text:
        meta["dossier_type"] = "bsc"

    # 6. Equipment matching
    detected_eqs = []
    if re.search(r"tsi\s*8380", lower_text):
        detected_eqs.append(as_dict("tsi_8380"))
    if re.search(r"tsi\s*8375", lower_text):
        detected_eqs.append(as_dict("tsi_8375"))
    if re.search(r"kanomax|climomaster|6501", lower_text):
        detected_eqs.append(as_dict("kanomax_6501"))
    if re.search(r"ati|6d|photometer|2i", lower_text):
        detected_eqs.append(as_dict("ati_6d_2i"))
    if re.search(r"khoi|smoke\s*pencil", lower_text):
        detected_eqs.append(as_dict("smoke"))
    if re.search(r"testo\s*(?:440|540|kit\s*lux)", lower_text):
        detected_eqs.append(as_dict("testo_440"))
    if re.search(r"extech|sdl470|uvc", lower_text):
        detected_eqs.append(as_dict("extech_sdl470"))
    if re.search(r"rion|nl[\s-]*53|do on", lower_text):
        detected_eqs.append(as_dict("rion_nl53"))
    if re.search(r"aerotrak|9350", lower_text):
        detected_eqs.append(as_dict("aerotrak_9350"))

    meta["equipments"] = detected_eqs
    return meta


def create_profile_template(meta: Dict[str, Any], output_path: str) -> Dict[str, Any]:
    """Builds a complete, minimal JSON profile for run.py --profile."""
    d_type = meta.get("dossier_type", "bsc")
    code = meta.get("code") or "Q326D131638"
    sample_date = meta.get("sample_date") or meta.get("reg_date") or "22/10/2026"
    cert_date = meta.get("cert_date") or sample_date

    # Setup profile structure matching Core/run.py expectations
    profile_data: Dict[str, Any] = {
        "code": code,
        "cert_date": cert_date,
        "sample_date": sample_date,
        "reg_date": meta.get("reg_date") or sample_date,
        "sample_time_vn": f"Ngày {sample_date} từ 09:30 đến 11:30",
        "sample_time_en": f"From {sample_date}",
        "sub_no": None,
        "is_multi_cert": False,
        "dossier_type": d_type,
        "client": {
            "name_vn": meta.get("client_name_vn") or "Công ty Cổ phần ...",
            "name_en": meta.get("client_name_en") or "Company Name ...",
            "address_vn": meta.get("address_vn") or "Địa chỉ khách hàng ...",
            "address_en": meta.get("address_en") or "Client Address ..."
        },
        "personnel": {
            "executor": meta.get("executor", "Phát"),
            "sign_gdv_vn": "GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
            "sign_gdv_en": "INSPECTOR\n\n\n\nNguyen Xuan Nam",
            "sign_leader_vn": "KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
            "sign_leader_en": "FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
        },
        "equipments": meta.get("equipments", []),
        "locations": [
            {
                "name_vn": "Vị trí thực hiện giám định",
                "name_en": "Inspection location",
                "clean_class": "ISO 5",
                "state_vn": "Trạng thái nghỉ",
                "state_en": "At rest"
            }
        ],
        "results_data": {}
    }

    # Preset results_data template according to dossier type
    if d_type == "bsc":
        profile_data["scope_vn"] = "Giám định tủ an toàn sinh học cấp 2"
        profile_data["scope_en"] = "Inspection of Biosafety Cabinet Class II"
        profile_data["standards_vn"] = "QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000"
        profile_data["standards_en"] = "QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000"
        profile_data["basis_vn"] = "EN 12469:2000"
        profile_data["basis_en"] = "EN 12469:2000"
        profile_data["results_data"] = {
            "cabinet": {
                "cabinet_name": "Tủ an toàn sinh học cấp 2",
                "brand": "ABT",
                "model": "ISA2 - 1000",
                "serial": "26080302-01",
                "device_id": "N/A"
            },
            "inflow": {
                "method": 2,
                "method_name": "Kanomax Climomaster 6501",
                "location": "Mép trên cửa kính",
                "width_mm": 930,
                "height_mm": 200,
                "readings": [
                    [0.24, 0.20, 0.20, 0.20, 0.22],
                    [0.60, 0.54, 0.62, 0.56, 0.64]
                ],
                "velocity": 0.40
            },
            "downflow": {
                "width_mm": 930,
                "depth_mm": 450,
                "readings": [
                    [0.32, 0.35, 0.35, 0.36],
                    [0.33, 0.34, 0.33, 0.32]
                ]
            },
            "lux": {
                "readings": [
                    [910, 930, 920, 950]
                ]
            },
            "noise": {
                "ambient": 52.0,
                "operating": 64.0
            },
            "hepa": {
                "downflow_leak_pct": 0.001,
                "exhaust_leak_pct": 0.001
            },
            "uvc": {
                "readings": [85.0, 90.0, 88.0]
            }
        }
    elif d_type == "fumehood":
        profile_data["scope_vn"] = "Giám định tủ hút khí độc"
        profile_data["scope_en"] = 'Measuring/ validating "Fume Hood"'
        profile_data["standards_vn"] = "ASHRAE 110, SEFA 1-2010"
        profile_data["standards_en"] = "ASHRAE 110, SEFA 1-2010"
        profile_data["basis_vn"] = "SEFA 1-2010, ASHRAE 110"
        profile_data["basis_en"] = "SEFA 1-2010, ASHRAE 110"
        profile_data["results_data"] = {
            "cabinet": {
                "cabinet_name": "Tủ hút khí độc",
                "brand": "Phượng Hải JSC",
                "model": "BL.GHA-18",
                "serial": "BLGHA260418-01",
                "device_id": "N/A"
            },
            "openings": [
                {
                    "width_mm": 1530,
                    "height_mm": 585,
                    "opening_pct": 100,
                    "readings": [
                        [0.52, 0.53, 0.51, 0.52, 0.55],
                        [0.51, 0.53, 0.52, 0.51, 0.51]
                    ]
                }
            ],
            "lux": {"ambient": 15.0, "readings": [[860, 870, 860]]},
            "noise": {"ambient": 51.0, "operating": 70.0}
        }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(profile_data, f, ensure_ascii=False, indent=2)

    return profile_data


def process_scan_intake(dossier_dir: str, mst_override: str = "", force_ocr: bool = False,
                        store: Any = None) -> Dict[str, Any]:
    """
    Main intake pipeline for a dossier folder:
    1. Runs Windows OCR on scan PDF(s) -> scan_ocr.md (0 credit); kết quả lưu ở ThucHien/_checkpoint/ocr.json + ocr.xlsx,
       lần sau dùng lại nếu file scan chưa đổi (force_ocr=True để OCR lại)
    2. Extracts metadata and generates profile.json (~300 tokens)
    3. Splits PDF into E-doc (BBLM & Ghi Chep)
    """
    dossier_dir = os.path.abspath(dossier_dir)
    print(f"\n{'='*70}")
    print(f"[Intake Helper] Processing scan intake for: {dossier_dir}")
    print(f"{'='*70}")

    # Checkpoint: OCR chỉ chạy 1 lần, chạy lại thì đọc ThucHien/_checkpoint/ocr.json
    if store is None:
        store = CheckpointStore(dossier_dir=dossier_dir, force=["ocr"] if force_ocr else [])
    elif force_ocr:
        store.force.add("ocr")
    if force_ocr:
        store.force_refresh(["ocr"])

    # Chụp lại bản ghi tách-scan trước khi begin() các giai đoạn đứng trước xóa nó (để còn biết có bỏ qua được không)
    prev_split = None if store.is_forced("split_scan") else store.get_stage("split_scan")

    # Find scan PDFs
    pdf_files = sorted(
        os.path.join(dossier_dir, f) for f in os.listdir(dossier_dir)
        if f.lower().endswith(".pdf") and not f.startswith("TBKQ")
    )

    combined_text = ""
    ocr_md_path = os.path.join(dossier_dir, "scan_ocr.md")

    if pdf_files:
        print(f"Found {len(pdf_files)} scan PDF(s): {[os.path.basename(p) for p in pdf_files]}")
        store.begin("ocr", combined_files_hash(pdf_files, extra=f"dpi={OCR_DPI}"))
        n_cached = 0
        n_pages = 0
        try:
            for pdf_p in pdf_files:
                print(f"OCR (Windows offline, 0 credit; dùng lại checkpoint nếu file chưa đổi): {os.path.basename(pdf_p)}...")
                try:
                    if store.get_ocr(pdf_p, OCR_DPI, force=force_ocr) is not None:
                        n_cached += 1
                    pages = ocr_scanned_pdf_pages(pdf_p, dpi=OCR_DPI, store=store, force=force_ocr)
                    n_pages += len(pages)
                    txt = pages_to_markdown(pdf_p, pages)
                    combined_text += f"\n\n# --- File: {os.path.basename(pdf_p)} ---\n" + txt
                except Exception as e:
                    print(f"[Warning] OCR failed on {pdf_p}: {e}")

            # Save Markdown OCR text
            with open(ocr_md_path, "w", encoding="utf-8") as f:
                f.write(combined_text)
            print(f"[Success] Saved full local OCR text to: {ocr_md_path}")
            store.finish(
                "ocr",
                outputs=[ocr_md_path, store.path("ocr.json"), store.path("ocr.xlsx")],
                data={"files": [os.path.basename(p) for p in pdf_files], "pages": n_pages, "reused": n_cached},
                note=f"dùng lại cache {n_cached}/{len(pdf_files)} file" if n_cached else "",
            )
        except BaseException as e:
            store.fail("ocr", e)
            raise
    else:
        print("[Info] No scan PDF found directly in root, checking scan_ocr.md...")
        if os.path.isfile(ocr_md_path):
            with open(ocr_md_path, "r", encoding="utf-8") as f:
                combined_text = f.read()

    # Determine default code from directory name
    dir_name = os.path.basename(dossier_dir)
    m_dir = re.search(r"(\d{6})", dir_name)
    def_code = f"Q326D{m_dir.group(1)}" if m_dir else "Q326D131638"

    # Extract metadata (checkpoint metadata.json: cùng văn bản OCR thì không trích lại / không tra MST lại)
    text_hash = sha256_text(f"{combined_text}|{mst_override}|{def_code}")
    meta = None
    if not store.is_forced("intake"):
        saved = store.read_json("metadata.json", None)
        if saved and saved.get("text_hash") == text_hash and isinstance(saved.get("metadata"), dict):
            meta = saved["metadata"]
            print("[Intake] Dùng lại metadata đã trích (văn bản OCR không đổi).")
    store.begin("intake", text_hash)
    try:
        if meta is None:
            meta = extract_metadata_from_text(combined_text, default_code=def_code, mst_override=mst_override)
            store.write_json("metadata.json", {"text_hash": text_hash, "metadata": meta})
        print(f"[Metadata Extracted]: Code={meta['code']}, Type={meta['dossier_type']}, Client={meta['client_name_vn']}")

        # Generate profile.json
        profile_path = os.path.join(dossier_dir, "profile.json")
        if not os.path.isfile(profile_path):
            create_profile_template(meta, profile_path)
            print(f"[Success] Generated starter profile JSON: {profile_path}")
        else:
            print(f"[Info] profile.json already exists at: {profile_path}")
        store.finish("intake", outputs=[profile_path, store.path("metadata.json")],
                     data={"code": meta.get("code"), "dossier_type": meta.get("dossier_type")})
    except BaseException as e:
        store.fail("intake", e)
        raise

    # Split scan PDFs into TBKQ/Edoc
    split_hash = combined_files_hash(pdf_files, extra=str(meta.get("code")))
    if store.entry_reusable(prev_split, split_hash):
        store.restore_stage("split_scan", prev_split)
        print("[Info] Tách scan đã làm xong và file scan không đổi -> bỏ qua (dùng --force-stage split_scan để làm lại).")
    else:
        store.begin("split_scan", split_hash)
        try:
            from Core.Tools.scan_splitter import split_dossier_scans
            edoc_target = os.path.join(dossier_dir, "TBKQ", "Edoc")
            sp = split_dossier_scans(dossier_dir, output_dir=edoc_target, code_override=meta["code"],
                                     store=store, force_ocr=force_ocr)
            store.finish("split_scan", outputs=[sp.get("bblm_path"), sp.get("ghi_chep_path")],
                         data={"bblm_pages": sp.get("bblm_pages"), "ghi_chep_pages": sp.get("ghi_chep_pages")})
        except Exception as e:
            store.fail("split_scan", e)
            print(f"[Warning] E-doc split encountered an issue: {e}")

    print(f"\n[Intake Completed] Everything prepared in: {dossier_dir}")
    print(f"-> Local OCR text: {ocr_md_path}")
    print(f"-> Starter Profile: {profile_path}")
    print(f"{'='*70}\n")

    return {
        "ocr_md_path": ocr_md_path,
        "checkpoint_dir": store.dir,
        "profile_path": profile_path,
        "metadata": meta
    }


def main():
    parser = argparse.ArgumentParser(description="QUATEST 3 Automated Intake Helper (Cách 2 & Cách 3)")
    parser.add_argument("--scan", type=str, help="Process dossier folder (run local OCR, extract metadata, generate profile.json)")
    parser.add_argument("--ocr", type=str, help="Run local OCR on a single PDF/image file and output markdown")
    parser.add_argument("--mst", type=str, default="", help="Mã số thuế doanh nghiệp cần đối soát (tùy chọn)")
    parser.add_argument("--force-ocr", action="store_true", help="Bỏ qua checkpoint, OCR lại toàn bộ file scan")
    args = parser.parse_args()

    if args.ocr:
        txt = convert_document(args.ocr)
        print(txt)
    elif args.scan:
        process_scan_intake(args.scan, mst_override=args.mst, force_ocr=args.force_ocr)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
