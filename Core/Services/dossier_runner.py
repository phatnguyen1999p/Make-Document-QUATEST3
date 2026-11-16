# -*- coding: utf-8 -*-
"""
End-to-End Orchestrator Pipeline for QUATEST 3 Inspection Dossiers.
Executes the full generation workflow:
1. Build BCKQ (Internal Inspection Report BM3030.03)
2. Build TBKQ VN & EN (Official Certificates)
3. Convert all deliverables to PDF and sync to TBKQ/, Output/, and Z:
4. Perform automated QA Verification (qa_check.py)
"""

import os
import re
from typing import Dict, Any, Iterable

from ..Models.dossier import Dossier, assign_cert_numbers
from .checkpoint import CheckpointStore, combined_files_hash
from ..Builders.bckq_builder import BCKQBuilder
from ..Builders.tbkq_builder import TBKQBuilder
from ..Builders.handlers.bsc import BSCHandler
from ..Tools.bsc_calc import build_bsc_calculation_workbook, verify_bsc_calculations, print_mismatch_warnings
# PDFExporter imported lazily to avoid circular imports
from ..Tools.qa_check import check_single_docx


class DossierRunner:
    """Pipeline runner for complete dossier document processing."""

    def __init__(self, dossier: Dossier):
        self.dossier = dossier
        self.store = None          # CheckpointStore (ThucHien/_checkpoint), tạo trong run_all
        self._cur_stage = None     # giai đoạn đang chạy, để ghi "failed" nếu pipeline văng lỗi

    def _begin(self, stage: str, inputs_hash: str = None) -> None:
        self._cur_stage = stage
        if self.store is not None:
            self.store.begin(stage, inputs_hash)

    def _finish(self, stage: str, outputs: Iterable = (), data: Dict[str, Any] = None, note: str = "",
                status: str = "done") -> None:
        if self.store is not None:
            self.store.finish(stage, outputs=outputs, data=data, note=note, status=status)
        self._cur_stage = None

    def _dossier_fingerprint(self) -> str:
        """Mã băm dữ liệu hồ sơ (để biết dữ liệu đầu vào của BCKQ/TBKQ có đổi so với lần dựng trước không)."""
        try:
            import dataclasses
            import json
            from .checkpoint import sha256_text
            return sha256_text(json.dumps(dataclasses.asdict(self.dossier), sort_keys=True, default=str, ensure_ascii=False))
        except Exception as _err:
            print(f"[Warning] dossier_runner fingerprint {type(_err).__name__}: {_err}")
            return ""

    @classmethod
    def run_batch(cls, dossiers, **run_kwargs):
        """
        Chạy nhiều báo cáo kết quả của cùng hồ sơ: tự đánh số phụ /1, /2, /3 ... (1 báo cáo thì mã trần)
        rồi chạy run_all() cho từng báo cáo. Trả về danh sách kết quả theo thứ tự.
        """
        dossiers = assign_cert_numbers(list(dossiers))
        for d in dossiers:
            print(f"[DossierRunner] {d.code}: số chứng thư {d.cert_no_vn} / {d.cert_no_en}")
        # Hồ sơ nhiều tủ (BSC/LAF, tủ cấy, tủ hút): gộp BCKQ vào 1 file (người dùng cho phép 30/10/2026).
        # merge_bckq=False để giữ BCKQ từng tủ.
        from ..Builders.bckq_builder import is_cabinet
        merge_bckq = run_kwargs.pop("merge_bckq", True) and len(dossiers) > 1 and all(is_cabinet(d) for d in dossiers)
        results = [cls(d).run_all(build_bckq=not merge_bckq, **run_kwargs) for d in dossiers]
        if merge_bckq:
            try:
                bckq = BCKQBuilder.build_merged(dossiers)
                for r in results:
                    if isinstance(r, dict):
                        r["bckq_path"] = bckq
                if run_kwargs.get("export_pdf", True):
                    from ..Exporters.pdf_exporter import PDFExporter
                    PDFExporter(dossiers[0]).export_files([bckq])
                if run_kwargs.get("run_qa", True):
                    from ..Tools.qa_check import check_single_docx
                    print(f"[QA] {os.path.basename(bckq)}: {check_single_docx(bckq, code=dossiers[0].code).get('status')}")
            except Exception as _err:
                print(f"[Warning] Gộp BCKQ lỗi ({type(_err).__name__}: {_err}); chạy lại với merge_bckq=False để có BCKQ từng tủ.")
        # Gom tem các tủ vào 1 file, 3 tủ / trang (người dùng chốt 30/10/2026); file tem lẻ từng tủ vẫn giữ
        tems = [r.get("tem_tu_path") for r in results if isinstance(r, dict) and r.get("tem_tu_path")]
        if len(tems) > 1:
            try:
                from ..Builders.tag_context import build_tem_tu_sheet
                d0 = dossiers[0]
                sheet = build_tem_tu_sheet(tems, os.path.join(d0.tbkq_dir, f"{d0.code.replace('Q326D', '')} - tem tu.docx"))
                for r in results:
                    if isinstance(r, dict):
                        r["tem_tu_sheet_path"] = sheet
                # Gom xong thì bỏ file tem lẻ vừa sinh trong lần chạy này (người dùng yêu cầu dọn file 30/10/2026)
                if sheet and os.path.isfile(sheet):
                    for t in tems:
                        try:
                            if os.path.abspath(t) != os.path.abspath(sheet):
                                os.remove(t)
                        except OSError as _e:
                            print(f"[Warning] Không xóa được tem lẻ {t}: {_e}")
                    for r in results:
                        if isinstance(r, dict):
                            r["tem_tu_path"] = sheet
                if run_kwargs.get("export_pdf", True):
                    from ..Exporters.pdf_exporter import PDFExporter
                    PDFExporter(d0).export_files([sheet])
            except Exception as _err:
                print(f"[Warning] Gom tem tủ lỗi ({type(_err).__name__}: {_err}); vẫn còn file tem từng tủ.")
        return results

    def run_all(self, export_pdf: bool = True, run_qa: bool = True, resume: bool = False,
                force_stages: Iterable[str] = (), build_bckq: bool = True) -> Dict[str, Any]:
        """
        Chạy toàn bộ pipeline và ghi tiến độ từng giai đoạn vào ThucHien/_checkpoint/manifest.json.

        resume=True: bỏ qua giai đoạn đã xong, dữ liệu đầu vào không đổi và file đầu ra còn đủ
        (hiện áp dụng cho split_scan; các giai đoạn dựng BCKQ/TBKQ/PDF/QA chỉ được ghi nhận).
        force_stages: tên giai đoạn bắt buộc làm lại (ocr, split_scan, ...), vd ["ocr"].
        """
        try:
            self.store = CheckpointStore(thuchien_dir=self.dossier.thuchien_dir, force=list(force_stages))
            if force_stages:
                self.store.force_refresh(force_stages)
        except Exception as _err:
            print(f"[Warning] Không khởi tạo được checkpoint ({type(_err).__name__}: {_err}); chạy không ghi tiến độ.")
            self.store = None
        try:
            return self._run_all_impl(export_pdf, run_qa, resume, build_bckq)
        except BaseException as e:
            if self.store is not None and self._cur_stage:
                self.store.fail(self._cur_stage, e)
            raise

    def _run_all_impl(self, export_pdf: bool, run_qa: bool, resume: bool, build_bckq: bool = True) -> Dict[str, Any]:
        """
        Executes complete pipeline:
        BCKQ -> TBKQ VN -> TBKQ EN -> PDF Conversion -> QA Check
        """
        print(f"\n{'='*70}")
        print(f"STARTING AUTOMATED DOSSIER RUNNER FOR: {self.dossier.code} ({self.dossier.dossier_type.upper()})")
        print(f"{'='*70}")

        results = {
            "dossier": self.dossier.code,
            "bckq_path": None,
            "tbkq_vn_path": None,
            "tbkq_en_path": None,
            "pdf_paths": [],
            "qa_results": {},
            "edoc_bblm_path": None,
            "edoc_ghichep_path": None
        }

        # 0. Split input scan PDFs into TBKQ/Edoc
        print("\n--- [Step 0/4] Splitting Raw Scan PDFs (/TBKQ/Edoc) ---")
        try:
            from ..Tools.scan_splitter import split_dossier_scans, find_candidate_scan_pdfs
            edoc_target = self.dossier.edoc_dir
            try:
                split_hash = combined_files_hash(find_candidate_scan_pdfs(self.dossier.input_dir), extra=self.dossier.code)
            except Exception as _err:
                print(f"[Warning] Không tính được mã băm file scan ({type(_err).__name__}: {_err})")
                split_hash = None
            prev = self.store.get_stage("split_scan") if self.store is not None else None
            if resume and split_hash and prev and self.store.should_skip("split_scan", split_hash):
                print("[Resume] Đã tách scan xong và file scan không đổi -> bỏ qua.")
                edoc_res = {
                    "bblm_path": (prev.get("data") or {}).get("bblm_path"),
                    "ghi_chep_path": (prev.get("data") or {}).get("ghi_chep_path"),
                    "code_6": (prev.get("data") or {}).get("code_6"),
                }
                for k in ("bblm_path", "ghi_chep_path"):
                    if edoc_res[k] and not os.path.isabs(edoc_res[k]):
                        edoc_res[k] = os.path.normpath(os.path.join(self.store.base_dir, edoc_res[k]))
            else:
                self._begin("split_scan", split_hash)
                edoc_res = split_dossier_scans(
                    dossier_dir=self.dossier.input_dir,
                    output_dir=edoc_target,
                    code_override=self.dossier.code,
                    store=self.store,
                )
                rel = (lambda p: os.path.relpath(p, self.store.base_dir) if (p and self.store is not None) else p)
                self._finish(
                    "split_scan",
                    outputs=[edoc_res.get("bblm_path"), edoc_res.get("ghi_chep_path"), edoc_res.get("khac_path")],
                    data={"bblm_path": rel(edoc_res.get("bblm_path")), "ghi_chep_path": rel(edoc_res.get("ghi_chep_path")),
                          "khac_path": rel(edoc_res.get("khac_path")),
                          "code_6": edoc_res.get("code_6")},
                )
            results["edoc_bblm_path"] = edoc_res.get("bblm_path")
            results["edoc_ghichep_path"] = edoc_res.get("ghi_chep_path")
            results["edoc_khac_path"] = edoc_res.get("khac_path")

            # Synchronize to network directory if available
            if self.dossier.network_dir and os.path.isdir(self.dossier.network_dir):
                import shutil
                net_edoc = os.path.join(self.dossier.network_dir, "TBKQ", "Edoc")
                os.makedirs(net_edoc, exist_ok=True)
                for edoc_p in [edoc_res.get("bblm_path"), edoc_res.get("ghi_chep_path"), edoc_res.get("khac_path")]:
                    if edoc_p and os.path.isfile(edoc_p):
                        shutil.copy2(edoc_p, os.path.join(net_edoc, os.path.basename(edoc_p)))
                # Clean up redundant legacy misnamed files on network
                code_6 = edoc_res.get("code_6")
                if code_6 and os.path.isdir(net_edoc):
                    valid_net_names = {f"{code_6} - BBLM.pdf", f"{code_6} - Ghi Chep.pdf", f"{code_6} - Khac.pdf"}
                    for nf in os.listdir(net_edoc):
                        if nf in valid_net_names:
                            continue
                        if code_6 in nf and (nf.endswith(" - BBLM.pdf") or nf.endswith(" - Ghi Chep.pdf")):
                            try:
                                os.remove(os.path.join(net_edoc, nf))
                            except OSError:
                                pass
        except Exception as e:
            if self.store is not None and self._cur_stage == "split_scan":
                self.store.fail("split_scan", e)
                self._cur_stage = None
            print(f"[Warning] Scan PDF splitting encountered an issue: {e}")

        # Automated Computerized Layout Generation (Step 2 - Layout vi tính)
        is_bsc = self.dossier.dossier_type.lower() in ("bsc", "biological_safety_cabinet", "biosafety_cabinet", "laf", "laf_bsc", "tu cay", "tu_cay", "tủ cấy")
        is_fumehood = self.dossier.dossier_type.lower() in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút")

        if is_bsc and not self.dossier.layouts:
            downflow_data = self.dossier.results_data.get("downflow", {})
            if downflow_data and downflow_data.get("readings"):
                try:
                    from ..Tools.bsc_calc import generate_bsc_downflow_grid_image
                    layout_p = os.path.join(self.dossier.tbkq_dir, "layout_downflow.png")
                    w = downflow_data.get("width_mm", 930)
                    d = downflow_data.get("depth_mm", 450)
                    r = downflow_data.get("readings", [])
                    generate_bsc_downflow_grid_image(layout_p, width_mm=w, depth_mm=d, velocity_grid=r)
                    self.dossier.layouts.append({
                        "path": layout_p,
                        "caption_vn": "Sơ đồ đo tốc độ gió cấp qua lọc HEPA (Downflow)",
                        "caption_en": "Downflow air velocity measuring layout"
                    })
                    print(f"[Computerized Layout] Automatically generated: {layout_p}")
                except Exception as e:
                    print(f"[Warning] Failed to generate automated BSC layout: {e}")

        # 1. Build BCKQ
        print("\n--- [Step 1/4] Generating Internal Report (BCKQ BM3030.03) ---")
        dossier_fp = self._dossier_fingerprint()
        bckq_file = None
        if build_bckq:
            self._begin("bckq", dossier_fp)
            bckq_builder = BCKQBuilder(self.dossier)
            bckq_file = bckq_builder.build()
            results["bckq_path"] = bckq_file
            self._finish("bckq", outputs=[bckq_file])
        else:
            print("[DossierRunner] BCKQ gộp chung các tủ (run_batch) -> bỏ BCKQ riêng của báo cáo này.")

        # 2. Build TBKQ (VN and EN)
        print("\n--- [Step 2/4] Generating Official Certificates (TBKQ VN & EN) ---")
        self._begin("tbkq", dossier_fp)
        tbkq_builder = TBKQBuilder(self.dossier)
        tbkq_res = tbkq_builder.build_all()
        if len(tbkq_res) == 4:
            tbkq_vn_file, tbkq_en_file, tbkq_vn_noeval, tbkq_en_noeval = tbkq_res
            results["tbkq_vn_path"] = tbkq_vn_file
            results["tbkq_en_path"] = tbkq_en_file
            results["tbkq_vn_noeval_path"] = tbkq_vn_noeval
            results["tbkq_en_noeval_path"] = tbkq_en_noeval
            created_tbkq = [tbkq_vn_file, tbkq_en_file, tbkq_vn_noeval, tbkq_en_noeval]
        else:
            tbkq_vn_file, tbkq_en_file = tbkq_res
            results["tbkq_vn_path"] = tbkq_vn_file
            results["tbkq_en_path"] = tbkq_en_file
            created_tbkq = [tbkq_vn_file, tbkq_en_file]

        is_bsc = self.dossier.dossier_type.lower() in ("bsc", "biological_safety_cabinet", "biosafety_cabinet", "laf", "laf_bsc", "tu cay", "tu_cay", "tủ cấy")
        is_fumehood = self.dossier.dossier_type.lower() in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút")
        tem_tu_file = None
        excel_calc_file = None
        if is_fumehood:
            print("\n--- [Fume Hood Specialized] Generating Validation Seal (tem tu.docx) ---")
            try:
                from ..Builders.handlers.fumehood import FumeHoodHandler
                fh_h = FumeHoodHandler(self.dossier)
                tem_tu_file = fh_h.build_tem_tu()
                results["tem_tu_path"] = tem_tu_file
            except Exception as e:
                print(f"[Warning] Failed to generate Validation Seal: {e}")

            print("\n--- [Fume Hood Specialized] Generating Excel Calculation Workbook (.xlsx) ---")
            try:
                from ..Tools.fumehood_calc import build_fumehood_calculation_workbook, verify_fumehood_calculations, print_fumehood_mismatch_warnings
                sub_part = f" - {self.dossier.sub_no}" if (self.dossier.is_multi_cert and self.dossier.sub_no) else ""
                excel_name = f"{self.dossier.code}{sub_part} - TinhToanFumehood.xlsx"
                excel_target_thuchien = os.path.join(self.dossier.thuchien_dir, excel_name)

                # Prepare dossier data dictionary for Excel calculation generator
                c_data = self.dossier.results_data.copy()
                c_data["code"] = self.dossier.code
                c_data["sub_no"] = self.dossier.sub_no or ""
                c_data["reg_date"] = self.dossier.reg_date or self.dossier.results_data.get("reg_date") or self.dossier.sample_date or self.dossier.cert_date
                c_data["sample_date"] = self.dossier.sample_date or self.dossier.cert_date
                c_data["cert_date"] = self.dossier.cert_date
                c_data["photos"] = self.dossier.photos
                c_data["input_dir"] = self.dossier.input_dir

                m_time = re.search(r'từ\s*(\d{1,2}[:h]\d{2})\s*đến\s*(\d{1,2}[:h]\d{2})', self.dossier.sample_time_vn or '')
                if m_time:
                    c_data["start_time"] = m_time.group(1).replace('h', ':')
                    c_data["end_time"] = m_time.group(2).replace('h', ':')
                else:
                    c_data["start_time"] = self.dossier.results_data.get("start_time", "10:00")
                    c_data["end_time"] = self.dossier.results_data.get("end_time", "12:00")
                c_data["client_name_vn"] = self.dossier.client.name_vn
                c_data["scope_vn"] = self.dossier.scope_vn
                c_data["scope_en"] = self.dossier.scope_en
                c_data["client_address_vn"] = self.dossier.client.address_vn
                c_data["client_name_en"] = self.dossier.client.name_en
                c_data["client_address_en"] = self.dossier.client.address_en
                c_data["location_vn"] = self.dossier.location_desc_vn
                c_data["location_en"] = self.dossier.location_desc_en
                pers = self.dossier.personnel
                leader_vn = getattr(pers, "sign_leader_vn", "") if pers else ""
                gdv_vn = getattr(pers, "sign_gdv_vn", "") if pers else ""
                leader_en = getattr(pers, "sign_leader_en", "") if pers else ""
                gdv_en = getattr(pers, "sign_gdv_en", "") if pers else ""
                c_data["leader_name_vn"] = leader_vn.split("\n")[-1].strip() if leader_vn else "Lê Đăng Huyền"
                c_data["gdv_name_vn"] = gdv_vn.split("\n")[-1].strip() if gdv_vn else "Nguyễn Xuân Nam"
                c_data["leader_name_en"] = leader_en.split("\n")[-1].strip() if leader_en else "Le Dang Huyen"
                c_data["gdv_name_en"] = gdv_en.split("\n")[-1].strip() if gdv_en else "Nguyen Xuan Nam"
                c_data["equipments"] = [eq.name_vn for eq in self.dossier.equipments]
                c_data["equipments_en"] = [eq.name_en for eq in self.dossier.equipments]
                cab_info = self.dossier.results_data.get("cabinet", {})
                c_data["cabinet_name"] = cab_info.get("cabinet_name", "Tủ hút phá mẫu")
                c_data["brand"] = cab_info.get("brand", "")
                c_data["model"] = cab_info.get("model", "")
                c_data["serial"] = cab_info.get("serial", "")
                c_data["device_id"] = cab_info.get("device_id", "N/A")

                # Check independent calculation vs scan reference
                mismatches = verify_fumehood_calculations(c_data)
                if mismatches:
                    print_fumehood_mismatch_warnings(mismatches)
                    results["calculation_mismatches"] = mismatches

                build_fumehood_calculation_workbook(c_data, excel_target_thuchien)
                results["excel_calc_path"] = excel_target_thuchien
            except Exception as e:
                print(f"[Warning] Failed to generate Fume Hood calculation workbook: {e}")

        elif is_bsc:
            print("\n--- [BSC Specialized] Generating Validation Seal (tem tu.docx) ---")
            try:
                bsc_h = BSCHandler(self.dossier)
                tem_tu_file = bsc_h.build_tem_tu()
                results["tem_tu_path"] = tem_tu_file
            except Exception as e:
                print(f"[Warning] Failed to generate Validation Seal: {e}")

            print("\n--- [BSC Specialized] Generating Excel Calculation Workbook (.xlsx) ---")
            try:
                sub_part = f" - {self.dossier.sub_no}" if (self.dossier.is_multi_cert and self.dossier.sub_no) else ""
                loai_calc = "TuCay" if self.dossier.dossier_type.lower() in ("laf", "tu cay", "tu_cay", "tủ cấy") else "BSC"
                excel_name = f"{self.dossier.code}{sub_part} - TinhToan{loai_calc}.xlsx"
                excel_target_thuchien = os.path.join(self.dossier.thuchien_dir, excel_name)

                # Prepare dossier data dictionary for Excel calculation generator
                c_data = self.dossier.results_data.copy()
                c_data["code"] = self.dossier.code
                c_data["sub_no"] = self.dossier.sub_no or ""
                c_data["reg_date"] = self.dossier.reg_date or self.dossier.results_data.get("reg_date") or self.dossier.sample_date or self.dossier.cert_date
                c_data["sample_date"] = self.dossier.sample_date or self.dossier.cert_date
                c_data["cert_date"] = self.dossier.cert_date

                m_time = re.search(r'từ\s*(\d{1,2}[:h]\d{2})\s*đến\s*(\d{1,2}[:h]\d{2})', self.dossier.sample_time_vn or '')
                if m_time:
                    c_data["start_time"] = m_time.group(1).replace('h', ':')
                    c_data["end_time"] = m_time.group(2).replace('h', ':')
                else:
                    c_data["start_time"] = self.dossier.results_data.get("start_time", "11:00")
                    c_data["end_time"] = self.dossier.results_data.get("end_time", "16:00")
                c_data["client_name_vn"] = self.dossier.client.name_vn
                c_data["scope_vn"] = self.dossier.scope_vn
                c_data["scope_en"] = self.dossier.scope_en
                c_data["client_address_vn"] = self.dossier.client.address_vn
                c_data["client_name_en"] = self.dossier.client.name_en
                c_data["client_address_en"] = self.dossier.client.address_en
                c_data["location_vn"] = self.dossier.location_desc_vn
                c_data["location_en"] = self.dossier.location_desc_en
                pers = self.dossier.personnel
                leader_vn = getattr(pers, "sign_leader_vn", "") if pers else ""
                gdv_vn = getattr(pers, "sign_gdv_vn", "") if pers else ""
                leader_en = getattr(pers, "sign_leader_en", "") if pers else ""
                gdv_en = getattr(pers, "sign_gdv_en", "") if pers else ""
                c_data["leader_name_vn"] = leader_vn.split("\n")[-1].strip() if leader_vn else "Lê Đăng Huyền"
                c_data["gdv_name_vn"] = gdv_vn.split("\n")[-1].strip() if gdv_vn else "Nguyễn Xuân Nam"
                c_data["leader_name_en"] = leader_en.split("\n")[-1].strip() if leader_en else "Le Dang Huyen"
                c_data["gdv_name_en"] = gdv_en.split("\n")[-1].strip() if gdv_en else "Nguyen Xuân Nam"
                c_data["equipments"] = [eq.name_vn for eq in self.dossier.equipments]
                c_data["equipments_en"] = [eq.name_en for eq in self.dossier.equipments]
                cab_info = self.dossier.results_data.get("cabinet", {})
                default_cab_name = "Tủ cấy vô trùng" if self.dossier.dossier_type.lower() in ("laf", "tu cay", "tu_cay", "tủ cấy") else "Tủ an toàn sinh học cấp 2"
                c_data["cabinet_name"] = cab_info.get("cabinet_name", default_cab_name)
                c_data["brand"] = cab_info.get("brand", "")
                c_data["model"] = cab_info.get("model", "")
                c_data["serial"] = cab_info.get("serial", "")
                c_data["device_id"] = cab_info.get("device_id", "N/A")

                # Check independent calculation vs scan reference
                mismatches = verify_bsc_calculations(c_data)
                if mismatches:
                    print_mismatch_warnings(mismatches)
                    results["calculation_mismatches"] = mismatches

                build_bsc_calculation_workbook(c_data, excel_target_thuchien)
                results["excel_calc_path"] = excel_target_thuchien
            except Exception as e:
                print(f"[Warning] Failed to generate Excel calculation workbook: {e}")

        # 3. Export to PDF & Synchronize (Word DOCX & Excel XLSX)
        created_docx = [p for p in [bckq_file] + created_tbkq + ([tem_tu_file] if tem_tu_file else []) if p]
        if export_pdf:
            # Kết thúc giai đoạn tbkq (đã dựng xong TBKQ, tem tủ, Excel tính toán)
            self._finish("tbkq", outputs=created_tbkq + [tem_tu_file, results.get("excel_calc_path")])
            print("\n--- [Step 3/4] Exporting to PDF via Word & Excel COM & Syncing Folders ---")
            self._begin("pdf")
            from ..Exporters.pdf_exporter import PDFExporter
            exporter = PDFExporter(self.dossier)
            pdf_files = exporter.export_files(created_docx)

            # Export calculation Excel files in /ThucHien to /pdf/
            xlsx_to_export = []
            if results.get("excel_calc_path") and os.path.isfile(results["excel_calc_path"]):
                xlsx_to_export.append(results["excel_calc_path"])
            elif os.path.isdir(self.dossier.thuchien_dir):
                for f in os.listdir(self.dossier.thuchien_dir):
                    if f.endswith(".xlsx") and not f.startswith("~$") and ("tinhtoan" in f.lower() or "kn" in f.lower()):
                        xlsx_to_export.append(os.path.join(self.dossier.thuchien_dir, f))

            if xlsx_to_export:
                excel_pdfs = exporter.export_excel_files(xlsx_to_export)
                pdf_files.extend(excel_pdfs)

            results["pdf_paths"] = pdf_files
            self._finish("pdf", outputs=list(pdf_files or []))
        else:
            self._finish("tbkq", outputs=created_tbkq + [tem_tu_file, results.get("excel_calc_path")])
            print("\n--- [Step 3/4] PDF Export Skipped ---")
            self._finish("pdf", status="skipped", note="--no-pdf")

        # 4. Run Automated QA Check
        all_qa_clean = True
        if run_qa:
            print("\n--- [Step 4/4] Running Automated QA Verification (qa_check.py) ---")
            self._begin("qa")
            for docx_p in created_docx:
                qa_res = check_single_docx(docx_p, code=self.dossier.code)
                status = qa_res.get("status")
                results["qa_results"][os.path.basename(docx_p)] = status
                if status != "CLEAN":
                    all_qa_clean = False
            self._finish("qa", data={"results": results["qa_results"]})
        else:
            self._finish("qa", status="skipped", note="--no-qa")

        print(f"\n{'='*70}")
        if all_qa_clean:
            print(f"PIPELINE COMPLETED SUCCESSFULLY FOR: {self.dossier.code}")
        else:
            print(f"PIPELINE COMPLETED WITH QA WARNINGS/FAILURES FOR: {self.dossier.code}")
        print(f"  Step 2 (ThucHien): {self.dossier.thuchien_dir}")
        print(f"  Step 3 (TBKQ):     {self.dossier.tbkq_dir}")
        print(f"  Edoc (Scan):       {self.dossier.edoc_dir}")
        print(f"  PDF Output:        {self.dossier.pdf_dir}")
        print(f"{'='*70}\n")

        return results
