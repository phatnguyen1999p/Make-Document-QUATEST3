# -*- coding: utf-8 -*-
"""
PDF Conversion and Export Dispatcher for QUATEST 3 Documents.
Uses Word & Excel COM for 100% pixel-perfect PDF rendering and automatically dispatches
to target directory structure: Input/<dossier>/pdf/, Output/, and Z: network drive.
"""

import os
import gc
import shutil
import time
import subprocess
from typing import List, Optional, Dict, Any

from ..Models.dossier import Dossier
from ..Services.upchungthu_bridge import (
    is_tbkq_file,
    process_tbkq_pdf,
    apply_tbkq_background,
    verify_tbkq_pdf,
)


# --- Khong dung may in tem HPL ----------------------------------------------
# Word/Excel dan trang PDF theo may in mac dinh. Neu may in mac dinh la may in tem
# (ten co "HPL"), file PDF bi chia thanh hang tram trang nho (vd phieu tinh toan 132115
# ra 965 trang). Truoc khi mo Word/Excel, tam doi may in mac dinh sang may in khac
# (uu tien "Microsoft Print to PDF"), va tra lai may in cu khi tien trinh Python ket thuc.
# May in tem tren may N2: "HPRT HT330 - ZPL #2" (nguoi dung goi la HPL). Regex, khong phan biet hoa/thuong.
LABEL_PRINTER_PATTERN = os.environ.get("QUATEST3_LABEL_PRINTER", r"HPL|HPRT|ZPL")
_printer_restore = {"done": False, "original": None}


def avoid_label_printer() -> None:
    """Dam bao may in mac dinh khong phai may in tem HPL truoc khi xuat PDF qua Word/Excel COM."""
    if _printer_restore["done"]:
        return
    try:
        import win32print
    except ImportError:
        return
    try:
        current = win32print.GetDefaultPrinter()
    except Exception as e:
        print(f"[Warning] Khong doc duoc may in mac dinh: {e}")
        return
    import re
    is_label = lambda n: re.search(LABEL_PRINTER_PATTERN, n or "", re.IGNORECASE) is not None
    if not is_label(current):
        print(f"[PDFExporter] May in mac dinh: {current} (khong phai may in tem)")
        _printer_restore["done"] = True
        return
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    names = [p[2] for p in win32print.EnumPrinters(flags)]
    ok = [n for n in names if not is_label(n)]
    target = next((n for n in ok if n.lower() == "microsoft print to pdf"), None) \
        or next((n for n in ok if "pdf" in n.lower()), None) \
        or next((n for n in ok if not any(k in n.lower() for k in ("onenote", "fax", "xps"))), None)
    if not target:
        raise RuntimeError(f"May in mac dinh la '{current}' (may in tem) va khong co may in nao khac de xuat PDF.")
    win32print.SetDefaultPrinter(target)
    _printer_restore.update(done=True, original=current)
    print(f"[PDFExporter] Tam doi may in mac dinh '{current}' -> '{target}' de xuat PDF (khong dung may in tem).")
    import atexit
    atexit.register(_restore_default_printer)


def _restore_default_printer() -> None:
    orig = _printer_restore.get("original")
    if not orig:
        return
    try:
        import win32print
        win32print.SetDefaultPrinter(orig)
        print(f"[PDFExporter] Da tra lai may in mac dinh: {orig}")
    except Exception as e:
        print(f"[Warning] Khong tra lai duoc may in mac dinh '{orig}': {e}")


def _is_process_running(image_name: str) -> bool:
    """True if a process with this image name (e.g. WINWORD.EXE) is currently running."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
            capture_output=True, text=True, timeout=10
        ).stdout
        return image_name.lower() in out.lower()
    except Exception as e:
        print(f"[Warning] Cannot query running processes ({image_name}): {e}")
        return True  # unknown -> assume running, so we never touch the user's session


def _force_kill_office_enabled() -> bool:
    """Force-killing Office is opt-in only: set QUATEST3_FORCE_KILL_OFFICE=1 to enable."""
    return os.environ.get("QUATEST3_FORCE_KILL_OFFICE", "").strip() == "1"


def clear_word_resiliency():
    """Clears Office Resiliency (DisabledItems and DocumentRecovery) to prevent modal crash dialogs.

    Only runs when no WINWORD.EXE is running, so recovery data of the user's open documents is never lost.
    """
    if _is_process_running("WINWORD.EXE"):
        print("[PDFExporter] Word is running - skipping Resiliency cleanup to protect open documents.")
        return
    import winreg
    for sub in [
        r"Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems",
        r"Software\Microsoft\Office\16.0\Word\Resiliency\DocumentRecovery"
    ]:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub, 0, winreg.KEY_ALL_ACCESS) as k:
                while True:
                    try:
                        name = winreg.EnumValue(k, 0)[0]
                        winreg.DeleteValue(k, name)
                    except OSError:
                        break
        except FileNotFoundError:
            pass  # key does not exist: nothing to clean
        except Exception as e:
            print(f"[Warning] Could not clear Word Resiliency key {sub}: {e}")


def kill_orphaned_word_processes():
    """Releases stuck Word processes - NEVER force-closes the user's Word by default.

    The exporter starts its own isolated Word instance (DispatchEx), so killing is normally unnecessary.
    Set QUATEST3_FORCE_KILL_OFFICE=1 to restore the old force-kill behaviour (unsaved work may be lost).
    """
    if _force_kill_office_enabled():
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True, timeout=5)
        except Exception as e:
            print(f"[Warning] taskkill WINWORD.EXE failed: {e}")
    elif _is_process_running("WINWORD.EXE"):
        print("[PDFExporter] Word is already open - leaving it untouched "
              "(set QUATEST3_FORCE_KILL_OFFICE=1 only if a stuck Word blocks export).")
    clear_word_resiliency()


def kill_orphaned_excel_processes():
    """Releases stuck Excel processes - NEVER force-closes the user's Excel by default (see Word variant)."""
    if _force_kill_office_enabled():
        try:
            subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"], capture_output=True, timeout=5)
        except Exception as e:
            print(f"[Warning] taskkill EXCEL.EXE failed: {e}")
    elif _is_process_running("EXCEL.EXE"):
        print("[PDFExporter] Excel is already open - leaving it untouched "
              "(set QUATEST3_FORCE_KILL_OFFICE=1 only if a stuck Excel blocks export).")


def repaginate_with_fields(doc) -> None:
    """
    Phân trang rồi cập nhật trường ở header/footer (PAGE, SECTIONPAGES…) và phân trang lại.
    Chỉ Repaginate() thì khi chuỗi keepNext đẩy dòng bảng sang trang sau, Word giữ SECTIONPAGES tính
    ở lượt bố cục trước -> footer lệch ('Trang 3/5' trong khi phụ lục 4 trang; 133266 ngày 02/11/2026).
    """
    doc.Repaginate()
    for sec in doc.Sections:
        for hf in list(sec.Headers) + list(sec.Footers):
            if hf.Exists:
                hf.Range.Fields.Update()
    doc.Repaginate()


def convert_docx_to_pdf(
    docx_path: str,
    pdf_path: Optional[str] = None,
    add_bg: bool = True,
    cert_date: Optional[str] = None
) -> str:
    """
    Converts a single .docx file to .pdf using Microsoft Word COM.
    Returns the path to the created PDF.
    """
    import win32com.client

    abs_docx = os.path.abspath(docx_path)
    if not os.path.isfile(abs_docx):
        raise FileNotFoundError(f"File not found: {abs_docx}")

    if not pdf_path:
        pdf_path = os.path.splitext(abs_docx)[0] + ".pdf"
    abs_pdf = os.path.abspath(pdf_path)

    avoid_label_printer()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(abs_docx, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
        repaginate_with_fields(doc)
        doc.SaveAs2(abs_pdf, FileFormat=17)  # wdFormatPDF = 17
        doc.Close(SaveChanges=0)
    finally:
        doc = None
        word.Quit()
        del word
        gc.collect()

    if add_bg and is_tbkq_file(abs_docx):
        try:
            process_tbkq_pdf(abs_pdf, cert_date=cert_date, add_bg=True, print_log=True)
        except Exception as e:
            print(f"[Warning] UpChungThu bridge error: {e}")

    return abs_pdf


def convert_xlsx_to_pdf(xlsx_path: str, pdf_path: Optional[str] = None) -> str:
    """
    Converts a calculation workbook .xlsx to .pdf using Microsoft Excel COM.
    If draft TBKQ sheets exist, exports the primary calculation sheet ('Bảng tính') to avoid page bloat.
    """
    import win32com.client

    abs_xlsx = os.path.abspath(xlsx_path)
    if not os.path.isfile(abs_xlsx):
        raise FileNotFoundError(f"File not found: {abs_xlsx}")

    if not pdf_path:
        pdf_path = os.path.splitext(abs_xlsx)[0] + ".pdf"
    abs_pdf = os.path.abspath(pdf_path)
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)

    avoid_label_printer()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_xlsx, ReadOnly=True)
        try:
            wb.Application.CalculateFull()
        except Exception:
            pass
        has_tbkq_sheets = any("tbkq" in s.Name.lower() for s in wb.Worksheets)
        bang_tinh_sheet = None
        for s in wb.Worksheets:
            if s.Name.lower() in ("bảng tính", "bang tinh", "tinhtoan", "tính toán", "calc"):
                bang_tinh_sheet = s
                break

        if has_tbkq_sheets and bang_tinh_sheet:
            bang_tinh_sheet.ExportAsFixedFormat(0, abs_pdf)
        else:
            wb.ExportAsFixedFormat(0, abs_pdf)
        wb.Close(SaveChanges=0)
    finally:
        s = None
        bang_tinh_sheet = None
        wb = None
        excel.Quit()
        del excel
        gc.collect()

    return abs_pdf


class PDFExporter:
    """Batch exporter and sync dispatcher for Dossier deliverables (Word DOCX and Excel XLSX)."""

    def __init__(self, dossier: Dossier):
        self.dossier = dossier

    def export_files(self, docx_paths: List[str]) -> List[str]:
        """Converts provided docx files to PDF and synchronizes across all target folders."""
        d = self.dossier
        d.ensure_output_directories()
        created_pdfs = []

        kill_orphaned_word_processes()
        time.sleep(0.5)

        import win32com.client
        avoid_label_printer()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            for fpath in docx_paths:
                abs_docx = os.path.abspath(fpath)
                if not os.path.isfile(abs_docx):
                    continue

                fname = os.path.basename(abs_docx)
                pdf_name = os.path.splitext(fname)[0] + ".pdf"
                pdf_target = os.path.join(d.pdf_dir, pdf_name)
                pdf_out = os.path.join(d.output_dir, pdf_name)

                print(f"[PDFExporter] Converting DOCX {fname} -> PDF in /pdf/ folder...")
                doc = None
                try:
                    doc = word.Documents.Open(abs_docx, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
                    repaginate_with_fields(doc)
                    pages = doc.ComputeStatistics(2)
                    doc.SaveAs2(os.path.abspath(pdf_target), FileFormat=17)
                    print(f"  [OK] {pdf_name} (Pages: {pages}) -> {pdf_target}")

                    # UpChungThu Integration: Insert Background & Run QA for TBKQ deliverables
                    if is_tbkq_file(fname):
                        try:
                            print(f"[UpChungThu] Processing TBKQ with background & QA checks: {pdf_name}...")
                            process_tbkq_pdf(
                                pdf_target,
                                cert_date=d.cert_date,
                                add_bg=True,
                                print_log=True
                            )
                        except Exception as uct_err:
                            print(f"  [Warning] UpChungThu processing error for {pdf_name}: {uct_err}")

                    # Sync copies to Z network drive (Output folder no longer used per user rule)
                    if d.output_dir:
                        try:
                            shutil.copy2(pdf_target, pdf_out)
                        except Exception as copy_err:
                            print(f"  [Warning] Could not copy {pdf_name} to output dir: {copy_err}")
                    if d.network_dir and os.path.isdir(d.network_dir):
                        try:
                            net_pdf = os.path.join(d.network_dir, "pdf")
                            os.makedirs(net_pdf, exist_ok=True)
                            shutil.copy2(pdf_target, os.path.join(net_pdf, pdf_name))
                        except Exception as copy_err:
                            print(f"  [Warning] Could not sync {pdf_name} to network drive: {copy_err}")

                    created_pdfs.append(pdf_target)
                except Exception as save_err:
                    print(f"  [Warning] Could not export {pdf_name} (may be open in another application): {save_err}")
                finally:
                    if doc is not None:
                        try:
                            doc.Close(SaveChanges=0)
                        except Exception:
                            pass
                    doc = None
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            del word
            gc.collect()

        return created_pdfs

    def export_excel_files(self, xlsx_paths: List[str]) -> List[str]:
        """Converts provided xlsx calculation files to PDF and synchronizes to /pdf/, Output/, and Z:."""
        d = self.dossier
        d.ensure_output_directories()
        created_pdfs = []
        if not xlsx_paths:
            return created_pdfs

        kill_orphaned_excel_processes()
        time.sleep(0.5)

        import win32com.client
        avoid_label_printer()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            for fpath in xlsx_paths:
                abs_xlsx = os.path.abspath(fpath)
                if not os.path.isfile(abs_xlsx):
                    continue

                fname = os.path.basename(abs_xlsx)
                pdf_name = os.path.splitext(fname)[0] + ".pdf"
                pdf_target = os.path.join(d.pdf_dir, pdf_name)
                pdf_out = os.path.join(d.output_dir, pdf_name)

                print(f"[PDFExporter] Converting Excel {fname} -> PDF in /pdf/ folder...")
                wb = None
                try:
                    wb = excel.Workbooks.Open(abs_xlsx, ReadOnly=True)
                    try:
                        wb.Application.CalculateFull()
                    except Exception:
                        pass
                    has_tbkq_sheets = any("tbkq" in s.Name.lower() for s in wb.Worksheets)
                    bang_tinh_sheet = None
                    for s in wb.Worksheets:
                        if s.Name.lower() in ("bảng tính", "bang tinh", "tinhtoan", "tính toán", "calc"):
                            bang_tinh_sheet = s
                            break

                    if has_tbkq_sheets and bang_tinh_sheet:
                        bang_tinh_sheet.ExportAsFixedFormat(0, os.path.abspath(pdf_target))
                    else:
                        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_target))

                    print(f"  [OK] Excel PDF: {pdf_name} -> {pdf_target}")
                    if d.output_dir:
                        try:
                            shutil.copy2(pdf_target, pdf_out)
                        except Exception as copy_err:
                            print(f"  [Warning] Could not copy {pdf_name} to output dir: {copy_err}")
                    if d.network_dir and os.path.isdir(d.network_dir):
                        try:
                            net_pdf = os.path.join(d.network_dir, "pdf")
                            os.makedirs(net_pdf, exist_ok=True)
                            shutil.copy2(pdf_target, os.path.join(net_pdf, pdf_name))
                        except Exception as copy_err:
                            print(f"  [Warning] Could not sync {pdf_name} to network drive: {copy_err}")

                    created_pdfs.append(pdf_target)
                except Exception as save_err:
                    print(f"  [Warning] Could not export Excel {pdf_name}: {save_err}")
                finally:
                    if wb is not None:
                        try:
                            wb.Close(SaveChanges=0)
                        except Exception:
                            pass
                    s = None
                    bang_tinh_sheet = None
                    wb = None
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
            del excel
            gc.collect()

        return created_pdfs

    def apply_background_to_dossier_pdfs(self, force_bg: bool = False) -> List[Dict[str, Any]]:
        """
        Processes all TBKQ PDFs in dossier's pdf/ folder:
        Applies UpChungThu background and runs QA checks, then syncs to network drive.
        """
        d = self.dossier
        results = []
        if not os.path.isdir(d.pdf_dir):
            return results

        for fname in sorted(os.listdir(d.pdf_dir)):
            if is_tbkq_file(fname) and fname.lower().endswith(".pdf"):
                pdf_target = os.path.join(d.pdf_dir, fname)
                try:
                    res = process_tbkq_pdf(
                        pdf_target,
                        cert_date=d.cert_date,
                        add_bg=True,
                        force_bg=force_bg,
                        print_log=True
                    )
                    results.append(res)
                    if d.network_dir and os.path.isdir(d.network_dir):
                        try:
                            net_pdf = os.path.join(d.network_dir, "pdf")
                            os.makedirs(net_pdf, exist_ok=True)
                            shutil.copy2(pdf_target, os.path.join(net_pdf, fname))
                        except Exception as copy_err:
                            print(f"  [Warning] Could not sync {fname} to network drive: {copy_err}")
                except Exception as e:
                    print(f"  [Error] Failed to process {fname}: {e}")

        return results
