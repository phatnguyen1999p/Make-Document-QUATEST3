# -*- coding: utf-8 -*-
"""
Core/Builders/handlers/mtld.py
==============================
Handler & Builder cho hồ sơ Quan trắc môi trường lao động (QTMTLĐ):
1. Sinh nội dung mục "KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC" (5 nhóm giải pháp: Kỹ thuật,
   Tổ chức lao động, Y tế và sức khỏe, Phương tiện bảo vệ cá nhân / BHLĐ, Các giải pháp khác)
   bằng AI (LLM API) hoặc bộ sinh ngữ cảnh biến thiên (Contextual Seeded Composer)
   với giọng văn nhẹ nhàng, mang tính ghi nhận nỗ lực của đơn vị và khuyến nghị hướng
   duy trì / khắc phục hoàn thiện thêm, đảm bảo các hồ sơ không trùng lặp rập khuôn.
2. Đồng bộ bảng Tổng hợp kết quả quan trắc (Tổng số mẫu, Số mẫu đạt, Số mẫu không đạt
   định dạng 2 chữ số `01`, `02`, `00`) giữa các bảng chi tiết của TBKQ và bảng tổng hợp của BCKQDG.
3. Dựng báo cáo TBKQ (`FORM - TBKQ - viet - mtld - 2026.docx`) và BCKQ nội bộ
   (`FORM - BCKQ - mtld - 2026.docx`, mẫu `BM3030.03 HL:01/8/2025`).
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from Core.Tools.mtld_calc import (
    ROOT_DIR,
    TEMPLATE_MTLD_DIR,
    build_complete_chemical_db,
    lookup_chemical_limit,
    populate_hoi_khi_doc_xlsx,
)


# ==============================================================================
# 1. SINH KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC (AI CALL + CONTEXTUAL VARIATION ENGINE)
# ==============================================================================

RECOMMENDATION_SYSTEM_PROMPT = """Bạn là chuyên gia tư vấn vệ sinh lao động và quan trắc môi trường lao động tại Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3).
Nhiệm vụ của bạn là viết mục "KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC" gồm 5 nhóm giải pháp cho Báo cáo kết quả quan trắc môi trường lao động của khách hàng.

YÊU CẦU BẮT BUỘC VỀ GIỌNG VĂN VÀ NỘI DUNG:
1. Giọng văn nhẹ nhàng, lịch sự, mang tính ghi nhận sự quan tâm và điều kiện kiểm soát hiện tại của doanh nghiệp, đồng thời gợi mở hướng duy trì và từng bước khắc phục, cải thiện ("mang tính ghi nhận và sẽ khắc phục"). Tuyệt đối không dùng từ ngữ nặng nề, phê phán hay áp đặt.
2. Diễn đạt tự nhiên, linh hoạt theo đặc thù ngành nghề, vị trí đo và các yếu tố thực tế của từng hồ sơ để mỗi hồ sơ có sắc thái câu chữ riêng, không rập khuôn giống hệt nhau.
3. Nếu tất cả các chỉ tiêu đều Đạt: ghi nhận môi trường làm việc tại thời điểm quan trắc nhìn chung đáp ứng tốt quy chuẩn hiện hành, đồng thời khuyến nghị duy trì bảo dưỡng định kỳ để giữ vững điều kiện làm việc thuận lợi.
4. Nếu có vị trí chưa đạt (ví dụ: Nhiệt độ, Độ ẩm, Tốc độ gió, Ánh sáng, Tiếng ồn, Bụi, Hơi khí độc, Tư thế lao động Ecgonomy...): ghi nhận doanh nghiệp đã có trang bị thông gió/chiếu sáng/bảo hộ bước đầu, nêu nhẹ nhàng các vị trí cần lưu ý thêm và đề xuất giải pháp cải thiện khả thi.
5. Trả về ĐÚNG định dạng JSON gồm 5 khóa:
   - "ky_thuat": chuỗi gồm 2-4 gạch đầu dòng ("- ...") cho mục Kỹ thuật
   - "to_chuc_lao_dong": chuỗi gồm 2-3 gạch đầu dòng ("- ...") cho mục Tổ chức lao động
   - "y_te_va_suc_khoe": chuỗi gồm 2-3 gạch đầu dòng ("- ...") cho mục Y tế và sức khỏe
   - "bao_ho_lao_dong": chuỗi gồm 2-3 gạch đầu dòng ("- ...") cho mục Phương tiện bảo vệ cá nhân (BHLĐ)
   - "giai_phap_khac": chuỗi gồm 1-2 gạch đầu dòng ("- ...") cho mục Các giải pháp khác
Chỉ trả về JSON thuần, không bọc markdown nếu không cần thiết."""


def _call_llm_for_recommendations(
    user_prompt: str,
    timeout: int = 20,
) -> Optional[Dict[str, str]]:
    """
    Thử gọi LLM API (Gemini / OpenAI / Anthropic) nếu trong môi trường có cấu hình API key.
    Nếu không có key hoặc lỗi mạng, trả về None để chuyển sang bộ sinh ngữ cảnh biến thiên.
    """
    gemini_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    openai_key = os.environ.get("OPENAI_API_KEY")
    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")

    raw_text = None
    try:
        if gemini_key:
            url = (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                f"gemini-2.5-flash:generateContent?key={gemini_key}"
            )
            payload = {
                "systemInstruction": {"parts": [{"text": RECOMMENDATION_SYSTEM_PROMPT}]},
                "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
                "generationConfig": {
                    "temperature": 0.75,
                    "responseMimeType": "application/json",
                },
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                raw_text = data["candidates"][0]["content"]["parts"][0]["text"]

        elif openai_key:
            url = "https://api.openai.com/v1/chat/completions"
            payload = {
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": RECOMMENDATION_SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.75,
                "response_format": {"type": "json_object"},
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {openai_key}",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                raw_text = data["choices"][0]["message"]["content"]

        elif anthropic_key:
            url = "https://api.anthropic.com/v1/messages"
            payload = {
                "model": "claude-3-5-haiku-latest",
                "max_tokens": 1024,
                "temperature": 0.75,
                "system": RECOMMENDATION_SYSTEM_PROMPT,
                "messages": [{"role": "user", "content": user_prompt}],
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "x-api-key": anthropic_key,
                    "anthropic-version": "2023-06-01",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                raw_text = data["content"][0]["text"]
    except Exception:
        return None

    if not raw_text:
        return None

    try:
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_text.strip(), flags=re.MULTILINE)
        parsed = json.loads(cleaned)
        required_keys = (
            "ky_thuat",
            "to_chuc_lao_dong",
            "y_te_va_suc_khoe",
            "bao_ho_lao_dong",
            "giai_phap_khac",
        )
        if all(k in parsed and isinstance(parsed[k], str) and parsed[k].strip() for k in required_keys):
            return {k: parsed[k].strip() for k in required_keys}
    except Exception:
        return None
    return None


def _pick_variant(seed_str: str, salt: str, options: List[str]) -> str:
    """Chọn một phương án câu văn ổn định nhưng khác biệt giữa các mã hồ sơ / tên công ty."""
    h = int(hashlib.sha256(f"{seed_str}::{salt}".encode("utf-8")).hexdigest(), 16)
    return options[h % len(options)]


def _compose_contextual_recommendations(
    company_name: str,
    industry_desc: str,
    factor_summary: List[Dict[str, Any]],
    failing_details: Optional[Dict[str, List[str]]] = None,
    dossier_code: str = "",
    variation_seed: str = "",
) -> Dict[str, str]:
    """
    Bộ sinh khuyến nghị biến thiên theo ngữ cảnh hồ sơ (Contextual Seeded Composer):
    - Đảm bảo giọng văn nhẹ nhàng, mang tính ghi nhận hiện trạng và định hướng khắc phục.
    - Mỗi hồ sơ (khác mã hồ sơ, tên đơn vị, ngành nghề, tổ hợp yếu tố đạt/chưa đạt)
      sẽ kết hợp các cấu trúc câu khác nhau, không bị trùng lặp cứng nhắc.
    """
    failing_details = failing_details or {}
    seed_key = f"{dossier_code}|{company_name}|{industry_desc}|{variation_seed}"

    # Phân loại các yếu tố có mẫu chưa đạt và các yếu tố đạt 100%
    failed_factors: List[str] = []
    monitored_factors: List[str] = []
    for item in factor_summary:
        f_name = str(item.get("name", "")).strip()
        monitored_factors.append(f_name)
        fail_cnt = int(item.get("fail", 0) or 0)
        if fail_cnt > 0:
            failed_factors.append(f_name)

    for f_name, locs in failing_details.items():
        if locs and f_name not in failed_factors:
            failed_factors.append(f_name)

    has_temp_fail = any("nhiệt độ" in f.lower() for f in failed_factors)
    has_hum_fail = any("độ ẩm" in f.lower() for f in failed_factors)
    has_wind_fail = any("gió" in f.lower() for f in failed_factors)
    has_light_fail = any("ánh sáng" in f.lower() or "chiếu sáng" in f.lower() for f in failed_factors)
    has_noise_fail = any("tiếng ồn" in f.lower() or "ồn" in f.lower() for f in failed_factors)
    has_dust_fail = any("bụi" in f.lower() for f in failed_factors)
    has_chem_fail = any(
        any(k in f.lower() for k in ("hơi", "khí", "toluen", "xylen", "aceton", "axit", "dung môi", "hóa chất"))
        for f in failed_factors
    )
    has_ergo_note = any("tư thế" in f.lower() or "ecgonomy" in f.lower() for f in monitored_factors)

    def _loc_phrase(keyword: str) -> str:
        for k, locs in failing_details.items():
            if keyword in k.lower() and locs:
                sample_locs = ", ".join(locs[:3])
                more = f" và {len(locs) - 3} vị trí khác" if len(locs) > 3 else ""
                return f" (như tại {sample_locs}{more})"
        return ""

    # --- 1. KỸ THUẬT ---
    kt_lines: List[str] = []
    if not failed_factors:
        kt_lines.append(
            _pick_variant(
                seed_key,
                "kt_all_pass_1",
                [
                    "- Ghi nhận hệ thống thông gió, chiếu sáng và kiểm soát môi trường tại các khu vực làm việc hiện đang vận hành ổn định, đáp ứng tốt các quy chuẩn hiện hành; đơn vị tiếp tục duy trì lịch vệ sinh, bảo dưỡng định kỳ.",
                    "- Qua kết quả quan trắc, điều kiện kỹ thuật tại nhà xưởng nhìn chung được kiểm soát tốt; khuyến nghị đơn vị tiếp tục duy trì chế độ kiểm tra định kỳ hệ thống quạt thông gió, điều hòa và đèn chiếu sáng để giữ vững chất lượng môi trường.",
                    "- Hiện trạng thông gió, chiếu sáng và thiết bị kỹ thuật tại cơ sở đang phát huy hiệu quả tích cực; đơn vị duy trì công tác bảo trì, vệ sinh lưới lọc gió và máng đèn thường xuyên nhằm bảo đảm độ ổn định lâu dài.",
                    "- Ghi nhận sự đầu tư đồng bộ về hạ tầng thông gió và chiếu sáng tại các vị trí làm việc; đề nghị đơn vị tiếp tục duy trì quy trình bảo dưỡng phòng ngừa đối với hệ thống kỹ thuật hiện hữu.",
                ],
            )
        )
        kt_lines.append(
            _pick_variant(
                seed_key,
                "kt_all_pass_2",
                [
                    "- Tiếp tục theo dõi định kỳ độ ồn và độ rung của máy móc thiết bị trong quá trình vận hành để kịp thời tra dầu mỡ, căn chỉnh khi có dấu hiệu hao mòn cơ khí.",
                    "- Duy trì kiểm tra định kỳ các chụp hút cục bộ và hệ thống dẫn khí tại khu vực sản xuất để bảo đảm lưu lượng trao đổi không khí luôn thông thoáng.",
                    "- Kết hợp vệ sinh công nghiệp thường xuyên tại các khu vực thao tác nhằm hạn chế bụi tích tụ và giữ không gian làm việc luôn sạch sẽ, thoáng mát.",
                ],
            )
        )
    else:
        kt_lines.append(
            _pick_variant(
                seed_key,
                "kt_ack_intro",
                [
                    "- Ghi nhận đơn vị đã trang bị hệ thống thông gió, chiếu sáng và các thiết bị kiểm soát môi trường cơ bản tại khu vực sản xuất; bên cạnh các vị trí đạt chuẩn, một số khu vực đặc thù sẽ được rà soát để hoàn thiện thêm.",
                    "- Nhìn chung các khu vực làm việc đã được doanh nghiệp quan tâm bố trí hệ thống kỹ thuật hỗ trợ; để điều kiện vi khí hậu và môi trường đồng đều hơn giữa các vị trí, khuyến nghị từng bước cải thiện thêm một số hạng mục sau.",
                    "- Qua đợt quan trắc, phần lớn các vị trí được duy trì điều kiện kỹ thuật tương đối ổn định; đối với một vài điểm cục bộ còn chịu ảnh hưởng từ đặc thù thiết bị hoặc thời tiết, đơn vị có thể kết hợp khắc phục từng bước.",
                ],
            )
        )
        if has_temp_fail or has_hum_fail or has_wind_fail:
            loc_str = _loc_phrase("nhiệt") or _loc_phrase("ẩm") or _loc_phrase("gió")
            kt_lines.append(
                _pick_variant(
                    seed_key,
                    "kt_vkh",
                    [
                        f"- Đối với yếu tố vi khí hậu{loc_str}: đơn vị xem xét điều chỉnh lại hướng quạt làm mát, tăng cường lưu thông không khí tự nhiên hoặc bổ sung quạt thông gió cục bộ để hỗ trợ tản nhiệt và cân bằng độ ẩm hiệu quả hơn.",
                        f"- Tại một số điểm có nhiệt độ/độ ẩm hoặc tốc độ gió chưa thật sự tối ưu{loc_str}: khuyến nghị vệ sinh định kỳ hệ thống quạt hút, điều chỉnh lưu lượng gió hợp lý và hạn chế luồng nhiệt bức xạ trực tiếp từ máy móc sang vị trí đứng thao tác.",
                        f"- Về điều kiện vi khí hậu{loc_str}: ghi nhận đặc thù khu vực có phát sinh nhiệt/ẩm cục bộ; đơn vị sẽ từng bước rà soát, bổ sung quạt đối lưu hoặc tối ưu chế độ vận hành hệ thống làm mát vào các khung giờ nắng nóng cao điểm.",
                    ],
                )
            )
        if has_light_fail:
            loc_str = _loc_phrase("sáng")
            kt_lines.append(
                _pick_variant(
                    seed_key,
                    "kt_light",
                    [
                        f"- Đối với độ rọi chiếu sáng{loc_str}: khuyến nghị vệ sinh chụp đèn, thay thế các bóng đèn đã suy giảm quang thông và bố trí thêm đèn chiếu sáng cục bộ tại bàn thao tác cần độ chính xác thị giác cao.",
                        f"- Về chiếu sáng tại vị trí làm việc{loc_str}: đơn vị ghi nhận và sẽ tiến hành rà soát, lau chùi bụi bẩn trên máng đèn, đồng thời điều chỉnh góc chiếu hoặc bổ sung bóng đèn LED tại khu vực bị khuất bóng.",
                        f"- Nhằm nâng cao tiện nghi thị giác tại vị trí có độ sáng còn thấp{loc_str}: đề xuất kiểm tra định kỳ hệ thống đèn hiện hữu, kết hợp tận dụng ánh sáng tự nhiên và bổ sung nguồn sáng tại vùng thao tác chi tiết.",
                    ],
                )
            )
        if has_noise_fail:
            loc_str = _loc_phrase("ồn")
            kt_lines.append(
                _pick_variant(
                    seed_key,
                    "kt_noise",
                    [
                        f"- Đối với các vị trí có mức tiếng ồn cao do đặc thù cơ khí/máy nén{loc_str}: khuyến nghị duy trì bảo dưỡng, tra dầu bôi trơn, siết chặt các chi tiết lắp ghép và kiểm tra đệm giảm chấn, vỏ che chắn âm của thiết bị.",
                        f"- Về yếu tố tiếng ồn tại khu vực vận hành máy{loc_str}: ghi nhận tiếng ồn phát sinh chủ yếu từ chu trình hoạt động liên tục của thiết bị; đơn vị sẽ chú trọng bảo trì định kỳ các bộ phận chuyển động và bố trí che chắn cách âm phù hợp.",
                        f"- Để giảm thiểu tiếng ồn lan truyền từ cụm máy công suất lớn{loc_str}: đơn vị tiếp tục duy trì lịch kiểm tra kỹ thuật, thay thế chi tiết mài mòn và cân nhắc bổ sung tấm tiêu âm/đệm cao su chống rung tại chân máy.",
                    ],
                )
            )
        if has_dust_fail or has_chem_fail:
            loc_str = _loc_phrase("bụi") or _loc_phrase("hơi")
            kt_lines.append(
                _pick_variant(
                    seed_key,
                    "kt_dust_chem",
                    [
                        f"- Tại khu vực có phát sinh bụi hoặc hơi hóa chất{loc_str}: duy trì hiệu quả hệ thống chụp hút cục bộ, định kỳ vệ sinh túi lọc bụi/than hoạt tính và đậy kín nắp thùng chứa dung môi khi không sử dụng.",
                        f"- Đối với việc kiểm soát bụi và hơi khí tại nguồn phát sinh{loc_str}: khuyến nghị kiểm tra lưu lượng hệ thống hút khí thải cục bộ, kết hợp vệ sinh bề mặt theo phương pháp ẩm để hạn chế phát tán ra không gian xung quanh.",
                    ],
                )
            )

    # --- 2. TỔ CHỨC LAO ĐỘNG ---
    tc_lines: List[str] = [
        _pick_variant(
            seed_key,
            "tc_main",
            [
                "- Tiếp tục duy trì việc bố trí ca kíp và thời gian làm việc - nghỉ ngơi hợp lý; tạo điều kiện để người lao động có khoảng nghỉ ngắn thư giãn mắt và cơ xương khớp giữa ca làm việc.",
                "- Ghi nhận đơn vị đã sắp xếp quy trình làm việc tương đối khoa học; khuyến nghị duy trì chế độ giải lao giữa giờ phù hợp, đặc biệt tại các khâu đòi hỏi sự tập trung cao hoặc đứng/ngồi thao tác liên tục.",
                "- Duy trì tổ chức phân công lao động hợp lý giữa các bộ phận, kết hợp luân phiên công việc hoặc bố trí thời gian nghỉ chuyển tiếp tại khu vực thoáng mát đối với các vị trí vận hành máy liên tục.",
                "- Sắp xếp mặt bằng thao tác gọn gàng, lối đi thông thoáng; duy trì lịch vệ sinh 5S tại nơi làm việc và nhắc nhở tắt các thiết bị phát nhiệt/phát ồn khi không có nhu cầu vận hành.",
            ],
        )
    ]
    if has_ergo_note:
        tc_lines.append(
            _pick_variant(
                seed_key,
                "tc_ergo",
                [
                    "- Đối với các vị trí làm việc có tư thế đứng hoặc cúi/với thao tác lặp lại: khuyến nghị hướng dẫn người lao động điều chỉnh chiều cao mặt bàn/ghế phù hợp tầm vóc và kết hợp vận động nhẹ giữa giờ.",
                    "- Quan tâm hỗ trợ công cụ nâng đỡ, xe đẩy tay tại các khâu tiếp liệu và khuyến khích người lao động thay đổi tư thế linh hoạt để giảm mỏi cơ cục bộ.",
                ],
            )
        )
    else:
        tc_lines.append(
            _pick_variant(
                seed_key,
                "tc_sub",
                [
                    "- Tăng cường phổ biến, hướng dẫn người lao động giữ gìn vệ sinh chung tại vị trí sản xuất và chủ động phản hồi sớm khi thiết bị thông gió, chiếu sáng có biểu hiện cần bảo trì.",
                    "- Duy trì biển chỉ dẫn an toàn - vệ sinh lao động trực quan tại các khu vực vận hành máy móc, kho hóa chất và khu vực kỹ thuật.",
                ],
            )
        )

    # --- 3. Y TẾ VÀ SỨC KHỎE ---
    yt_lines: List[str] = [
        _pick_variant(
            seed_key,
            "yt_1",
            [
                "- Tiếp tục duy trì chương trình khám sức khỏe định kỳ và khám phát hiện bệnh nghề nghiệp hàng năm cho người lao động theo đúng quy định hiện hành để theo dõi sức khỏe lâu dài.",
                "- Ghi nhận sự quan tâm của đơn vị trong công tác chăm sóc sức khỏe nhân sự; khuyến nghị tiếp tục tổ chức khám sức khỏe định kỳ và lưu trữ hồ sơ theo dõi sức khỏe cho từng bộ phận.",
                "- Duy trì tổ chức khám sức khỏe định kỳ hàng năm, đồng thời chú trọng tư vấn sức khỏe và khám chuyên khoa đối với người lao động làm việc tại các bộ phận sản xuất đặc thù.",
            ],
        ),
        _pick_variant(
            seed_key,
            "yt_2",
            [
                "- Đảm bảo cung cấp đầy đủ nước uống sạch, mát tại các khu vực làm việc và duy trì tủ thuốc sơ cấp cứu với đầy đủ vật tư y tế thiết yếu tại nhà xưởng.",
                "- Duy trì bố trí điểm cấp nước uống hợp vệ sinh gần khu vực sản xuất để người lao động kịp thời bù nước, đặc biệt trong những ngày thời tiết nắng nóng.",
                "- Thường xuyên kiểm tra, bổ sung danh mục thuốc và dụng cụ tại tủ thuốc sơ cấp cứu của các bộ phận để sẵn sàng hỗ trợ chăm sóc sức khỏe ban đầu.",
            ],
        ),
    ]

    # --- 4. PHƯƠNG TIỆN BẢO VỆ CÁ NHÂN (BHLĐ) ---
    bh_lines: List[str] = [
        _pick_variant(
            seed_key,
            "bh_1",
            [
                "- Ghi nhận đơn vị đã trang bị phương tiện bảo vệ cá nhân cho người lao động; khuyến nghị tiếp tục duy trì cấp phát định kỳ và nhắc nhở nhân viên sử dụng đúng quy cách trong suốt ca làm việc.",
                "- Tiếp tục duy trì việc trang bị đầy đủ trang phục, giày bảo hộ, khẩu trang và các phương tiện bảo vệ cá nhân phù hợp với từng vị trí công việc tại nhà máy.",
                "- Phát huy nền tảng tuân thủ an toàn hiện có, đơn vị tiếp tục rà soát bổ sung kịp thời các trang thiết bị bảo hộ lao động khi bị cũ, hỏng để đảm bảo hiệu quả bảo vệ tối ưu.",
            ],
        )
    ]
    if has_noise_fail:
        bh_lines.append(
            _pick_variant(
                seed_key,
                "bh_noise",
                [
                    "- Tại các khu vực có mức ồn cao, đơn vị tiếp tục trang bị nút tai hoặc chụp tai chống ồn có độ giảm âm phù hợp và hướng dẫn người lao động đeo đúng cách khi vận hành thiết bị.",
                    "- Đặc biệt lưu ý nhắc nhở người lao động sử dụng thường xuyên nút tai/chụp tai chống ồn khi làm việc gần các cụm máy có độ ồn lớn.",
                ],
            )
        )
    elif has_dust_fail or has_chem_fail:
        bh_lines.append(
            _pick_variant(
                seed_key,
                "bh_dust_chem",
                [
                    "- Chú trọng trang bị khẩu trang lọc bụi/lọc hơi hữu cơ, găng tay và kính bảo hộ chuyên dụng cho nhân viên tiếp xúc trực tiếp với hóa chất hoặc bụi sản xuất.",
                    "- Hướng dẫn người lao động kiểm tra độ kín khít của khẩu trang và bảo quản phương tiện bảo vệ cá nhân tại nơi khô ráo, sạch sẽ sau mỗi ca làm việc.",
                ],
            )
        )
    else:
        bh_lines.append(
            _pick_variant(
                seed_key,
                "bh_general",
                [
                    "- Thường xuyên hướng dẫn, khích lệ người lao động tự giác bảo quản và sử dụng đúng phương tiện bảo vệ cá nhân tương ứng với từng công đoạn sản xuất.",
                    "- Kết hợp kiểm tra định kỳ tình trạng sử dụng trang bị bảo hộ tại hiện trường nhằm tạo thói quen làm việc an toàn, chuyên nghiệp.",
                ],
            )
        )

    # --- 5. CÁC GIẢI PHÁP KHÁC ---
    khac_lines: List[str] = [
        _pick_variant(
            seed_key,
            "khac_1",
            [
                "- Tiếp tục duy trì thực hiện quan trắc môi trường lao động định kỳ hàng năm theo quy định tại Nghị định 44/2016/NĐ-CP để kịp thời theo dõi diễn biến môi trường và chủ động cải thiện điều kiện làm việc.",
                "- Duy trì công tác quan trắc môi trường lao động định kỳ hằng năm kết hợp huấn luyện an toàn, vệ sinh lao động nhằm nâng cao nhận thức tự bảo vệ sức khỏe cho toàn thể cán bộ công nhân viên.",
                "- Định kỳ hàng năm tổ chức quan trắc, đánh giá lại các yếu tố môi trường lao động và cập nhật hồ sơ vệ sinh môi trường lao động để không ngừng hoàn thiện môi trường làm việc an toàn, thân thiện.",
                "- Kết hợp duy trì quan trắc môi trường lao động định kỳ hàng năm với phong trào xanh - sạch - đẹp, góp phần xây dựng môi trường làm việc ngày càng tiện nghi và bền vững.",
            ],
        )
    ]

    return {
        "ky_thuat": "\n".join(kt_lines),
        "to_chuc_lao_dong": "\n".join(tc_lines),
        "y_te_va_suc_khoe": "\n".join(yt_lines),
        "bao_ho_lao_dong": "\n".join(bh_lines),
        "giai_phap_khac": "\n".join(khac_lines),
    }


def generate_mtld_recommendations(
    company_name: str,
    industry_desc: str = "",
    factor_summary: Optional[List[Dict[str, Any]]] = None,
    failing_details: Optional[Dict[str, List[str]]] = None,
    dossier_code: str = "",
    use_ai: bool = True,
    variation_seed: str = "",
) -> Dict[str, str]:
    """
    Hàm chính sinh mục "KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC" cho hồ sơ QTMTLĐ.
    - Nếu `use_ai=True` và có API key (GEMINI_API_KEY / GOOGLE_API_KEY / OPENAI_API_KEY / ANTHROPIC_API_KEY),
      gọi trực tiếp LLM để viết khuyến nghị tự nhiên, nhẹ nhàng, mang tính ghi nhận và sẽ khắc phục.
    - Nếu chạy offline hoặc không có API key, tự động dùng Contextual Seeded Composer
      để sinh nội dung biến thiên riêng cho từng hồ sơ (không trùng lặp giữa các hồ sơ).
    """
    factor_summary = factor_summary or []
    failing_details = failing_details or {}

    if use_ai:
        summary_lines = [
            f"- {f.get('name')}: Tổng {f.get('total', 0)} mẫu, Đạt {f.get('pass', 0)} mẫu, Không đạt {f.get('fail', 0)} mẫu"
            for f in factor_summary
        ]
        fail_lines = [
            f"- {k}: {', '.join(v)}" for k, v in failing_details.items() if v
        ]
        user_prompt = (
            f"Mã hồ sơ: {dossier_code}\n"
            f"Tên đơn vị: {company_name}\n"
            f"Ngành nghề / Quy trình sản xuất: {industry_desc or 'Sản xuất công nghiệp / Văn phòng'}\n"
            f"Tóm tắt kết quả quan trắc:\n" + ("\n".join(summary_lines) or "- Các yếu tố cơ bản đạt yêu cầu") + "\n"
            f"Chi tiết các vị trí chưa đạt cần lưu ý:\n" + ("\n".join(fail_lines) or "- Không có vị trí vượt ngưỡng")
        )
        ai_res = _call_llm_for_recommendations(user_prompt)
        if ai_res is not None:
            return ai_res

    return _compose_contextual_recommendations(
        company_name=company_name,
        industry_desc=industry_desc,
        factor_summary=factor_summary,
        failing_details=failing_details,
        dossier_code=dossier_code,
        variation_seed=variation_seed,
    )


# ==============================================================================
# 2. CẬP NHẬT BẢNG KHUYẾN NGHỊ & ĐỒNG BỘ BẢNG TỔNG HỢP TRONG WORD (TBKQ & BCKQ)
# ==============================================================================

def _set_cell_text_preserve_font(
    cell: docx.table._Cell,
    text: str,
    font_name: str = "Times New Roman",
    font_size_pt: float = 13.0,
    bold: bool = False,
    italic: bool = False,
    align: Optional[WD_ALIGN_PARAGRAPH] = None,
) -> None:
    """Ghi nội dung vào ô bảng Word với font Times New Roman chuẩn."""
    lines = [ln for ln in text.split("\n") if ln.strip() != ""] or [""]
    # Giữ đoạn đầu tiên, xóa các đoạn thừa
    while len(cell.paragraphs) > 1:
        p_del = cell.paragraphs[-1]._element
        p_del.getparent().remove(p_del)

    for idx, ln in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        for r in p.runs:
            r.text = ""
        run = p.runs[0] if p.runs else p.add_run()
        run.text = ln
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = bold
        run.font.italic = italic
        if align is not None:
            p.alignment = align


def apply_mtld_recommendations_to_doc(
    doc: docx.Document,
    recommendations: Dict[str, str],
) -> bool:
    """
    Tìm bảng 'KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC' trong tài liệu TBKQ (hỗ trợ cả 3 kiểu bảng:
    1 cột dọc gộp ô như Q326D033771, 2 cột dọc như 1126N2.25DG, và 5 cột ngang như 891N2.25DG)
    và điền nội dung khuyến nghị vào đúng từng mục với font Times New Roman 13pt,
    giữ nguyên định dạng gạch đầu dòng (`–` / ListParagraph) của biểu mẫu gốc.
    """
    from copy import deepcopy
    from docx.oxml import OxmlElement

    key_map = [
        (("kỹ thuật",), "ky_thuat"),
        (("tổ chức lao động",), "to_chuc_lao_dong"),
        (("y tế", "sức khỏe"), "y_te_va_suc_khoe"),
        (("bảo vệ cá nhân", "bảo hộ lao động", "bhlđ"), "bao_ho_lao_dong"),
        (("giải pháp khác", "khác"), "giai_phap_khac"),
    ]

    updated = False
    for tbl in doc.tables:
        if len(tbl.rows) == 0:
            continue
        first_cell_txt = tbl.rows[0].cells[0].text.strip().lower()

        # Kiểu 1: Bảng 1 cột (hoặc các cột đã merge ngang thành 1 ô mỗi dòng, như FORM - TBKQ - viet - mtld - 2026.docx)
        if "khuyến nghị" in first_cell_txt and len(tbl.rows) >= 6:
            is_single_box_rows = (
                len(tbl.columns) == 1
                or tbl.rows[1].cells[0]._tc is tbl.rows[1].cells[-1]._tc
            )
            if is_single_box_rows:
                from docx.oxml.ns import qn

                # Đặt lề trên/dưới ô (tblCellMar top/bottom) gọn gàng và chống cắt đôi từng mục (cantSplit)
                tblPr = tbl._tbl.tblPr
                if tblPr is not None:
                    old_mar = tblPr.find(qn("w:tblCellMar"))
                    if old_mar is not None:
                        tblPr.remove(old_mar)
                    new_mar = OxmlElement("w:tblCellMar")
                    for side, val in (("top", "20"), ("bottom", "20"), ("left", "108"), ("right", "108")):
                        m_el = OxmlElement(f"w:{side}")
                        m_el.set(qn("w:w"), val)
                        m_el.set(qn("w:type"), "dxa")
                        new_mar.append(m_el)
                    tblPr.append(new_mar)

                for r_idx in range(1, len(tbl.rows)):
                    trPr = tbl.rows[r_idx]._tr.get_or_add_trPr()
                    if trPr.find(qn("w:cantSplit")) is None:
                        trPr.append(OxmlElement("w:cantSplit"))
                    cell = tbl.rows[r_idx].cells[0]
                    heading_txt = cell.paragraphs[0].text.strip() if cell.paragraphs else ""
                    h_low = heading_txt.lower()
                    for kw_tuple, rec_key in key_map:
                        if any(kw in h_low for kw in kw_tuple) and rec_key in recommendations:
                            # Lưu lại pPr mẫu của dòng gạch đầu dòng thứ 2 (chứa numPr bullet '–' và hanging indent)
                            bullet_pPr = None
                            if len(cell.paragraphs) > 1 and cell.paragraphs[1]._p.pPr is not None:
                                bullet_pPr = deepcopy(cell.paragraphs[1]._p.pPr)

                            # Xóa trHeight cố định (nếu có) để ô co giãn vừa khít nội dung
                            if tbl.rows[r_idx]._tr.trPr is not None:
                                tr_h = tbl.rows[r_idx]._tr.trPr.find(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trHeight"
                                )
                                if tr_h is not None:
                                    tbl.rows[r_idx]._tr.trPr.remove(tr_h)

                            # Giữ lại đoạn tiêu đề đầu tiên (in đậm), xóa các đoạn nội dung bên dưới
                            while len(cell.paragraphs) > 1:
                                p_del = cell.paragraphs[-1]._element
                                p_del.getparent().remove(p_del)
                            if cell.paragraphs:
                                p0 = cell.paragraphs[0]
                                p0.paragraph_format.space_before = Pt(0)
                                p0.paragraph_format.space_after = Pt(0)
                                p0.paragraph_format.line_spacing = 1.05
                                for r in p0.runs:
                                    r.font.name = "Times New Roman"
                                    r.font.size = Pt(13)
                                    r.font.bold = True
                            lines = [
                                ln.strip()
                                for ln in recommendations[rec_key].split("\n")
                                if ln.strip()
                            ]
                            for ln in lines:
                                clean_ln = re.sub(r"^[-–+•]\s*", "", ln)
                                p_new = cell.add_paragraph()
                                if bullet_pPr is not None:
                                    if p_new._p.pPr is not None:
                                        p_new._p.remove(p_new._p.pPr)
                                    p_new._p.insert(0, deepcopy(bullet_pPr))
                                p_new.paragraph_format.space_before = Pt(0)
                                p_new.paragraph_format.space_after = Pt(0)
                                p_new.paragraph_format.line_spacing = 1.05
                                run = p_new.add_run(clean_ln)
                                run.font.name = "Times New Roman"
                                run.font.size = Pt(13)
                                run.font.bold = False
                            updated = True
                            break
                if updated:
                    # Xóa ngắt trang cứng (<w:br w:type="page"/>) thừa ngay sau bảng Khuyến nghị
                    # để khối chữ ký đi liền tự nhiên (hoặc tự sang trang nếu không đủ chỗ nhờ cantSplit + keepNext)
                    nxt = tbl._tbl.getnext()
                    while nxt is not None and nxt.tag.endswith("}p"):
                        for br_el in list(nxt.iter(qn("w:br"))):
                            if br_el.get(qn("w:type")) == "page":
                                br_el.getparent().remove(br_el)
                        nxt = nxt.getnext()
                    # Giữ bảng chữ ký (bảng cuối tài liệu) luôn nằm nguyên khối trên cùng 1 trang
                    if len(doc.tables) > 0:
                        sig_tbl = doc.tables[-1]
                        if sig_tbl is not tbl and len(sig_tbl.rows) >= 2:
                            for s_idx, s_row in enumerate(sig_tbl.rows):
                                s_trPr = s_row._tr.get_or_add_trPr()
                                if s_trPr.find(qn("w:cantSplit")) is None:
                                    s_trPr.append(OxmlElement("w:cantSplit"))
                                if s_idx < len(sig_tbl.rows) - 1:
                                    _set_row_keep_next(s_row, True)
                    break

        if len(tbl.columns) < 2:
            continue

        # Kiểu 2: Bảng 2 cột dọc (Cột 0 là tên nhóm giải pháp, Cột 1 là nội dung)
        # BẮT BUỘC kiểm tra trước xem bảng có >= 3 dòng khớp tên nhóm giải pháp không
        # (tránh nhận nhầm Table 1 chứa "TRUNG TÂM KỸ THUẬT..." ở trang 2!)
        candidate_pairs: List[Tuple[int, str]] = []
        for r_idx, row in enumerate(tbl.rows):
            if row.cells[0]._tc is row.cells[1]._tc:
                continue
            c0 = row.cells[0].text.strip().lower()
            if "trung tâm kỹ thuật" in c0 or "ủy ban tiêu chuẩn" in c0:
                continue
            for kw_tuple, rec_key in key_map:
                if any(kw in c0 for kw in kw_tuple) and rec_key in recommendations:
                    candidate_pairs.append((r_idx, rec_key))
                    break
        if len(candidate_pairs) >= 3:
            for r_idx, rec_key in candidate_pairs:
                _set_cell_text_preserve_font(
                    tbl.rows[r_idx].cells[1],
                    recommendations[rec_key],
                    font_size_pt=13.0,
                    bold=False,
                    align=WD_ALIGN_PARAGRAPH.LEFT,
                )
            updated = True
            break

        # Kiểu 3: Bảng 2 dòng x 5 cột ngang (mẫu 891N2.25DG / 930N2.25DG)
        if len(tbl.rows) == 2 and len(tbl.columns) == 5:
            headers = [c.text.strip().lower() for c in tbl.rows[0].cells]
            if any("kỹ thuật" in h for h in headers) and any("tổ chức" in h for h in headers):
                for c_idx, h_txt in enumerate(headers):
                    for kw_tuple, rec_key in key_map:
                        if any(kw in h_txt for kw in kw_tuple) and rec_key in recommendations:
                            _set_cell_text_preserve_font(
                                tbl.rows[1].cells[c_idx],
                                recommendations[rec_key],
                                font_size_pt=13.0,
                                bold=False,
                                align=WD_ALIGN_PARAGRAPH.LEFT,
                            )
                            updated = True
                            break
                if updated:
                    # Xóa ngắt trang cứng (<w:br w:type="page"/>) thừa ngay sau bảng Khuyến nghị
                    # để khối chữ ký đi liền tự nhiên (hoặc tự sang trang nếu không đủ chỗ nhờ cantSplit + keepNext)
                    nxt = tbl._tbl.getnext()
                    while nxt is not None and nxt.tag.endswith("}p"):
                        for br_el in list(nxt.iter(qn("w:br"))):
                            if br_el.get(qn("w:type")) == "page":
                                br_el.getparent().remove(br_el)
                        nxt = nxt.getnext()
                    # Giữ bảng chữ ký (bảng cuối tài liệu) luôn nằm nguyên khối trên cùng 1 trang
                    if len(doc.tables) > 0:
                        sig_tbl = doc.tables[-1]
                        if sig_tbl is not tbl and len(sig_tbl.rows) >= 2:
                            for s_idx, s_row in enumerate(sig_tbl.rows):
                                s_trPr = s_row._tr.get_or_add_trPr()
                                if s_trPr.find(qn("w:cantSplit")) is None:
                                    s_trPr.append(OxmlElement("w:cantSplit"))
                                if s_idx < len(sig_tbl.rows) - 1:
                                    _set_row_keep_next(s_row, True)
                    break

    return updated


def _set_row_keep_next(row: docx.table._Row, enabled: bool = True) -> None:
    """
    Bật/tắt keep_with_next (<w:keepNext/>) cho MỌI thẻ <w:p> trực tiếp trong `row._tr`
    (bao gồm cả các ô <w:tc> có <w:vMerge/> continue mà `row.cells` của python-docx bỏ qua).
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    for p_el in row._tr.iter(qn("w:p")):
        pPr = p_el.get_or_add_pPr()
        kn = pPr.find(qn("w:keepNext"))
        if enabled:
            if kn is None:
                pPr.append(OxmlElement("w:keepNext"))
        else:
            if kn is not None:
                pPr.remove(kn)


def ensure_mtld_table_headers_and_cantsplit(
    doc: docx.Document,
    detach_summary_rows: bool = True,
) -> None:
    """
    Đặt thuộc tính lặp lại dòng tiêu đề (`<w:tblHeader/>`) khi sang trang cho toàn bộ
    các bảng kết quả quan trắc và bảng tổng hợp trong TBKQ / BCKQ, đồng thời:
    1. Nếu `detach_summary_rows=True`: tách các dòng không viền 'Tổng hợp kết quả quan trắc:'
       nằm dưới dòng 'Tổng cộng' ra thành các đoạn văn ngay sau bảng (để Word không bao giờ
       lặp lại header bảng trống phía trên dòng chữ 'Tổng số mẫu...').
    2. Đặt `keep_with_next` cho các dòng tiêu đề bảng (`r0..header_end_idx`) để tiêu đề bảng
       không bao giờ nằm mồ côi ở cuối trang mà không có dòng dữ liệu.
    3. Đặt `keep_with_next` cho dòng dữ liệu cuối cùng trước 'Tổng cộng' để dòng 'Tổng cộng'
       luôn đi kèm ít nhất 1 dòng dữ liệu khi sang trang.
    4. Đặt `<w:cantSplit/>` cho các dòng dữ liệu và `<w:noWrap/>` cho các ô kết quả dạng `< 0,94`.
    """
    from copy import deepcopy
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.table import Table

    # 1. Xóa các đoạn văn trống thừa (<w:p> rỗng không chứa ngắt trang hay sectPr) trong body
    #    từ sau bảng thông tin chung (Table 2) / mục I. CÁC YẾU TỐ trở đi
    #    (loại bỏ cụm 5 dòng trống trước mục IV. HƠI KHÍ ĐỘC và 21 dòng trống cuối tài liệu).
    body = doc.element.body
    seen_section_one = False
    tbl_count_seen = 0
    for child in list(body):
        if child.tag.endswith("}tbl"):
            tbl_count_seen += 1
            if tbl_count_seen >= 2:
                seen_section_one = True
        elif child.tag.endswith("}p"):
            p_txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            p_upper = p_txt.upper()
            if (
                p_txt.startswith("I.")
                or "CÁC YẾU TỐ" in p_upper
                or "KẾT QUẢ QUAN TRẮC" in p_upper
            ):
                seen_section_one = True
            has_page_br = any(
                br.get(qn("w:type")) == "page" for br in child.iter(qn("w:br"))
            )
            pPr = child.find(qn("w:pPr"))
            has_sect_pr = pPr is not None and pPr.find(qn("w:sectPr")) is not None
            has_drawing = child.find(".//" + qn("w:drawing")) is not None
            if seen_section_one and not p_txt and not has_page_br and not has_sect_pr and not has_drawing:
                body.remove(child)

    # 2. Giữ các đoạn tiêu đề mục (I., II., III., IV., 1., 2. — kể cả đánh số tự động w:numPr)
    #    luôn dính với bảng bên dưới và chuẩn hóa khoảng cách đoạn (space_before = 6pt, space_after = 2pt)
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        pPr = p._p.find(qn("w:pPr"))
        has_num_pr = pPr is not None and pPr.find(qn("w:numPr")) is not None
        if t and (re.match(r"^([IVX]+\.|[0-9]+\.)", t) or (has_num_pr and len(t) < 90)):
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)

    for tbl in list(doc.tables):
        if len(tbl.rows) < 3:
            continue
        top_txt = " ".join(c.text.strip().lower() for r in tbl.rows[:3] for c in r.cells)
        is_result_or_summary = any(
            k in top_txt
            for k in (
                "giới hạn cho phép",
                "vị trí quan trắc",
                "tổng hợp kết quả quan trắc",
                "yếu tố đo, kiểm tra",
                "số mẫu đạt",
            )
        )
        if not is_result_or_summary:
            continue

        # Nếu là bảng kết quả chi tiết có các dòng không viền "Tổng hợp kết quả quan trắc:" / "Ghi chú:"
        # nằm dưới dòng "Tổng cộng", tách các dòng đó sang một bảng không viền riêng ngay sau `tbl._tbl`
        # (giữ nguyên 100% định dạng cột, căn lề và bảng con Ghi chú, đồng thời ngăn Word lặp lại
        # <w:tblHeader/> phía trên dòng chữ "Tổng số mẫu..." khi sang trang).
        total_row_idx = -1
        for r_idx, row in enumerate(tbl.rows):
            row_txt = " ".join(c.text.strip().lower() for c in row.cells)
            if "tổng cộng" in row_txt:
                total_row_idx = r_idx
                break

        has_detached_summary = False
        if (
            detach_summary_rows
            and "giới hạn cho phép" in top_txt
            and total_row_idx >= 0
            and total_row_idx < len(tbl.rows) - 1
        ):
            summary_tbl_el = OxmlElement("w:tbl")
            if tbl._tbl.tblPr is not None:
                summary_tbl_el.append(deepcopy(tbl._tbl.tblPr))
            tbl_grid = tbl._tbl.find(qn("w:tblGrid"))
            if tbl_grid is not None:
                summary_tbl_el.append(deepcopy(tbl_grid))

            rows_to_move = [tbl.rows[r_idx]._tr for r_idx in range(total_row_idx + 1, len(tbl.rows))]
            kept_any_row = False
            for tr_el in rows_to_move:
                tbl._tbl.remove(tr_el)
                row_text_all = "".join(t.text or "" for t in tr_el.iter(qn("w:t"))).strip()
                has_nested_tbl = tr_el.find(".//" + qn("w:tbl")) is not None
                # Bỏ qua các dòng trắng thừa (như dòng r7 rỗng ngăn cách hoặc dòng rỗng cuối bảng)
                if not row_text_all and not has_nested_tbl:
                    continue
                # Trong ô chứa bảng con Ghi chú, xóa dòng <w:p> trống thừa nằm sau bảng con <w:tbl>
                if has_nested_tbl:
                    for tc_el in tr_el.findall(qn("w:tc")):
                        seen_nested = False
                        for tc_child in list(tc_el):
                            if tc_child.tag.endswith("}tbl"):
                                seen_nested = True
                            elif seen_nested and tc_child.tag.endswith("}p"):
                                p_t = "".join(x.text or "" for x in tc_child.iter(qn("w:t"))).strip()
                                # Word bắt buộc mỗi <w:tc> phải kết thúc bằng ít nhất 1 <w:p>;
                                # thu nhỏ chiều cao dòng <w:p> kết thúc sau bảng con về 1pt để triệt tiêu khoảng trắng
                                if not p_t:
                                    pPr = tc_child.get_or_add_pPr()
                                    sp = pPr.find(qn("w:spacing"))
                                    if sp is None:
                                        sp = OxmlElement("w:spacing")
                                        pPr.append(sp)
                                    sp.set(qn("w:before"), "0")
                                    sp.set(qn("w:after"), "0")
                                    sp.set(qn("w:line"), "20")
                                    sp.set(qn("w:lineRule"), "exact")

                trPr = tr_el.get_or_add_trPr()
                th = trPr.find(qn("w:tblHeader"))
                if th is not None:
                    trPr.remove(th)
                # Gỡ chiều cao cố định quá lớn trên dòng chứa Ghi chú / dòng tổng hợp để ôm sát chữ
                trHeight = trPr.find(qn("w:trHeight"))
                if trHeight is not None and has_nested_tbl:
                    trPr.remove(trHeight)
                if trPr.find(qn("w:cantSplit")) is None:
                    trPr.append(OxmlElement("w:cantSplit"))
                summary_tbl_el.append(tr_el)
                kept_any_row = True

            if kept_any_row:
                tbl._tbl.addnext(summary_tbl_el)
                summary_tbl = Table(summary_tbl_el, doc)
                for s_idx, s_row in enumerate(summary_tbl.rows):
                    if s_idx < len(summary_tbl.rows) - 1:
                        _set_row_keep_next(s_row, True)
                has_detached_summary = True

        # Xác định dòng cuối của phần header bảng (dòng có chữ STT / TT / Số mẫu đạt / Silic)
        header_end_idx = 0
        for r_idx in range(min(4, len(tbl.rows))):
            r_txt = " ".join(c.text.strip().lower() for c in tbl.rows[r_idx].cells)
            if any(
                k in r_txt
                for k in ("stt", "vị trí quan trắc", "số mẫu đạt", "tổng số mẫu")
            ):
                header_end_idx = r_idx

        num_rows = len(tbl.rows)
        for r_idx, row in enumerate(tbl.rows):
            trPr = row._tr.get_or_add_trPr()
            # Lặp lại các dòng tiêu đề bảng khi bảng tràn sang trang mới
            if r_idx <= header_end_idx:
                if trPr.find(qn("w:tblHeader")) is None:
                    trPr.append(OxmlElement("w:tblHeader"))
                # Giữ header luôn đi kèm ít nhất dòng dữ liệu đầu tiên (không để mồ côi header ở cuối trang)
                _set_row_keep_next(row, True)
            # Giữ dòng dữ liệu cuối cùng luôn dính với dòng "Tổng cộng" (và dính với khối Tổng hợp kết quả)
            elif num_rows >= header_end_idx + 3 and r_idx == num_rows - 2:
                _set_row_keep_next(row, True)
            elif has_detached_summary and r_idx == num_rows - 1:
                _set_row_keep_next(row, True)

            # Chống cắt đôi dòng bảng giữa 2 trang
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))

            # Chống gãy dòng kiểu "<\n0,94" trong các ô kết quả hẹp
            if r_idx > header_end_idx:
                for cell in row.cells:
                    ctxt = cell.text.strip()
                    if ctxt.startswith("<") and len(ctxt) <= 10:
                        tcPr = cell._tc.get_or_add_tcPr()
                        if tcPr.find(qn("w:noWrap")) is None:
                            tcPr.append(OxmlElement("w:noWrap"))


def _replace_paragraph_with_page_numpages(
    p: docx.text.paragraph.Paragraph,
    prefix: str = "Trang ",
    font_size_pt: float = 13.0,
    italic: bool = False,
) -> None:
    """Thay nội dung đoạn văn bằng trường động `Trang {PAGE}/{NUMPAGES}` chuẩn Word."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    for r in list(p.runs):
        p._p.remove(r._r)

    def _add_run_text(txt: str) -> None:
        r = p.add_run(txt)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size_pt)
        r.font.italic = italic

    def _add_field(field_name: str) -> None:
        for fld_type, instr in (("begin", None), (None, f" {field_name} "), ("separate", None), ("end", None)):
            r = p.add_run()
            r.font.name = "Times New Roman"
            r.font.size = Pt(font_size_pt)
            r.font.italic = italic
            if fld_type is not None:
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), fld_type)
                r._r.append(fc)
            elif instr is not None:
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = instr
                r._r.append(it)

    _add_run_text(prefix)
    _add_field("PAGE")
    _add_run_text("/")
    _add_field("NUMPAGES")


def fix_mtld_header_and_page_fields(
    doc: docx.Document,
    dossier_code: str,
    issue_date_str: str,
    is_bckq: bool = False,
) -> None:
    """
    Chuẩn hóa Header & Footer của TBKQ / BCKQ:
    - Ở TBKQ: cập nhật đúng ô `cells[1]` = `<mã>/MTLĐ` và `cells[3]` = `<ngày>` trên 1 dòng
      (không làm biến dạng chiều cao header), và đổi số trang ở Footer thành `{PAGE}/{NUMPAGES}`.
    - Ở BCKQ: đổi dòng `Trang 2/2` cứng trong Header thành `{PAGE}/{NUMPAGES}`.
    """
    for sec in doc.sections:
        for hdr in (sec.header, sec.first_page_header):
            if hdr is None:
                continue
            for tbl in hdr.tables:
                if not is_bckq and len(tbl.rows) >= 1 and len(tbl.rows[0].cells) >= 4:
                    r0 = tbl.rows[0]
                    _set_cell_text_preserve_font(r0.cells[0], "Số:", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.LEFT)
                    _set_cell_text_preserve_font(r0.cells[1], f"{dossier_code}/MTLĐ", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.LEFT)
                    _set_cell_text_preserve_font(r0.cells[2], "Ngày", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
                    _set_cell_text_preserve_font(r0.cells[3], issue_date_str, font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
                elif is_bckq and len(tbl.rows) >= 1:
                    last_cell = tbl.rows[0].cells[-1]
                    for p in last_cell.paragraphs:
                        if "Trang" in p.text:
                            _replace_paragraph_with_page_numpages(p, prefix="Trang ", font_size_pt=11.0, italic=True)

        for ftr in (sec.footer, sec.first_page_footer):
            if ftr is None:
                continue
            for p in ftr.paragraphs:
                if "Trang" in p.text:
                    _replace_paragraph_with_page_numpages(p, prefix="Trang ", font_size_pt=13.0, italic=False)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def format_count_2digits(val: int) -> str:
    """Định dạng số lượng mẫu trong bảng tổng hợp theo chuẩn 2 chữ số (00, 01, 02... 15)."""
    if val <= 0:
        return "00"
    return f"{val:02d}"


def sync_mtld_summary_table(
    doc: docx.Document,
    factor_summary: List[Dict[str, Any]],
) -> bool:
    """
    Đồng bộ bảng Tổng hợp kết quả quan trắc (Table có tiêu đề 'TỔNG HỢP KẾT QUẢ QUAN TRẮC MÔI TRƯỜNG LAO ĐỘNG'
    hoặc cột 'Tổng số mẫu', 'Số mẫu đạt', 'Số mẫu không đạt') trong cả TBKQ và BCKQDG.
    Mỗi phần tử trong `factor_summary`:
      {"name": "Nhiệt độ", "total": 5, "pass": 4, "fail": 1, "display_name": "..."}
    """
    target_tbl = None
    for tbl in doc.tables:
        if len(tbl.rows) < 4 or len(tbl.columns) < 5:
            continue
        top_text = " ".join(
            c.text.strip().lower() for r in tbl.rows[:3] for c in r.cells
        )
        if "tổng hợp kết quả quan trắc" in top_text or (
            "tổng số mẫu" in top_text and "số mẫu đạt" in top_text
        ):
            target_tbl = tbl
            break

    if target_tbl is None:
        return False

    by_key: Dict[str, Dict[str, Any]] = {}
    for item in factor_summary:
        k = str(item.get("name", "")).strip().lower()
        if k:
            by_key[k] = item

    sum_tot = 0
    sum_pas = 0
    sum_fal = 0

    for r_idx in range(1, len(target_tbl.rows)):
        row = target_tbl.rows[r_idx]
        c1_txt = row.cells[1].text.strip()
        f_txt = c1_txt.lower()
        if not f_txt or f_txt in ("yếu tố đo, kiểm tra", "bụi", "hơi khí độc"):
            continue

        if "tổng cộng" in f_txt or "tổng cộng" in row.cells[0].text.strip().lower():
            if len(row.cells) >= 8:
                _set_cell_text_preserve_font(
                    row.cells[2],
                    format_count_2digits(sum_tot),
                    font_size_pt=13.0,
                    bold=True,
                    align=WD_ALIGN_PARAGRAPH.CENTER,
                )
                _set_cell_text_preserve_font(
                    row.cells[4],
                    format_count_2digits(sum_pas),
                    font_size_pt=13.0,
                    bold=True,
                    align=WD_ALIGN_PARAGRAPH.CENTER,
                )
                _set_cell_text_preserve_font(
                    row.cells[6],
                    format_count_2digits(sum_fal),
                    font_size_pt=13.0,
                    bold=True,
                    align=WD_ALIGN_PARAGRAPH.CENTER,
                )
            continue

        matched_item = by_key.get(f_txt)
        if matched_item is None:
            clean_f = re.sub(r"^[-–•\s]+", "", f_txt).strip()
            matched_item = by_key.get(clean_f)
        if matched_item is None:
            for k, v in by_key.items():
                clean_k = re.sub(r"^[-–•\s]+", "", k).strip()
                if clean_k and (clean_k == f_txt or clean_k in f_txt or f_txt in clean_k):
                    matched_item = v
                    break

        if matched_item is not None:
            if matched_item.get("display_name"):
                _set_cell_text_preserve_font(
                    row.cells[1],
                    str(matched_item["display_name"]),
                    font_size_pt=13.0,
                    align=WD_ALIGN_PARAGRAPH.LEFT,
                )
            tot = int(matched_item.get("total", 0))
            pas = int(matched_item.get("pass", 0))
            fal = int(matched_item.get("fail", max(0, tot - pas)))
            sum_tot += tot
            sum_pas += pas
            sum_fal += fal

            if len(row.cells) >= 8:
                # Dòng Bụi có chia cột Silic (2,4,6) và Khác (3,5,7)
                if "bụi" in f_txt and row.cells[2]._tc is not row.cells[3]._tc:
                    _set_cell_text_preserve_font(row.cells[2], "–", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[3], format_count_2digits(tot), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[4], "–", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[5], format_count_2digits(pas), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[6], "–", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[7], format_count_2digits(fal), font_size_pt=13.0, bold=(fal > 0), align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    _set_cell_text_preserve_font(row.cells[2], format_count_2digits(tot), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[4], format_count_2digits(pas), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                    _set_cell_text_preserve_font(row.cells[6], format_count_2digits(fal), font_size_pt=13.0, bold=(fal > 0), align=WD_ALIGN_PARAGRAPH.CENTER)
            elif len(row.cells) >= 5:
                factor_cell_idx = 2 if len(row.cells) >= 6 else 1
                _set_cell_text_preserve_font(row.cells[factor_cell_idx + 1], format_count_2digits(tot), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                _set_cell_text_preserve_font(row.cells[factor_cell_idx + 2], format_count_2digits(pas), font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
                _set_cell_text_preserve_font(row.cells[factor_cell_idx + 3], format_count_2digits(fal), font_size_pt=13.0, bold=(fal > 0), align=WD_ALIGN_PARAGRAPH.CENTER)

    return True


def build_mtld_reports_from_template(
    dossier_code: str,
    company_name: str,
    company_address: str,
    inspection_date_str: str,
    factor_summary: List[Dict[str, Any]],
    failing_details: Optional[Dict[str, List[str]]] = None,
    industry_desc: str = "",
    out_tbkq_path: Optional[Path] = None,
    out_bckq_path: Optional[Path] = None,
    use_ai_recommendations: bool = True,
    issue_date_str: Optional[str] = None,
) -> Tuple[Path, Path, Dict[str, str]]:
    """
    Khởi tạo bộ báo cáo TBKQ và BCKQ nội bộ từ bộ Template chuẩn 2026:
    - Tự động cập nhật Mã hồ sơ, Ngày ban hành, Tên & Địa chỉ khách hàng, Ngày quan trắc.
    - Tự động đồng bộ bảng Tổng hợp kết quả quan trắc giữa TBKQ và BCKQ nội bộ.
    - Tự động sinh mục "KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC" bằng AI / Contextual Seeded Composer
      với giọng văn nhẹ nhàng, mang tính ghi nhận và sẽ khắc phục.
    """
    if issue_date_str is None:
        issue_date_str = date.today().strftime("%d/%m/%Y")

    tpl_tbkq = TEMPLATE_MTLD_DIR / "FORM - TBKQ - viet - mtld - 2026.docx"
    tpl_bckq = TEMPLATE_MTLD_DIR / "FORM - BCKQ - mtld - 2026.docx"

    if out_tbkq_path is None:
        out_tbkq_path = ROOT_DIR / "Input" / dossier_code / "TBKQ" / f"TBKQ - viet - mtld - {dossier_code}.docx"
    if out_bckq_path is None:
        out_bckq_path = ROOT_DIR / "Input" / dossier_code / "ThucHien" / f"BCKQ - 1 - mtld - {dossier_code}.docx"

    out_tbkq_path = Path(out_tbkq_path)
    out_bckq_path = Path(out_bckq_path)
    out_tbkq_path.parent.mkdir(parents=True, exist_ok=True)
    out_bckq_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. Sinh khuyến nghị giải pháp khắc phục
    recs = generate_mtld_recommendations(
        company_name=company_name,
        industry_desc=industry_desc,
        factor_summary=factor_summary,
        failing_details=failing_details,
        dossier_code=dossier_code,
        use_ai=use_ai_recommendations,
    )

    # 2. Dựng TBKQ
    doc_tbkq = docx.Document(tpl_tbkq)
    fix_mtld_header_and_page_fields(doc_tbkq, dossier_code, issue_date_str, is_bckq=False)

    # Cập nhật Trang bìa (Table 0) & Số/Ngày trang trong (Table 1) & Thông tin chung (Table 2)
    try:
        dt_issue = datetime.strptime(issue_date_str, "%d/%m/%Y").date()
    except Exception:
        dt_issue = date.today()
    try:
        dt_insp = datetime.strptime(inspection_date_str, "%d/%m/%Y").date()
    except Exception:
        dt_insp = dt_issue

    if len(doc_tbkq.tables) > 0:
        t0 = doc_tbkq.tables[0]
        if len(t0.rows) >= 12:
            _set_cell_text_preserve_font(t0.rows[6].cells[2], f"Ngày {dt_issue.day:02d}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(t0.rows[6].cells[3], f"Tháng {dt_issue.month:02d}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(t0.rows[6].cells[4], f"Năm {dt_issue.year}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(t0.rows[10].cells[0], company_name.upper(), font_size_pt=16.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(t0.rows[11].cells[1], company_address, font_size_pt=13.0, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

    if len(doc_tbkq.tables) > 1:
        t1 = doc_tbkq.tables[1]
        if len(t1.rows) >= 2:
            _set_cell_text_preserve_font(t1.rows[1].cells[0], f"Số: {dossier_code}/MTLĐ", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(
                t1.rows[1].cells[1],
                f"TPHCM, ngày {dt_issue.day:02d} tháng {dt_issue.month:02d} năm {dt_issue.year}",
                font_size_pt=13.0,
                italic=True,
                align=WD_ALIGN_PARAGRAPH.CENTER,
            )

    if len(doc_tbkq.tables) > 2:
        t2 = doc_tbkq.tables[2]
        if len(t2.rows) >= 6:
            _set_cell_text_preserve_font(
                t2.rows[4].cells[0],
                f"Do Ông: Nguyễn Xuân Nam làm đại diện đã tiến hành quan trắc môi trường lao động tại:\n"
                f"{company_name}\n"
                f"Địa chỉ: {company_address}",
                font_size_pt=13.0,
            )
            _set_cell_text_preserve_font(
                t2.rows[5].cells[0],
                f"Ngày {dt_insp.day:02d} tháng {dt_insp.month:02d} năm {dt_insp.year}",
                font_size_pt=13.0,
            )

    sync_mtld_summary_table(doc_tbkq, factor_summary)
    apply_mtld_recommendations_to_doc(doc_tbkq, recs)
    ensure_mtld_table_headers_and_cantsplit(doc_tbkq)
    doc_tbkq.save(out_tbkq_path)

    # 3. Dựng BCKQ nội bộ (BM3030.03 HL:01/8/2025)
    doc_bckq = docx.Document(tpl_bckq)
    fix_mtld_header_and_page_fields(doc_bckq, dossier_code, issue_date_str, is_bckq=True)
    if len(doc_bckq.tables) > 0:
        tb0 = doc_bckq.tables[0]
        if len(tb0.rows) >= 8:
            _set_cell_text_preserve_font(tb0.rows[0].cells[8], dossier_code, font_size_pt=13.0, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
            _set_cell_text_preserve_font(tb0.rows[0].cells[-1], "", font_size_pt=13.0, bold=False)
            _set_cell_text_preserve_font(tb0.rows[2].cells[-1], company_name.upper(), font_size_pt=13.0, bold=False)
            _set_cell_text_preserve_font(tb0.rows[3].cells[-1], company_address, font_size_pt=13.0, bold=False)
            _set_cell_text_preserve_font(tb0.rows[6].cells[-1], f"Nhà máy của {company_name}, {company_address}", font_size_pt=13.0, bold=False)
            _set_cell_text_preserve_font(tb0.rows[7].cells[-1], f"Bắt đầu từ 08:30 đến 15:30 ngày {inspection_date_str}", font_size_pt=13.0, bold=False)

    for p in doc_bckq.paragraphs:
        if "Q326" in p.text:
            new_txt = re.sub(r"Q3\d{2}[A-D]\d{6}", dossier_code, p.text)
            for r in p.runs:
                r.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            run.text = new_txt
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)

    sync_mtld_summary_table(doc_bckq, factor_summary)
    ensure_mtld_table_headers_and_cantsplit(doc_bckq)
    doc_bckq.save(out_bckq_path)

    return out_tbkq_path, out_bckq_path, recs


class MTLDHandler:
    """
    Domain Handler cho hồ sơ Quan trắc Môi trường Lao động (QTMTLĐ) tích hợp trực tiếp vào `TBKQBuilder`.
    """

    def __init__(self, dossier: Any):
        self.dossier = dossier

    def build_tbkq(self, is_en: bool = False, has_assessment: bool = True) -> str:
        dossier_code = getattr(self.dossier, "dossier_no", "") or getattr(self.dossier, "code", "")
        company_name = getattr(self.dossier, "client_name", "") or ""
        company_address = getattr(self.dossier, "client_address", "") or ""
        insp_date = getattr(self.dossier, "inspection_date", "") or date.today().strftime("%d/%m/%Y")
        issue_date = getattr(self.dossier, "cert_date", "") or date.today().strftime("%d/%m/%Y")
        factor_summary = getattr(self.dossier, "factor_summary", []) or []
        failing_details = getattr(self.dossier, "failing_details", {}) or {}
        industry_desc = getattr(self.dossier, "industry_desc", "") or ""

        out_tbkq, _, _ = build_mtld_reports_from_template(
            dossier_code=dossier_code,
            company_name=company_name,
            company_address=company_address,
            inspection_date_str=str(insp_date),
            factor_summary=factor_summary,
            failing_details=failing_details,
            industry_desc=industry_desc,
            issue_date_str=str(issue_date),
        )
        return str(out_tbkq)

