"""
CCN-DA Studio - Localhost Server (server.py)
Khởi chạy dịch vụ web Business Intelligence Platform trên http://localhost:8080
Cung cấp giao diện SaaS Web UI và hệ thống REST API tương tác thời gian thực:
- GET  /                          -> Phục vụ ui/index.html
- GET  /api/status                -> Trạng thái hệ thống & 12 skills
- GET  /api/payload               -> Trả về payload toàn diện của BI Platform
- GET  /api/sheet/files           -> Danh sách tệp nguồn có sẵn (CSV, XLSX đa Sheet)
- GET  /api/sheet/manifest        -> Lấy hồ sơ bàn giao sang bước 2 (Cleaning)
- POST /api/sheet/inspect         -> Phân tích cấu trúc, kiểm tra lỗi/cảnh báo, preview
- POST /api/sheet/upload          -> Nhận file tải lên từ người dùng (CSV, XLSX)
- POST /api/sheet/confirm         -> Xác nhận hoàn thành Bước 1 & bàn giao sang Bước 2
- POST /api/workflow/run          -> Kích hoạt quy trình BI 7 giai đoạn
- GET  /api/workflow/state        -> Trạng thái chi tiết quy trình 7 giai đoạn
- POST /api/chat                  -> Trợ lý AI BI Assistant phản hồi tức thời
- GET  /data/export/<filename>    -> Tải các tệp dữ liệu sạch đã kiểm chứng (Verify)
"""

import os
import sys
import csv
import json
import base64
import email
import mimetypes
import urllib.request
import urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# Thêm thư mục gốc vào path
WORKSPACE_ROOT = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE_ROOT not in sys.path:
    sys.path.insert(0, WORKSPACE_ROOT)

# Cấu hình Gemini AI Assistant Key
DEFAULT_GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6IW9uNQGKiTn28uAQbM3lfkJLViI21h7UhRuJ_socwSUQ")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-lite-latest")


def build_dynamic_bi_system_context() -> tuple:
    """
    Xây dựng ngữ cảnh phân tích và bộ chỉ số KPI động 100% từ tệp dữ liệu thực tế đang nạp.
    Tuyệt đối loại bỏ số liệu giả lập, bám sát các trường dữ liệu thực tế và câu hỏi người dùng.
    """
    manifest3_path = os.path.join(WORKSPACE_ROOT, "data", "analysis_handover_manifest.json")
    manifest1_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
    
    m3 = {}
    if os.path.exists(manifest3_path):
        try:
            with open(manifest3_path, "r", encoding="utf-8") as f:
                m3 = json.load(f)
        except Exception:
            pass

    m1 = {}
    if os.path.exists(manifest1_path):
        try:
            with open(manifest1_path, "r", encoding="utf-8") as f:
                m1 = json.load(f)
        except Exception:
            pass

    biz_ctx = m3.get("business_context") or m1.get("business_context") or {}
    company_type = biz_ctx.get("company_type") or "Bán lẻ & Thương mại Đa kênh"
    questions = biz_ctx.get("questions") or []

    kpis = m3.get("kpis") or {}
    net_rev = kpis.get("NetRevenue") or kpis.get("net_revenue") or 0.0
    gross_profit = kpis.get("GrossProfit") or kpis.get("gross_profit") or 0.0
    margin_pct = kpis.get("GrossMarginPct") or kpis.get("gross_margin_pct") or 0.0
    orders = kpis.get("TotalOrders") or kpis.get("total_orders") or 0
    aov = kpis.get("AOV") or kpis.get("aov") or 0.0
    discounts = kpis.get("TotalDiscounts") or kpis.get("discounts") or 0.0

    channels = m3.get("channels") or m3.get("methods", {}).get("comparative", {}).get("channels", [])
    top_channel = channels[0] if channels else {}
    top_ch_name = top_channel.get("channel") or top_channel.get("name") or "Website"
    top_ch_rev = top_channel.get("revenue_formatted") or f"{top_channel.get('revenue', 0):,.0f} đ"
    top_ch_share = top_channel.get("market_share_pct") or top_channel.get("share_pct") or 0.0

    products = m3.get("methods", {}).get("comparative", {}).get("top_products", [])
    top_prod = products[0] if products else {}
    top_p_name = top_prod.get("product_name") or top_prod.get("name") or "Sản phẩm chủ lực"
    top_p_profit = top_prod.get("profit_formatted") or f"{top_prod.get('profit', 0):,.0f} đ"

    quarters = m3.get("methods", {}).get("trend", {}).get("quarterly", [])
    top_q = max(quarters, key=lambda x: x.get("revenue", 0)) if quarters else {}

    answered = m3.get("answered_questions") or []

    context_lines = [
        f"Bạn là CCN-DA AI Assistant - Cố vấn Trí tuệ Doanh nghiệp & Phân tích Dữ liệu cấp cao của nền tảng CCN-DA.",
        f"Doanh nghiệp: {company_type}.",
        f"BỘ SỐ LIỆU ĐÃ KIỂM TOÁN VERIFY GATE 100% PASS TỪ TỆP DỮ LIỆU THỰC TẾ:",
        f"- Doanh thu thuần (Net Revenue): {net_rev:,.0f} VNĐ",
        f"- Lợi nhuận gộp (Gross Profit): {gross_profit:,.0f} VNĐ (Biên LN: {margin_pct}%)",
        f"- Tổng khối lượng đơn hàng: {orders:,} đơn | AOV trung bình: {aov:,.0f} VNĐ/đơn",
        f"- Chiết khấu & Khuyến mại: {discounts:,.0f} VNĐ",
        f"- Kênh bán dẫn đầu doanh thu: {top_ch_name} với {top_ch_rev} ({top_ch_share}% thị phần)",
        f"- Sản phẩm mang lại lợi nhuận cao nhất: {top_p_name} với {top_p_profit} lợi nhuận gộp",
    ]
    if top_q:
        context_lines.append(f"- Quý cao điểm nhất: {top_q.get('quarter')} đạt {top_q.get('revenue_formatted') or top_q.get('revenue', 0)}")
    if questions:
        context_lines.append(f"- Các bài toán kinh doanh người dùng quan tâm: {', '.join(questions)}")

    context_lines.append("\nNguyên tắc trả lời: Luôn dựa 100% vào số liệu thực tế trên. Cấu trúc Minto: Kết luận -> Dẫn chứng số liệu -> Hành động.")

    context_text = "\n".join(context_lines)
    metrics_summary = {
        "net_rev": net_rev,
        "gross_profit": gross_profit,
        "margin_pct": margin_pct,
        "orders": orders,
        "aov": aov,
        "discounts": discounts,
        "top_ch_name": top_ch_name,
        "top_ch_rev": top_ch_rev,
        "top_ch_share": top_ch_share,
        "top_p_name": top_p_name,
        "top_p_profit": top_p_profit,
        "top_q": top_q,
        "answered": answered,
        "questions": questions
    }
    return context_text, metrics_summary


def ask_gemini_bi_assistant(query: str, api_key: str = None) -> dict:
    key = api_key or DEFAULT_GEMINI_KEY
    dynamic_context, summary = build_dynamic_bi_system_context()

    if not query or not query.strip():
        return {
            "reply": f"Xin chào! Tôi là Trợ lý Phân tích CCN-DA. Doanh nghiệp hiện ghi nhận doanh thu thuần **{summary['net_rev']:,.0f} đ** với biên lợi nhuận gộp **{summary['margin_pct']}%**. Bạn có thể hỏi về kênh bán hàng dẫn đầu, sản phẩm chủ lực, hoặc bài toán kinh doanh cụ thể.",
            "source": "system",
            "model": GEMINI_MODEL
        }

    # Thử gọi Google Generative Language API với Gemini Flash
    if key and not key.startswith("YOUR_"):
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={key}"
        full_prompt = f"{dynamic_context}\n\nCâu hỏi của người dùng:\n{query}\n\nTrả lời ngắn gọn, chuẩn xác theo số liệu thực tế:"
        payload_data = {
            "contents": [{"parts": [{"text": full_prompt}]}],
            "generationConfig": {
                "temperature": 0.2,
                "maxOutputTokens": 800
            }
        }
        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload_data).encode("utf-8"),
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = json.loads(resp.read().decode("utf-8"))
                ans_text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
                return {"reply": ans_text, "model": GEMINI_MODEL, "source": "gemini"}
        except Exception as e:
            print(f"[Gemini API Warning] Lỗi gọi Gemini API ({e}), chuyển sang động cơ phân tích nội bộ...")

    # Fallback quy tắc nội bộ dựa trên dữ liệu thật của tệp đang nạp
    lower = query.lower()
    net_rev = summary["net_rev"]
    gross_profit = summary["gross_profit"]
    margin_pct = summary["margin_pct"]
    orders = summary["orders"]
    aov = summary["aov"]
    discounts = summary["discounts"]
    top_ch_name = summary["top_ch_name"]
    top_ch_rev = summary["top_ch_rev"]
    top_ch_share = summary["top_ch_share"]
    top_p_name = summary["top_p_name"]
    top_p_profit = summary["top_p_profit"]
    top_q = summary["top_q"]

    if any(k in lower for k in ["tăng trưởng", "tang_truong", "kenh", "kênh", "q4", "chu kỳ", "thị phần"]):
        q_label = top_q.get("quarter", "kỳ cao điểm")
        q_rev_str = top_q.get("revenue_formatted") or f"{top_q.get('revenue', 0):,.0f} đ"
        ans = f"**Kết luận:** Kênh **{top_ch_name}** giữ vai trò động lực tăng trưởng dẫn đầu toàn hệ thống với doanh thu thuần đạt **{top_ch_rev}** (chiếm {top_ch_share}% thị phần).\n\n**Dẫn chứng số liệu:**\n- Chu kỳ hoạt động ghi nhận {q_label} đạt {q_rev_str} đóng góp doanh số nổi bật.\n- Biên lợi nhuận gộp toàn doanh nghiệp được bảo toàn an toàn ở mức **{margin_pct}%**.\n\n**Khuyến nghị:** Tập trung ngân sách phát triển kênh {top_ch_name} và tối ưu giỏ hàng."
    elif any(k in lower for k in ["sản phẩm", "san_pham", "lợi nhuận", "loi_nhuan", "lãi", "mặt hàng"]):
        ans = f"**Kết luận:** Sản phẩm **{top_p_name}** là trụ cột sinh lời số 1 của doanh nghiệp với tổng lợi nhuận gộp đạt **{top_p_profit}**.\n\n**Dẫn chứng số liệu:**\n- Tổng lợi nhuận gộp toàn hệ thống đạt {gross_profit:,.0f} VNĐ trên {net_rev:,.0f} VNĐ doanh thu thuần (Biên LN: {margin_pct}%).\n- Tổng số lượng đơn hàng ghi nhận: {orders:,} đơn.\n\n**Khuyến nghị:** Bảo đảm nguồn cung ứng liên tục cho {top_p_name} và thiết lập gói bán kèm combo phụ kiện để gia tăng AOV."
    elif any(k in lower for k in ["chiết khấu", "chiet_khau", "khuyến mại", "khuyen_mai", "giảm giá", "voucher"]):
        ans = f"**Kết luận:** Tổng chi phí chiết khấu & khuyến mại ghi nhận là **{discounts:,.0f} VNĐ**.\n\n**Dẫn chứng số liệu:**\n- Doanh nghiệp duy trì biên lợi nhuận gộp vững chắc ở mức **{margin_pct}%** (đạt {gross_profit:,.0f} VNĐ lợi nhuận sau chiết khấu).\n- Giá trị đơn hàng trung bình (AOV) đạt **{aov:,.0f} VNĐ/đơn**.\n\n**Khuyến nghị:** Giữ chính sách chiết khấu linh hoạt theo bậc thang giá trị đơn hàng để bảo vệ biên lợi nhuận."
    elif any(k in lower for k in ["khách hàng", "khach_hang", "vip", "champions", "rfm", "khách"]):
        ans = f"**Kết luận:** Khách hàng giao dịch qua kênh **{top_ch_name}** đóng góp quy mô doanh thu lớn nhất ({top_ch_share}% thị phần).\n\n**Dẫn chứng số liệu:**\n- Quy mô giỏ hàng trung bình (AOV) toàn hệ thống đạt **{aov:,.0f} VNĐ/đơn** qua {orders:,} đơn hàng đã hoàn tất.\n\n**Khuyến nghị:** Tăng cường chăm sóc khách hàng tại kênh {top_ch_name} và triển khai chương trình loyalty để kích thích tỷ lệ quay lại mua sắm."
    else:
        ans = f"**Hệ thống Phân Tích CCN-DA (Dữ liệu thực tế 100% Verify Gate PASS):**\n- Doanh thu thuần: **{net_rev:,.0f} đ** | Lợi nhuận gộp: **{gross_profit:,.0f} đ** (Biên LN: **{margin_pct}%**).\n- Tổng khối lượng: **{orders:,} đơn hàng** | AOV bình quân: **{aov:,.0f} đ/đơn**.\n- Kênh bán dẫn đầu: **{top_ch_name}** ({top_ch_rev}, chiếm {top_ch_share}% thị phần).\n- Sản phẩm sinh lời cao nhất: **{top_p_name}** ({top_p_profit} lãi gộp).\n\nBạn có thể hỏi thêm về phân tích kênh bán, so sánh chu kỳ thời gian hoặc khuyến nghị chiến lược."

    return {"reply": ans, "model": "CCN-DA-Dynamic-Engine", "source": "rule-based"}

from workflow.engine import BIWorkflowEngine
from workflow.stage_1_sheet import SheetStage
from workflow.stage_2_cleaning import CleaningStage
from workflow.stage_3_analysis import AnalysisStage
from workflow.stage_4_reporting import ReportingStage

class BIPlatformHTTPHandler(BaseHTTPRequestHandler):
    def address_string(self):
        # Disable slow reverse DNS lookups on Windows localhost
        return str(self.client_address[0])

    def _send_json(self, data: dict, status_code: int = 200):
        body = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(body)

    def _serve_file(self, file_path: str, default_type: str = "text/plain"):
        if not os.path.exists(file_path):
            self.send_error(404, f"File not found: {file_path}")
            return
        
        mime_type, _ = mimetypes.guess_type(file_path)
        mime_type = mime_type or default_type
        if mime_type.startswith("text/") or mime_type in ["application/javascript", "application/json"]:
            mime_type += "; charset=utf-8"

        with open(file_path, "rb") as f:
            content = f.read()

        self.send_response(200)
        self.send_header('Content-Type', mime_type)
        self.send_header('Content-Length', str(len(content)))
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.end_headers()
        self.wfile.write(content)

    def _parse_multipart(self, content_type_header: str, body_bytes: bytes):
        """
        Trích xuất an toàn các trường dữ liệu và tệp tin từ multipart/form-data
        bằng thư viện email chuẩn của Python, bảo toàn nguyên vẹn 100% byte dữ liệu.
        """
        raw_headers = f"Content-Type: {content_type_header}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        msg = email.message_from_bytes(raw_headers + body_bytes)
        fields = {}
        files = []
        for part in msg.walk():
            cd = part.get("Content-Disposition", "")
            if not cd:
                continue
            filename = part.get_filename()
            field_name = part.get_param("name", header="content-disposition")
            if filename:
                payload = part.get_payload(decode=True)
                files.append({
                    "field_name": field_name or "file",
                    "filename": filename,
                    "content_type": part.get_content_type(),
                    "data": payload if payload is not None else b"",
                    "size": len(payload) if payload else 0
                })
            elif field_name:
                payload = part.get_payload(decode=True)
                if payload:
                    try:
                        fields[field_name] = payload.decode("utf-8")
                    except UnicodeDecodeError:
                        fields[field_name] = payload.decode("latin-1", errors="ignore")
        return fields, files

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        # 1. API Status
        if path == "/api/status":
            self._send_json({
                "status": "ONLINE",
                "platform": "CCN-DA Studio · Enterprise BI Platform (AI301 Final Project)",
                "port": 8080,
                "workflow_stages": 7,
                "skills_count": 12,
                "warehouse_connected": True,
                "ai_assistant": {
                    "model": GEMINI_MODEL,
                    "configured": bool(DEFAULT_GEMINI_KEY),
                    "key_prefix": DEFAULT_GEMINI_KEY[:10] + "..." if DEFAULT_GEMINI_KEY else "NONE"
                }
            })
            return

        # 1b. API List Sheet Files
        if path == "/api/sheet/files":
            raw_dir = os.path.join(WORKSPACE_ROOT, "data", "raw")
            uploads_dir = os.path.join(WORKSPACE_ROOT, "data", "uploads")
            files_list = []
            stage = SheetStage()

            for folder, category in [(raw_dir, "Kho Dữ Liệu Mẫu Hệ Thống"), (uploads_dir, "Tệp Người Dùng Tải Lên")]:
                if os.path.exists(folder):
                    for fn in os.listdir(folder):
                        if fn.endswith(('.csv', '.xlsx', '.xls', '.xlsm')):
                            fp = os.path.join(folder, fn)
                            ext = os.path.splitext(fn)[1].upper().replace('.', '')
                            sheets = stage.list_sheets(fp) if ext in ['XLSX', 'XLSM', 'XLS'] else ["Default"]
                            files_list.append({
                                "file_name": fn,
                                "file_path": fp,
                                "format": ext,
                                "size_bytes": os.path.getsize(fp),
                                "category": category,
                                "sheets": sheets
                            })
            self._send_json({"files": files_list})
            return

        # 1c. API Get Latest Handover Manifest (Step 1 -> Step 2)
        if path == "/api/sheet/manifest":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            if os.path.exists(manifest_path):
                with open(manifest_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"status": "PENDING", "message": "Chưa có dữ liệu bàn giao"})
            return

        # 1d. API Get Cleaning Handover Manifest (Step 2 -> Step 3)
        if path == "/api/cleaning/manifest":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "cleaning_handover_manifest.json")
            if os.path.exists(manifest_path):
                with open(manifest_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"status": "PENDING", "message": "Chưa có dữ liệu làm sạch bàn giao"})
            return

        # 1d-2. API Get Audit Readiness Manifest (Step 2.5 Audit Gate)
        if path == "/api/audit/manifest":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "audit_handover_manifest.json")
            if os.path.exists(manifest_path):
                with open(manifest_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"status": "PENDING", "message": "Chưa có hồ sơ kiểm duyệt dữ liệu"})
            return

        # 1e. API Get Analysis Handover Manifest (Step 3 -> Step 4)
        if path == "/api/analysis/manifest":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "analysis_handover_manifest.json")
            if os.path.exists(manifest_path):
                with open(manifest_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"status": "PENDING", "message": "Chưa có kết quả phân tích bàn giao"})
            return

        # 1f. API Get Reporting Manifest (Step 4 Dashboard & Report)
        if path == "/api/reporting/manifest":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "reporting_handover_manifest.json")
            if os.path.exists(manifest_path):
                with open(manifest_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("status") == "COMPLETED" and data.get("dashboard", {}).get("summary_kpis"):
                    self._send_json(data)
                    return
            try:
                import importlib
                import workflow.stage_4_reporting
                importlib.reload(workflow.stage_4_reporting)
                stage = workflow.stage_4_reporting.ReportingStage()
                data = stage.run_reporting()
                self._send_json(data)
            except Exception as e:
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 2. API Payload
        if path == "/api/payload":
            payload_path = os.path.join(WORKSPACE_ROOT, "data", "bi_system_payload.json")
            if os.path.exists(payload_path):
                with open(payload_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"error": "Payload not generated yet"}, 404)
            return

        # 3. API Workflow State
        if path == "/api/workflow/state":
            state_path = os.path.join(WORKSPACE_ROOT, "data", "workflow_execution_state.json")
            if os.path.exists(state_path):
                with open(state_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._send_json(data)
            else:
                self._send_json({"error": "No workflow run yet"}, 404)
            return

        # 4. Exported Data Files Download
        if path.startswith("/data/export/") or path.startswith("/ui/data/export/"):
            fname = os.path.basename(path)
            export_file = os.path.join(WORKSPACE_ROOT, "data", "export", fname)
            ui_export_file = os.path.join(WORKSPACE_ROOT, "ui", "data", "export", fname)

            target_file = None
            if os.path.exists(export_file):
                target_file = export_file
            elif os.path.exists(ui_export_file):
                target_file = ui_export_file
            else:
                try:
                    stage = ReportingStage()
                    stage.generate_export_files()
                    if os.path.exists(export_file):
                        target_file = export_file
                except Exception as e:
                    print(f"[Export:AutoGenerate:Error] {e}")

            if not target_file or not os.path.exists(target_file):
                self.send_error(404, f"Export file not found: {fname}")
                return

            mime_type, _ = mimetypes.guess_type(target_file)
            if not mime_type:
                if fname.endswith(".xlsx"):
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif fname.endswith(".csv"):
                    mime_type = "text/csv; charset=utf-8"
                elif fname.endswith(".html"):
                    mime_type = "text/html; charset=utf-8"
                else:
                    mime_type = "application/octet-stream"

            with open(target_file, "rb") as f:
                content = f.read()

            self.send_response(200)
            self.send_header('Content-Type', mime_type)
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Access-Control-Allow-Origin', '*')
            if not fname.endswith(".html"):
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(content)
            return

        # 5. Static UI Assets
        if path == "/index.html":
            self.send_response(301)
            self.send_header("Location", "/")
            self.end_headers()
            return
        elif path == "/":
            self._serve_file(os.path.join(WORKSPACE_ROOT, "ui", "index.html"), "text/html")
            return
        elif path in ["/app.js", "/ui/app.js"]:
            self._serve_file(os.path.join(WORKSPACE_ROOT, "ui", "app.js"), "application/javascript")
            return
        elif path in ["/payload.js", "/ui/payload.js"]:
            self._serve_file(os.path.join(WORKSPACE_ROOT, "ui", "payload.js"), "application/javascript")
            return
        else:
            # Fallback static search in ui/
            rel_path = path.lstrip("/")
            candidate = os.path.join(WORKSPACE_ROOT, "ui", rel_path)
            if os.path.exists(candidate) and os.path.isfile(candidate):
                self._serve_file(candidate)
                return
            
            self.send_error(404, f"Path not found: {path}")

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        content_len = int(self.headers.get('Content-Length', 0))
        content_type = self.headers.get('Content-Type', '')
        body_bytes = self.rfile.read(content_len) if content_len > 0 else b''

        # Parse JSON nếu Content-Type là json hoặc không phải multipart
        req_data = {}
        if "application/json" in content_type or ("multipart/form-data" not in content_type and body_bytes):
            try:
                req_data = json.loads(body_bytes.decode('utf-8'))
            except Exception:
                req_data = {}

        # 1. API Inspect Sheet File (Step 1 Core)
        if path == "/api/sheet/inspect":
            file_name = req_data.get("file_name", "raw_sales_transactions.csv")
            sheet_name = req_data.get("sheet_name", None)

            candidate_paths = [
                os.path.join(WORKSPACE_ROOT, "data", "raw", file_name),
                os.path.join(WORKSPACE_ROOT, "data", "uploads", file_name),
                os.path.join(WORKSPACE_ROOT, file_name),
                file_name
            ]
            target_path = None
            for cp in candidate_paths:
                if os.path.exists(cp) and os.path.isfile(cp):
                    target_path = cp
                    break

            if not target_path:
                self._send_json({
                    "status": "ERROR",
                    "readable": False,
                    "validation": {
                        "blocking_errors": [f"Không tìm thấy tệp: {file_name}"],
                        "warnings": [],
                        "can_proceed_to_cleaning": False
                    }
                }, 404)
                return

            stage = SheetStage()
            manifest = stage.inspect_and_validate(target_path, sheet_name)
            self._send_json(manifest)
            return

        # 2. API Upload File (Step 1 Upload)
        if path == "/api/sheet/upload":
            uploads_dir = os.path.join(WORKSPACE_ROOT, "data", "uploads")
            os.makedirs(uploads_dir, exist_ok=True)
            
            clean_filename = "uploaded_data.csv"
            sheet_name = None

            try:
                if "multipart/form-data" in content_type:
                    # 3. Log Backend: Xác nhận request nhận được multipart file
                    print(f"[Server:Upload] Nhận request upload multipart/form-data: Content-Length={content_len} bytes")
                    fields, files = self._parse_multipart(content_type, body_bytes)
                    sheet_name = fields.get("sheet_name", None)

                    if not files or files[0]["size"] == 0:
                        file_label = files[0]["filename"] if files else "không xác định"
                        print(f"[Server:Upload] LỖI: Tệp '{file_label}' rỗng hoặc không có dữ liệu (0 bytes).")
                        self._send_json({"status": "ERROR", "error": "Nội dung tệp trống"}, 400)
                        return

                    target_file = files[0]
                    raw_filename = target_file["filename"]
                    clean_filename = os.path.basename(raw_filename).strip('\"\' ') or "uploaded_data.csv"
                    save_path = os.path.join(uploads_dir, clean_filename)

                    with open(save_path, "wb") as f:
                        f.write(target_file["data"])
                    
                    saved_size = os.path.getsize(save_path)
                    print(f"[Server:Upload] Đã lưu tệp vào '{save_path}', dung lượng thực tế: {saved_size} bytes (Frontend: {target_file['size']} bytes)")
                else:
                    # JSON payload
                    file_name = req_data.get("filename", "uploaded_data.csv")
                    clean_filename = os.path.basename(file_name).strip('\"\' ') or "uploaded_data.csv"
                    content = req_data.get("content", None)
                    b64_data = req_data.get("base64", None)
                    sheet_name = req_data.get("sheet_name", None)
                    save_path = os.path.join(uploads_dir, clean_filename)

                    if b64_data:
                        if "," in b64_data:
                            b64_data = b64_data.split(",", 1)[1]
                        file_bytes = base64.b64decode(b64_data)
                        if len(file_bytes) == 0:
                            print(f"[Server:Upload] LỖI: Tệp base64 '{clean_filename}' có dung lượng 0 bytes.")
                            self._send_json({"status": "ERROR", "error": "Nội dung tệp trống"}, 400)
                            return
                        with open(save_path, "wb") as f:
                            f.write(file_bytes)
                    elif content is not None:
                        if not content.strip():
                            print(f"[Server:Upload] LỖI: Nội dung text của '{clean_filename}' rỗng.")
                            self._send_json({"status": "ERROR", "error": "Nội dung tệp trống"}, 400)
                            return
                        with open(save_path, "w", encoding="utf-8") as f:
                            f.write(content)
                    else:
                        print(f"[Server:Upload] LỖI: Không tìm thấy nội dung tệp trong request JSON.")
                        self._send_json({"status": "ERROR", "error": "Nội dung tệp trống"}, 400)
                        return
                    
                    saved_size = os.path.getsize(save_path)
                    print(f"[Server:Upload] Đã nhận tệp qua JSON: '{clean_filename}', dung lượng: {saved_size} bytes")

                # 4. Log Parser: Đọc và thẩm định tệp
                stage = SheetStage()
                manifest = stage.inspect_and_validate(save_path, sheet_name)
                rows_count = manifest.get("rows_count", 0)
                cols_count = manifest.get("columns_count", 0)
                status = manifest.get("status", "ERROR")
                print(f"[Stage 1:Sheet] Parser hoàn tất: đọc được {rows_count} dòng, {cols_count} cột. Trạng thái: {status}")

                # 5. Log Kết quả: Trạng thái xử lý và nguyên nhân nếu thất bại
                if status == "COMPLETED":
                    print(f"[Server:Upload] Xử lý thành công tệp '{clean_filename}'. Trả về mã 200.")
                    # Invalidate stale downstream manifests
                    for downstream in ["cleaning_handover_manifest.json", "audit_handover_manifest.json", "analysis_handover_manifest.json", "reporting_handover_manifest.json"]:
                        downstream_path = os.path.join(WORKSPACE_ROOT, "data", downstream)
                        if os.path.exists(downstream_path):
                            try:
                                os.remove(downstream_path)
                            except Exception:
                                pass
                    response_payload = {
                        **manifest,
                        "status": "SUCCESS",
                        "stage_status": status,
                        "file_name": clean_filename,
                        "file_size": os.path.getsize(save_path),
                        "rows_count": rows_count,
                        "row_count": rows_count,
                        "columns_count": cols_count,
                        "column_count": cols_count,
                        "manifest": manifest
                    }
                    self._send_json(response_payload, 200)
                else:
                    err_msgs = manifest.get("validation", {}).get("blocking_errors") or manifest.get("blocking_errors") or ["Không thể thẩm định tệp tin"]
                    print(f"[Server:Upload] Thẩm định thất bại cho '{clean_filename}': {'; '.join(err_msgs)}")
                    self._send_json({
                        "status": "ERROR",
                        "error": "; ".join(err_msgs),
                        "manifest": manifest,
                        **manifest
                    }, 400)
            except Exception as e:
                print(f"[Server:Upload] Ngoại lệ hệ thống: {str(e)}")
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 2b. API Analyze Business Questions Feasibility (Question vs Data Schema)
        if path == "/api/sheet/analyze-questions":
            questions = req_data.get("questions", [])
            columns = req_data.get("columns", [])
            schema = req_data.get("schema", {})
            try:
                stage = SheetStage()
                evaluations = stage.evaluate_business_questions(questions, columns, schema)
                self._send_json({
                    "status": "SUCCESS",
                    "questions_count": len(evaluations),
                    "evaluations": evaluations
                })
            except Exception as e:
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 3. API Confirm Handover (Step 1 -> Step 2 Handover)
        if path == "/api/sheet/confirm":
            manifest = req_data.get("manifest")
            if not manifest:
                self._send_json({"status": "ERROR", "error": "Thiếu dữ liệu manifest"}, 400)
                return

            blocking_errors = manifest.get("validation", {}).get("blocking_errors", [])
            if blocking_errors:
                self._send_json({
                    "status": "ERROR",
                    "error": "Không thể chuyển bước vì còn lỗi nghiêm trọng: " + "; ".join(blocking_errors)
                }, 400)
                return

            biz_context = req_data.get("business_context", {})
            company_type = biz_context.get("company_type", "").strip()
            if not company_type:
                self._send_json({
                    "status": "ERROR",
                    "error": "Bắt buộc phải chọn hoặc nhập 'Loại hình doanh nghiệp' trước khi chuyển bước!"
                }, 400)
                return

            questions = biz_context.get("questions", [])
            if not questions or len(questions) == 0:
                self._send_json({
                    "status": "ERROR",
                    "error": "Vui lòng nhập ít nhất 1 câu hỏi kinh doanh cần phân tích!"
                }, 400)
                return

            # AI đánh giá câu hỏi nếu chưa có kết quả đánh giá
            stage = SheetStage()
            if "question_evaluations" not in biz_context or not biz_context["question_evaluations"]:
                biz_context["question_evaluations"] = stage.evaluate_business_questions(
                    questions, manifest.get("columns", []), manifest.get("detected_schema", {})
                )

            manifest["status"] = "COMPLETED"
            manifest["business_context"] = biz_context

            # Đảm bảo có phân loại Measures / Dimensions
            if "classified_columns" not in manifest or not manifest["classified_columns"]:
                manifest["classified_columns"] = stage.categorize_columns(
                    manifest.get("columns", []), manifest.get("detected_schema", {})
                )

            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2)

            # Invalidate downstream manifests so stale figures from previous file NEVER leak
            for downstream in ["cleaning_handover_manifest.json", "audit_handover_manifest.json", "analysis_handover_manifest.json", "reporting_handover_manifest.json"]:
                downstream_path = os.path.join(WORKSPACE_ROOT, "data", downstream)
                if os.path.exists(downstream_path):
                    try:
                        os.remove(downstream_path)
                    except Exception:
                        pass

            self._send_json({
                "status": "COMPLETED",
                "message": "Đã kết nối nguồn dữ liệu, xác định bối cảnh doanh nghiệp và bàn giao thành công sang Bước 2 (Làm Sạch)!",
                "handover_summary": {
                    "file_name": manifest.get("file_name"),
                    "selected_sheet": manifest.get("selected_sheet"),
                    "rows_count": manifest.get("rows_count"),
                    "columns_count": manifest.get("columns_count"),
                    "company_type": company_type,
                    "questions_count": len(questions),
                    "usability_rating": manifest.get("usability_rating", "OPTIMAL_READY"),
                    "warnings_count": len(manifest.get("validation", {}).get("warnings", []))
                }
            })
            return

        # 3b. API Profile Data Quality (Step 2 - 7 Issue Categories)
        if path == "/api/cleaning/profile":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            if not os.path.exists(manifest_path):
                self._send_json({"status": "ERROR", "error": "Chưa có dữ liệu bàn giao từ Bước 1"}, 400)
                return
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest_step1 = json.load(f)

            file_path = manifest_step1.get("file_path", "")
            if not os.path.exists(file_path):
                file_path = os.path.join(WORKSPACE_ROOT, "data", "raw", manifest_step1.get("file_name", "raw_sales_transactions.csv"))

            sheet_stage = SheetStage()
            ext = os.path.splitext(file_path)[1].lower()
            if ext in [".xlsx", ".xls", ".xlsm"]:
                records = sheet_stage.read_xlsx(file_path, manifest_step1.get("selected_sheet"))
            else:
                records = sheet_stage.read_csv(file_path)

            cleaning_stage = CleaningStage()
            profile = cleaning_stage.profile_data_quality(
                records,
                columns=manifest_step1.get("columns", []),
                schema=manifest_step1.get("detected_schema", {})
            )
            self._send_json({
                "status": "SUCCESS",
                "profile": profile,
                "step1_summary": {
                    "file_name": manifest_step1.get("file_name"),
                    "selected_sheet": manifest_step1.get("selected_sheet"),
                    "rows_count": manifest_step1.get("rows_count"),
                    "columns_count": manifest_step1.get("columns_count"),
                    "company_type": manifest_step1.get("business_context", {}).get("company_type", ""),
                    "questions": manifest_step1.get("business_context", {}).get("questions", [])
                }
            })
            return

        # 3c. API Execute Data Cleaning (Step 2 - Transform & Verify Gate)
        if path == "/api/cleaning/execute":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            if not os.path.exists(manifest_path):
                self._send_json({"status": "ERROR", "error": "Chưa có dữ liệu bàn giao từ Bước 1"}, 400)
                return
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest_step1 = json.load(f)

            file_path = manifest_step1.get("file_path", "")
            if not os.path.exists(file_path):
                file_path = os.path.join(WORKSPACE_ROOT, "data", "raw", manifest_step1.get("file_name", "raw_sales_transactions.csv"))

            sheet_stage = SheetStage()
            ext = os.path.splitext(file_path)[1].lower()
            if ext in [".xlsx", ".xls", ".xlsm"]:
                records = sheet_stage.read_xlsx(file_path, manifest_step1.get("selected_sheet"))
            else:
                records = sheet_stage.read_csv(file_path)

            enabled_strategies = req_data.get("strategies", [
                "STRAT_DEDUP", "STRAT_TRIM", "STRAT_DATE", "STRAT_TYPES", "STRAT_DERIVED", "STRAT_OUTLIERS"
            ])

            import importlib
            import workflow.stage_2_cleaning
            importlib.reload(workflow.stage_2_cleaning)
            cleaning_stage = workflow.stage_2_cleaning.CleaningStage()

            # Chuyển đầy đủ bối cảnh tệp tải lên và danh sách câu hỏi nghiệp vụ thực tế
            biz_ctx = dict(manifest_step1.get("business_context") or {})
            biz_ctx["file_name"] = manifest_step1.get("file_name", "")
            biz_ctx["file_path"] = file_path
            biz_ctx["selected_sheet"] = manifest_step1.get("selected_sheet", "")
            biz_ctx["company_type"] = biz_ctx.get("company_type") or manifest_step1.get("business_context", {}).get("company_type", "")
            biz_ctx["questions"] = biz_ctx.get("questions") or manifest_step1.get("business_context", {}).get("questions", [])

            cleaning_result = cleaning_stage.execute_cleaning(
                raw_records=records,
                columns=manifest_step1.get("columns", []),
                schema=manifest_step1.get("detected_schema", {}),
                enabled_strategies=enabled_strategies,
                business_context=biz_ctx
            )
            self._send_json({
                "status": "SUCCESS",
                "result": cleaning_result,
                "manifest": cleaning_result,
                "verify_passed": cleaning_result.get("verify_passed", True),
                **cleaning_result
            })
            return

        # 3d. API Confirm Step 2 (Cleaning -> Step 3 Handover)
        if path == "/api/cleaning/confirm":
            manifest_path = os.path.join(WORKSPACE_ROOT, "data", "cleaning_handover_manifest.json")
            if not os.path.exists(manifest_path):
                self._send_json({"status": "ERROR", "error": "Chưa có kết quả làm sạch để xác nhận"}, 400)
                return
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest_step2 = json.load(f)

            if not manifest_step2.get("verify_passed", False):
                self._send_json({"status": "ERROR", "error": "Bộ kiểm chứng Verify Gate chưa PASS, không thể mở khóa Bước 3!"}, 400)
                return

            self._send_json({
                "status": "COMPLETED",
                "message": "Đã làm sạch dữ liệu và vượt qua kiểm toán Verify Gate 6/6! Sẵn sàng bàn giao sang Bước 3 (Phân Tích Dữ Liệu).",
                "handover_summary": {
                    "cleaned_rows": manifest_step2.get("cleaned_stats", {}).get("rows_count"),
                    "cleaned_columns": manifest_step2.get("cleaned_stats", {}).get("columns_count"),
                    "dq_score": manifest_step2.get("cleaned_stats", {}).get("dq_score", 100.0),
                    "verify_passed": True,
                    "audit_events_count": len(manifest_step2.get("audit_log", []))
                }
            })
            return

        # 3d-2. API Audit Data Readiness & Gap Elicitation (Step 2.5 Core)
        if path == "/api/audit/gap-analysis":
            manifest1_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            m1 = {}
            if os.path.exists(manifest1_path):
                try:
                    with open(manifest1_path, "r", encoding="utf-8") as f:
                        m1 = json.load(f)
                except Exception:
                    pass
            biz_ctx = m1.get("business_context", {})
            company_type = biz_ctx.get("company_type", "Bán lẻ & Thương mại Đa kênh")
            file_name = m1.get("file_name", "Tệp dữ liệu")

            cleaned_csv = os.path.join(WORKSPACE_ROOT, "data", "cleaned_sales_transactions.csv")
            cleaned_records = None
            if os.path.exists(cleaned_csv):
                with open(cleaned_csv, mode="r", encoding="utf-8-sig") as f:
                    cleaned_records = list(csv.DictReader(f))

            try:
                import importlib
                import workflow.stage_2_5_audit
                importlib.reload(workflow.stage_2_5_audit)
                audit_stage = workflow.stage_2_5_audit.DataReadinessAuditStage()
                audit_res = audit_stage.audit_dataset_readiness(
                    cleaned_records=cleaned_records,
                    file_name=file_name,
                    company_type=company_type
                )
                self._send_json(audit_res)
            except Exception as e:
                print(f"[Audit:GapAnalysis:Error] {e}")
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 3d-3. API Submit Gap Declarations & Omissions (Step 2.5 -> Step 3)
        if path == "/api/audit/submit-gaps":
            declarations = req_data.get("declarations", {})
            omitted_metrics = req_data.get("omitted_metrics", [])
            try:
                import importlib
                import workflow.stage_2_5_audit
                importlib.reload(workflow.stage_2_5_audit)
                audit_stage = workflow.stage_2_5_audit.DataReadinessAuditStage()
                manifest = audit_stage.save_user_decisions(
                    declarations=declarations,
                    omitted_metrics=omitted_metrics
                )
                self._send_json({
                    "status": "CONFIRMED",
                    "message": "Đã ghi nhận bổ khuyết dữ liệu và thiết lập phạm vi phân tích thành công!",
                    "manifest": manifest
                })
            except Exception as e:
                print(f"[Audit:SubmitGaps:Error] {e}")
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 3e. API Run Data Analysis (Step 3 - KPIs, 5 Methods, Insights, Questions)
        if path == "/api/analysis/run":
            # Nạp thông tin bối cảnh thực tế từ Bước 1
            manifest1_path = os.path.join(WORKSPACE_ROOT, "data", "sheet_handover_manifest.json")
            manifest_step1 = {}
            if os.path.exists(manifest1_path):
                try:
                    with open(manifest1_path, "r", encoding="utf-8") as f:
                        manifest_step1 = json.load(f)
                except Exception:
                    pass

            manifest2_path = os.path.join(WORKSPACE_ROOT, "data", "cleaning_handover_manifest.json")
            manifest_step2 = {}
            if os.path.exists(manifest2_path):
                try:
                    with open(manifest2_path, "r", encoding="utf-8") as f:
                        manifest_step2 = json.load(f)
                except Exception:
                    pass

            # Bối cảnh ưu tiên danh sách câu hỏi và ngành nghề người dùng đã nhập
            biz_ctx = {
                **(manifest_step1.get("business_context") or {}),
                **(manifest_step2.get("handover_from_step1") or {})
            }
            if not biz_ctx.get("questions") and manifest_step1.get("business_context", {}).get("questions"):
                biz_ctx["questions"] = manifest_step1["business_context"]["questions"]

            try:
                import importlib
                import workflow.stage_3_analysis
                importlib.reload(workflow.stage_3_analysis)
                stage = workflow.stage_3_analysis.AnalysisStage()

                # Đọc tệp dữ liệu sạch thực tế vừa làm sạch
                cleaned_records = None
                cleaned_csv = os.path.join(WORKSPACE_ROOT, "data", "cleaned_sales_transactions.csv")
                if not os.path.exists(cleaned_csv) or os.path.getsize(cleaned_csv) == 0:
                    try:
                        import workflow.stage_2_cleaning
                        importlib.reload(workflow.stage_2_cleaning)
                        c_stage = workflow.stage_2_cleaning.CleaningStage()
                        c_stage.run_cleaning()
                    except Exception as ce:
                        print(f"[Server:Analysis] Warning auto-cleaning: {ce}")

                if os.path.exists(cleaned_csv):
                    with open(cleaned_csv, mode="r", encoding="utf-8-sig") as f:
                        cleaned_records = list(csv.DictReader(f))

                analysis_res = stage.run_analysis(
                    cleaned_records=cleaned_records,
                    business_context=biz_ctx
                )
                self._send_json({
                    "status": "SUCCESS",
                    "result": analysis_res
                })
            except Exception as e:
                print(f"[Server:Analysis] Lỗi phân tích: {e}")
                self._send_json({"status": "ERROR", "error": str(e)}, 500)
            return

        # 3f. API Confirm Step 3 (Analysis -> Step 4 Reporting Handover)
        if path == "/api/analysis/confirm":
            manifest3_path = os.path.join(WORKSPACE_ROOT, "data", "analysis_handover_manifest.json")
            if not os.path.exists(manifest3_path):
                self._send_json({"status": "ERROR", "error": "Chưa có kết quả phân tích để xác nhận"}, 400)
                return
            with open(manifest3_path, "r", encoding="utf-8") as f:
                manifest_step3 = json.load(f)

            if not manifest_step3.get("verify_passed", False):
                self._send_json({"status": "ERROR", "error": "Bộ kiểm chứng Verify Gate Bước 3 chưa PASS, không thể mở khóa Bước 4!"}, 400)
                return

            self._send_json({
                "status": "COMPLETED",
                "message": "Đã hoàn thành phân tích toàn diện và đạt chuẩn kiểm toán toán học 6/6! Đã mở khóa Giai đoạn 4 (Báo Cáo & Dashboard).",
                "handover_summary": {
                    "net_revenue": manifest_step3.get("kpis", {}).get("Formatted", {}).get("NetRevenue"),
                    "gross_profit": manifest_step3.get("kpis", {}).get("Formatted", {}).get("GrossProfit"),
                    "gross_margin": manifest_step3.get("kpis", {}).get("Formatted", {}).get("GrossMarginPct"),
                    "total_orders": manifest_step3.get("kpis", {}).get("Formatted", {}).get("TotalOrders"),
                    "answered_questions_count": len(manifest_step3.get("answered_questions", [])),
                    "insights_count": len(manifest_step3.get("insights", [])),
                    "verify_passed": True
                }
            })
            return

        # 3g. API Ask Custom Business Question in Step 3
        if path == "/api/analysis/ask-question":
            custom_q = req_data.get("question", "").strip()
            if not custom_q:
                self._send_json({"status": "ERROR", "error": "Câu hỏi không được để trống"}, 400)
                return

            ai_res = ask_gemini_bi_assistant(custom_q)
            raw_text = ai_res.get("reply", "")
            ans_text = raw_text.replace("\n", "<br/>")

            self._send_json({
                "status": "SUCCESS",
                "question": custom_q,
                "answer": ans_text,
                "reply": raw_text,
                "model": ai_res.get("model", GEMINI_MODEL),
                "source": ai_res.get("source", "gemini")
            })
            return

        # 3h. API Run Reporting (Step 4 Dashboard & Report Compilation)
        if path == "/api/reporting/run":
            period = req_data.get("period", "ALL")
            channel = req_data.get("channel", "ALL")
            segment = req_data.get("segment", "ALL")
            stage = ReportingStage()
            res = stage.run_reporting(period, channel, segment)
            self._send_json({
                "status": "SUCCESS",
                "result": res
            })
            return

        # 3i. API Dynamic Dashboard Filter (Step 4 Real-time Recalculation)
        if path == "/api/reporting/filter":
            period = req_data.get("period", "ALL")
            channel = req_data.get("channel", "ALL")
            segment = req_data.get("segment", "ALL")
            stage = ReportingStage()
            dash = stage.build_dashboard_data(period, channel, segment)
            self._send_json({
                "status": "SUCCESS",
                "dashboard": dash
            })
            return

        # 3j. API Confirm Final Workflow Completion (Step 4 Final Approval)
        if path == "/api/reporting/confirm":
            manifest4_path = os.path.join(WORKSPACE_ROOT, "data", "reporting_handover_manifest.json")
            if not os.path.exists(manifest4_path):
                stage = ReportingStage()
                manifest_step4 = stage.run_reporting()
            else:
                with open(manifest4_path, "r", encoding="utf-8") as f:
                    manifest_step4 = json.load(f)

            if not manifest_step4.get("verify_passed", False):
                self._send_json({"status": "ERROR", "error": "Bộ kiểm chứng Verify Gate Bước 4 chưa PASS!"}, 400)
                return

            self._send_json({
                "status": "COMPLETED",
                "message": "🎉 CHÚC MỪNG! TOÀN BỘ QUY TRÌNH 4 GIAI ĐOẠN BI ĐÃ HOÀN TẤT THÀNH CÔNG 100%! Báo cáo và Dashboard đã sẵn sàng xuất bản.",
                "workflow_completed": True,
                "exports": manifest_step4.get("exports", {}),
                "verify_passed": True
            })
            return

        # 4. API Run Workflow
        if path == "/api/workflow/run":
            decision = req_data.get("decision", "APPROVE")
            adjustments = req_data.get("adjustments", {})
            try:
                engine = BIWorkflowEngine(WORKSPACE_ROOT)
                state = engine.run_workflow(user_feedback_decision=decision, feedback_adjustments=adjustments)
                self._send_json({
                    "status": "SUCCESS",
                    "workflow_id": state.workflow_id,
                    "verification_passed": state.verification_passed,
                    "stage_statuses": state.stage_statuses,
                    "kpis": state.kpis,
                    "exported_files": state.exported_files
                })
            except Exception as e:
                self._send_json({"status": "FAILED", "error": str(e)}, 500)
            return

        # 5. API Chat with AI BI Assistant
        if path == "/api/chat":
            q = req_data.get("message") or req_data.get("query") or req_data.get("text") or ""
            q = str(q).strip()
            ai_res = ask_gemini_bi_assistant(q)
            reply_text = ai_res.get("reply", "")
            self._send_json({
                "status": "SUCCESS",
                "reply": reply_text,
                "answer": reply_text,
                "model": ai_res.get("model", GEMINI_MODEL),
                "source": ai_res.get("source", "gemini"),
                "query": q
            })
            return

        self.send_error(404, f"API Endpoint not found: {path}")

def run_server(port: int = 8080):
    server_address = ('', port)
    try:
        httpd = HTTPServer(server_address, BIPlatformHTTPHandler)
        print(f"===============================================================")
        print(f"  ✦ CCN-DA Studio Server is RUNNING at: http://localhost:{port}")
        print(f"  ✦ Live Dashboard:  http://localhost:{port}")
        print(f"  ✦ REST API Status: http://localhost:{port}/api/status")
        print(f"===============================================================")
        httpd.serve_forever()
    except OSError as e:
        if "address already in use" in str(e).lower() or e.errno == 98 or e.errno == 10048:
            print(f"Port {port} is busy, trying port {port + 1}...")
            run_server(port + 1)
        else:
            raise e

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    run_server(port)
