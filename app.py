"""
사고조사보고서 입력/발송 앱
- 폼 입력 → docx 보고서 생성 → 네이버 SMTP로 이메일 발송
- docx + 위치도 + 현장사진(최대 4개) 모두 첨부
"""
import os
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from flask import Flask, render_template, request, jsonify

from report_generator import create_accident_report

# ─────────────────────────────────────────────────────────────
# 환경설정
# ─────────────────────────────────────────────────────────────
NAVER_USER = os.getenv("NAVER_USER", "no-reply-ansansafe@naver.com")
NAVER_PASS = os.getenv("NAVER_PASS", "")

SMTP_HOST = "smtp.naver.com"
SMTP_PORT = 465

ALLOWED_EXT = {"jpg", "jpeg", "png"}
MAX_FILE_MB = 5
MAX_TOTAL_MB = 25

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB


# ─────────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────────
def get_ext(filename: str) -> str:
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""


def allowed_file(filename: str) -> bool:
    return get_ext(filename) in ALLOWED_EXT


def save_upload(file_storage, dest_path: str) -> int:
    """업로드 파일을 디스크에 저장 후 바이트 수 반환. 검증 실패 시 ValueError."""
    if not allowed_file(file_storage.filename):
        raise ValueError(f"허용되지 않은 파일 형식: {file_storage.filename}")

    file_storage.save(dest_path)
    size_mb = os.path.getsize(dest_path) / (1024 * 1024)
    if size_mb > MAX_FILE_MB:
        os.remove(dest_path)
        raise ValueError(f"파일이 너무 큽니다({size_mb:.1f}MB): {file_storage.filename}")
    return os.path.getsize(dest_path)


def attach_path(msg: MIMEMultipart, file_path: str, display_name: str, mime_main="application", mime_sub="octet-stream"):
    """디스크 경로의 파일을 메일에 첨부."""
    with open(file_path, "rb") as f:
        part = MIMEBase(mime_main, mime_sub)
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        "attachment",
        filename=("utf-8", "", display_name),
    )
    msg.attach(part)


def build_html_body(data: dict) -> str:
    """이메일 본문(HTML) — 표 형식 요약"""
    actions_html = "".join(
        f"<tr><td style='padding:6px 10px;border:1px solid #ddd;width:80px'>{a['time']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd'>{a['content']}</td></tr>"
        for a in data["actions"]
    )
    if not actions_html:
        actions_html = "<tr><td colspan='2' style='padding:6px 10px;border:1px solid #ddd;color:#888'>입력 없음</td></tr>"

    return f"""
    <html><body style="font-family:'맑은 고딕',sans-serif;color:#222;line-height:1.6">
      <h2 style="border-bottom:2px solid #333;padding-bottom:6px">사고조사보고서</h2>
      <p style="color:#555;font-size:13px">상세 양식은 첨부 docx 파일을 확인하세요.</p>
      <table style="border-collapse:collapse;width:100%;margin-top:10px">
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;width:140px;text-align:left">사고 일시</th>
            <td style="padding:8px;border:1px solid #ddd">{data['accident_datetime']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">사고 장소</th>
            <td style="padding:8px;border:1px solid #ddd">{data['accident_place']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">단지번호</th>
            <td style="padding:8px;border:1px solid #ddd">{data['complex_no']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">인명피해</th>
            <td style="padding:8px;border:1px solid #ddd">{data['damage_person']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">재산피해</th>
            <td style="padding:8px;border:1px solid #ddd">{data['damage_property']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">출동자</th>
            <td style="padding:8px;border:1px solid #ddd">{data['responders']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">원인</th>
            <td style="padding:8px;border:1px solid #ddd">{data['cause']}</td></tr>
        <tr><th style="background:#f4f4f4;padding:8px;border:1px solid #ddd;text-align:left">조치결과</th>
            <td style="padding:8px;border:1px solid #ddd">{data['measures']}</td></tr>
      </table>

      <h3 style="margin-top:24px">조치사항</h3>
      <table style="border-collapse:collapse;width:100%">
        <tr><th style="background:#f4f4f4;padding:6px 10px;border:1px solid #ddd;width:80px">시간</th>
            <th style="background:#f4f4f4;padding:6px 10px;border:1px solid #ddd;text-align:left">내용</th></tr>
        {actions_html}
      </table>

      <p style="margin-top:24px;font-size:12px;color:#888">
        본 메일은 사고조사보고 시스템에서 자동 발송되었습니다.<br/>
        제출 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
      </p>
    </body></html>
    """


# ─────────────────────────────────────────────────────────────
# 라우트
# ─────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    """배포 환경 헬스체크용"""
    return {"status": "ok"}, 200


@app.route("/submit", methods=["POST"])
def submit():
    if not NAVER_PASS:
        return jsonify({"success": False, "message": "서버에 NAVER_PASS 환경변수가 설정되지 않았습니다."}), 500

    try:
        # ── 사고일시 조립 ──
        y = request.form.get("year", "").strip()
        mo = request.form.get("month", "").strip().zfill(2)
        d = request.form.get("day", "").strip().zfill(2)
        h = request.form.get("hour", "").strip().zfill(2)
        mi = request.form.get("minute", "").strip().zfill(2)
        accident_dt = f"{y}-{mo}-{d} {h}:{mi}"

        # ── 일반 필드 ──
        place = request.form.get("place", "").strip()
        complex_no = request.form.get("complex_no", "").strip()
        damage_person = request.form.get("damage_person", "없음").strip()
        damage_property = request.form.get("damage_property", "없음").strip()
        responders = request.form.get("responders", "").strip()
        cause = request.form.get("cause", "").strip()
        measures = request.form.get("measures", "").strip()
        recipient = request.form.get("recipient", "").strip()

        if not recipient or "@" not in recipient:
            return jsonify({"success": False, "message": "올바른 이메일 주소를 입력하세요."}), 400
        if not place or not complex_no:
            return jsonify({"success": False, "message": "사고 장소와 단지번호는 필수입니다."}), 400

        # ── 조치사항 ──
        a_hours = request.form.getlist("action_hour[]")
        a_minutes = request.form.getlist("action_minute[]")
        a_contents = request.form.getlist("action_content[]")
        actions = []
        for hh, mm, cc in zip(a_hours, a_minutes, a_contents):
            cc = cc.strip()
            if cc:
                actions.append({"time": f"{hh.zfill(2)}:{mm.zfill(2)}", "content": cc})

        # ── 임시 폴더에서 모든 파일 작업 (자동 정리) ──
        with tempfile.TemporaryDirectory(prefix="accident_report_") as tmpdir:

            # 위치도 저장
            location_img_path = None
            loc_file = request.files.get("location_img")
            if loc_file and loc_file.filename:
                location_img_path = os.path.join(tmpdir, f"location.{get_ext(loc_file.filename)}")
                save_upload(loc_file, location_img_path)

            # 현장사진 저장 (최대 4개)
            site_img_paths = []
            site_files = request.files.getlist("site_imgs[]")[:4]
            for i, f in enumerate(site_files, start=1):
                if f and f.filename:
                    p = os.path.join(tmpdir, f"site_{i}.{get_ext(f.filename)}")
                    save_upload(f, p)
                    site_img_paths.append(p)

            # 합계 용량 검증
            total_mb = sum(
                os.path.getsize(p) for p in [location_img_path, *site_img_paths] if p
            ) / (1024 * 1024)
            if total_mb > MAX_TOTAL_MB:
                return jsonify({"success": False, "message": f"첨부 합계가 {MAX_TOTAL_MB}MB를 초과합니다."}), 400

            # ── docx 보고서 생성 ──
            docx_data = {
                "accident_datetime": accident_dt,
                "accident_place": place,
                "complex_no": complex_no,
                "damage_person": damage_person,
                "damage_property": damage_property,
                "responders": responders,
                "actions": actions,
                "cause": cause,
                "measures": measures,
                "location_img": location_img_path,
                "site_imgs": site_img_paths,
            }
            dt_tag = accident_dt.replace(":", "").replace("-", "").replace(" ", "_")
            docx_filename = f"사고조사보고서_{dt_tag}.docx"
            docx_path = os.path.join(tmpdir, docx_filename)
            create_accident_report(docx_data, docx_path)

            # ── 이메일 본문 ──
            body_data = {
                "accident_datetime": accident_dt,
                "accident_place": place,
                "complex_no": complex_no,
                "damage_person": damage_person,
                "damage_property": damage_property,
                "responders": responders,
                "cause": cause,
                "measures": measures,
                "actions": actions,
            }

            msg = MIMEMultipart()
            msg["From"] = NAVER_USER
            msg["To"] = recipient
            msg["Subject"] = f"[사고조사보고서] {accident_dt} {place}"
            msg.attach(MIMEText(build_html_body(body_data), "html", "utf-8"))

            # ── 첨부: docx (먼저) ──
            attach_path(
                msg, docx_path, docx_filename,
                mime_main="application",
                mime_sub="vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            # ── 첨부: 위치도 ──
            if location_img_path:
                ext = get_ext(location_img_path)
                attach_path(msg, location_img_path, f"위치도.{ext}",
                            mime_main="image", mime_sub=ext)

            # ── 첨부: 현장사진 ──
            for i, p in enumerate(site_img_paths, start=1):
                ext = get_ext(p)
                attach_path(msg, p, f"현장사진{i}.{ext}",
                            mime_main="image", mime_sub=ext)

            # ── SMTP 발송 ──
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15) as server:
                server.login(NAVER_USER, NAVER_PASS)
                server.send_message(msg)

        return jsonify({"success": True, "message": f"{recipient} 로 발송 완료 (보고서 docx 포함)"})

    except ValueError as ve:
        return jsonify({"success": False, "message": str(ve)}), 400
    except smtplib.SMTPAuthenticationError:
        return jsonify({"success": False, "message": "네이버 SMTP 인증 실패. IMAP/SMTP 사용 설정과 비밀번호를 확인하세요."}), 500
    except Exception as e:
        return jsonify({"success": False, "message": f"발송 중 오류: {e}"}), 500


if __name__ == "__main__":
    # 로컬 테스트 시에만 실행 (Render 등 배포 환경에서는 gunicorn이 처리)
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
