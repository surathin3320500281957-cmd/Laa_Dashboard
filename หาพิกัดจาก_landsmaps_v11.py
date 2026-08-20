# -*- coding: utf-8 -*-
"""
สคริปต์ดึงพิกัดแปลงที่ดินจาก landsmaps.dol.go.th โดยอัตโนมัติ (เวอร์ชัน v11)
ใช้ Playwright ควบคุมเบราว์เซอร์ กรอกจังหวัด/อำเภอ/เลขที่โฉนด แล้วอ่านพิกัดที่เว็บแสดงผล

ใหม่ในเวอร์ชันนี้ (ทดสอบว่า rate limit ผูกกับ session/คุกกี้ หรือผูกกับ IP):
0h. แทนที่จะแค่ "พักรอเฉยๆ" ทุก 20 แปลงแบบ v10 → เปลี่ยนเป็น "ปิดเบราว์เซอร์แล้วเปิดใหม่
    ทั้งหมด" (เซสชัน/คุกกี้ใหม่เอี่ยม) ทุก 20 แปลง ผ่านฟังก์ชัน open_fresh_session()
    - ถ้า rate limit เป็นแบบ session-based → วิธีนี้รีเซ็ตตัวนับได้ทันที เร็วกว่า v10 มาก
    - ถ้าเป็นแบบ IP-based → วิธีนี้ช่วยไม่ได้ ต้องพึ่ง DELAY_SECONDS (ยังคงไว้เป็น fallback)
0h. ลดเวลาพักจาก 150 วิ (v10) เหลือ 30 วิ ก่อนเปิดเซสชันใหม่ (BREAK_PAUSE_SECONDS)
    เพราะเน้นพึ่งการรีเซ็ตเซสชันเป็นหลัก ไม่ใช่พึ่งเวลารอเฉยๆ

พื้นหลัง: สังเกตพบว่า v8/v9 ต่างเวอร์ชันกันแต่เจอ "found รัวๆ ~24 ครั้งแรก แล้วหยุดสนิท
ตลอดที่เหลือ (เว็บแจ้ง 'ไม่พบ' ตรงๆ ไม่มี popup บัง)" เหมือนกันทั้งคู่ — ชี้ว่าสาเหตุอยู่ที่
เว็บตรวจจับรูปแบบการเข้าถึงแบบบอทแล้วเริ่มตอบผลลอมๆ ไม่ใช่บั๊กจาก popup ที่เคยแก้ไปก่อนหน้า

ของเดิมจาก v10 (เพิ่มดีเลย์ + พัก แต่ยังไม่ได้ทดสอบว่ารีสตาร์ตเซสชันช่วยไหม):
0g. เพิ่มเวลาหน่วงระหว่างคำขอจาก 4 วิ เป็น 18 วิ (DELAY_SECONDS)
0g. เพิ่มการพักยาว 2.5 นาที ทุกๆ 20 แปลงที่ค้น (BREAK_EVERY_N / BREAK_DURATION_SECONDS)
    เพื่อลดโอกาสโดนระบบตรวจจับรูปแบบการเข้าถึงแบบบอท

⚠️ ผลคือใช้เวลารันนานขึ้นมาก — ประมาณ 6-7 ชั่วโมงสำหรับ 621 แปลง (เดิม ~2 ชม.)
   แนะนำรันข้ามคืน หรือแบ่งรันหลายรอบ (ปิดคอมได้ระหว่างพัก แต่ต้องรอจนจบ 1 รอบการรันจริงๆ
   ก่อนปิด ไม่งั้นจะสะดุดกลางทาง — ใช้ระบบ resume รันต่อได้ในรอบถัดไป)

ของเดิมจาก v9 (แก้ปัญหา popup "ข่าวประกาศ" เด้งขึ้นมาซ้ำระหว่างรัน):
0f. เดิม (v5-v8) ปิด #modal_news แค่ครั้งเดียวตอนเปิดเว็บครั้งแรก แต่พบว่า popup นี้
    เด้งขึ้นมาใหม่ได้อีกกลางทางระหว่างรัน (พบตอนรันถึงแปลงที่ ~426) ทำให้แปลงหลังจากนั้น
    ทั้งหมดกลายเป็น timeout ทั้งที่ไม่ใช่เพราะหาไม่เจอจริง — เพิ่มการเช็ค/ปิด popup นี้ซ้ำ
    ทั้งก่อนเริ่มค้นหาทุกแปลง และระหว่างรอผลลัพธ์ในทุกรอบของ polling loop

ของเดิมจาก v8 (แก้ปัญหา popup SweetAlert2 ปิดไม่สำเร็จ วนซ้ำไม่รู้จบ):
0e. ใช้ selector เจาะจง button.swal2-confirm แทนการหาจากข้อความ "รับทราบ"
0e. บังคับคลิกทะลุ (force=True) กันปัญหาโดน overlay ของ SweetAlert2 บัง
0e. จำกัดจำนวนครั้งที่พยายามปิด (สูงสุด 6 ครั้ง) ถ้าเกินนี้แสดงว่ามีปัญหาอื่น จะหยุดลองปิดแล้วรอเฉยๆ

ของเดิมจาก v7 (คอยเช็คซ้ำทุกครึ่งวินาทีระหว่างรอผลลัพธ์):
0b. ปิด #modal_news ด้วยปุ่ม button[data-dismiss="modal"] ที่แท้จริง (แทนการเดา selector หลายแบบ)

ของเดิมจาก v4 (ปิดกั้น popup ตำแหน่งที่ตั้ง):
0. ปิดกั้น popup ขอสิทธิ์ตำแหน่งที่ตั้ง (geolocation) ของเบราว์เซอร์ไม่ให้เด้งขึ้นมาเลย

ของเดิมจาก v3 (แก้ปัญหาการดัก error กรณีไม่พบข้อมูล):
1. ดัก JavaScript alert/confirm/prompt ที่เว็บอาจเด้งขึ้นมาตอนค้นไม่เจอ — กดปิดให้อัตโนมัติ
   (ถ้าไม่ดักไว้ สคริปต์จะค้างรอตลอดกาลตอนเจอ popup แบบนี้)
2. แยกสถานะผลลัพธ์ชัดเจนขึ้น:
   - found        = พบพิกัดสำเร็จ
   - not_found    = เว็บแจ้งชัดเจนว่าไม่พบข้อมูล (ผ่าน popup alert)
   - timeout      = รอนานเกินกำหนดแต่ไม่มีทั้งพิกัดและไม่มี popup แจ้งเตือน (เว็บอาจช้า/ค้าง)
   - error        = เกิดข้อผิดพลาดทางเทคนิคอื่นๆ (เช่น หา element ไม่เจอ, จังหวัด/อำเภอผิด)

วิธีใช้: เหมือน v2 ทุกประการ (อ่านจาก input.csv, บันทึกทีละแถว, resume ได้)
"""

import csv
import os
import time
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

INPUT_FILE = "input.csv"
OUTPUT_FILE = "ผลลัพธ์พิกัด.csv"
DELAY_SECONDS = 18         # หน่วงเวลานานขึ้นมาก (จาก 4 วิ เป็น 18 วิ) กันโดน rate limit
BREAK_EVERY_N = 20         # ปิด-เปิดเบราว์เซอร์ใหม่ (เซสชันใหม่) ทุกๆ กี่แปลง
BREAK_PAUSE_SECONDS = 30   # พักสั้นๆ ก่อนเปิดเซสชันใหม่ (ให้เว็บ "ลืม" เซสชันเก่าจริงๆ)
WAIT_AFTER_PROVINCE = 1500
RESULT_TIMEOUT = 10000
FIELDNAMES = ["province", "district", "deed_no", "lat", "lng", "status", "note"]


def load_input_records():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"ไม่พบไฟล์ {INPUT_FILE} — กรุณาสร้างไฟล์นี้ในโฟลเดอร์เดียวกับสคริปต์ "
            f"มีคอลัมน์: province, district, deed_no"
        )
    with open(INPUT_FILE, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return [row for row in reader]


def load_already_done():
    done_keys = set()
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                key = (row["province"], row["district"], row["deed_no"])
                done_keys.add(key)
    return done_keys


def ensure_output_header():
    if not os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()


def append_result(row):
    with open(OUTPUT_FILE, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writerow(row)


def select_district_fuzzy(page, district_name):
    options = page.eval_on_selector_all(
        "#cbamphur option", "els => els.map(e => e.textContent.trim())"
    )
    target = district_name.strip()

    for opt in options:
        if opt == target:
            page.select_option("#cbamphur", label=opt)
            return True

    for opt in options:
        opt_without_code = opt.split("-", 1)[-1].strip() if "-" in opt else opt
        if opt_without_code == target or opt.endswith(target):
            page.select_option("#cbamphur", label=opt)
            return True

    return False


def open_fresh_session(p):
    """เปิดเบราว์เซอร์ + context ใหม่ทั้งหมด (คุกกี้/เซสชันใหม่เอี่ยม) แล้วเข้าเว็บ+ปิด popup
    ข่าวประกาศให้พร้อมใช้งาน — เรียกทั้งตอนเริ่มรันครั้งแรก และเรียกซ้ำทุกๆ BREAK_EVERY_N แปลง
    เพื่อทดสอบว่า rate limit ของเว็บผูกกับ session/คุกกี้หรือไม่ (ถ้าผูกกับ IP อย่างเดียว
    วิธีนี้จะไม่ช่วย ต้องพึ่ง DELAY_SECONDS เป็นหลัก)"""
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(
        permissions=["geolocation"],
        geolocation={"latitude": 13.7563, "longitude": 100.5018},
    )
    page = context.new_page()

    dialog_state = {"message": None}

    def handle_dialog(dialog):
        dialog_state["message"] = dialog.message
        print(f"   💬 เว็บเด้งข้อความ: {dialog.message}")
        dialog.dismiss()

    page.on("dialog", handle_dialog)

    print("   กำลังเปิดเว็บ landsmaps.dol.go.th (เซสชันใหม่) ...")
    page.goto("https://landsmaps.dol.go.th/", timeout=60000)
    page.wait_for_timeout(3000)

    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
    except Exception:
        pass

    try:
        news_modal = page.locator("#modal_news")
        if news_modal.is_visible(timeout=3000):
            close_btn = page.locator('#modal_news button[data-dismiss="modal"]')
            close_btn.click(timeout=3000)
            print("   ปิด popup ข่าวประกาศสำเร็จ")
            page.wait_for_timeout(500)
        else:
            print("   ไม่มี popup ข่าวประกาศขึ้นมารอบนี้")
    except Exception as e:
        print(f"   (ไม่พบ popup ข่าวประกาศ หรือปิดไม่สำเร็จ: {e})")

    return browser, context, page, dialog_state


def run():
    all_records = load_input_records()
    already_done = load_already_done()
    ensure_output_header()

    pending = [
        r for r in all_records
        if (r["province"], r["district"], r["deed_no"]) not in already_done
    ]

    print(f"รายการทั้งหมด: {len(all_records)} | ทำไปแล้ว: {len(already_done)} | เหลือ: {len(pending)}")
    if not pending:
        print("ทำครบทุกรายการแล้วครับ ไม่ต้องรันต่อ")
        return

    with sync_playwright() as p:
        browser, context, page, dialog_state = open_fresh_session(p)

        total = len(pending)
        for i, rec in enumerate(pending):
            province = rec["province"]
            district = rec["district"]
            deed_no = str(rec["deed_no"])

            print(f"\n[{i+1}/{total}] ค้นหา: {province} / {district} / โฉนดเลขที่ {deed_no}")

            row_result = {
                "province": province, "district": district, "deed_no": deed_no,
                "lat": "", "lng": "", "status": "", "note": "",
            }
            dialog_state["message"] = None  # เคลียร์ก่อนเริ่มค้นหาแปลงใหม่ทุกครั้ง

            # เช็คปิด popup "ข่าวประกาศ" ก่อนเริ่มทุกแปลงด้วย เผื่อมันเด้งขึ้นมา
            # ระหว่างที่รอ delay ของแปลงก่อนหน้า (ไม่ใช่แค่เช็คตอนรอผลลัพธ์เท่านั้น)
            try:
                news_modal = page.locator("#modal_news")
                if news_modal.count() > 0 and news_modal.is_visible():
                    page.locator('#modal_news button[data-dismiss="modal"]').click(timeout=1000, force=True)
                    print("   (ปิด popup ข่าวประกาศก่อนเริ่มแปลงนี้)")
                    page.wait_for_timeout(500)
            except Exception:
                pass

            try:
                page.select_option("#cbprovince", label=province)
                page.wait_for_timeout(WAIT_AFTER_PROVINCE)

                district_ok = select_district_fuzzy(page, district)
                if not district_ok:
                    row_result["status"] = "error"
                    row_result["note"] = f"ไม่พบอำเภอ '{district}' ในตัวเลือกของจังหวัดนี้"
                    print(f"   ⚠️ {row_result['note']}")
                    append_result(row_result)
                    if i < total - 1:
                        if (i + 1) % BREAK_EVERY_N == 0:
                            print(f"\n   🔄 ปิด-เปิดเบราว์เซอร์ใหม่ (ทำมาแล้ว {i+1} แปลง) เพื่อรีเซ็ตเซสชัน กันโดน rate limit...")
                            browser.close()
                            time.sleep(BREAK_PAUSE_SECONDS)
                            browser, context, page, dialog_state = open_fresh_session(p)
                        else:
                            time.sleep(DELAY_SECONDS)
                    continue

                page.wait_for_timeout(500)
                page.fill("#faketxtparcelno", "")
                page.fill("#faketxtparcelno", deed_no)
                try:
                    page.click("#btnSearch", timeout=5000)
                except PWTimeoutError:
                    # ปุ่มอาจถูกบังโดย element อื่นที่ไม่คาดคิด — บังคับคลิกทะลุไปเลย
                    print("   (ปุ่มค้นหาถูกบัง กำลังบังคับคลิกทะลุ...)")
                    page.click("#btnSearch", force=True, timeout=5000)

                # รอผลลัพธ์แบบ "คอยเช็คซ้ำทุกครึ่งวินาที" จนกว่าจะเจอพิกัด หรือหมดเวลา
                # ระหว่างรอ ถ้า popup disclaimer ("รับทราบ") โผล่ขึ้นมาเมื่อไหร่ก็กดปิดทันที
                # (ไม่ใช้วิธีลองครั้งเดียวแบบเดิม เพราะ popup อาจขึ้นช้ากว่าที่คาดไว้)
                href = None
                dismiss_count = 0
                MAX_DISMISS_ATTEMPTS = 6  # ถ้าเกินนี้ยังปิดไม่ได้ แสดงว่ามีปัญหาอื่น หยุดลองปิดแล้วรอเฉยๆ
                poll_deadline = time.time() + (RESULT_TIMEOUT / 1000)
                while time.time() < poll_deadline:
                    # เช็คและปิด popup "ข่าวประกาศ" (#modal_news) ถ้าโผล่ขึ้นมาใหม่ระหว่างรัน
                    # (ก่อนหน้านี้ปิดแค่ครั้งเดียวตอนเปิดเว็บ แต่พบว่า popup นี้เด้งขึ้นมาซ้ำได้
                    # กลางทางระหว่างรันด้วย เช่น ทุกๆ ช่วงเวลาหนึ่ง ไม่ใช่แค่ตอนเปิดเว็บครั้งแรก)
                    try:
                        news_modal = page.locator("#modal_news")
                        if news_modal.count() > 0 and news_modal.is_visible():
                            news_close_btn = page.locator('#modal_news button[data-dismiss="modal"]')
                            news_close_btn.click(timeout=1000, force=True)
                            print("   (ปิด popup ข่าวประกาศที่โผล่ขึ้นมาใหม่ระหว่างรัน)")
                            page.wait_for_timeout(500)
                    except Exception:
                        pass

                    # เช็คและปิด popup SweetAlert2 ("รับทราบ") ถ้าโผล่ขึ้นมา ณ จังหวะนี้
                    # ใช้ class .swal2-confirm ตรงๆ (แม่นกว่าหาจากข้อความ) + บังคับคลิกทะลุ
                    if dismiss_count < MAX_DISMISS_ATTEMPTS:
                        try:
                            ack_btn = page.locator("button.swal2-confirm")
                            if ack_btn.count() > 0 and ack_btn.first.is_visible():
                                ack_btn.first.click(timeout=1000, force=True)
                                dismiss_count += 1
                                print(f"   (ปิด popup รับทราบ ครั้งที่ {dismiss_count})")
                                page.wait_for_timeout(800)  # รอ animation ปิดให้จบก่อนเช็คซ้ำ
                        except Exception:
                            pass

                    # เช็คว่ามีลิงก์พิกัดโผล่มาหรือยัง
                    try:
                        link_locator = page.locator('a[href*="google.com/maps?q="]')
                        if link_locator.count() > 0 and link_locator.first.is_visible():
                            href = link_locator.first.get_attribute("href")
                            break
                    except Exception:
                        pass

                    page.wait_for_timeout(500)

                try:
                    if href is None:
                        raise PWTimeoutError("polling timeout")
                    coords_part = href.split("q=")[-1]
                    lat, lng = coords_part.split(",")[0], coords_part.split(",")[1]

                    row_result["lat"] = lat.strip()
                    row_result["lng"] = lng.strip()
                    row_result["status"] = "found"
                    print(f"   ✅ พบพิกัด: {lat}, {lng}")

                except PWTimeoutError:
                    # รอจนหมดเวลาแล้วไม่มีพิกัดขึ้น — เช็คว่ามี popup แจ้งเตือนไหม
                    if dialog_state["message"]:
                        row_result["status"] = "not_found"
                        row_result["note"] = dialog_state["message"]
                        print(f"   ℹ️ ไม่พบข้อมูล (เว็บแจ้ง: {dialog_state['message']})")
                    else:
                        row_result["status"] = "timeout"
                        row_result["note"] = "รอเกินเวลากำหนด ไม่มีพิกัดและไม่มีข้อความแจ้งเตือน"
                        print(f"   ⏱️ {row_result['note']}")

            except Exception as e:
                row_result["status"] = "error"
                row_result["note"] = str(e)
                print(f"   ⚠️ เกิดข้อผิดพลาด: {e}")

            append_result(row_result)

            try:
                page.click("text=ปิดหน้าต่าง", timeout=2000)
            except Exception:
                pass

            if i < total - 1:
                # ปิด-เปิดเบราว์เซอร์ใหม่ (เซสชัน/คุกกี้ใหม่) เป็นระยะทุกๆ BREAK_EVERY_N แปลง
                # เพื่อทดสอบ+ป้องกัน rate limit (พบว่าค้นถี่ๆ ติดกันเกิน ~24 ครั้ง เว็บเริ่ม
                # ตอบ "ไม่พบ" ปลอมๆ ทุกคำขอ — ถ้าเป็น session-based วิธีนี้จะรีเซ็ตตัวนับได้)
                if (i + 1) % BREAK_EVERY_N == 0:
                    print(f"\n   🔄 ปิด-เปิดเบราว์เซอร์ใหม่ (ทำมาแล้ว {i+1} แปลง) เพื่อรีเซ็ตเซสชัน กันโดน rate limit...")
                    browser.close()
                    time.sleep(BREAK_PAUSE_SECONDS)
                    browser, context, page, dialog_state = open_fresh_session(p)
                else:
                    time.sleep(DELAY_SECONDS)

        browser.close()

    print(f"\n=== เสร็จสิ้นรอบนี้ === ผลลัพธ์สะสมอยู่ที่ {OUTPUT_FILE}")


if __name__ == "__main__":
    run()
