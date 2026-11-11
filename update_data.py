"""
update_data.py — สคริปต์แปลงข้อมูล Excel → JSON สำหรับ SLA Dashboard (index.html)

วิธีใช้:
    1. แก้ตัวแปร EXCEL_FILE, REF_DATE_STR ด้านล่างให้ตรงกับไฟล์/วันที่ที่จะอัปเดต
    2. รัน: python3 update_data.py
    3. สคริปต์จะสร้างไฟล์ AF_new.json / NA_new.json / MN_new.json / ZM_new.json / RJ_new.json
       พร้อมพิมพ์รายงานสรุปมาตรฐาน (ตาม README หัวข้อ 9.1) ออกทางหน้าจอ
    4. ตรวจสอบตัวเลขในรายงานให้ตรงกับที่คาดไว้ก่อน แล้วค่อยเขียนทับ AF/NA/ZM/RJ/MNA/MZM/MRJ ใน index.html
       (การ integrate เข้า index.html แนะนำให้ทำแยกเป็นอีกขั้นตอน ไม่ auto-write ในสคริปต์นี้
        เพื่อให้มีจังหวะตรวจสอบก่อนเขียนทับไฟล์จริงเสมอ ตาม README 9.1)

⚠️ อ่าน README.md หัวข้อ 7, 7.1, 8, 8.1-8.5, 9 ก่อนใช้สคริปต์นี้เสมอ — มีบั๊ก/ข้อควรระวังที่เจอมาแล้ว
   หลายจุดที่สคริปต์นี้แก้ไว้ให้แล้ว (ดู comment แต่ละจุด) แต่โครงสร้าง sheet ต้นทางอาจเปลี่ยนได้อีก
   ต้องตรวจสอบ dtype/ชื่อคอลัมน์ทุกรอบตามที่ README เตือนไว้

ประวัติ: สคริปต์นี้รวบรวม logic จากการอัปเดตข้อมูลหลายรอบ (18/21/25-08-69) รวมบั๊กที่เจอและแก้แล้ว:
  - ปี พ.ศ./ค.ศ. ปนกันในคอลัมน์ datetime เดียวกัน (CONTRACT_DATE vs DATE_01-11)
  - update field ต้อง coalesce 3 ชั้น
  - NA ต้องยึดฐานจาก sheet ต้นทาง ไม่ใช่ filter จาก AF (MNA/MZM/MRJ ก็ยึดหลักเดียวกัน)
  - days_lag/ck_custcare_st ล้ำหน้า 1 สถานะ (แก้ฝั่ง JS ใน index.html ไม่ใช่ตรงนี้)
  - safe_start_con(): วันทำสัญญาของ ZM/RJ มีทั้ง serial number และ datetime แล้วแต่ sheet/รอบข้อมูล
  - safe_yyyymmdd_or_text(): due_first/change_status ของ ZM/RJ สลับบทบาทกันได้ระหว่างรอบข้อมูล
  - sheet_status (RJ) ห้าม hardcode ต้องอ่านจาก CUSTCARE_STATUS_NAME จริงทุกครั้ง
  - guarantee ของ ZM/RJ เปลี่ยนชื่อคอลัมน์จาก CONSTRUCTION_AGENCY.1 เป็น ST_GUARANTEE
  - บ้านเลขที่เพี้ยนเป็นวันที่ (Google Sheets auto-convert) ต้องเช็ค+แก้ทุกรอบ
"""

import pandas as pd
import datetime
import json
import re

# ═══════════════════════════════════════════════════════════
# ตั้งค่าก่อนรันทุกครั้ง
# ═══════════════════════════════════════════════════════════
EXCEL_FILE = '/mnt/user-data/uploads/ส_ญญาย_งไม_ส_งมอบในระบบ-SLA.xlsx'  # ⚠️ แก้ path ให้ตรงกับไฟล์ที่อัปโหลดใหม่ทุกครั้ง
REF_DATE_STR = '2026-11-03'
OLD_INDEX_HTML = '/home/claude/index.html'  # ไฟล์ dashboard เดิม (เอาไว้เทียบผลลัพธ์)

# ชื่อ sheet ปัจจุบัน — เคยเปลี่ยนมาแล้ว 1 ครั้ง (รอจัดจ้าง → ขอลดค่าซ่อม) เช็คชื่อ sheet จริงก่อนรันเสมอ
SHEET_AF = 'AF'
SHEET_CK = 'สำหรับกรอง-ตรวจสอบ'   # → NA, และ note ของ AF
SHEET_ZM = 'ซ่อมเสร็จ'            # → ZM (ดำเนินการเสร็จสิ้น)
SHEET_RJ = 'ขอลดค่าซ่อม'          # → RJ (ประสงค์ขอส่วนลด) — เดิมชื่อ "รอจัดจ้าง"

# ── ข้อมูลหน่วยงานอื่น (main tab ใหม่ 25-08-69) — header row 4 (header=3) ทั้ง 3 sheet
SHEET_MNA = 'หน่วยอื่นไม่ดำเนินงาน'   # → MNA (ไม่ดำเนินการ(หน่วยงานอื่น))
SHEET_MZM = 'หน่วยอื่นซ่อมเสร็จ'      # → MZM (ดำเนินการเสร็จสิ้น(หน่วยงานอื่น))
SHEET_MRJ = 'หน่วยอื่นขอลดค่าซ่อม'    # → MRJ (ประสงค์ขอส่วนลด(หน่วยงานอื่น))
# ⚠️ guarantee ของ 3 sheet นี้อยู่ในคอลัมน์ 'CONSTRUCTION_AGENCY.1' (ไม่ใช่ 'ST_GUARANTEE' แบบ ZM/RJ เดิม) — เช็คทุกรอบ

REF = datetime.date.fromisoformat(REF_DATE_STR)

STAGES11 = ['รับคำร้องแจ้งซ่อม','สำรวจซ่อม','ยืนยันการสำรวจซ่อม','ประมาณราคา',
            'ตรวจสอบและรับรองราคา','อนุมัติสรุปค่าก่อสร้าง','รอการจัดจ้าง','จัดหาผู้รับจ้าง',
            'ดำเนินการซ่อม','ดำเนินการเสร็จสิ้น','ขอเบิกเงินค่าจ้าง']


# ═══════════════════════════════════════════════════════════
# Helper functions (แปลงชนิดข้อมูล / วันที่)
# ═══════════════════════════════════════════════════════════
def s(v):
    if pd.isna(v): return ''
    return str(v).strip()

def num_int(v):
    if pd.isna(v): return ''
    return str(int(v))

def num_comma(v, dec=0):
    if pd.isna(v): return ''
    return f'{v:,.0f}' if dec == 0 else f'{v:,.2f}'

def dt_dmy_be(dt):
    """datetime -> 'DD/MM/YY' (พ.ศ.) — auto-detect ว่าปีเก็บเป็น ค.ศ. หรือ พ.ศ.ตรงๆ"""
    if pd.isna(dt): return ''
    dt = pd.Timestamp(dt)
    y = dt.year
    if y < 2100:
        y += 543
    return f'{dt.day:02d}/{dt.month:02d}/{y % 100:02d}'

def yyyymmdd_dmy_be(v):
    """ตัวเลข YYYYMMDD (ค.ศ.) -> 'DD/MM/YY' (พ.ศ.)"""
    if pd.isna(v): return ''
    sv = str(int(v))
    if len(sv) != 8: return ''
    y, mo, d = int(sv[:4]), int(sv[4:6]), int(sv[6:8])
    return f'{d:02d}/{mo:02d}/{(y + 543) % 100:02d}'

def yyyymmdd_to_date(v):
    if pd.isna(v): return None
    sv = str(int(v))
    if len(sv) != 8: return None
    return datetime.date(int(sv[:4]), int(sv[4:6]), int(sv[6:8]))

def parse_dmy(s_):
    """'DD/MM/YY' (พ.ศ.) -> datetime.date"""
    if not s_: return None
    p = s_.split('/')
    if len(p) != 3: return None
    d, mo, yy = int(p[0]), int(p[1]), int(p[2])
    return datetime.date(1957 + yy, mo, d)

def safe_yyyymmdd_or_text(v):
    """รองรับทั้งกรณีเป็นตัวเลข YYYYMMDD และข้อความ (ZM/RJ due_first/change_status สลับบทบาทกันได้)"""
    if pd.isna(v): return ''
    if isinstance(v, str): return v.strip()
    try:
        return yyyymmdd_dmy_be(v)
    except Exception:
        return s(v)

def safe_start_con(v):
    """วันทำสัญญาของ ZM/RJ — รองรับทั้ง serial number ดิบ (int) และ datetime object ที่ parse มาแล้ว
    ⚠️ ห้ามใช้ pd.Timedelta(days=...) บวกตรงๆ จะ overflow เงียบๆ กับตัวเลข serial ใหญ่ (~244000+)
       ต้องใช้ pd.to_datetime(..., origin='1899-12-30', unit='D') เท่านั้น"""
    if pd.isna(v): return ''
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return dt_dmy_be(v)
    try:
        dt = pd.to_datetime(v, origin='1899-12-30', unit='D')
        return dt_dmy_be(dt)
    except Exception:
        return ''

def status_lag(af_rec):
    """วันค้างสถานะปัจจุบันจริง — ใช้ d0N ของสถานะที่ update อยู่จริง ไม่ใช่ raw days_lag ตรงๆ"""
    if af_rec is None: return None
    upd = af_rec.get('update', '')
    if upd in STAGES11:
        idx = STAGES11.index(upd) + 1
        dt = parse_dmy(af_rec.get(f'd{idx:02d}', ''))
        if dt:
            return (REF - dt).days
    dl = af_rec.get('days_lag', '')
    return int(dl) if str(dl).strip().isdigit() else None

def house_corrupted(house_val):
    """เช็คบ้านเลขที่เพี้ยนเป็นวันที่ (Google Sheets auto-convert เลข 2 ตัวคั่น / เช่น 8/2)"""
    hv = str(house_val or '')
    return bool(re.match(r'^\d{4}-\d{2}-\d{2}', hv)) or '00:00:00' in hv

# รายการบ้านเลขที่เพี้ยนที่เคยเจอและยืนยันค่าที่ถูกต้องแล้ว (แก้อัตโนมัติทุกรอบที่เจอซ้ำ)
KNOWN_HOUSE_FIXES = {
    '200199759': '2/8',   # ยืนยันจาก number_format "d/m" ในไฟล์ xlsx ต้นฉบับ + serial 46236
    '200302294': '2/6',   # 03-11-69: HOUSE_NO_ACT เป็น datetime 2026-06-02 + number_format "d/m" → "2/6" (เพื่อนบ้านในโครงการ 2/34, 2/39 ...)
}


# ═══════════════════════════════════════════════════════════
# 1) AF
# ═══════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════
# วันที่/เวลาอัปเดตที่แสดงผลบน Dashboard (ป้าย "อัปเดต: ...")
# ⚠️ ห้ามพิมพ์วันที่เองแล้ว — ให้อ่านจากบรรทัดที่ 2 (แถวที่ 2, คอลัมน์ A) ของ sheet AF เสมอ
#    ตัวอย่างข้อความจริง: "สัญญายังไม่บันทึกส่งมอบในระบบ F/E (CT_NOTSEND_CUSTCARE  ณ 18-8-69 11:00.45 น)"
#    → ดึงวันที่/เวลาออกมาด้วย regex แล้ว format เป็น "DD-MM-YYYY(พ.ศ.) HH:MM น."
# ═══════════════════════════════════════════════════════════
def extract_update_banner():
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_AF, header=None, nrows=2)
    raw_text = s(df_raw.iloc[1, 0])
    m = re.search(r'ณ\s*(\d{1,2})-(\d{1,2})-(\d{2})\s+(\d{1,2}):(\d{2})(?:[.:](\d{2}))?\s*น', raw_text)
    if not m:
        print(f'⚠️ ไม่พบรูปแบบวันที่ในบรรทัดที่ 2 ของ AF sheet: {raw_text!r} — ต้องใส่ป้าย "อัปเดต:" เองด้วยมือรอบนี้')
        return {'raw': raw_text, 'date_display': '', 'time_display': ''}
    d, mo, y2, hh, mi, _ = m.groups()
    date_display = f'{int(d):02d}-{int(mo):02d}-25{y2}'   # ปีใน sheet เป็น พ.ศ. 2 หลักอยู่แล้ว (เช่น 69 → 2569)
    time_display = f'{int(hh):02d}:{mi} น.'
    return {'raw': raw_text, 'date_display': date_display, 'time_display': time_display}

def extract_sheet_dates():
    """วันที่ข้อมูลของแต่ละ sheet ย่อย จากบรรทัดแรก-สอง (ข้อความที่ตามหลัง "ณ" เช่น "ณ 6-11-69 ") — ถ้าไม่มี = '' (หน้าเว็บจะไม่แสดงป้าย)"""
    out = {}
    for key, sh in [('na', SHEET_CK), ('zm', SHEET_ZM), ('rj', SHEET_RJ),
                    ('mna', SHEET_MNA), ('mzm', SHEET_MZM), ('mrj', SHEET_MRJ)]:
        try:
            df_raw = pd.read_excel(EXCEL_FILE, sheet_name=sh, header=None, nrows=2)
            text = ' '.join(s(v) for v in df_raw.iloc[:2, :6].values.ravel() if s(v))
        except Exception as e:
            print(f'⚠️ อ่านวันที่ข้อมูลจาก sheet {sh} ไม่ได้: {e}')
            text = ''
        m = re.search(r'ณ\s*(\d{1,2})-(\d{1,2})-(\d{2,4})', text)
        if m:
            d, mo, y = m.groups()
            y = y if len(y) == 4 else '25' + y
            out[key] = f'{int(d):02d}-{int(mo):02d}-{y}'
        else:
            out[key] = ''
    return out

def build_af():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_AF, header=2)
    # ⚠️ กันแถวขยะ: บางรอบไฟล์ต้นทางมีแถว header ซ้ำปนมาใน data (CUST_CODE เป็นข้อความ 'CUST_CODE' แทนตัวเลข) — ตัดทิ้งก่อนประมวลผลเสมอ
    junk_mask = pd.to_numeric(df['CUST_CODE'], errors='coerce').isna() & df['CUST_CODE'].notna()
    if junk_mask.any():
        print(f'⚠️ พบแถวขยะใน AF sheet (CUST_CODE ไม่ใช่ตัวเลข) {junk_mask.sum()} แถว — ตัดทิ้งก่อนประมวลผล')
        df = df[~junk_mask].reset_index(drop=True)
    df_ck = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CK, header=2)
    ck_lookup = {}       # cust -> ปรับสรุปการUPDATE (fallback สำหรับ update field)
    note_lookup = {}     # cust -> หมายเหตุ (join เข้า AF โดยตรง ให้ OVNA ได้ด้วยอัตโนมัติ)
    for _, r in df_ck.iterrows():
        c = num_int(r['รหัสลูกค้า'])
        if c:
            ck_lookup[c] = s(r.get('ปรับสรุปการUPDATE'))
            note_lookup[c] = s(r.get('หมายเหตุ'))

    records = []
    fallback_count = 0
    for _, r in df.iterrows():
        cust = num_int(r['CUST_CODE'])
        if not cust:
            continue

        ck_raw = s(r.get('CK_CUSTCARE_ST'))
        if ck_raw:
            parts = ck_raw.split(' ', 1)
            ck_grp = parts[1] if (len(parts) == 2 and parts[0].isdigit() and len(parts[0]) == 8) else ck_raw
        else:
            ck_grp = 'ไม่เคยซ่อม'

        maxd = yyyymmdd_to_date(r.get('MAX_STATUS_DATE'))
        days_lag = str((REF - maxd).days) if maxd else ''

        # update field: coalesce 3 ชั้น (สำคัญมาก — ห้ามใช้แค่ CUSTCARE_STATUS_NAME อย่างเดียว)
        custcare_status = s(r.get('CUSTCARE_STATUS_NAME'))
        if custcare_status:
            upd = custcare_status
        elif cust in ck_lookup and ck_lookup[cust]:
            upd = ck_lookup[cust]
        else:
            upd = s(r.get('NEXT_CUSTCARE'))
            fallback_count += 1

        house = s(r.get('HOUSE_NO_ACT'))
        if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
            house = KNOWN_HOUSE_FIXES[cust]

        rec = {
            'dept': s(r.get('DEPT_SHORT_NAME')), 'bp': s(r.get('BP_SHORT_NAME')), 'proj': s(r.get('PROJ_NAME')),
            'house': house, 'cust': cust, 'cust_name': s(r.get('CUSTNAME1DISPLAY')),
            'change_status': s(r.get('CHANGE_STATUS')), 'status': s(r.get('NEXT_CUSTCARE')), 'update': upd,
            'guarantee': s(r.get('ST_GUARANTEE')), 'confirm_flag': s(r.get('CONFIRM_FLAG')),
            'next_custcare': s(r.get('NEXT_CUSTCARE')),
            'contract_date': yyyymmdd_dmy_be(r.get('CONTRACT_DATE')),
            'days_lag': days_lag,
            'ck_custcare_st': ck_raw, 'ck_custcare_st_grp': ck_grp,
            'd01': dt_dmy_be(r.get('DATE_01')), 'd02': dt_dmy_be(r.get('DATE_02')), 'd03': dt_dmy_be(r.get('DATE_03')),
            'd04': dt_dmy_be(r.get('DATE_04')), 'd05': dt_dmy_be(r.get('DATE_05')), 'd06': dt_dmy_be(r.get('DATE_06')),
            'd07': dt_dmy_be(r.get('DATE_07')), 'd08': dt_dmy_be(r.get('DATE_08')), 'd09': dt_dmy_be(r.get('DATE_09')),
            'd10': dt_dmy_be(r.get('DATE_10')), 'd11': dt_dmy_be(r.get('DATE_11')),
            'aging09': num_int(r.get('AGING_09')), 'ck_aging09': s(r.get('CK_AGING_09')),
            'contract_type': s(r.get('CONTRACT_TYPE1')),
            'note': note_lookup.get(cust, ''),   # ← join จาก สำหรับกรอง-ตรวจสอบ โดยตรง (ไม่ใช่แค่ NA)
        }
        records.append(rec)

    return records, fallback_count


# ═══════════════════════════════════════════════════════════
# 2) NA (ไม่ดำเนินการ สคล/สคจ ชี้แจง) — ยึดฐานจาก sheet สำหรับกรอง-ตรวจสอบ ไม่ใช่ filter จาก AF
# ═══════════════════════════════════════════════════════════
def build_na(af_by_cust, na_fields):
    df_ck = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CK, header=2)
    na_new = []
    found_cnt = not_found_cnt = 0
    for _, r in df_ck.iterrows():
        cust = num_int(r['รหัสลูกค้า'])
        if not cust:
            continue
        note = s(r.get('หมายเหตุ'))
        if cust in af_by_cust:
            rec = dict(af_by_cust[cust])
            found_cnt += 1
        else:
            upd = s(r.get('ปรับสรุปการUPDATE'))
            rec = {k: '' for k in na_fields}
            rec.update({
                'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
                'house': (KNOWN_HOUSE_FIXES[cust] if cust in KNOWN_HOUSE_FIXES and house_corrupted(s(r.get('บ้านเลขที่'))) else s(r.get('บ้านเลขที่'))), 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
                'change_status': s(r.get('CHANGE_STATUS')), 'status': upd, 'update': upd,
                'confirm_flag': s(r.get('CONFIRM_FLAG')), 'next_custcare': upd,
                'ck_custcare_st_grp': 'ไม่เคยซ่อม',
            })
            not_found_cnt += 1
        rec['note'] = note
        na_new.append(rec)
    return na_new, found_cnt, not_found_cnt


# ═══════════════════════════════════════════════════════════
# 4) ZM (ดำเนินการเสร็จสิ้น) และ RJ (ประสงค์ขอส่วนลด)
# ═══════════════════════════════════════════════════════════
def build_zm_rj_base(sheet, p_col, af_by_cust, q_col='นัดรับมอบ'):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=3)
    recs = []
    for _, r in df.iterrows():
        cust = s(int(r['รหัสลูกค้า'])) if pd.notna(r['รหัสลูกค้า']) else ''
        if not cust:
            continue
        af_r = af_by_cust.get(cust)
        house = s(r.get('บ้านเลขที่'))
        if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
            house = KNOWN_HOUSE_FIXES[cust]
        rec = {
            'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
            'house': house, 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
            'start_con': safe_start_con(r.get('วันที่ทำสัญญา')),
            'due_first': safe_yyyymmdd_or_text(r.get('วันที่นัด\nส่งมอบครั้งแรก')),
            'due_new': yyyymmdd_dmy_be(r.get('วันที่นัด\nส่งมอบครั้งใหม่')),
            'change_status': safe_yyyymmdd_or_text(r.get('CHANGE_STATUS')),
            'agency': s(r.get('CONSTRUCTION_AGENCY')),
            'guarantee': s(r.get('ST_GUARANTEE')),   # ⚠️ ชื่อคอลัมน์เปลี่ยนจาก CONSTRUCTION_AGENCY.1 แล้ว
            'o_letter': s(r.get('ส่งจดหมาย')), 'p_extra': s(r.get(p_col)), 'q_appt': s(r.get(q_col)),
            'note': s(r.get('หมายเหตุ')),
            'sheet_custcare_status': s(r.get('CUSTCARE_STATUS_NAME')),  # ห้าม hardcode! อ่านค่าจริงเสมอ
            'contract_type': af_r.get('contract_type', '') if af_r else '',
            'contract_date': af_r.get('contract_date', '') if af_r else '',
            'delivery_status': 'พบใน CT_NotSend' if af_r else 'ไม่พบใน CT_NotSend',
        }
        recs.append((rec, af_r))
    return recs

def build_zm_rj_base_v2(sheet, p_col, af_by_cust, q_col, guarantee_col='ST_GUARANTEE'):
    """เหมือน build_zm_rj_base ทุกประการ แต่รองรับชื่อคอลัมน์ guarantee ที่ต่างกันได้
    (sheet 'หน่วยอื่น...' ใช้ 'CONSTRUCTION_AGENCY.1' แทน 'ST_GUARANTEE')"""
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=3)
    recs = []
    for _, r in df.iterrows():
        cust = s(int(r['รหัสลูกค้า'])) if pd.notna(r['รหัสลูกค้า']) else ''
        if not cust:
            continue
        af_r = af_by_cust.get(cust)
        house = s(r.get('บ้านเลขที่'))
        if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
            house = KNOWN_HOUSE_FIXES[cust]
        rec = {
            'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
            'house': house, 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
            'start_con': safe_start_con(r.get('วันที่ทำสัญญา')),
            'due_first': safe_yyyymmdd_or_text(r.get('วันที่นัด\nส่งมอบครั้งแรก')),
            'due_new': yyyymmdd_dmy_be(r.get('วันที่นัด\nส่งมอบครั้งใหม่')),
            'change_status': safe_yyyymmdd_or_text(r.get('CHANGE_STATUS')),
            'agency': s(r.get('CONSTRUCTION_AGENCY')),
            'guarantee': s(r.get(guarantee_col)),
            'o_letter': s(r.get('ส่งจดหมาย')), 'p_extra': s(r.get(p_col)), 'q_appt': s(r.get(q_col)),
            'note': s(r.get('หมายเหตุ')),
            'sheet_custcare_status': s(r.get('CUSTCARE_STATUS_NAME')),
            'contract_type': af_r.get('contract_type', '') if af_r else '',
            'contract_date': af_r.get('contract_date', '') if af_r else '',
            'delivery_status': 'พบใน CT_NotSend' if af_r else 'ไม่พบใน CT_NotSend',
        }
        recs.append((rec, af_r))
    return recs


# ═══════════════════════════════════════════════════════════
# 5) ข้อมูลหน่วยงานอื่น — MNA / MZM / MRJ (main tab ใหม่, โครงสร้างเดียวกับ NA/ZM/RJ)
# ═══════════════════════════════════════════════════════════
def build_mna(af_by_cust, na_fields):
    df_ck = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_MNA, header=3)
    mna_new = []
    found_cnt = not_found_cnt = 0
    for _, r in df_ck.iterrows():
        cust = num_int(r['รหัสลูกค้า'])
        if not cust:
            continue
        note = s(r.get('หมายเหตุ'))
        if cust in af_by_cust:
            rec = dict(af_by_cust[cust])
            found_cnt += 1
        else:
            upd = s(r.get('ปรับสรุปการUPDATE'))
            rec = {k: '' for k in na_fields}
            rec.update({
                'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
                'house': (KNOWN_HOUSE_FIXES[cust] if cust in KNOWN_HOUSE_FIXES and house_corrupted(s(r.get('บ้านเลขที่'))) else s(r.get('บ้านเลขที่'))), 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
                'change_status': s(r.get('CHANGE_STATUS')), 'status': upd, 'update': upd,
                'confirm_flag': s(r.get('CONFIRM_FLAG')), 'next_custcare': upd,
                'ck_custcare_st_grp': 'ไม่เคยซ่อม',
            })
            not_found_cnt += 1
        rec['note'] = note
        mna_new.append(rec)
    return mna_new, found_cnt, not_found_cnt

def build_mzm(af_by_cust):
    pairs = build_zm_rj_base_v2(SHEET_MZM, 'แจ้งซ่อมเพิ่มเติม', af_by_cust, q_col='นัดรับมอบ', guarantee_col='CONSTRUCTION_AGENCY.1')
    MZM = []
    for rec, af_r in pairs:
        sl = status_lag(af_r)
        rec['status_lag'] = str(sl) if sl is not None else ''
        del rec['sheet_custcare_status']
        MZM.append(rec)
    return MZM

def build_mrj(af_by_cust):
    pairs = build_zm_rj_base_v2(SHEET_MRJ, 'ใบตอบรับ', af_by_cust, q_col='ขอส่วนลด', guarantee_col='CONSTRUCTION_AGENCY.1')
    MRJ = []
    for rec, af_r in pairs:
        rec['sheet_status'] = rec.pop('sheet_custcare_status')
        if af_r:
            rec['af_status'] = af_r.get('update', '')
            sl = status_lag(af_r)
            rec['af_status_lag'] = str(sl) if sl is not None else ''
        else:
            rec['af_status'] = ''
            rec['af_status_lag'] = ''
        MRJ.append(rec)
    return MRJ


def build_zm(af_by_cust):
    pairs = build_zm_rj_base(SHEET_ZM, 'แจ้งซ่อมเพิ่มเติม', af_by_cust)
    ZM = []
    for rec, af_r in pairs:
        sl = status_lag(af_r)
        rec['status_lag'] = str(sl) if sl is not None else ''
        del rec['sheet_custcare_status']  # ZM ไม่ใช้ฟิลด์นี้
        ZM.append(rec)
    return ZM

def build_rj(af_by_cust):
    pairs = build_zm_rj_base(SHEET_RJ, 'ใบตอบรับ', af_by_cust, q_col='ขอส่วนลด')
    RJ = []
    for rec, af_r in pairs:
        rec['sheet_status'] = rec.pop('sheet_custcare_status')  # ← อ่านจาก sheet จริง ไม่ hardcode
        if af_r:
            rec['af_status'] = af_r.get('update', '')
            sl = status_lag(af_r)
            rec['af_status_lag'] = str(sl) if sl is not None else ''
        else:
            rec['af_status'] = ''
            rec['af_status_lag'] = ''
        RJ.append(rec)
    return RJ


# ═══════════════════════════════════════════════════════════
# Main — build ทั้งหมด + รายงานสรุปมาตรฐาน (README หัวข้อ 9.1)
# ═══════════════════════════════════════════════════════════
def main():
    print(f'=== อัปเดตข้อมูลจากไฟล์: {EXCEL_FILE} ===')
    print(f'=== วันที่อ้างอิง (REF_DATE): {REF_DATE_STR} — ต้องตรงกับ MAX_STATUS_DATE ล่าสุดใน AF ===\n')

    update_banner = extract_update_banner()
    sheet_dates = extract_sheet_dates()
    print('\n--- วันที่ข้อมูลของแต่ละ sheet ย่อย (ว่าง = ไฟล์ต้นทางไม่ระบุ → ไม่แสดงป้าย) ---')
    for k, v in sheet_dates.items():
        print(f'{k}: {v or "-"}' + (f'   (AF: {update_banner["date_display"]})' if v and v != update_banner['date_display'] else ''))
    print(f'=== ป้าย "อัปเดต:" ดึงจากบรรทัดที่ 2 ของ AF sheet: {update_banner["raw"]!r} ===')
    print(f'    -> แสดงผลเป็น: อัปเดต: {update_banner["date_display"]}  {update_banner["time_display"]}\n')

    AF, af_fallback = build_af()
    af_by_cust = {r['cust']: r for r in AF}
    print(f'AF: {len(AF)} รายการ | update fallback (ไม่มีทั้ง CUSTCARE_STATUS_NAME และ ปรับสรุปการUPDATE): {af_fallback}')

    NA, na_found, na_not_found = build_na(af_by_cust, list(AF[0].keys()))
    print(f'NA: {len(NA)} รายการ | พบใน AF: {na_found} | ไม่พบ: {na_not_found}')

    ZM = build_zm(af_by_cust)
    zm_delivered = sum(1 for r in ZM if r['delivery_status'] == 'ไม่พบใน CT_NotSend')
    print(f'ZM: {len(ZM)} รายการ | ไม่พบใน CT_NotSend(หลุดจาก AF): {zm_delivered}')

    RJ = build_rj(af_by_cust)
    rj_changed = sum(1 for r in RJ if r['af_status'] and r['af_status'] != r['sheet_status'])
    print(f'RJ: {len(RJ)} รายการ | สถานะเปลี่ยนจาก sheet_status แล้ว: {rj_changed}')

    MNA, mna_found, mna_not_found = build_mna(af_by_cust, list(AF[0].keys()))
    print(f'MNA (ไม่ดำเนินการ หน่วยงานอื่น): {len(MNA)} รายการ | พบใน AF: {mna_found} | ไม่พบ: {mna_not_found}')

    MZM = build_mzm(af_by_cust)
    print(f'MZM (ดำเนินการเสร็จสิ้น หน่วยงานอื่น): {len(MZM)} รายการ')

    MRJ = build_mrj(af_by_cust)
    print(f'MRJ (ประสงค์ขอส่วนลด หน่วยงานอื่น): {len(MRJ)} รายการ')

    # ⚠️ เช็ครหัสลูกค้า (join key) ไม่ซ้ำ + join AF — กันเคสรหัสถูกปัดเป็น 200000000 ในไฟล์ต้นทาง (เจอ 21-9-69: ผ่าน "เป็นตัวเลข" แต่เหลือ 2 ค่า)
    print('\n--- เช็ครหัสลูกค้าไม่ซ้ำ (distinct = จำนวนแถว ถ้าปกติ) + join AF ---')
    for name, arr in [('NA', NA), ('ZM', ZM), ('RJ', RJ), ('MNA', MNA), ('MZM', MZM), ('MRJ', MRJ)]:
        distinct = len({r['cust'] for r in arr})
        in_af = sum(1 for r in arr if r['cust'] in af_by_cust)
        flag = '' if distinct == len(arr) else '  ❌ รหัสซ้ำ/เพี้ยน — หยุด ตรวจไฟล์ต้นทางก่อน integrate'
        print(f'{name}: {len(arr)} แถว | รหัสไม่ซ้ำ {distinct} | พบใน AF {in_af}{flag}')

    # เช็คบ้านเลขที่เพี้ยน
    print('\n--- เช็คบ้านเลขที่เพี้ยนเป็นวันที่ ---')
    for name, arr in [('AF', AF), ('NA', NA), ('ZM', ZM), ('RJ', RJ), ('MNA', MNA), ('MZM', MZM), ('MRJ', MRJ)]:
        bad = [(r['cust'], r['house']) for r in arr if house_corrupted(r.get('house', ''))]
        print(f'{name}: {len(bad)} จุด' + (f' -> {bad}' if bad else ''))
        if bad:
            print('  ⚠️ เจอ cust ใหม่ที่ไม่อยู่ใน KNOWN_HOUSE_FIXES — ต้องตรวจสอบค่าที่ถูกต้องเองก่อนแก้ (ดู README หัวข้อ 7 ข้อ 5)')

    # no-op / สถานะจริงปนอยู่ (README หัวข้อ 4.3)
    print('\n--- เช็ค record ที่มีสถานะจริงปนอยู่ (README 4.3) ---')
    ovna = [r for r in AF if r['status'] == 'ไม่ดำเนินการ']
    print(f'OVNA (derive จาก AF.filter): {len(ovna)} | ในสถานะจริง (ต้องเป็น 0 เสมอ): {sum(1 for r in ovna if r["update"] in STAGES11)}')
    print(f'NA: {len(NA)} | ในสถานะจริงปนอยู่: {sum(1 for r in NA if r["update"] in STAGES11)}')

    # เทียบกับข้อมูลเดิมในไฟล์ (ถ้ามี)
    try:
        with open(OLD_INDEX_HTML, encoding='utf-8') as f:
            old_content = f.read()
        old_af = json.loads(re.search(r'const AF=(\[.*?\]);', old_content).group(1))
        old_by_cust = {str(r['cust']): r for r in old_af}
        common = [c for c in af_by_cust if c in old_by_cust]
        print(f'\n--- เทียบกับ AF เดิมในไฟล์ ---')
        print(f'ตรงกัน: {len(common)}/{len(AF)} (เดิมมี {len(old_af)})')
        for f in ['dept', 'update', 'contract_type', 'd01', 'guarantee']:
            match = sum(1 for c in common if str(old_by_cust[c][f]) == str(af_by_cust[c][f]))
            print(f'  {f}: {match}/{len(common)} ({match/len(common)*100:.1f}%)')
    except Exception as e:
        print(f'\n(ข้ามการเทียบกับไฟล์เดิม: {e})')

    # เขียนไฟล์ผลลัพธ์
    for name, data in [('AF', AF), ('NA', NA), ('ZM', ZM), ('RJ', RJ),
                        ('MNA', MNA), ('MZM', MZM), ('MRJ', MRJ)]:
        with open(f'{name}_new.json', 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
    with open('UPDATE_BANNER.json', 'w', encoding='utf-8') as f:
        json.dump(update_banner, f, ensure_ascii=False)
    with open('SHEET_DATES.json', 'w', encoding='utf-8') as f:
        json.dump(sheet_dates, f, ensure_ascii=False)
    print('\n=== เขียนไฟล์ AF_new.json / NA_new.json / MN_new.json / ZM_new.json / RJ_new.json / UPDATE_BANNER.json เรียบร้อย ===')
    print('=== ตรวจสอบตัวเลขข้างบนให้ครบตามรายงานสรุปมาตรฐาน (README 9.1) ก่อน integrate เข้า index.html ===')


if __name__ == '__main__':
    main()
