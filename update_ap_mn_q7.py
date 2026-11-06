import os
import shutil
import re
import pandas as pd
from app import app, db, NORMALIZE_PERSON
from models import ApartmentRecord

EXCEL_PATH = '/Users/qweasdzxcbm/Downloads/Premium Building Quận 7.xlsx'

def norm(s):
    if not s or pd.isna(s): return ''
    return re.sub(r'\s+', ' ', str(s).strip().replace('\n', ' ')).upper()

def clean_person(p):
    if not p or pd.isna(p): return None
    p = str(p).strip().upper()
    return NORMALIZE_PERSON.get(p, p)

STATUS_MAP = {
    'RESEARCH': 'Research',
    'PLAN B': 'Plan B',
    'PLAN A': 'Plan A',
    'DEAL': 'Deal',
    'DONE': 'Done',
    'CODE': 'Code',
    'UNQUALIFIED': 'Unqualified',
    'PENDING': 'Pending',
    'LOST': 'Lost',
    'REJECT': 'Reject',
}

def clean_status(s):
    if not s or pd.isna(s): return None
    s = str(s).strip().upper()
    return STATUS_MAP.get(s, s.title())

def clean_num_units(v):
    if pd.isna(v) or v is None: return None
    s = str(v).strip()
    if s == '900': return '800-1000 căn'
    if s == '290': return '<300 căn'
    if s == '500': return '300-500 căn'
    return s

def clean_int(v):
    if pd.isna(v) or v is None: return None
    try:
        return int(float(str(v).replace(',', '').replace("'", "")))
    except (ValueError, TypeError):
        return None

def clean_str(v):
    if pd.isna(v) or v is None: return None
    s = str(v).strip()
    return s if s and s.lower() != 'nan' else None

def clean_upper(v):
    s = clean_str(v)
    return s.upper() if s else None

# Mapping of Excel row index -> DB record ID (or None for brand new records)
ROW_TO_DB_ID = {
    0: 44,    # SUNRISE CITY CENTRAL
    1: None,  # SUNRISE CITY NORTH (new)
    2: None,  # SUNRISE CITY SOUTH (new)
    3: 273,   # SUNRISE CITYVIEW
    4: 235,   # THE PANORAMA (stt 237)
    5: 236,   # MỸ KHÁNH
    6: 237,   # MỸ PHÚC (Q7)
    7: 238,   # PHÚ MỸ
    8: 239,   # V-STAR
    9: 240,   # GARDEN COURT 2
    10: 241,  # MINH THÀNH
    11: 242,  # MỸ PHÚ
    12: 243,  # BELLEZA LÔ C1
    13: 244,  # HOÀNG ANH GIA LAI 2
    14: 245,  # SKY GARDEN
    15: 246,  # HIM LAM RIVER SIDE
    16: 247,  # HIM LAM RIVER SIDE 2
    17: 248,  # RIVERSIDE RESIDENCE (stt 251)
    18: 249,  # H18 - RIVER PARK
    19: 250,  # GREEN VIEW (4 blocks, 0 screens)
    20: 251,  # CẢNH VIÊN 3
    21: 252,  # HOÀNG ANH GIA LAI 1 BLOCK A3
    22: 253,  # THE BRIDGE VIEW
    23: 254,  # RIVIERA POINT
    24: 255,  # QUỐC CƯỜNG GIA LAI
    25: 256,  # HOÀNG ANH THANH BÌNH
    26: 257,  # MỸ KHANG
    27: 260,  # GRANDVIEW
    28: 261,  # ERA TOWN
    29: 262,  # STAR HILL (C14B, 6 blocks, 24 screens)
    30: 263,  # LUXCITY
    31: 264,  # TULIP TOWER
    32: 265,  # GREEN VALLEY
    33: 266,  # HOÀNG QUỐC VIỆT
    34: 267,  # LONG SƠN
    35: 268,  # AN GIA SKYLINE
    36: 537,  # RIVERPARK RESIDENCE (ID 537, 3 blocks)
    37: 272,  # NAM PHÚC
    38: 275,  # SCENIC VALLEY
    39: 276,  # AN PHÚ LONG LAND 1
    40: 277,  # NGỌC LAN (stt 280)
    41: 299,  # SCENIC VALLEY 2
    42: 278,  # ECO GREEN SÀI GÒN
    43: 280,  # LUX GARDEN
    44: 281,  # COSMO I
    45: 282,  # URBAN HILL
    46: 283,  # FLORITA
    47: 284,  # SUNSHINE SKY CITY
    48: 285,  # LAVIDA PLUS
    49: 286,  # THE VIEW RIVIERA POINT
    50: 287,  # GOLDEN KING
    51: 289,  # Q7 SAIGON RIVERSIDE COMPLEX
    52: 290,  # ASCENTIA
    53: 294,  # ANTONIA
    54: 292,  # HAPPY VALLEY PREMIER
    55: 293,  # RIVER PANORAMA (stt 296)
    56: 295,  # BELLEZA LÔ D, E
    57: 296,  # JAMONA CITY
    58: 297,  # M-ONE
    59: 298,  # THE ERA TOWN KHU A
    60: 300,  # RICHLANE RESIDENCES
    61: 301,  # CẢNH VIÊN 1
    62: 489,  # AN GIA RIVERSIDE
    63: 490,  # GARDEN PLAZA 1&2
    64: 507,  # BENHILL
    65: 508,  # THE LEGACY
    66: 509,  # GREEN VIEW (1 block, 15 screens)
    67: None, # DOCKLANDS (new)
    68: None, # CARDINAL COURT (new)
    69: None, # THE AURORA (new)
    70: None, # THE INFINITY RIVIERA PONIT (new)
    71: 536,  # MIDTOWN THE PEAK (Midtown M8)
    72: 535,  # MIDTOWN THE SIGNATURE (Midtown M7)
    73: 534,  # MIDTOWN THE SYMPHONY (Midtown M6)
    74: 533,  # MIDTOWN THE GRANDE (Midtown M5)
    75: 327,  # SAIGON MIA
    76: 346,  # SUNRISE RIVERSIDE LÔ G
    77: 350,  # SUNRISE RIVERSIDE LÔ E
    78: 348,  # SAIGON SOUTH
}

# Duplicate records in DB that are redundant copies of records above
DUPLICATE_DB_IDS_TO_REMOVE = [
    531,  # "Panorama" (dup of 235)
    527,  # "River Panorama" (dup of 293)
    532,  # "Ngọc Lan" (dup of 277)
    512,  # "Riverside Residence" (dup of 248)
]

def run_update():
    db_path = os.path.join(os.path.dirname(__file__), 'instance', 'database.db')
    backup_path = os.path.join(os.path.dirname(__file__), 'instance', 'database.db.bak')
    print(f"Creating backup of database: {backup_path}")
    shutil.copy2(db_path, backup_path)
    
    df = pd.read_excel(EXCEL_PATH)
    print(f"Loaded Excel with {len(df)} rows.")

    with app.app_context():
        # Current max STT in AP MN
        max_stt = db.session.query(db.func.max(ApartmentRecord.stt)).filter_by(
            category='AP', region='MN'
        ).scalar() or 0
        print(f"Current max STT in AP MN: {max_stt}")
        
        # Remove duplicate records
        print(f"Removing {len(DUPLICATE_DB_IDS_TO_REMOVE)} duplicate ghost records...")
        for dup_id in DUPLICATE_DB_IDS_TO_REMOVE:
            dup_rec = db.session.get(ApartmentRecord, dup_id)
            if dup_rec:
                print(f"  Deleting duplicate ID={dup_rec.id} ({dup_rec.building_name})")
                db.session.delete(dup_rec)
        db.session.flush()

        updated_count = 0
        inserted_count = 0

        for idx, row in df.iterrows():
            target_id = ROW_TO_DB_ID.get(idx)
            
            bname = norm(row['Tên chung cư'])
            num_blocks = clean_int(row['Số Block'])
            district = clean_upper(row['Quận'])
            city = clean_upper(row['Thành phố'])
            status = clean_status(row['Trạng thái'])
            
            pt1 = clean_person(row['Người PT1'])
            pt2 = clean_person(row['Người PT2'])
            pt3 = clean_person(row['Người PT3'])
            pts = [p for p in (pt1, pt2, pt3) if p]
            pic = ', '.join(pts) if pts else None
            
            classification = clean_upper(row['Phân loại'])
            num_units = clean_num_units(row['SL Căn hộ / VP'])
            must_have = 'Must have' if clean_str(row['Must have']) else None
            price_range = clean_str(row['Giá bán'])
            infrastructure = clean_str(row['CSVC'])
            occupancy = clean_upper(row['Tỷ lệ lấp đầy'])
            previous_operator = clean_upper(row['Đơn vị cũ'])
            
            screens_in = clean_int(row['Màn trong thang'])
            screens_out = clean_int(row['Màn ngoài thang'])
            total_screens = clean_int(row['Tổng SL màn'])
            digital_standee = clean_int(row['Digital Standee'])
            digital_standee_status = clean_str(row['Trạng thái Standee'])
            sl_led = clean_int(row['SL Led'])
            
            if target_id is not None:
                rec = db.session.get(ApartmentRecord, target_id)
                if not rec:
                    raise ValueError(f"Record with ID {target_id} not found!")
                
                rec.building_name = bname
                rec.num_blocks = num_blocks
                rec.district = district
                rec.city = city
                rec.status = status
                rec.person_in_charge = pic
                rec.classification = classification
                rec.num_units = num_units
                rec.must_have = must_have
                rec.price_range = price_range
                rec.infrastructure = infrastructure
                rec.occupancy = occupancy
                rec.previous_operator = previous_operator
                rec.screens_in_elevator = screens_in
                rec.screens_outside_elevator = screens_out
                rec.total_screens = total_screens
                rec.digital_standee = digital_standee
                rec.digital_standee_status = digital_standee_status
                rec.sl_led = sl_led
                
                # If record didn't have an STT (e.g. Midtown M5-M8, River Park Residence), assign sequentially
                if rec.stt is None:
                    max_stt += 1
                    rec.stt = max_stt
                
                updated_count += 1
            else:
                max_stt += 1
                new_rec = ApartmentRecord(
                    category='AP',
                    region='MN',
                    stt=max_stt,
                    building_name=bname,
                    num_blocks=num_blocks,
                    district=district,
                    city=city,
                    status=status,
                    person_in_charge=pic,
                    classification=classification,
                    num_units=num_units,
                    must_have=must_have,
                    price_range=price_range,
                    infrastructure=infrastructure,
                    occupancy=occupancy,
                    previous_operator=previous_operator,
                    screens_in_elevator=screens_in,
                    screens_outside_elevator=screens_out,
                    total_screens=total_screens,
                    digital_standee=digital_standee,
                    digital_standee_status=digital_standee_status,
                    sl_led=sl_led,
                )
                db.session.add(new_rec)
                inserted_count += 1
                
        db.session.commit()
        print(f"Successfully updated {updated_count} records and added {inserted_count} new records!")
        
        # Verify final count
        final_count = ApartmentRecord.query.filter_by(category='AP', region='MN').count()
        print(f"Total AP MN records in DB now: {final_count}")

if __name__ == '__main__':
    run_update()
