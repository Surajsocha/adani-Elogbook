"""
Seed realistic dummy data for AHP Dry System Logbook.
Run: python seed_dry_dummy.py
"""
import os, django, random
from datetime import date, timedelta, time

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from users.models import CustomUser
from ahp_logbook.models import (
    AHPDryLogEntry, ESPDryDeashing, DryEquipmentStatus
)

# ── Helpers ──────────────────────────────────────────────────────────────────
def t(h, m=0):
    """Return a time object, safely handling minute/hour overflow."""
    total_min = h * 60 + m
    total_min %= (24 * 60)
    return time(total_min // 60, total_min % 60)

def hrs(start, stop):
    """Return float hours between two time objects."""
    s = start.hour * 60 + start.minute
    e = stop.hour * 60 + stop.minute
    diff = (e - s) % (24 * 60)
    return round(diff / 60, 2)

SHIFTS = ['A', 'B', 'C']

ESP_PASSES = [
    'AB-1', 'CD-1', 'AB-2', 'CD-2', 'AB-3', 'CD-3',
    'AB-4', 'CD-4', 'AB-5', 'CD-5', 'AB-6', 'CD-6'
]
DRY_EQUIPMENTS = [
    'COMP-1', 'COMP-2', 'COMP-3', 'COMP-4',
    'CT-FAN', 'CT-PUMP-A', 'CT-PUMP-B',
    'CLASS-A', 'CLASS-B'
]

# Fetch users
try:
    admin_user = CustomUser.objects.get(employee_id='ADMIN001')
    supervisor = CustomUser.objects.get(employee_id='SUP001')
    operator   = CustomUser.objects.get(employee_id='OPR001')
except CustomUser.DoesNotExist:
    print("ERROR: Run python setup_demo.py first to create users!")
    exit(1)

# ── Sample entries: last 5 days × 3 shifts ──────────────────────────────────
ENTRIES = []
today = date.today()

for day_offset in range(5, 0, -1):
    d = today - timedelta(days=day_offset)
    for shift in SHIFTS:
        ENTRIES.append({'date': d, 'shift': shift})

STATUSES = ['approved', 'approved', 'submitted', 'draft']

created_count = 0
skipped_count = 0

for entry_def in ENTRIES:
    d, shift = entry_def['date'], entry_def['shift']

    if AHPDryLogEntry.objects.filter(date=d, shift=shift).exists():
        print(f"  SKIP  {d} {shift} – already exists")
        skipped_count += 1
        continue

    status = random.choice(STATUSES)

    entry = AHPDryLogEntry.objects.create(
        date=d,
        shift=shift,
        document_no='ADTPS/AHP/OPN/F/01',
        shift_incharge=supervisor,
        field_operator=operator,
        ic1_initial=round(random.uniform(1000, 1200), 2),
        ic1_final=round(random.uniform(1200, 1400), 2),
        ic2_initial=round(random.uniform(800, 1000), 2),
        ic2_final=round(random.uniform(1000, 1200), 2),
        fine_ash_silo_level=round(random.uniform(40, 80), 2),
        coarse_ash_silo_level=round(random.uniform(30, 70), 2),
        mt300_ash_silo_level=round(random.uniform(20, 60), 2),
        fly_ash_quantity=round(random.uniform(150, 300), 2),
        observations=random.choice([
            'Dry system running normally. No issues observed.',
            'Compressor 2 vibration high, informed mechanical.',
            'Ash silo level maintained within limits.',
            'Routine inspection done.',
        ]),
        remarks='All OK' if status == 'approved' else '',
        operator_signature='Op1',
        status=status,
        prepared_by=operator,
        approved_by=supervisor if status in ('approved',) else None,
        approval_remarks='Approved as normal.' if status == 'approved' else '',
    )

    # ── ESP Dry Deashing cycles ─────────────────────────────────────────────
    for unit in ['1', '2']:
        for esp_pass in ESP_PASSES:
            for cycle in [1, 2, 3]:
                if random.random() < 0.8:  # 80% chance of having cycle data
                    sh = random.randint(6, 21)
                    sm = random.choice([0, 15, 30, 45])
                    start = t(sh, sm)
                    dur = random.randint(30, 90)  # minutes
                    stop = t(sh, sm + dur)
                    ESPDryDeashing.objects.create(
                        log_entry=entry,
                        unit=unit,
                        esp_pass=esp_pass,
                        cycle=cycle,
                        start_time=start,
                        stop_time=stop,
                        total_time=round(dur / 60, 2),
                    )
                else:
                    ESPDryDeashing.objects.create(
                        log_entry=entry,
                        unit=unit,
                        esp_pass=esp_pass,
                        cycle=cycle,
                        start_time=None,
                        stop_time=None,
                        total_time=None,
                    )

    # ── Equipment Status ───────────────────────────────────────────────
    for equipment in DRY_EQUIPMENTS:
        for cycle in [1, 2, 3]:
            if random.random() < 0.7:
                sh = random.randint(6, 20)
                sm = random.choice([0, 15, 30])
                start = t(sh, sm)
                dur = random.randint(60, 180)
                stop = t(sh, sm + dur)
                DryEquipmentStatus.objects.create(
                    log_entry=entry,
                    equipment=equipment,
                    cycle=cycle,
                    start_time=start,
                    stop_time=stop,
                    running_hrs=round(dur / 60, 2),
                    availability_status=random.choice(['R', 'R', 'A', 'B']),
                )
            else:
                DryEquipmentStatus.objects.create(
                    log_entry=entry,
                    equipment=equipment,
                    cycle=cycle,
                    start_time=None,
                    stop_time=None,
                    running_hrs=None,
                    availability_status=random.choice(['R', 'A', 'B', 'P']),
                )

    print(f"  CREATED  {d} Shift-{shift} | Status={status}")
    created_count += 1

print()
print(f"=== Dry System Dummy Data Seeding Complete ===")
print(f"  Created : {created_count} entries")
print(f"  Skipped : {skipped_count} (already existed)")
