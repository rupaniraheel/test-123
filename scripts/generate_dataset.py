#!/usr/bin/env python3
"""
Generates a realistic, internally-consistent Oracle-style dataset for the
Vital Pakistan Trust (VPT) Consumption Analysis dashboard.

Output: data/VPT_Consumption_Dataset.xlsx

The dataset is synthetic but built to behave like real ERP data:
  * 27 months (Jul-2024 .. Sep-2026)
  * master data: donors, projects, items, suppliers, warehouses, calendar
  * transactions: requisitions/POs, GRNs with batch + expiry, inventory issues
  * consumption is drawn FEFO from received batches, so stock never goes negative
  * deliberate, documented data-quality defects (unattributed project codes,
    missing prices, duplicate GRNs) so the dashboard's exception logic is testable
"""
import random
from collections import deque, defaultdict
from datetime import date, timedelta
import calendar
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(20261030)

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "VPT_Consumption_Dataset.xlsx")

START = date(2024, 7, 1)
END = date(2026, 9, 30)


def month_list(start, end):
    out, y, m = [], start.year, start.month
    while (y, m) <= (end.year, end.month):
        out.append((y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


MONTHS = month_list(START, END)
FY = lambda y, m: f"FY{y if m >= 7 else y - 1}-{str((y + 1) if m >= 7 else y)[2:]}"


# --------------------------------------------------------------------------
# 1. Reference / master data
# --------------------------------------------------------------------------
donors = [
    ("DNR-01", "UNICEF Pakistan", "Multilateral", "UPAK/SND/2024/118", "PKR"),
    ("DNR-02", "World Health Organization", "Multilateral", "WHO-PAK-2024-0442", "PKR"),
    ("DNR-03", "The Global Fund", "Multilateral", "GF-PAK-TB-2023-09", "PKR"),
    ("DNR-04", "PPAF", "National NGO/INGO", "PPAF-SR-2024-217", "PKR"),
    ("DNR-05", "Engro Foundation", "Corporate", "EF-HLTH-2024-06", "PKR"),
    ("DNR-06", "Sindh Health Department", "Government", "SHD-PHC-2024-31", "PKR"),
    ("DNR-07", "Vital Pakistan Trust (Self-funded)", "Own source", "VPT-CORE", "PKR"),
]

projects = [
    # id, name, donor, district, facility, start, end, budget, sector, scale
    ("PRJ-001", "Maternal & Child Health Programme - Korangi", "DNR-01", "Karachi - Korangi",
     "VPT Korangi MCH Centre", date(2024, 7, 1), date(2026, 12, 31), 92_500_000, "MCH", 1.35),
    ("PRJ-002", "Nutrition Stabilisation Centre - Malir", "DNR-01", "Karachi - Malir",
     "VPT Malir NSC", date(2024, 7, 1), date(2026, 12, 31), 61_000_000, "Nutrition", 0.95),
    ("PRJ-003", "TB DOTS Expansion - Karachi South", "DNR-03", "Karachi - South",
     "VPT Lyari DOTS Centre", date(2024, 7, 1), date(2026, 9, 30), 74_000_000, "TB", 1.10),
    ("PRJ-004", "Primary Healthcare Support - Tharparkar", "DNR-04", "Tharparkar",
     "VPT Mithi & Diplo BHUs", date(2024, 10, 1), date(2026, 12, 31), 58_000_000, "PHC", 0.85),
    ("PRJ-005", "Flood Response & Health Recovery - Dadu", "DNR-02", "Dadu",
     "VPT Johi Mobile Health Unit", date(2024, 7, 1), date(2025, 12, 31), 46_500_000, "Emergency", 0.70),
    ("PRJ-006", "Rural Health Centre Strengthening - Jamshoro", "DNR-06", "Jamshoro",
     "VPT Manjhand RHC", date(2025, 1, 1), date(2026, 12, 31), 39_000_000, "PHC", 0.65),
    ("PRJ-007", "Community Health Worker Programme - Bin Qasim", "DNR-05", "Karachi - Bin Qasim",
     "VPT Bin Qasim CHW Hub", date(2024, 7, 1), date(2026, 6, 30), 27_500_000, "Community", 0.55),
    ("PRJ-008", "Emergency Obstetric Care - Lyari", "DNR-07", "Karachi - Lyari",
     "VPT Lyari EmOC Centre", date(2025, 1, 1), date(2026, 12, 31), 33_000_000, "MCH", 0.60),
    ("PRJ-009", "Hepatitis B & C Screening - Keamari", "DNR-06", "Karachi - Keamari",
     "VPT Keamari Screening Clinic", date(2024, 7, 1), date(2025, 6, 30), 21_000_000, "PHC", 0.45),
    ("PRJ-010", "Thalassaemia Support Programme - Gadap", "DNR-07", "Karachi - Gadap",
     "VPT Gadap Thalassaemia Unit", date(2024, 7, 1), date(2026, 12, 31), 44_000_000, "Chronic care", 0.75),
]

# id, name, category, sub-category, UOM, pack size, unit cost PKR, consumable, shelf life months
items = [
    # ---- Medicines
    ("MED-0001", "Paracetamol 500 mg Tablet", "Medicines", "Analgesic / Antipyretic", "Tablet", 1000, 1.15, "Yes", 36),
    ("MED-0002", "Amoxicillin 250 mg Capsule", "Medicines", "Antibiotic", "Capsule", 1000, 4.20, "Yes", 36),
    ("MED-0003", "Amoxicillin + Clavulanic Acid 625 mg", "Medicines", "Antibiotic", "Tablet", 100, 38.50, "Yes", 30),
    ("MED-0004", "Ceftriaxone 1 g Injection", "Medicines", "Antibiotic", "Vial", 1, 96.00, "Yes", 24),
    ("MED-0005", "Metronidazole 400 mg Tablet", "Medicines", "Antibiotic", "Tablet", 1000, 2.10, "Yes", 36),
    ("MED-0006", "Ciprofloxacin 500 mg Tablet", "Medicines", "Antibiotic", "Tablet", 200, 6.80, "Yes", 36),
    ("MED-0007", "ORS Sachet (Low Osmolarity)", "Medicines", "Rehydration", "Sachet", 100, 22.50, "Yes", 30),
    ("MED-0008", "Ibuprofen 400 mg Tablet", "Medicines", "Analgesic / Antipyretic", "Tablet", 1000, 2.40, "Yes", 36),
    ("MED-0009", "Normal Saline 0.9% 500 ml", "Medicines", "IV Fluid", "Bottle", 20, 118.00, "Yes", 24),
    ("MED-0010", "Ringer's Lactate 500 ml", "Medicines", "IV Fluid", "Bottle", 20, 132.00, "Yes", 24),
    ("MED-0011", "Dextrose 5% 500 ml", "Medicines", "IV Fluid", "Bottle", 20, 124.00, "Yes", 24),
    ("MED-0012", "Tetanus Toxoid Vaccine 0.5 ml", "Medicines", "Vaccine", "Dose", 10, 96.00, "Yes", 18),
    ("MED-0013", "Misoprostol 200 mcg Tablet", "Medicines", "Obstetric", "Tablet", 60, 26.00, "Yes", 30),
    ("MED-0014", "Oxytocin 10 IU Injection", "Medicines", "Obstetric", "Ampoule", 10, 34.50, "Yes", 18),
    ("MED-0015", "Magnesium Sulphate 50% 10 ml", "Medicines", "Obstetric", "Ampoule", 10, 41.00, "Yes", 24),
    ("MED-0016", "Ferrous Sulphate + Folic Acid Tablet", "Medicines", "Supplement", "Tablet", 1000, 1.85, "Yes", 36),
    ("MED-0017", "Zinc Sulphate 20 mg Dispersible Tablet", "Medicines", "Supplement", "Tablet", 500, 3.10, "Yes", 30),
    ("MED-0018", "Vitamin A 200,000 IU Capsule", "Medicines", "Supplement", "Capsule", 100, 12.40, "Yes", 30),
    ("MED-0019", "Rifampicin + Isoniazid 150/75 mg", "Medicines", "Anti-TB", "Tablet", 672, 18.60, "Yes", 24),
    ("MED-0020", "Isoniazid 300 mg Tablet", "Medicines", "Anti-TB", "Tablet", 1000, 3.90, "Yes", 36),
    ("MED-0021", "Ethambutol 400 mg Tablet", "Medicines", "Anti-TB", "Tablet", 500, 7.40, "Yes", 36),
    ("MED-0022", "Pyrazinamide 500 mg Tablet", "Medicines", "Anti-TB", "Tablet", 500, 8.10, "Yes", 36),
    ("MED-0023", "Mebendazole 100 mg Tablet", "Medicines", "Deworming", "Tablet", 500, 2.60, "Yes", 36),
    ("MED-0024", "Artemether + Lumefantrine 20/120 mg", "Medicines", "Antimalarial", "Tablet", 30, 118.00, "Yes", 24),
    ("MED-0025", "Salbutamol Inhaler 100 mcg", "Medicines", "Respiratory", "Inhaler", 1, 268.00, "Yes", 24),
    ("MED-0026", "Hydrocortisone 100 mg Injection", "Medicines", "Emergency", "Vial", 1, 52.00, "Yes", 24),
    ("MED-0027", "Adrenaline 1 mg/ml Injection", "Medicines", "Emergency", "Ampoule", 10, 68.00, "Yes", 18),
    ("MED-0028", "Atropine 0.6 mg/ml Injection", "Medicines", "Emergency", "Ampoule", 10, 39.00, "Yes", 18),
    ("MED-0029", "Diclofenac 75 mg Injection", "Medicines", "Analgesic / Antipyretic", "Ampoule", 10, 28.50, "Yes", 24),
    ("MED-0030", "Insulin Regular 100 IU/ml", "Medicines", "Chronic care", "Vial", 1, 540.00, "Yes", 18),
    ("MED-0031", "Deferasirox 500 mg Tablet", "Medicines", "Chronic care", "Tablet", 30, 1_450.00, "Yes", 24),
    ("MED-0032", "Folic Acid 5 mg Tablet", "Medicines", "Supplement", "Tablet", 1000, 1.05, "Yes", 36),
    ("MED-0033", "Calcium Carbonate + Vitamin D3", "Medicines", "Supplement", "Tablet", 500, 6.20, "Yes", 30),
    ("MED-0034", "Tenofovir + Emtricitabine 300/200 mg", "Medicines", "Antiviral", "Tablet", 30, 890.00, "Yes", 24),
    # ---- Medical / surgical consumables
    ("CON-0001", "Examination Gloves, Latex (Pair)", "Medical Consumables", "PPE", "Pair", 100, 14.50, "Yes", 48),
    ("CON-0002", "Surgical Gloves, Sterile 7.5 (Pair)", "Medical Consumables", "PPE", "Pair", 50, 62.00, "Yes", 48),
    ("CON-0003", "Disposable Syringe 5 ml with Needle", "Medical Consumables", "Injection", "Piece", 100, 21.50, "Yes", 60),
    ("CON-0004", "Disposable Syringe 10 ml with Needle", "Medical Consumables", "Injection", "Piece", 100, 27.00, "Yes", 60),
    ("CON-0005", "IV Cannula 22G", "Medical Consumables", "IV therapy", "Piece", 50, 46.00, "Yes", 48),
    ("CON-0006", "IV Giving Set", "Medical Consumables", "IV therapy", "Piece", 50, 84.00, "Yes", 48),
    ("CON-0007", "Cotton Wool Roll 500 g", "Medical Consumables", "Dressing", "Roll", 10, 245.00, "Yes", 48),
    ("CON-0008", "Gauze Piece Sterile 10x10 cm", "Medical Consumables", "Dressing", "Piece", 100, 9.80, "Yes", 48),
    ("CON-0009", "Adhesive Bandage 10 cm x 4 m", "Medical Consumables", "Dressing", "Roll", 12, 96.00, "Yes", 48),
    ("CON-0010", "Surgical Face Mask, 3-ply", "Medical Consumables", "PPE", "Piece", 50, 7.20, "Yes", 48),
    ("CON-0011", "Nitrile Examination Gloves (Box 100)", "Medical Consumables", "PPE", "Box", 1, 1_280.00, "Yes", 48),
    ("CON-0012", "Alcohol Swab", "Medical Consumables", "Antisepsis", "Piece", 100, 2.40, "Yes", 36),
    ("CON-0013", "Povidone Iodine 10% 100 ml", "Medical Consumables", "Antisepsis", "Bottle", 12, 168.00, "Yes", 36),
    ("CON-0014", "Urine Collection Bag 2000 ml", "Medical Consumables", "Urology", "Piece", 25, 78.00, "Yes", 48),
    ("CON-0015", "Disposable Delivery Kit", "Medical Consumables", "MCH", "Kit", 1, 640.00, "Yes", 36),
    ("CON-0016", "Umbilical Cord Clamp", "Medical Consumables", "MCH", "Piece", 100, 11.50, "Yes", 60),
    ("CON-0017", "Suture Polyglactin 2-0", "Medical Consumables", "Surgical", "Piece", 12, 310.00, "Yes", 36),
    ("CON-0018", "Blood Transfusion Set", "Medical Consumables", "Transfusion", "Piece", 25, 195.00, "Yes", 48),
    ("CON-0019", "Hand Sanitiser 500 ml", "Medical Consumables", "Antisepsis", "Bottle", 12, 385.00, "Yes", 24),
    ("CON-0020", "Sputum Collection Container", "Medical Consumables", "Laboratory", "Piece", 100, 26.00, "Yes", 60),
    # ---- Laboratory
    ("LAB-0001", "Hemoglobin Reagent (Cyanmethemoglobin)", "Laboratory", "Haematology", "Bottle", 1, 4_850.00, "Yes", 18),
    ("LAB-0002", "Rapid Diagnostic Test - Malaria", "Laboratory", "RDT", "Test", 25, 148.00, "Yes", 18),
    ("LAB-0003", "Rapid Diagnostic Test - Hepatitis B", "Laboratory", "RDT", "Test", 25, 172.00, "Yes", 18),
    ("LAB-0004", "Rapid Diagnostic Test - Hepatitis C", "Laboratory", "RDT", "Test", 25, 188.00, "Yes", 18),
    ("LAB-0005", "Glucose Test Strips", "Laboratory", "Biochemistry", "Strip", 50, 42.00, "Yes", 12),
    ("LAB-0006", "Blood Collection Tube EDTA 4 ml", "Laboratory", "Phlebotomy", "Piece", 100, 27.50, "Yes", 24),
    ("LAB-0007", "Microscope Slide, Frosted", "Laboratory", "Consumable", "Piece", 50, 5.60, "Yes", 60),
    ("LAB-0008", "Ziehl-Neelsen Stain Kit", "Laboratory", "Microbiology", "Kit", 1, 3_950.00, "Yes", 18),
    # ---- Nutrition
    ("NUT-0001", "RUTF Plumpy'Nut 92 g Sachet", "Nutrition", "Therapeutic food", "Sachet", 150, 218.00, "Yes", 24),
    ("NUT-0002", "F-75 Therapeutic Milk 400 g", "Nutrition", "Therapeutic food", "Tin", 24, 690.00, "Yes", 24),
    ("NUT-0003", "F-100 Therapeutic Milk 400 g", "Nutrition", "Therapeutic food", "Tin", 24, 760.00, "Yes", 24),
    ("NUT-0004", "Ready-to-Use Supplementary Food (RUSF)", "Nutrition", "Supplementary food", "Sachet", 150, 145.00, "Yes", 24),
    ("NUT-0005", "Multiple Micronutrient Powder", "Nutrition", "Supplement", "Sachet", 30, 58.00, "Yes", 24),
    ("NUT-0006", "Therapeutic Zinc + ORS Combo Pack", "Nutrition", "Supplement", "Kit", 20, 96.00, "Yes", 24),
    ("NUT-0007", "Infant Formula, Therapeutic (Lactogen-free)", "Nutrition", "Therapeutic food", "Tin", 12, 1_240.00, "Yes", 18),
    # ---- Equipment (low volume, capital)
    ("EQP-0001", "Digital Weighing Scale, Infant", "Equipment", "MCH equipment", "Unit", 1, 18_500.00, "No", 120),
    ("EQP-0002", "Digital Weighing Scale, Adult", "Equipment", "MCH equipment", "Unit", 1, 22_000.00, "No", 120),
    ("EQP-0003", "MUAC Measuring Tape", "Equipment", "Nutrition equipment", "Unit", 1, 650.00, "No", 60),
    ("EQP-0004", "BP Apparatus, Aneroid", "Equipment", "Clinical equipment", "Unit", 1, 8_900.00, "No", 120),
    ("EQP-0005", "Stethoscope, Dual Head", "Equipment", "Clinical equipment", "Unit", 1, 4_200.00, "No", 120),
    ("EQP-0006", "Digital Thermometer", "Equipment", "Clinical equipment", "Unit", 1, 1_450.00, "No", 60),
    ("EQP-0007", "Pulse Oximeter, Fingertip", "Equipment", "Clinical equipment", "Unit", 1, 6_800.00, "No", 84),
    ("EQP-0008", "Nebuliser Machine", "Equipment", "Clinical equipment", "Unit", 1, 14_500.00, "No", 120),
    ("EQP-0009", "Examination Couch, Folding", "Equipment", "Furniture", "Unit", 1, 34_000.00, "No", 120),
    ("EQP-0010", "Cold Box, Vaccine Carrier 5 L", "Equipment", "Cold chain", "Unit", 1, 12_800.00, "No", 84),
    # ---- General / administrative
    ("GEN-0001", "Nitrile Gloves, Utility (Pair)", "General Items", "Cleaning & safety", "Pair", 50, 32.00, "Yes", 36),
    ("GEN-0002", "Bleach Solution 5% 1 L", "General Items", "Cleaning & safety", "Bottle", 12, 285.00, "Yes", 24),
    ("GEN-0003", "A4 Paper Ream 80 gsm", "General Items", "Stationery", "Ream", 5, 1_650.00, "Yes", 120),
    ("GEN-0004", "Registers & Health Cards (Pack)", "General Items", "Stationery", "Pack", 1, 890.00, "Yes", 120),
    ("GEN-0005", "Sharps Disposal Container 5 L", "General Items", "Waste management", "Piece", 1, 420.00, "Yes", 60),
    ("GEN-0006", "Biohazard Waste Bag (Pack 50)", "General Items", "Waste management", "Pack", 1, 560.00, "Yes", 36),
    ("GEN-0007", "Printed IEC Material (Booklet)", "General Items", "Awareness", "Piece", 100, 38.00, "Yes", 60),
    ("GEN-0008", "Hand Wash Soap 100 g", "General Items", "Cleaning & safety", "Piece", 24, 118.00, "Yes", 36),
]

suppliers = [
    ("SUP-001", "Getz Pharma (Pvt) Ltd", "Karachi", "0712345-6", "30 days", 21),
    ("SUP-002", "Muller & Phipps Pakistan", "Karachi", "0789012-3", "45 days", 28),
    ("SUP-003", "Searle Pakistan Ltd", "Karachi", "0654321-8", "30 days", 24),
    ("SUP-004", "Highnoon Laboratories Ltd", "Lahore", "0891234-5", "45 days", 35),
    ("SUP-005", "AGP Limited", "Karachi", "0456789-1", "30 days", 26),
    ("SUP-006", "Ferozsons Laboratories Ltd", "Rawalpindi", "0234567-8", "45 days", 38),
    ("SUP-007", "Surgimed Pakistan", "Karachi", "0367890-2", "15 days", 18),
    ("SUP-008", "Medisure Distributors", "Karachi", "0523456-7", "30 days", 22),
    ("SUP-009", "Lab Sciences & Services", "Karachi", "0678123-4", "30 days", 30),
    ("SUP-010", "NutriCare Importers", "Karachi", "0945678-9", "60 days", 55),
    ("SUP-011", "Al-Hamd Traders (General Items)", "Karachi", "0198765-3", "15 days", 12),
    ("SUP-012", "Pakistan Medical Equipment Co.", "Islamabad", "0823456-1", "45 days", 42),
]

warehouses = [
    ("WH-01", "Central Medical Store - Karachi", "Karachi", "Central warehouse"),
    ("WH-02", "Korangi Project Store", "Karachi - Korangi", "Project store"),
    ("WH-03", "Malir Project Store", "Karachi - Malir", "Project store"),
    ("WH-04", "Lyari Project Store", "Karachi - Lyari", "Project store"),
    ("WH-05", "Tharparkar Field Store", "Tharparkar", "Field store"),
    ("WH-06", "Dadu Field Store", "Dadu", "Field store"),
    ("WH-07", "Jamshoro Field Store", "Jamshoro", "Field store"),
    ("WH-08", "Bin Qasim CHW Hub Store", "Karachi - Bin Qasim", "Project store"),
]

CATEGORY_SUPPLIER = {
    "Medicines": ["SUP-001", "SUP-002", "SUP-003", "SUP-004", "SUP-005", "SUP-006"],
    "Medical Consumables": ["SUP-007", "SUP-008", "SUP-002"],
    "Laboratory": ["SUP-009", "SUP-008"],
    "Nutrition": ["SUP-010", "SUP-002"],
    "Equipment": ["SUP-012", "SUP-008"],
    "General Items": ["SUP-011", "SUP-008"],
}

BUYERS = ["A. Rehman", "S. Fatima", "M. Iqbal", "H. Ali"]
SERVICE_POINTS = {
    "PRJ-001": ["OPD", "Antenatal Clinic", "Labour Room", "EPI Corner", "Pharmacy"],
    "PRJ-002": ["NSC Ward", "OTP Corner", "Stabilisation Ward", "Follow-up Clinic"],
    "PRJ-003": ["DOTS Centre", "Sputum Lab", "Contact Tracing"],
    "PRJ-004": ["BHU OPD", "Mobile Clinic", "EPI Corner", "MNCH Corner"],
    "PRJ-005": ["Mobile Health Unit", "Relief Camp Clinic", "Dispensary"],
    "PRJ-006": ["RHC OPD", "Emergency Room", "MNCH Corner"],
    "PRJ-007": ["Community Outreach", "CHW Kit Replenishment"],
    "PRJ-008": ["EmOC Triage", "Labour Room", "Operation Theatre"],
    "PRJ-009": ["Screening Clinic", "Lab Corner"],
    "PRJ-010": ["Thalassaemia Day Care", "Transfusion Bay", "Pharmacy"],
}

# which categories each project draws on and how heavily
PROJECT_PROFILE = {
    "PRJ-001": {"Medicines": 1.0, "Medical Consumables": 1.0, "Laboratory": 0.5, "Nutrition": 0.55, "Equipment": 0.05, "General Items": 0.35},
    "PRJ-002": {"Medicines": 0.65, "Medical Consumables": 0.6, "Laboratory": 0.35, "Nutrition": 1.35, "Equipment": 0.06, "General Items": 0.3},
    "PRJ-003": {"Medicines": 0.85, "Medical Consumables": 0.55, "Laboratory": 0.95, "Nutrition": 0.15, "Equipment": 0.03, "General Items": 0.3},
    "PRJ-004": {"Medicines": 0.85, "Medical Consumables": 0.75, "Laboratory": 0.3, "Nutrition": 0.7, "Equipment": 0.07, "General Items": 0.4},
    "PRJ-005": {"Medicines": 0.8, "Medical Consumables": 0.5, "Laboratory": 0.15, "Nutrition": 0.45, "Equipment": 0.04, "General Items": 0.45},
    "PRJ-006": {"Medicines": 0.7, "Medical Consumables": 0.7, "Laboratory": 0.25, "Nutrition": 0.4, "Equipment": 0.07, "General Items": 0.35},
    "PRJ-007": {"Medicines": 0.5, "Medical Consumables": 0.55, "Laboratory": 0.1, "Nutrition": 0.5, "Equipment": 0.02, "General Items": 0.55},
    "PRJ-008": {"Medicines": 0.9, "Medical Consumables": 0.95, "Laboratory": 0.3, "Nutrition": 0.2, "Equipment": 0.05, "General Items": 0.3},
    "PRJ-009": {"Medicines": 0.35, "Medical Consumables": 0.4, "Laboratory": 1.0, "Nutrition": 0.05, "Equipment": 0.03, "General Items": 0.25},
    "PRJ-010": {"Medicines": 1.0, "Medical Consumables": 0.6, "Laboratory": 0.7, "Nutrition": 0.1, "Equipment": 0.04, "General Items": 0.3},
}

# item-level demand weight (units per month at project scale 1.0)
ITEM_DEMAND = {
    "MED-0001": 9500, "MED-0002": 4200, "MED-0003": 260, "MED-0004": 420, "MED-0005": 2600,
    "MED-0006": 520, "MED-0007": 5400, "MED-0008": 3200, "MED-0009": 380, "MED-0010": 320,
    "MED-0011": 260, "MED-0012": 320, "MED-0013": 90, "MED-0014": 110, "MED-0015": 55,
    "MED-0016": 6800, "MED-0017": 3100, "MED-0018": 700, "MED-0019": 900, "MED-0020": 600,
    "MED-0021": 420, "MED-0022": 400, "MED-0023": 2200, "MED-0024": 210, "MED-0025": 160,
    "MED-0026": 70, "MED-0027": 55, "MED-0028": 45, "MED-0029": 240, "MED-0030": 95,
    "MED-0031": 130, "MED-0032": 5400, "MED-0033": 900, "MED-0034": 70,
    "CON-0001": 7800, "CON-0002": 700, "CON-0003": 5600, "CON-0004": 2900, "CON-0005": 520,
    "CON-0006": 300, "CON-0007": 110, "CON-0008": 6400, "CON-0009": 260, "CON-0010": 3600,
    "CON-0011": 26, "CON-0012": 4200, "CON-0013": 160, "CON-0014": 180, "CON-0015": 145,
    "CON-0016": 1300, "CON-0017": 75, "CON-0018": 45, "CON-0019": 120, "CON-0020": 950,
    "LAB-0001": 14, "LAB-0002": 620, "LAB-0003": 380, "LAB-0004": 360, "LAB-0005": 480,
    "LAB-0006": 1150, "LAB-0007": 900, "LAB-0008": 6,
    "NUT-0001": 1650, "NUT-0002": 210, "NUT-0003": 230, "NUT-0004": 950, "NUT-0005": 2300,
    "NUT-0006": 780, "NUT-0007": 60,
    "EQP-0001": 1.2, "EQP-0002": 1.0, "EQP-0003": 4.0, "EQP-0004": 1.5, "EQP-0005": 2.0,
    "EQP-0006": 2.5, "EQP-0007": 1.2, "EQP-0008": 0.6, "EQP-0009": 0.4, "EQP-0010": 0.8,
    "GEN-0001": 1100, "GEN-0002": 130, "GEN-0003": 14, "GEN-0004": 22, "GEN-0005": 140,
    "GEN-0006": 90, "GEN-0007": 1900, "GEN-0008": 260,
}

# seasonal multipliers by calendar month for communicable / nutrition items
SEASON = {
    1: 1.25, 2: 1.15, 3: 1.00, 4: 0.90, 5: 0.85, 6: 0.95,
    7: 1.15, 8: 1.30, 9: 1.25, 10: 1.05, 11: 0.95, 12: 1.05,
}
NUTRITION_SEASON = {1: 1.35, 2: 1.25, 3: 1.05, 4: 0.90, 5: 0.85, 6: 0.95,
                    7: 1.10, 8: 1.25, 9: 1.30, 10: 1.10, 11: 1.00, 12: 1.15}


def month_scale(y, m):
    """overall demand growth + one-off shocks (outbreak, flood, funding cut)"""
    idx = (y - 2024) * 12 + (m - 1) - 6  # months since Jul-2024
    growth = 1 + 0.004 * idx  # ~5% p.a. programme growth
    shock = 1.0
    if (y, m) in [(2025, 8), (2025, 9)]:
        shock *= 1.28  # monsoon / diarrhoea surge
    if (y, m) in [(2026, 3)]:
        shock *= 0.78  # funding interruption in Q4 FY25-26
    if (y, m) in [(2026, 7)]:
        shock *= 1.12  # scale-up after new grant
    return growth * shock


def project_active(pid, y, m):
    p = next(p for p in projects if p[0] == pid)
    first = date(y, m, 1)
    last = date(y, m, calendar.monthrange(y, m)[1])
    return p[5] <= last and p[6] >= first


# --------------------------------------------------------------------------
# 2. Demand plan -> purchase orders -> receipts -> FEFO consumption
# --------------------------------------------------------------------------
item_by_id = {i[0]: i for i in items}
wc = {w[0]: w for w in warehouses}
PROJECT_WH = {"PRJ-001": "WH-02", "PRJ-002": "WH-03", "PRJ-003": "WH-04", "PRJ-004": "WH-05",
              "PRJ-005": "WH-06", "PRJ-006": "WH-07", "PRJ-007": "WH-08", "PRJ-008": "WH-04",
              "PRJ-009": "WH-01", "PRJ-010": "WH-01"}

po_rows, grn_rows, issue_rows = [], [], []
po_seq = defaultdict(int)
grn_seq = defaultdict(int)
iss_seq = defaultdict(int)
batch_seq = defaultdict(int)

CITY_PREFIX = {"Karachi": "KHI", "Tharparkar": "TPK", "Dadu": "DDU", "Jamshoro": "JMS"}


def fy_code(d):
    return f"FY{(d.year if d.month >= 7 else d.year - 1)}"


def next_po_no(d, buyer_idx):
    key = fy_code(d)
    po_seq[key] += 1
    return f"VPT-PO-{key[2:]}-{po_seq[key]:05d}"


def next_grn_no(d):
    key = fy_code(d)
    grn_seq[key] += 1
    return f"GRN-{key[2:]}-{grn_seq[key]:05d}"


def next_issue_no(d):
    key = f"{d.year}{d.month:02d}"
    iss_seq[key] += 1
    return f"ISS-{key}-{iss_seq[key]:04d}"


def inflate(base, d):
    """annual price inflation of ~9%, compounded from Jul-2024"""
    years = (d - date(2024, 7, 1)).days / 365.25
    return round(base * (1.09 ** years), 2)


def round_to_pack(qty, pack):
    if pack <= 1:
        return max(1, int(round(qty)))
    return max(pack, int(round(qty / pack)) * pack)


DEMAND_SCALE = 0.62

# build demand lines
demand = defaultdict(dict)  # (pid, item) -> {(y,m): qty}
for pid, pname, *_rest in projects:
    scale = next(p for p in projects if p[0] == pid)[9]
    profile = PROJECT_PROFILE[pid]
    for y, m in MONTHS:
        if not project_active(pid, y, m):
            continue
        ms = month_scale(y, m)
        for iid, irow in item_by_id.items():
            cat = irow[2]
            wt = profile.get(cat, 0)
            if wt <= 0:
                continue
            base = ITEM_DEMAND[iid]
            if base < 5:  # equipment: occasional, lumpy
                if random.random() > 0.10 * wt:
                    continue
            else:
                # not every item every month, but most months for core items
                if random.random() > min(0.94, 0.55 + 0.4 * wt):
                    continue
            seas = 1.0
            if cat == "Nutrition":
                seas = NUTRITION_SEASON[m]
            elif cat in ("Medicines", "Medical Consumables", "Laboratory"):
                seas = SEASON[m]
            noise = random.uniform(0.75, 1.3)
            qty = base * DEMAND_SCALE * scale * wt * ms * seas * noise
            if base < 5:
                qty = max(1, round(qty))
            else:
                qty = round_to_pack(qty, irow[5])
            if qty <= 0:
                continue
            demand[(pid, iid)][(y, m)] = qty

# purchase + receipt generation
for (pid, iid), months in demand.items():
    irow = item_by_id[iid]
    pack, base_cost, shelf = irow[5], irow[6], irow[8]
    cat = irow[2]
    wh = PROJECT_WH[pid]
    supplier = random.choice(CATEGORY_SUPPLIER[cat])
    lead = next(s[5] for s in suppliers if s[0] == supplier)
    prev_qty = None
    for (y, m), qty in sorted(months.items()):
        # order 30-45 days ahead, cover 1-2 months of requirement
        po_date = date(y, m, 1) - timedelta(days=random.randint(28, 45))
        if po_date < date(2024, 6, 1):
            po_date = date(2024, 6, 1)
        cover = 1.06 if cat != "Equipment" else 1.0
        order_qty = round_to_pack(qty * cover, pack)
        if prev_qty and random.random() < 0.10:  # order several months at once
            order_qty = round_to_pack(order_qty * random.uniform(1.6, 2.2), pack)
        prev_qty = order_qty
        unit = inflate(base_cost, po_date)
        unit = round(unit * random.uniform(0.94, 1.09), 2)  # market / supplier variance
        if random.random() < 0.004:
            unit = 0.0  # donated stock, priced at zero
        req_no = f"REQ-{po_date.year}-{random.randint(1, 9999):04d}"
        po_no = next_po_no(po_date, 0)
        po_rows.append([
            po_no, po_date, req_no, po_date - timedelta(days=random.randint(3, 12)),
            pid, supplier, iid, irow[1], cat, order_qty, irow[4],
            unit, round(order_qty * unit, 2), "PKR",
            random.choice(["Received"] * 12 + ["Partially Received"] * 4 + ["Cancelled"]),
            random.choice(BUYERS), lead,
        ])

        # receipts: 1-2 GRNs, delivered after lead time (or not at all if cancelled)
        status = po_rows[-1][14]
        if status == "Cancelled":
            continue
        n_grn = random.choice([1, 1, 1, 2])
        remaining = order_qty
        grn_base = po_date + timedelta(days=lead + random.randint(-5, 10))
        for g in range(n_grn):
            if g == n_grn - 1:
                rec = remaining
            else:
                rec = round_to_pack(remaining * random.uniform(0.35, 0.6), pack)
                if rec >= remaining:
                    rec = remaining
            if rec <= 0:
                continue
            remaining -= rec
            g_date = grn_base + timedelta(days=g * random.randint(4, 14))
            key = (f"{iid[:3]}", g_date.strftime("%y%m"))
            batch_seq[key] += 1
            batch = f"{iid[:3]}-{g_date.strftime('%y%m')}-{batch_seq[key]:03d}"
            mfg = g_date - timedelta(days=random.randint(30, 150))
            exp = date(mfg.year + shelf // 12, ((mfg.month + shelf % 12 - 1) % 12) + 1, 1)
            exp = exp - timedelta(days=1)
            # a small share of near-expiry stock, as happens in real stores
            if random.random() < 0.03:
                exp = g_date + timedelta(days=random.randint(45, 150))
            grn_rows.append([
                next_grn_no(g_date), g_date, po_no, pid, iid, irow[1], rec, irow[4],
                wh, batch, mfg, exp, unit, "Accepted",
            ])
            if remaining <= 0:
                break

# opening stock so that Jul-2024 consumption is possible
opening_date = date(2024, 6, 20)
for (pid, iid), months in demand.items():
    first_key = min(months)
    qty = months[first_key]
    irow = item_by_id[iid]
    pack, unit, shelf = irow[5], irow[6], irow[8]
    wh = PROJECT_WH[pid]
    batch_seq[(iid[:3], "2406")] += 1
    batch = f"{iid[:3]}-2406-{batch_seq[(iid[:3], '2406')]:03d}"
    exp = date(2026, 6, 30) if shelf >= 24 else date(2025, 12, 31)
    grn_rows.append([
        next_grn_no(opening_date), opening_date, "OPENING", pid, iid, irow[1],
        round_to_pack(qty * 1.7, pack), irow[4], wh, batch,
        date(2024, 1, 15), exp, inflate(unit, opening_date), "Accepted",
    ])

grn_rows.sort(key=lambda r: r[1])

# ---- FEFO issue engine -------------------------------------------------
stock = defaultdict(list)  # (pid, item) -> list of batch dicts
for g in grn_rows:
    _grn, gdate, _po, pid, iid, _nm, rec, _uom, wh, batch, mfg, exp, cost, _st = g
    stock[(pid, iid)].append({"batch": batch, "qty": float(rec), "exp": exp, "cost": cost,
                              "grn": _grn, "wh": wh})

UNATTR_PROJECT = "GEN-0000"
shortfalls = []

for (pid, iid), months in sorted(demand.items()):
    irow = item_by_id[iid]
    carry = 0.0
    for (y, m), qty in sorted(months.items()):
        need = qty + carry
        carry = 0.0
        batches = sorted(stock[(pid, iid)], key=lambda b: (b["exp"], b["grn"]))
        days_in_month = calendar.monthrange(y, m)[1]
        used_total = 0.0
        split_no = 0
        for b in batches:
            if need <= 0:
                break
            if b["qty"] <= 0:
                continue
            take = min(b["qty"], need)
            # spread the issue across the month
            frac = take / max(need, 0.0001)
            day = min(days_in_month, max(1, int(round(frac * days_in_month * random.uniform(0.5, 1.3)))))
            idate = date(y, m, min(day + split_no * 3, days_in_month))
            b["qty"] -= take
            need -= take
            used_total += take
            split_no += 1
            # deliberate defect: ~2.2% of issue lines lose their project attribution
            charged = pid
            if random.random() < 0.022:
                charged = UNATTR_PROJECT
            sp = random.choice(SERVICE_POINTS.get(pid, ["OPD"]))
            cost = b["cost"] if b["cost"] else inflate(irow[6], idate)
            issue_rows.append([
                next_issue_no(idate), idate, charged, pid, iid, irow[1], cat_of := irow[2],
                take, irow[4], cost, round(take * cost, 2), b["wh"], b["grn"], b["batch"],
                b["exp"], random.choice(["Project Issue"] * 12 + ["Department Issue"] * 2
                                         + ["Inter-store Transfer"]),
                sp, random.choice(BUYERS),
            ])
        if need > 0:
            shortfalls.append([pid, iid, f"{y}-{m:02d}", round(need, 2)])
            carry = need  # unmet demand rolls into next month

# short-dated / donated receipts near the end of the window: stock that could not be
# consumed in time and now sits in store as expiry risk (added after the FEFO engine, so
# it is deliberately not issued)
AS_AT_LOCAL = max(r[1] for r in issue_rows)
donation_pool = [i for i in items if i[2] in ("Medicines", "Nutrition", "Medical Consumables")]
random.shuffle(donation_pool)
donation_records = []
for n, irow in enumerate(donation_pool[:11]):
    iid = irow[0]
    pid = random.choice([p[0] for p in projects if p[6] >= date(2026, 12, 1)])
    wh = PROJECT_WH[pid]
    rec_date = date(2026, random.choice([7, 8, 9]), random.randint(1, 25))
    base = ITEM_DEMAND[iid]
    qty = round_to_pack(base * random.uniform(1.5, 4.0), irow[5])
    expired = n < 4
    if expired:
        exp = AS_AT_LOCAL - timedelta(days=random.randint(8, 60))
        rec_date = exp - timedelta(days=random.randint(20, 60))
    else:
        exp = AS_AT_LOCAL + timedelta(days=random.randint(12, 150))
    cost = 0.0 if n in (0, 5) else inflate(irow[6], rec_date)
    batch_seq[(iid[:3], rec_date.strftime("%y%m"))] += 1
    batch = f"{iid[:3]}-{rec_date.strftime('%y%m')}-{batch_seq[(iid[:3], rec_date.strftime('%y%m'))]:03d}"
    grn_rows.append([
        next_grn_no(rec_date), rec_date, "DONATION-2026", pid, iid, irow[1], qty, irow[4],
        wh, batch, rec_date - timedelta(days=random.randint(30, 120)), exp, cost, "Accepted",
    ])
    donation_records.append([rec_date, iid, irow[1], qty, batch, exp, cost])

grn_rows.sort(key=lambda r: r[1])

# duplicate GRN defect (8 rows) - typical double posting
dupes = random.sample(range(len(grn_rows)), 8)
for idx in sorted(dupes, reverse=True):
    g = list(grn_rows[idx])
    g[5] = g[5] + " "  # trailing space in item description
    grn_rows.insert(idx + 1, g)

issue_rows.sort(key=lambda r: (r[1], r[0]))
po_rows.sort(key=lambda r: (r[1], r[0]))

# ---- budget & service delivery ----------------------------------------
budgets = []
cats = ["Medicines", "Medical Consumables", "Laboratory", "Nutrition", "Equipment", "General Items", "Human Resource", "Programme Operations"]
weights = [0.27, 0.13, 0.05, 0.14, 0.06, 0.05, 0.22, 0.08]
for pid, pname, donor, dist, fac, sdate, edate, budget, sector, scale in projects:
    for fyr in ["FY2024-25", "FY2025-26", "FY2026-27"]:
        # project budget spreads unevenly across years
        share = {"FY2024-25": 0.42, "FY2025-26": 0.38, "FY2026-27": 0.20}[fyr]
        if sdate > date(2025, 6, 30) and fyr == "FY2024-25":
            continue
        if edate < date(2025, 7, 1) and fyr != "FY2024-25":
            continue
        for c, w in zip(cats, weights):
            b = round(budget * share * w * random.uniform(0.92, 1.08), -3)
            budgets.append([pid, fyr, c, b, "PKR"])

service = []
for pid, pname, donor, dist, fac, sdate, edate, budget, sector, scale in projects:
    for y, m in MONTHS:
        if not project_active(pid, y, m):
            continue
        ms = month_scale(y, m)
        ben = int(scale * random.uniform(2600, 4200) * ms)
        service.append([pid, date(y, m, 1), f"{y}-{m:02d}", f"Jul-{y}" if m >= 7 else f"Jul-{y-1}",
                        ben, int(ben * random.uniform(1.4, 2.3)),
                        int(ben * random.uniform(0.18, 0.34)), sector])

# --------------------------------------------------------------------------
# 3. Workbook
# --------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F3B57")
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FMT = "dd-mmm-yyyy"
MONEY = '#,##0.00'
INT = '#,##0'

wb = Workbook()
wb.remove(wb.active)


def write_sheet(title, headers, rows, widths, numfmt=None, freeze="A2", table_name=None, note=None):
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(row=1, column=1, value=note).font = Font(name="Calibri", size=9, italic=True, color="595959")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(len(headers), 8))
        ws.row_dimensions[1].height = 18
        r0 = 2
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=r0, column=c, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[r0].height = 30
    for ri, row in enumerate(rows, start=r0 + 1):
        for ci, v in enumerate(row, start=1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            if isinstance(v, date):
                cell.number_format = DATE_FMT
            elif numfmt and ci in numfmt:
                cell.number_format = numfmt[ci]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if table_name and rows:
        ref = f"A{r0}:{get_column_letter(len(headers))}{r0 + len(rows)}"
        t = Table(displayName=table_name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
    ws.freeze_panes = f"A{r0 + 1}" if freeze == "A2" else freeze
    ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(headers))}{r0 + len(rows)}"
    ws.sheet_view.showGridLines = False
    return ws


# README
ws = wb.create_sheet("README")
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 100
rows = [
    ("VITAL PAKISTAN TRUST", ""),
    ("Consumption Analysis Dashboard - Sample Dataset", ""),
    ("", ""),
    ("Purpose", "Sample / test dataset for building and validating the Consumption Analysis dashboard (BRD v1.0). Mirrors the structure of Oracle ERP extracts: procurement, receipts and inventory issues by project and product."),
    ("Data type", "SYNTHETIC. Generated to be realistic in volume, seasonality, pricing and behaviour. No real patient, beneficiary or supplier transaction is represented; supplier and organisation names are used only for realism and should be replaced before external circulation."),
    ("Coverage", "27 months: 01-Jul-2024 to 30-Sep-2026 (FY2024-25, FY2025-26 and Q1 FY2026-27)."),
    ("Volume", f"{len(po_rows):,} PO lines | {len(grn_rows):,} GRN lines | {len(issue_rows):,} consumption (issue) lines | {len(budgets):,} budget lines | {len(service):,} service-delivery lines"),
    ("Currency", "PKR (all values, no tax gross-up applied)"),
    ("", ""),
    ("How the data links", "dim_project.ProjectID -> fact_purchase_order.ProjectID / fact_receipts.ProjectID / fact_consumption.ProjectID; fact_purchase_order.PONumber -> fact_receipts.PONumber -> fact_consumption.GRNNumber (batch traceability); dim_item.ItemID and dim_supplier.SupplierID behave as conformed dimensions."),
    ("Stock logic", "Consumption is issued FEFO from received batches, so batch-level stock never goes negative. Unmet demand is carried forward and appears in Data_Quality_Notes as shortfall."),
    ("Timing note", "The last two months (Aug-Sep 2026) contain purchase orders not yet fully received at the extract date, so quantities issued are lower than the underlying demand. This is intentional and demonstrates the 'open commitment / in-transit' reporting requirement (FR-31)."),
    ("Known data defects", "Left in deliberately so that the dashboard's data-quality and exception logic can be demonstrated: see sheet Data_Quality_Notes."),
    ("Sheets", "02_Dim_Donor, 03_Dim_Project, 04_Dim_Item, 05_Dim_Supplier, 06_Dim_Warehouse, 07_Dim_Calendar, 08_Fact_PurchaseOrder, 09_Fact_Receipts, 10_Fact_Consumption, 11_Fact_Budget, 12_Fact_ServiceDelivery, 13_Data_Dictionary, 14_Data_Quality_Notes"),
    ("Refresh in production", "Daily incremental extract from Oracle into staging, then full reload of the reporting model before 07:00 PKT; monthly close snapshot frozen on working day 3."),
]
r = 1
for label, text in rows:
    if label == "VITAL PAKISTAN TRUST":
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(name="Calibri", size=18, bold=True, color="1F3B57")
    elif label == "Consumption Analysis Dashboard - Sample Dataset":
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(name="Calibri", size=13, color="595959")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    else:
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(name="Calibri", size=10, bold=True)
        c.alignment = Alignment(vertical="top")
        t = ws.cell(row=r, column=3, value=text)
        t.font = Font(name="Calibri", size=10)
        t.alignment = Alignment(wrap_text=True, vertical="top")
        if text and len(text) > 110:
            ws.row_dimensions[r].height = 15 * (len(text) // 105 + 1)
    r += 1

write_sheet("02_Dim_Donor",
            ["DonorID", "DonorName", "DonorType", "GrantReference", "Currency", "ReportingFrequency"],
            [list(d) + [random.choice(["Monthly", "Quarterly"])] for d in donors],
            [11, 34, 18, 24, 10, 20], table_name="tblDonor")

write_sheet("03_Dim_Project",
            ["ProjectID", "ProjectName", "DonorID", "DonorName", "Sector", "District", "Facility",
             "StartDate", "EndDate", "Status", "ApprovedBudgetPKR", "DefaultWarehouseID"],
            [[p[0], p[1], p[2], next(d[1] for d in donors if d[0] == p[2]), p[8], p[3], p[4],
              p[5], p[6], "Active" if p[6] >= date(2026, 9, 30) else "Closed",
              p[7], PROJECT_WH[p[0]]] for p in projects],
            [11, 52, 11, 28, 14, 22, 30, 13, 13, 9, 18, 20],
            numfmt={11: INT}, table_name="tblProject")

item_rows = []
for i in items:
    item_rows.append([i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7],
                      round(i[6] * random.uniform(0.25, 0.45), 2),
                      random.choice(["FEFO", "FEFO", "FIFO"]),
                      "Active" if random.random() > 0.04 else "Inactive"])
write_sheet("04_Dim_Item",
            ["ItemID", "ItemDescription", "Category", "SubCategory", "UOM", "PackSize",
             "StandardUnitCostPKR", "IsConsumable", "ReorderLevel", "IssueMethod", "ItemStatus"],
            item_rows, [11, 44, 22, 24, 10, 9, 18, 12, 14, 11, 11],
            numfmt={7: MONEY, 9: INT}, table_name="tblItem")

supplier_rows = [list(s) + ["Active"] for s in suppliers]
supplier_rows[9][6] = "Blacklisted"   # deliberate: a blacklisted supplier still present in PO history
write_sheet("05_Dim_Supplier",
            ["SupplierID", "SupplierName", "City", "NTN", "PaymentTerms", "LeadTimeDays", "SupplierStatus"],
            supplier_rows,
            [11, 40, 14, 14, 15, 14, 16], table_name="tblSupplier")

write_sheet("06_Dim_Warehouse",
            ["WarehouseID", "WarehouseName", "District", "WarehouseType", "InCharge"],
            [list(w) + [random.choice(["Store Officer", "Pharmacist", "Project Coordinator"])] for w in warehouses],
            [13, 40, 24, 20, 20], table_name="tblWarehouse")

cal_rows = []
for y, m in MONTHS:
    d = date(y, m, 1)
    cal_rows.append([d, d.strftime("%b-%Y"), y, m, d.strftime("%B"), f"Q{((m - 1) % 3) + 1}",
                     FY(y, m), "Month-end close WD3" if m else "", f"{y}-{m:02d}"])
write_sheet("07_Dim_Calendar",
            ["MonthStartDate", "MonthLabel", "Year", "MonthNumber", "MonthName", "Quarter",
             "FiscalYear", "CloseConvention", "PeriodKey"],
            cal_rows, [15, 12, 7, 13, 12, 8, 12, 22, 11], table_name="tblCalendar")

write_sheet("08_Fact_PurchaseOrder",
            ["PONumber", "PODate", "RequisitionNo", "RequisitionDate", "ProjectID", "SupplierID",
             "ItemID", "ItemDescription", "Category", "QuantityOrdered", "UOM", "UnitPricePKR",
             "LineTotalPKR", "Currency", "POStatus", "Buyer", "LeadTimeDays"],
            po_rows,
            [20, 13, 15, 15, 11, 11, 11, 40, 22, 15, 10, 13, 15, 9, 18, 13, 12],
            numfmt={10: INT, 12: MONEY, 13: MONEY}, table_name="tblPO")

write_sheet("09_Fact_Receipts",
            ["GRNNumber", "GRNDate", "PONumber", "ProjectID", "ItemID", "ItemDescription",
             "QuantityReceived", "UOM", "WarehouseID", "BatchNumber", "ManufactureDate",
             "ExpiryDate", "UnitCostPKR", "QCStatus"],
            grn_rows,
            [18, 13, 20, 11, 11, 40, 16, 10, 13, 18, 15, 13, 13, 11],
            numfmt={7: INT, 13: MONEY}, table_name="tblGRN")

write_sheet("10_Fact_Consumption",
            ["IssueID", "IssueDate", "ProjectIDCharged", "ProjectIDOriginal", "ItemID",
             "ItemDescription", "Category", "QuantityIssued", "UOM", "UnitCostPKR",
             "TotalCostPKR", "WarehouseID", "GRNNumber", "BatchNumber", "ExpiryDate",
             "IssueType", "ServicePoint", "IssuedBy"],
            issue_rows,
            [18, 13, 16, 16, 11, 40, 22, 15, 10, 13, 15, 13, 18, 18, 13, 20, 26, 12],
            numfmt={8: INT, 10: MONEY, 11: MONEY}, table_name="tblConsumption")

write_sheet("11_Fact_Budget",
            ["ProjectID", "FiscalYear", "ExpenseCategory", "BudgetPKR", "Currency"],
            budgets, [11, 12, 24, 16, 10], numfmt={4: INT}, table_name="tblBudget")

write_sheet("12_Fact_ServiceDelivery",
            ["ProjectID", "MonthStartDate", "PeriodKey", "FiscalYear", "BeneficiariesServed",
             "PatientVisits", "Admissions", "Sector"],
            service, [11, 15, 11, 12, 20, 14, 12, 16],
            numfmt={5: INT, 6: INT, 7: INT}, table_name="tblService")

# data dictionary
dd = [
    ("02_Dim_Donor", "DonorID", "Unique donor / grant holder", "Text", "DNR-01", "PK"),
    ("02_Dim_Donor", "DonorType", "Multilateral / Government / Corporate / Own source", "Text", "", ""),
    ("03_Dim_Project", "ProjectID", "Unique programme/project code; the principal reporting dimension", "Text", "PRJ-001", "PK"),
    ("03_Dim_Project", "ApprovedBudgetPKR", "Total approved project budget across life of project", "Number", "92,500,000", ""),
    ("03_Dim_Project", "Status", "Active / Closed as at reporting date", "Text", "", ""),
    ("04_Dim_Item", "ItemID", "Unique product code", "Text", "MED-0001", "PK"),
    ("04_Dim_Item", "Category", "Medicines / Medical Consumables / Laboratory / Nutrition / Equipment / General Items", "Text", "", ""),
    ("04_Dim_Item", "PackSize", "Units per pack; ordering is rounded to pack multiples", "Number", "1000", ""),
    ("04_Dim_Item", "StandardUnitCostPKR", "Reference cost per unit used for budgeting and variance", "Number", "1.15", ""),
    ("04_Dim_Item", "IsConsumable", "Yes = consumed and re-orderable; No = capital equipment", "Text", "", ""),
    ("08_Fact_PurchaseOrder", "PONumber", "Purchase order number; grain = one PO line", "Text", "VPT-PO-24-00042", "PK"),
    ("08_Fact_PurchaseOrder", "LineTotalPKR", "QuantityOrdered x UnitPricePKR", "Number", "", "Derived"),
    ("08_Fact_PurchaseOrder", "POStatus", "Received / Partially Received / Cancelled", "Text", "", ""),
    ("09_Fact_Receipts", "GRNNumber", "Goods receipt note; grain = one batch receipt line", "Text", "GRN-24-00117", "PK"),
    ("09_Fact_Receipts", "BatchNumber", "Batch / lot number for traceability", "Text", "MED-2407-004", ""),
    ("09_Fact_Receipts", "ExpiryDate", "Batch expiry date - drives FEFO and expiry risk reporting", "Date", "31-Jul-2027", ""),
    ("10_Fact_Consumption", "IssueID", "Inventory issue; grain = one item x batch x issue date", "Text", "ISS-202507-0031", "PK"),
    ("10_Fact_Consumption", "ProjectIDCharged", "Project the cost was actually charged to (may be GEN-0000 if unattributed)", "Text", "", ""),
    ("10_Fact_Consumption", "ProjectIDOriginal", "Project that physically consumed the item - used for data-quality correction", "Text", "", ""),
    ("10_Fact_Consumption", "QuantityIssued", "Quantity issued from store, in the item's base UOM", "Number", "", ""),
    ("10_Fact_Consumption", "UnitCostPKR", "Weighted / batch cost at time of issue", "Number", "", ""),
    ("10_Fact_Consumption", "TotalCostPKR", "QuantityIssued x UnitCostPKR - the consumption value used in all KPIs", "Number", "", "Derived"),
    ("10_Fact_Consumption", "IssueType", "Project Issue / Department Issue / Inter-store Transfer", "Text", "", ""),
    ("11_Fact_Budget", "ExpenseCategory", "Budget line category, aligned to item categories plus HR and operations", "Text", "", ""),
    ("12_Fact_ServiceDelivery", "BeneficiariesServed", "Unique beneficiaries served in the month (programme M&E data)", "Number", "", ""),
]
write_sheet("13_Data_Dictionary",
            ["Sheet", "Field", "Description", "Data Type", "Example", "Key"],
            [list(x) for x in dd], [26, 26, 74, 12, 22, 10], table_name="tblDataDictionary")

dq = [
    ["DQ-01", "Unattributed project on consumption", "fact_consumption",
     "Issue lines charged to GEN-0000 instead of a project",
     f"{sum(1 for r in issue_rows if r[2] == 'GEN-0000'):,} lines",
     "Cannot be attributed to a project; shown as 'Unattributed' in all project views. ProjectIDOriginal holds the true project for correction."],
    ["DQ-02", "Zero-value purchase lines", "fact_purchase_order",
     "In-kind / donated stock received at zero price",
     f"{sum(1 for r in po_rows if r[11] == 0):,} lines",
     "Distorts average unit price; exclude from price-variance calculations."],
    ["DQ-03", "Duplicate GRN posting", "fact_receipts",
     "Same GRN and batch posted twice with a trailing space in the item description",
     f"{len(dupes)} lines", "Inflates receipts and stock; deduplicate on GRNNumber + BatchNumber."],
    ["DQ-04", "Unmet demand / stock-out", "issue engine",
     "Demand that could not be issued because no stock was available; carried to the next month",
     f"{len(shortfalls):,} project-item-months",
     "Evidence of stock-out; drives the 'stock-out risk' and 'service interruption' metrics."],
    ["DQ-05", "Near-expiry batches", "fact_receipts",
     "Batches with expiry inside 150 days of receipt",
     f"{sum(1 for g in grn_rows if isinstance(g[11], date) and 0 <= (g[11] - g[1]).days <= 150):,} lines",
     "Feeds the expiry-risk report; action required by store in-charge."],
    ["DQ-06", "Cancelled purchase orders", "fact_purchase_order",
     "Orders cancelled after approval",
     f"{sum(1 for r in po_rows if r[14] == 'Cancelled'):,} lines",
     "Exclude from commitment and investment calculations."],
    ["DQ-07", "Inactive items still in master", "dim_item", "Item flagged inactive but present in history",
     f"{sum(1 for r in item_rows if r[10] == 'Inactive')} items", "Requires master-data clean-up before go-live."],
    ["DQ-08", "Blacklisted suppliers with transactions", "dim_supplier",
     "Supplier flagged blacklisted but still appears in PO history",
     f"{sum(1 for s in supplier_rows if s[6] == 'Blacklisted')} supplier(s)",
     "Blacklisted suppliers should not receive new purchase orders; procurement review required."],
]
ws = write_sheet("14_Data_Quality_Notes",
                 ["Ref", "Issue", "Sheet", "Description", "Volume in this dataset", "Dashboard handling"],
                 [list(x) for x in dq], [8, 34, 20, 60, 22, 66], table_name="tblDQ")
for row in ws.iter_rows(min_row=2, max_row=2 + len(dq), min_col=6, max_col=6):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")

# shortfall detail as its own sheet (useful for exception testing)
write_sheet("15_Stockout_Detail",
            ["ProjectID", "ItemID", "MonthKey", "UnmetQuantity"],
            shortfalls[:5000], [11, 11, 11, 15], numfmt={4: INT}, table_name="tblStockout")

wb.save(OUT)

import json
billable = sum(r[10] for r in issue_rows if r[15] != "Inter-store Transfer")
unattr = sum(r[10] for r in issue_rows if r[2] == "GEN-0000")
prod_cats = ["Medicines", "Medical Consumables", "Laboratory", "Nutrition", "Equipment", "General Items"]
prod_budget = sum(b[3] for b in budgets if b[2] in prod_cats)
summary = {
    "start": START.isoformat(), "end": END.isoformat(), "as_at": max(r[1] for r in issue_rows).isoformat(),
    "months": len(MONTHS),
    "consumption_all": round(sum(r[10] for r in issue_rows), 2),
    "consumption_billable": round(billable, 2),
    "purchase_value": round(sum(r[12] for r in po_rows if r[14] != "Cancelled"), 2),
    "budget_total": round(sum(b[3] for b in budgets), 2),
    "budget_product": round(prod_budget, 2),
    "unattributed_value": round(unattr, 2),
    "unattributed_pct": round(unattr / billable * 100, 2),
    "beneficiaries": int(sum(r[4] for r in service)),
    "po_lines": len(po_rows), "grn_lines": len(grn_rows), "issue_lines": len(issue_rows),
    "shortfalls": len(shortfalls),
    "projects": len(projects), "items": len(items), "suppliers": len(suppliers),
}
with open(os.path.join(OUT_DIR, "summary.json"), "w") as f:
    json.dump(summary, f, indent=2)
print(json.dumps(summary, indent=2))

print("saved", OUT)
print("PO lines", len(po_rows), "GRN lines", len(grn_rows), "Issues", len(issue_rows))
print("consumption value", round(sum(r[10] for r in issue_rows), 2))
print("purchase value", round(sum(r[12] for r in po_rows), 2))
