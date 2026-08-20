"""
Build a paste-ready IN_Shock block for a household-expenditure shock across all
nine jurisdictions.

    python scripts/make_super_template.py

The 32 categories supplied are Household Expenditure Classification headings,
not IOIG codes, so they have to be concorded before the model can strip them.

IMPORTANT: this concordance is expert judgement, NOT the ABS one. The
authoritative source is 'Industry and Product Concordance Tables 2023-24.xlsx'
from the ABS Input-Output *Methodology* page (not Data downloads), which carries
IOPG-to-Household-Expenditure-Classification and IOPC-to-HFCE tables. CLAUDE.md
is explicit that building bridging matrices by hand without it is guesswork.
Treat every weight below as a starting assumption to be replaced, and record it
on the Assumptions tab of any report.

Weights within a category sum to 1.0. All lines are household consumption, so
the column group is Q1_HFCE throughout.
"""
import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'output' / 'Super_shock_IN_Shock_template.csv'
REGIONS = ['NSW', 'Vic', 'QLD', 'SA', 'WA', 'Tas', 'NT', 'ACT', 'Aus']

# category -> [(IOIG code, IOIG name, weight), ...]
CONC = {
 'Other recreation durables, gardens and pets': [
    ('2502', 'Other manufactured products', 0.55),
    ('0103', 'Other agriculture', 0.25),
    ('1109', 'Other food product manufacturing', 0.20)],
 'Financial and other services': [
    ('6201', 'Finance', 0.70), ('6401', 'Auxiliary finance and insurance services', 0.30)],
 'Personal effects': [('2502', 'Other manufactured products', 1.00)],
 'Household textiles, glassware and utensils': [
    ('1303', 'Textile product manufacturing', 0.45),
    ('2001', 'Glass and glass product manufacturing', 0.20),
    ('2002', 'Ceramic product manufacturing', 0.15),
    ('2204', 'Other fabricated metal product manufacturing', 0.20)],
 'Tools and equipment for house and garden': [
    ('2204', 'Other fabricated metal product manufacturing', 0.60),
    ('2405', 'Specialised and other machinery and equipment manufacturing', 0.40)],
 'Postal services': [('5101', 'Postal and courier pick-up and delivery service', 1.00)],
 'Education': [
    ('8210', 'Arts, sports, adult and other education services', 0.50),
    ('8110', 'Technical, vocational and tertiary education services', 0.35),
    ('8010', 'Primary and secondary education services', 0.15)],
 'Vehicle operation (fuel, servicing, etc.)': [
    ('1701', 'Petroleum and coal product manufacturing', 0.55),
    ('9401', 'Automotive repair and maintenance', 0.45)],
 'Routine household maintenance': [
    ('7310', 'Building cleaning, pest control and other support services', 0.60),
    ('9402', 'Other repair and maintenance', 0.40)],
 'Clothing and footwear': [
    ('1305', 'Clothing manufacturing', 0.80), ('1306', 'Footwear manufacturing', 0.20)],
 'Newspapers, books and stationery': [
    ('5401', 'Publishing (except internet and music publishing)', 0.70),
    ('1502', 'Paper stationery and other converted paper product manufacturing', 0.30)],
 'Transport services, holidays and accommodation': [
    ('4401', 'Accommodation', 0.40), ('4901', 'Air and space transport', 0.35),
    ('4601', 'Road transport', 0.15),
    ('7210', 'Employment, travel agency and other administrative services', 0.10)],
 'Medical products and appliances': [
    ('1801', 'Human pharmaceutical and medicinal product manufacturing', 1.00)],
 'Recreational and cultural services': [
    ('9101', 'Sports and recreation', 0.40),
    ('8901', 'Heritage, creative and performing arts', 0.25),
    ('9201', 'Gambling', 0.25), ('5501', 'Motion picture and sound recording', 0.10)],
 'Furniture and furnishings': [('2501', 'Furniture manufacturing', 1.00)],
 'Audio-visual, photographic and IT equipment': [
    ('2401', 'Professional, scientific, computer and electronic equipment manufacturing', 1.00)],
 'Personal care': [
    ('9501', 'Personal services', 0.55),
    ('1804', 'Cleaning compounds and toiletry preparation manufacturing', 0.45)],
 'Rent': [('6701', 'Actual rent for housing', 1.00)],
 'Catering services': [('4501', 'Food and beverage services', 1.00)],
 'Telephone and internet services': [
    ('5801', 'Telecommunication services', 0.85),
    ('5701', 'Internet service providers and data processing', 0.15)],
 'Dwelling repairs/maintenance/improvements': [
    ('3201', 'Construction services', 0.65),
    ('3001', 'Residential building construction', 0.35)],
 'Electricity, gas and other fuels': [
    ('2605', 'Electricity transmission, distribution, on selling', 0.70),
    ('2701', 'Gas supply', 0.30)],
 'Health services': [('8401', 'Health care services', 1.00)],
 'Food and non-alcoholic beverages': [
    ('1109', 'Other food product manufacturing', 0.28),
    ('1101', 'Meat and meat product manufacturing', 0.18),
    ('1104', 'Fruit and vegetable product manufacturing', 0.12),
    ('1103', 'Dairy product manufacturing', 0.10),
    ('1107', 'Bakery product manufacturing', 0.10),
    ('1201', 'Soft drinks, cordials and syrup manufacturing', 0.10),
    ('1106', 'Grain mill and cereal product manufacturing', 0.06),
    ('1108', 'Sugar and confectionery manufacturing', 0.06)],
 'Household appliances': [('2404', 'Domestic appliance manufacturing', 1.00)],
 'Tobacco': [('1205', 'Wine, spirits and tobacco', 1.00)],
 'Insurance (cash-outlay basis)': [
    ('6301', 'Insurance and superannuation funds', 1.00)],
 'Social protection and aged care': [
    ('8601', 'Residential care and social assistance services', 1.00)],
 'Alcoholic beverages': [
    ('1205', 'Wine, spirits and tobacco', 0.60), ('1202', 'Beer manufacturing', 0.40)],
 'Rates, water and body corporate': [
    ('2801', 'Water supply, sewerage and drainage services', 0.45),
    ('7501', 'Public administration and regulatory services', 0.35),
    ('6702', 'Non-residential property operators and real estate services', 0.20)],
 'Motor vehicle purchase': [
    ('2301', 'Motor vehicles and parts; other transport equipment manufacturing', 1.00)],
 'Telephone equipment': [
    ('2401', 'Professional, scientific, computer and electronic equipment manufacturing', 1.00)],
}


def main():
    bad = [c for c, rows in CONC.items() if abs(sum(w for _, _, w in rows) - 1) > 1e-9]
    if bad:
        raise SystemExit(f'weights do not sum to 1: {bad}')
    per_region = sum(len(v) for v in CONC.values())
    OUT.parent.mkdir(exist_ok=True)
    with open(OUT, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Region', 'Description', 'IOIG code', 'Column group', 'Direct only',
                    'Year 1 $m', 'HEC category', 'Weight'])
        for reg in REGIONS:
            for cat, rows in CONC.items():
                for code, name, wt in rows:
                    w.writerow([reg, f'{cat} - {name}', code, 'Q1_HFCE', 'N', '',
                                cat, wt])
    print(f'{len(CONC)} categories -> {per_region} IOIG lines per jurisdiction')
    print(f'{len(REGIONS)} jurisdictions -> {per_region * len(REGIONS)} rows')
    print(f'wrote {OUT}')
    print('\nCategories needing a split (weights are judgement, replace with the ABS concordance):')
    for cat, rows in CONC.items():
        if len(rows) > 1:
            print(f'  {cat}\n      ' + ', '.join(f'{c} {w:.0%}' for c, _, w in rows))


if __name__ == '__main__':
    main()
