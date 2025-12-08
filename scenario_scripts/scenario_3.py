
import pandas as pd
from tqdm import tqdm
import time

# --- Start timer ---
start_time = time.time()

# --- Constants ---
PERCENT_NEW = 0.65

input_path = "new/Bidsheet Master Consolidate Landed 12052025.csv"

output_file = 'scenario_outputs/scenario 3 12052025.xlsx'

incumbent_col = "Normalized incumbent supplier"
valid_supplier_col = "Valid Supplier"
volume_col = "Annual Volume (per UOM)"
supplier_port_file = "Supplier Port per Part table 070925.csv"
freight_file = "Freight cost mutipliers table 071025v2.csv"
port_country_map = {
    'DALIAN': 'China', 
    'NINGBO': 'China', 
    'QINGDAO': 'China', 
    'QINGDAO2': 'China', 
    'SHANGHAI': 'China',
    'SHENZHEN': 'China', 
    'TIANJIN': 'China', 
    'XINGANG': 'China', 
    'XIAMEN': 'China',
    'AHMEDABAD': 'India', 
    'CHENNAI': 'India', 
    'DADRI': 'India', 
    'MUMBAI': 'India',
    'MUNDRA': 'India', 
    'NHAVA SHEVA': 'India',
    'SURABAYA': 'Indonesia',
    'PORT KLANG': 'Malaysia', 
    'PASIR GUDANG': 'Malaysia', 
    'TANJUNG PELAPAS': 'Malaysia',
    'BUSAN': 'South Korea',
    'KAOHSIUNG': 'Taiwan', 
    'KEELUNG': 'Taiwan', 
    'TAICHUNG': 'Taiwan', 
    'TAIPEI': 'Taiwan',
    'BANGKOK': 'Thailand',
    'LAEM CHABANG': 'Thailand',
    'HO CHI MINH CITY': 'Vietnam', 
    'VUNG TAU': 'Vietnam', 
    'HAI PHONG': 'Vietnam',
    'VIRGINIA': 'India'
}

freight_df = pd.read_csv(freight_file) 
supplier_port_df = pd.read_csv(supplier_port_file)
tariff_df_2 = pd.read_csv("tariff_part_level_cleaned.csv")


# Set indices for fast lookup
tariff_df_2.set_index('ROW ID #', inplace=True)
supplier_port_df.set_index('ROW ID #', inplace=True)
freight_df.set_index('Reference', inplace=True)



def get_supplier_info(row_id, supplier):
    """
    Returns a dict with:
    - Division
    - Port
    - FreightMultiplier
    - Tariff values (tariff_value, Metal Tariff, Metal Type, Country)
    
    Returns None if any info is missing.
    """
    try:
        # 1️⃣ Get supplier port and division from supplier_port_df
        supplier_row = supplier_port_df.loc[row_id]
        division = supplier_row['Division']
        port = supplier_row[supplier]
        country = port_country_map[port]
        # 2️⃣ Get freight multiplier from freight_df
        freight_multiplier = freight_df.loc[port, division]

        # 3️⃣ Get tariff info from tariff_df_2
        tariff_rows = tariff_df_2.loc[row_id]
        if isinstance(tariff_rows, pd.Series):
            tariff_row = tariff_rows
        else:
            # Filter to match the country
            tariff_row = tariff_rows[tariff_rows['Country'] == country]
            if tariff_row.empty:
                return None
            tariff_row = tariff_row.iloc[0]

        tariff_value = float(tariff_row['tariff_value'])
        metal_tariff = float(tariff_row['Metal Tariff'])
        metal_type = tariff_row['Metal Type']
        
        return {
            'row_id': row_id,
            'Division': division,
            'Port': port,
            'FreightMultiplier': float(freight_multiplier),
            'tariff_value': float(tariff_value),
            'Metal Tariff': float(metal_tariff),
            'Metal Type': metal_type,
            'Country': country
        }

    except KeyError:
        # Any missing row or column returns None
        return None


# --- Load files ---
output_reference_file_path = "new/outout-reference.csv"

print("Reading:", input_path)
df = pd.read_csv(input_path)
output_reference_df = pd.read_csv(output_reference_file_path)
print(f"Loaded {len(df)} rows\n")
# Calculate TOTAL_COST from actual data
TOTAL_COST = df['Landed Extended Cost USD'].sum()
THRESHOLD_COST = TOTAL_COST * PERCENT_NEW

print(TOTAL_COST)
print(f"Calculated TOTAL_COST from input data: ${TOTAL_COST:,.2f}")
print(f"THRESHOLD_COST ({PERCENT_NEW*100}%): ${THRESHOLD_COST:,.2f}")

# --- Identify R2 landed cost columns ---
r2_fob_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]

# --- Prepare suppliers ---
incumbent_suppliers = df[incumbent_col].unique()
suppliers = [col.split(" - R2")[0] for col in r2_fob_cols]

# --- PART ASSIGNMENT LOGIC (HEAVILY COMMENTED) ---
# We process all rows and classify them into:
#   1. No valid suppliers: Not awarded.
#   2. Incumbent did not bid, but minimum bid exists: Assign to min bid (contributes to 65% threshold).
#   3. Incumbent bid:
#       a. Incumbent is the minimum: Retain incumbent (does NOT contribute to 65% threshold).
#       b. Incumbent is NOT the minimum: Assign to min bid (contributes to 65% threshold).

no_valid_supplier_parts = []
must_assign_min_bid_parts = []  # Forced to min bid, always contributes to threshold
incumbent_retained_parts = []
candidate_new_supplier_parts = []  # Eligible for threshold assignment
net_new_supplier_list = set()
total_cost_not_awarded = 0

MANEK_EXTRA_VOLUME = 0
PUSHTI_EXTRA_VOLUME = 0

PARTS_NOT_TO_ASSIGN_TO_MANEK = ['619', '13908', '618', '13907', '620', '13909', '621', '13910', '13911', '622', '13912', '13913', '5574', '623', '13914', '5575', '5576', '624', '5577', '5578', '5579', '5634', '596', '13889', '13893', '13895', '594', '13887', '595', '13888', '599', '13892', '13894', '600', '13896', '593', '5562', '597', '13890', '598', '13891', '7902', '12297', '5564', '12298', '603', '5565', '584', '5554', '585', '13870', '586', '13871', '13872', '587', '13873', '13875', '13876', '13877', '5556', '588', '5557', '5548', '581', '13867', '582', '13868', '5550', '583', '5551', '5552', '628', '5589', '5590', '629', '5591', '630', '5592', '5593', '631', '5594', '12326', '5596', '5597', '5598', '5599', '5600', '5601', '5602', '5603', '5604', '13915', '13916', '5580', '13917', '5581', '625', '5582', '626', '5583', '5584', '627', '5585', '5586', '5588', '632', '5605', '5606', '633', '5607', '634', '5608', '635', '5609', '636', '5610', '637', '5611', '5612', '13515', '639', '5613', '640', '5614', '5615', '641', '5616', '642', '5617', '643', '5618', '5619', '644', '5620', '5621', '5635', '645', '5622', '5623', '5624', '5626', '646', '5627', '647', '5628', '648', '5629', '5630', '5631', '5632', '612', '13904', '605', '13897', '606', '13898', '604', '5566', '607', '13899', '608', '13900', '609', '13901', '610', '13902', '611', '13903', '613', '13905', '614', '13906', '5567', '615', '5568', '5569', '13279', '617', '5570', '5571', '5572', '13878', '13879', '13880', '590', '13881', '591', '13883', '13884', '13885', '592', '13886', '589', '5560', '13882', '5633', '12518', '9839']

SPEND_ON_MANEK = 0

# for idx, row in df.iterrows():
#     if str(row['ROW ID #']) in PARTS_NOT_TO_ASSIGN_TO_MANEK:
#         df.at[idx, 'Manek Metalcraft - R2 - Total landed cost per UOM (USD)'] = 0

for idx, row in df.iterrows():
    incumbent = row.get(incumbent_col)
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
    second_min_supplier = row.get("2nd Lowest Bid Landed Supplier")
    valid_supplier_count = row.get(valid_supplier_col, 0)

    # forced incumbents:
    # if row["ROW ID #"] in [9704, 11781, 8462]:
    #     incumbent_retained_parts.append({
    #         "index": idx,
    #         "row": row,
    #         "incumbent": incumbent,
    #         "reason": "Force to incumbent at WAPP"
    #     })
    #     continue

    # 1. No valid suppliers
    if valid_supplier_count == 0 or pd.isna(min_supplier):
        no_valid_supplier_parts.append({
            "index": idx,
            "row": row,
            "reason": "No valid suppliers"
        })
        total_cost_not_awarded += row.get('Landed Extended Cost USD', 0)
        continue

    # 2. Incumbent did not bid, but minimum bid exists
    if incumbent not in suppliers and pd.notna(min_supplier):
        landed_cost = row.get(f"{min_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        volume = pd.to_numeric(row.get(volume_col), errors='coerce')
        extended_cost = landed_cost * volume
        savings_usd = row.get(f"{min_supplier} - Final Landed USD savings vs baseline", 0)

        wapp_price = row.get('Volume-banded WAPP')
        row_id = row.get('ROW ID #')
        multiplier_info = get_supplier_info(row_id, incumbent)

        if not multiplier_info:
            print(f"No info found for {row_id}, {incumbent}")
            wapp_landed_cost = 999999
        else:
            if incumbent == 'KG Machinery':
                is_kg = True
            wapp_landed_cost = wapp_price * multiplier_info['FreightMultiplier'] + wapp_price * (multiplier_info['tariff_value'] + multiplier_info['Metal Tariff'])

        try:
            savings_usd = float(savings_usd)
        except:
            savings_usd = 0

        if wapp_landed_cost < landed_cost and incumbent != '-':
            incumbent_retained_parts.append({
                "index": idx,
                "row": row,
                "incumbent": incumbent,
                "reason": "Incumbent did not bid, but its WAPP landed is lower than Lowest Bid."
            })
        
        else:
            must_assign_min_bid_parts.append({
                "index": idx,
                "row": row,
                "extended_cost": extended_cost,
                "savings_usd": savings_usd,
                "min_supplier": min_supplier,
                "incumbent": incumbent,
                "reason": "Incumbent did not bid, using Final Minimum Bid Landed Supplier"
            })
        continue

    # 3. Incumbent bid
    incumbent_bid_val = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
    if incumbent_bid_val > 0:
        # a. Incumbent is the minimum
        if min_supplier == incumbent:
            incumbent_retained_parts.append({
                "index": idx,
                "row": row,
                "incumbent": incumbent,
                "reason": "Incumbent retained (lowest bid)"
            })
        # b. Incumbent is NOT the minimum
        else:

            if str(row.get("ROW ID #")) in PARTS_NOT_TO_ASSIGN_TO_MANEK and min_supplier == "Manek Metalcraft":
                if second_min_supplier == incumbent:
                    incumbent_retained_parts.append({
                        "index": idx,
                        "row": row,
                        "incumbent": incumbent,
                        "reason": "Incumbent Supplier retained over Manek MetalCraft (part restriction)"
                    })
                    continue
                else:
                    min_supplier = second_min_supplier
                    reason = "Avoid assigning to Manek MetalCraft (part restriction) for red brass; assigned to 2nd lowest bidder"

            elif min_supplier == "Manek Metalcraft" and SPEND_ON_MANEK > 3500000:
                if second_min_supplier == incumbent:
                    incumbent_retained_parts.append({
                        "index": idx,
                        "row": row,
                        "incumbent": incumbent,
                        "reason": "Incumbent Supplier retained over Manek MetalCraft (part restriction)"
                    })
                else:
                    min_supplier = second_min_supplier
                    reason = "Avoid assigning to Manek MetalCraft (part restriction) for red brass; assigned to 2nd lowest bidder"

                continue

            landed_cost = row.get(f"{min_supplier} - R2 - Total landed cost per UOM (USD)", 0)
            fob_cost = row.get(f"{min_supplier} - R2 - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)", 0)
            volume = pd.to_numeric(row.get(volume_col), errors='coerce')
            extended_cost = landed_cost * volume
            fob_extended_cost = fob_cost * volume
            savings_usd = row.get(f"{min_supplier} - Final Landed USD savings vs baseline", 0)

            has_incumbent_bid = pd.notna(incumbent_bid_val) and incumbent_bid_val > 0

            if ((min_supplier == 'Pushti Metal' and PUSHTI_EXTRA_VOLUME <= 1000000) or (min_supplier == 'Manek Metalcraft' and MANEK_EXTRA_VOLUME <= 6400000)) and incumbent in ['Mayank', 'Brass Pro Industrial'] :

                if min_supplier in ['Manek Metalcraft', 'Pushti Metal']:
                    if incumbent == 'Mayank':
                        if min_supplier == 'Manek Metalcraft':
                            reason = "Incumbent Supplier Mayank prefered over Manek MetalCraft"
                        elif min_supplier == 'Pushti Metal':
                            reason = "Incumbent Supplier Mayank prefered over Pushti Metal"
                        incumbent_retained_parts.append({
                            "index": idx,
                            "row": row,
                            "incumbent": incumbent,
                            "reason": reason
                        })
                        if min_supplier == 'Manek Metalcraft':
                            MANEK_EXTRA_VOLUME += row.get('Annual Volume (per UOM)')
                        elif min_supplier == 'Pushti Metal':
                            PUSHTI_EXTRA_VOLUME += row.get('Annual Volume (per UOM)')
                    else:
                        try:
                            savings_usd = float(savings_usd)
                        except:
                            savings_usd = 0
                        if min_supplier == 'Manek Metalcraft':
                            MANEK_EXTRA_VOLUME += row.get('Annual Volume (per UOM)')
                            SPEND_ON_MANEK += fob_extended_cost

                        candidate_new_supplier_parts.append({
                            "index": idx,
                            "row": row,
                            "extended_cost": extended_cost,
                            "savings_usd": savings_usd,
                            "min_supplier": min_supplier,
                            "incumbent": incumbent,
                            "reason": "Incumbent bid, but not lowest; eligible for new supplier assignment"
                        })
                
                else:
                    try:
                        savings_usd = float(savings_usd)
                    except:
                        savings_usd = 0
                    
                    if min_supplier == 'Manek Metalcraft':
                        SPEND_ON_MANEK += fob_extended_cost
                    candidate_new_supplier_parts.append({
                        "index": idx,
                        "row": row,
                        "extended_cost": extended_cost,
                        "savings_usd": savings_usd,
                        "min_supplier": min_supplier,
                        "incumbent": incumbent,
                        "reason": "Incumbent bid, but not lowest; eligible for new supplier assignment"
                    })
            else:
                try:
                    savings_usd = float(savings_usd)
                except:
                    savings_usd = 0
                candidate_new_supplier_parts.append({
                    "index": idx,
                    "row": row,
                    "extended_cost": extended_cost,
                    "savings_usd": savings_usd,
                    "min_supplier": min_supplier,
                    "incumbent": incumbent,
                    "reason": "Incumbent bid, but not lowest; eligible for new supplier assignment"
                })

    else:
        # Incumbent did not bid, but minimum bid exists (should already be handled above)
        pass

# --- Sort candidate new supplier parts by savings descending ---
candidate_new_supplier_parts.sort(key=lambda x: x["savings_usd"], reverse=True)

# --- Assign must-assign-min-bid parts first (these are forced, contribute to threshold) ---
decision_rows = []
new_supplier_spent = 0
selected_new_rows = set()

for part in must_assign_min_bid_parts:
    # if new_supplier_spent + part["extended_cost"] <= THRESHOLD_COST:
    #     decision_rows.append({
    #         "index": part["index"],
    #         "row": part["row"],
    #         "new_supplier": part["min_supplier"],
    #         "extended_cost": part["extended_cost"],
    #         "incumbent": part["incumbent"],
    #         "reason": part["reason"]
    #     })
    #     new_supplier_spent += part["extended_cost"]
    #     selected_new_rows.add(part["index"])
    # else:
    #     # If threshold exceeded, assign to incumbent (if possible)
    #     decision_rows.append({
    #         "index": part["index"],
    #         "row": part["row"],
    #         "new_supplier": part["incumbent"],
    #         "extended_cost": 0,
    #         "incumbent": part["incumbent"],
    #         "reason": "Threshold exceeded, retaining incumbent"
    #     })
    #     selected_new_rows.add(part["index"])
    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": part["min_supplier"],
        "extended_cost": part["extended_cost"],
        "incumbent": part["incumbent"],
        "reason": part["reason"]
    })
    new_supplier_spent += part["extended_cost"]
    selected_new_rows.add(part["index"])

# --- Assign candidate new supplier parts until threshold hit ---
for part in candidate_new_supplier_parts:
    if part["index"] in selected_new_rows:
        continue
    # if new_supplier_spent + part["extended_cost"] <= THRESHOLD_COST:
    #     decision_rows.append({
    #         "index": part["index"],
    #         "row": part["row"],
    #         "new_supplier": part["min_supplier"],
    #         "extended_cost": part["extended_cost"],
    #         "incumbent": part["incumbent"],
    #         "reason": part["reason"] + f" (within {PERCENT_NEW*100}% threshold)"
    #     })
    #     new_supplier_spent += part["extended_cost"]
    #     selected_new_rows.add(part["index"])
    # else:
    #     # Threshold exceeded, retain incumbent
    #     decision_rows.append({
    #         "index": part["index"],
    #         "row": part["row"],
    #         "new_supplier": part["incumbent"],
    #         "extended_cost": 0,
    #         "incumbent": part["incumbent"],
    #         "reason": "Threshold exceeded, retaining incumbent"
    #     })
    #     selected_new_rows.add(part["index"])

    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": part["min_supplier"],
        "extended_cost": part["extended_cost"],
        "incumbent": part["incumbent"],
        "reason": part["reason"]
    })
    new_supplier_spent += part["extended_cost"]
    selected_new_rows.add(part["index"])

# --- Assign all incumbent retained parts ---
for part in incumbent_retained_parts:
    if part["index"] in selected_new_rows:
        continue
    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": part["incumbent"],
        "extended_cost": 0,
        "incumbent": part["incumbent"],
        "reason": part["reason"]
    })
    selected_new_rows.add(part["index"])

# --- Assign all no valid supplier parts ---
for part in no_valid_supplier_parts:
    if part["index"] in selected_new_rows:
        continue
    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": "-",
        "extended_cost": 0,
        "incumbent": part["row"].get(incumbent_col),
        "reason": part["reason"]
    })
    selected_new_rows.add(part["index"])


# --- Step 4: Assign rest (fallback to incumbent or final bid supplier) ---
for idx, row in df.iterrows():
    if idx in selected_new_rows:
        continue

    incumbent = row.get(incumbent_col)
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
    valid_supplier_count = row.get(valid_supplier_col, 0)

    if valid_supplier_count == 0:
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": "-",
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "No valid suppliers"
        })
    elif incumbent in suppliers:
        incumbent_bid_val = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
        if incumbent_bid_val == 0:
            wapp_price = row.get('Volume-banded WAPP')
            multiplier_info = get_supplier_info(row.get('ROW ID #'), incumbent)

            incumbent_bid_val = wapp_price * multiplier_info['FreightMultiplier'] + wapp_price * (multiplier_info['tariff_value'] + multiplier_info['Metal Tariff'])

        if incumbent_bid_val > 0 and incumbent_bid_val < row.get(f"{min_supplier} - R2 - Total landed cost per UOM (USD)", 0):
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": incumbent,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": f"Incumbent did not bid, but its WAPP landed is lower than Lowest Bid."
            })
        else:
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": min_supplier,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": f"Forced to Lowest Bidder"
            })
    elif pd.notna(min_supplier):
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": min_supplier,
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "Incumbent did not bid, using Final Minimum Bid Landed Supplier"
        })
    else:
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": "-",
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "No valid bids"
        })


output_data = []
total_fob_savings_usd = 0
total_landed_savings_usd = 0
total_annual_revenue_discount = 0
incumbent_retained = 0
new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0

net_new_supplier_count = 0  
parts_where_no_bids = 0
def calculate_wapp_landed_savings(row, p_u):
    incumbent = row.get(incumbent_col)
    if incumbent == '-':
        return 0

    wapp_landed_cost = row.get('Volume-banded WAPP Landed Cost')
    wapp_price = row.get('Volume-banded WAPP')
    row_id = row.get('ROW ID #')
    multiplier_info = get_supplier_info(row_id, incumbent)
    new_wapp_landed_cost = wapp_price * multiplier_info['FreightMultiplier'] + wapp_price * (multiplier_info['tariff_value'] + multiplier_info['Metal Tariff'])
    new_wapp_landed_cost = round(new_wapp_landed_cost, 4)
    
    if p_u == 'pct':
        return float((wapp_landed_cost - new_wapp_landed_cost) / wapp_landed_cost)
    
    elif p_u == 'usd':
        return float(((wapp_landed_cost - new_wapp_landed_cost) / wapp_landed_cost) * row.get('Landed Extended Cost USD'))
    
    else:
        return new_wapp_landed_cost
    
print("\nBuilding final output rows...\n")
for decision in tqdm(decision_rows, total=len(decision_rows), desc="Finalizing"):
    row = decision["row"]

    idx = decision["index"]
    selected_supplier = decision["new_supplier"]
    reason = decision["reason"]
    incumbent = decision["incumbent"]

    # Get savings columns
    pct_col = f"{selected_supplier} - Final % savings vs baseline"
    usd_col = f"{selected_supplier} - Final USD savings vs baseline"
    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"

    # Reference supplier metadata
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
    def get_ref_value(col):
        return ref_row[col].values[0] if not ref_row.empty else "-"
    
    try:
        fob_savings_usd = float(row.get(usd_col)) if pd.notna(row.get(usd_col)) else 0
    except (ValueError, TypeError):
        fob_savings_usd = 0
    
    total_fob_savings_usd += fob_savings_usd 
        
    try:
        lcs = float(row.get(landed_usd_col)) if pd.notna(landed_usd_col) else 0
    except (ValueError, TypeError): 
        lcs = 0

    total_landed_savings_usd  += lcs
    if row.get("ROW ID #") == 66:
        stop = True
        
    landed_cost_key = f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"
    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

    if selected_supplier == '-': selected_supplier = incumbent

    if selected_supplier != incumbent:
        landed_cost = row.get(landed_cost_key)
        if not landed_cost:  # catches 0, None, '', False
            landed_cost = 0
    else:
        landed_cost = row.get(incumbent_key)
        if not landed_cost:
            landed_cost = calculate_wapp_landed_savings(row, 'new_landed_cost')

    result_extended_cost = landed_cost * row.get(volume_col, 0)
    if selected_supplier == "-":
        # parts_where_no_bids += 1
        output_row = {
            "ROW ID #": row.get("ROW ID #"),
            "Division": row.get("Division"),
            "Part #": row.get("Part #"),
            "Item Description": row.get("Item Description"),
            "Product Group": row.get("Product Group"),
            "Part Family": row.get("Part Family"),
            "Incumbent Supplier": incumbent,
            "Selected Supplier": incumbent,
            "Annual Volume (per UOM)": row.get("Annual Volume (per UOM)"),
            "Final quote per each FOB Port of Departure (USD)": row.get(f"{incumbent} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)", row.get('Volume-banded WAPP')),
            "FOB Savings %": "-",
            "FOB Savings USD": "-",
            "Landed Cost Savings %": row[landed_pct_col] if (landed_pct_col in row and row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(row, 'pct'),
            "Landed Cost Savings USD": row[landed_usd_col] if (landed_usd_col in row and row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(row, 'usd'),
            "Annual revenue discount USD": 0,
            "Reason": 'No valid bids in this round, so forced to incumbent supplier.',
            "Landed Extended Cost USD": result_extended_cost,
            "Is Totally New Supplier": "No",
            "Part Switched": "No",
            
            # "valid_supplier_count": row.get('Valid Supplier')
            "valid_supplier_count": row.get('Valid Supplier'),
        }
        output_data.append(output_row)
        continue

    # Note: Supplier metrics will be recalculated after rationalization
    # to ensure accuracy after any supplier reassignments

    
    output_row = {
        "ROW ID #": row.get("ROW ID #"),
        "Division": row.get("Division"),
        "Part #": row.get("Part #"),
        "Item Description": row.get("Item Description"),
        "Product Group": row.get("Product Group"),
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "Annual Volume (per UOM)": row.get('Annual Volume (per UOM)'),
        "Final quote per each FOB Port of Departure (USD)": row.get(f"{selected_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)", 0) if selected_supplier!=incumbent else row.get(f"{incumbent} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)", row.get('Volume-banded WAPP')),
        "FOB Savings %": row.get(pct_col, 0),
        "FOB Savings USD": row.get(usd_col, 0),
        "Landed Cost Savings %": row[landed_pct_col] if (landed_pct_col in row and row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(row, 'pct'),

        "Landed Cost Savings USD": row[landed_usd_col] if (landed_usd_col in row and row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(row, 'usd'),

        "Reason": reason,
        "Landed Extended Cost USD": result_extended_cost,
        "Is Totally New Supplier": "Yes" if selected_supplier not in incumbent_suppliers else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
        
        # "valid_supplier_count": row.get('Valid Supplier')
        "valid_supplier_count": row.get('Valid Supplier'),
    }
    if result_extended_cost == 0:
        stop=True
    output_data.append(output_row)

# --- BINZHOU ZELI1 REMOVAL LOGIC (BEFORE RATIONALIZATION) ---
print("\nApplying Binzhou Zeli removal logic for specific parts...")

binzhou_zeli_supplier = "Binzhou Zeli"
binzhou_reassignments = 0
binzhou_zeli_exclusion_parts = [
    "CGBSL-200-A1","CGDSL-200-A1","CGCSL-200CR-A1","CDCSL-200-A1", "CDCSL-300-A1","CGBSL-300-A1","CGDSL-300-A1","CGCSL-300CR-A1", "CGBSL-400-A1","CGDSL-400-A1","CDCSL-400-A1","CDCSL-200-SS1", "CGCSL-400CR-A1","CGBSL-200-SS1","CGCSL-200CR-SS1","CGDSL-200-SS1", "CDCSL-600-A1","CDCSL-300-SS1","CGDSL-600-A1","CGBSL-300-SS1", "CGCSL-600CR-A1","CGDSL-300-SS1","CGCSL-300CR-SS1","CDCSL-400-SS1", "CGBSL-400-SS1","CGDSL-400-SS1","CGCSL-400CR-SS1","CDCSL-600-SS1", "CGBSL-600-SS1","CGDSL-600-SS1","CGCSL-600CR-SS1"
]

print(f"Binzhou Zeli exclusion applies to {len(binzhou_zeli_exclusion_parts)} specific part numbers")
def find_best_alternative_to_binzhou(row, all_suppliers):
    """Find the best alternative supplier excluding Binzhou Zeli"""

    return 'Luxecasting', "Defaulted to Luxecasting"

    # incumbent = row.get("Incumbent Supplier", "")
    # valid_supplier_count = row.get("Valid Supplier", 0)
    
    # # If only one valid supplier and it's Binzhou Zeli, we have no choice
    #     # Check if Binzhou Zeli is the only bidder
    # binzhou_landed_cost = row.get(f"{binzhou_zeli_supplier} - R2 - Total landed cost per UOM (USD)", 0)

    # # Count other valid bidders
    # other_valid_bidders = 0
    # for supplier in all_suppliers:
    #     if supplier != binzhou_zeli_supplier:
    #         landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
    #         if pd.notna(landed_cost) and landed_cost > 0:
    #             other_valid_bidders += 1
        
    #     if other_valid_bidders == 0:
    #         return binzhou_zeli_supplier, "Only valid supplier available"
    
    # # # First preference: incumbent (if not Binzhou Zeli and has valid bid)
    # # if incumbent != binzhou_zeli_supplier and incumbent in all_suppliers:
    # #     incumbent_landed_cost = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
    # #     if pd.notna(incumbent_landed_cost) and incumbent_landed_cost > 0:
    # #         return incumbent, "Reassigned to incumbent (avoiding Binzhou Zeli)"
    
    # # Second preference: find lowest bidder excluding Binzhou Zeli
    # best_supplier = None
    # best_cost = float('inf')
    
    # for supplier in all_suppliers:
    #     if supplier == binzhou_zeli_supplier:
    #         continue
            
    #     landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
    #     if pd.notna(landed_cost) and landed_cost > 0 and landed_cost < best_cost:
    #         best_cost = landed_cost
    #         best_supplier = supplier
    
    # if best_supplier:
    #     return best_supplier, f"Reassigned to lowest bidder excluding Binzhou Zeli (${best_cost:.2f})"
    
    # # Fallback: if no other valid bidders, keep Binzhou Zeli
    # return binzhou_zeli_supplier, "No alternative suppliers available"

# Get all R2 landed cost columns for finding alternatives
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]
all_suppliers = ['Luxecasting']

# Create a lookup dictionary for faster DataFrame access (highly optimized)
print("Creating DataFrame lookup for performance optimization...")
df_lookup = {}
for idx, row in df.iterrows():
    row_id = row.get("ROW ID #")
    if row_id is not None:
        df_lookup[row_id] = row
print(f"DataFrame lookup created with {len(df_lookup)} entries")

# Apply Binzhou Zeli removal logic to output_data (only for specific parts)
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    part_number = row.get("Part #", "")
    
    # Only apply Binzhou Zeli removal logic to specific part numbers
    if current_supplier == binzhou_zeli_supplier and part_number in binzhou_zeli_exclusion_parts:
        # Find corresponding row in original dataframe using lookup
        row_id = row.get("ROW ID #")
        df_row = df_lookup.get(row_id)
        
        if df_row is not None:
            new_supplier, reason = find_best_alternative_to_binzhou(df_row, all_suppliers)
            print(calculate_wapp_landed_savings(df_row, 'pct'))
            if new_supplier != current_supplier:
                # Update the supplier assignment
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"Binzhou Zeli avoided: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    fob_col = f"{new_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
                    
                    
                    output_data[i]["Final quote per each FOB Port of Departure (USD)"] = df_row.get(fob_col, df_row.get('Volume-banded WAPP'))
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row[landed_pct_col] if (landed_pct_col in df_row and df_row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'pct')
                    output_data[i]["Landed Cost Savings USD"] = df_row[landed_usd_col] if (landed_usd_col in df_row and df_row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'usd')
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    
                    landed_cost_key = f"{new_supplier} - R2 - Total landed cost per UOM (USD)"
                    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

                    if new_supplier != incumbent:
                        landed_cost = df_row.get(landed_cost_key)
                        if not landed_cost:  # catches 0, None, '', False
                            landed_cost = 0
                    else:
                        landed_cost = df_row.get(incumbent_key)
                        if not landed_cost:
                            landed_cost = calculate_wapp_landed_savings(df_row, 'new_landed_cost')
                    result_extended_cost = landed_cost * df_row.get(volume_col, 0)

                    output_data[i]["Landed Extended Cost USD"] = result_extended_cost
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                
                binzhou_reassignments += 1

print(f"Binzhou Zeli removal complete: {binzhou_reassignments} parts reassigned")

### West Legend-MTD re-allocation logic starts here
# --- West Legend-MTD REMOVAL LOGIC (BEFORE RATIONALIZATION) ---
print("\nApplying West Legend-MTD removal logic for specific parts...")

west_legend_mtd_supplier = "West Legend-MTD"
west_legend_mtd_reassignments = 0
west_legend_mtd_exclusion_rows = [
    row['ROW ID #'] for row in df.itertuples() if getattr(row, "Selected Supplier", "") == west_legend_mtd_supplier
]

def find_best_alternative_to_west_legend_mtd(row, all_suppliers):
    """Find the best alternative supplier excluding West Legend-MTD"""
    incumbent = row.get("Incumbent Supplier", "")
    valid_supplier_count = row.get("Valid Supplier", 0)
    
    # If only one valid supplier and it's West Legend-MTD, we have no choice
    if valid_supplier_count != 1:
        # Check if West Legend-MTD is the only bidder
        west_legend_mtd_landed_cost = row.get(f"{west_legend_mtd_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        
        other_valid_bidders = 0
        for supplier in all_suppliers:
            if supplier != west_legend_mtd_supplier:
                landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
                if pd.notna(landed_cost) and landed_cost > 0:
                    other_valid_bidders += 1
        
        if other_valid_bidders == 0:
            return west_legend_mtd_supplier, "Only valid supplier available"
        
    # find lowest bidder excluding West Legend-MTD
    best_supplier = None
    best_cost = float('inf')
    
    for supplier in all_suppliers:
        if supplier == west_legend_mtd_supplier:
            continue
            
        landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.notna(landed_cost) and landed_cost > 0 and landed_cost < best_cost:
            best_cost = landed_cost
            best_supplier = supplier
    
    if best_supplier:
        return best_supplier, f"Reassigned to lowest bidder excluding West Legend-MTD (${best_cost:.2f})"
    
    # Fallback: if no other valid bidders, keep West Legend-MTD

    return west_legend_mtd_supplier, "No alternative suppliers available"


# Get all R2 landed cost columns for finding alternatives
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]

# Create a lookup dictionary for faster DataFrame access (highly optimized)
print("Creating DataFrame lookup for performance optimization...")
df_lookup = {}
for idx, row in df.iterrows():
    row_id = row.get("ROW ID #")
    if row_id is not None:
        df_lookup[row_id] = row
print(f"DataFrame lookup created with {len(df_lookup)} entries")

# Apply  West Legend-MTD removal logic to output_data (only for specific parts)
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    row_id = row.get("ROW ID #", "")
    
    # Only apply West Legend-MTD removal logic to specific part numbers
    if current_supplier == west_legend_mtd_supplier:
        # Find corresponding row in original dataframe using lookup
        row_id = row.get("ROW ID #")
        df_row = df_lookup.get(row_id)
        
        if df_row is not None:
            new_supplier, reason = find_best_alternative_to_west_legend_mtd(df_row, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"West Legend-MTD avoided: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    fob_col = f"{new_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
                    
                    
                    output_data[i]["Final quote per each FOB Port of Departure (USD)"] = df_row.get(fob_col, df_row.get('Volume-banded WAPP'))
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row[landed_pct_col] if (landed_pct_col in df_row and df_row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'pct')
                    output_data[i]["Landed Cost Savings USD"] = df_row[landed_usd_col] if (landed_usd_col in df_row and df_row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'usd')
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    
                    landed_cost_key = f"{new_supplier} - R2 - Total landed cost per UOM (USD)"
                    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

                    if new_supplier != incumbent:
                        landed_cost = df_row.get(landed_cost_key)
                        if not landed_cost:  # catches 0, None, '', False
                            landed_cost = 0
                    else:
                        landed_cost = df_row.get(incumbent_key)
                        if not landed_cost:
                            landed_cost = calculate_wapp_landed_savings(df_row, 'new_landed_cost')
                    result_extended_cost = landed_cost * df_row.get(volume_col, 0)

                    output_data[i]["Landed Extended Cost USD"] = result_extended_cost
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                
                west_legend_mtd_reassignments += 1

print(f"West Legend-MTD removal complete: {west_legend_mtd_reassignments} parts reassigned")


# --------- Removing Manek for red brass
manek_supplier= "Manek Metalcraft"
manek_reassignments = 0
def find_best_alternative_to_manek(row, all_suppliers):
    """Find the best alternative supplier excluding Manek Metalcraft"""
    incumbent = row.get("Incumbent Supplier", "")
    valid_supplier_count = row.get("Valid Supplier", 0)
    
    # If only one valid supplier and it's West Legend-MTD, we have no choice
    if valid_supplier_count != 1:
        # Check if Manek Metalcraft is the only bidder
        manek_landed_cost = row.get(f"{manek_supplier} - R2 - Total landed cost per UOM (USD)", 0)

        other_valid_bidders = 0
        for supplier in all_suppliers:
            if supplier != manek_supplier:
                landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
                if pd.notna(landed_cost) and landed_cost > 0:
                    other_valid_bidders += 1
        
        if other_valid_bidders == 0:
            return manek_supplier, "Only valid supplier available"
        
    # find lowest bidder excluding West Legend-MTD
    best_supplier = None
    best_cost = float('inf')
    
    for supplier in all_suppliers:
        if supplier == manek_supplier:
            continue
            
        landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.notna(landed_cost) and landed_cost > 0 and landed_cost < best_cost:
            best_cost = landed_cost
            best_supplier = supplier
    
    if best_supplier:
        return best_supplier, f"Reassigned to lowest bidder excluding Manek Metalcraft (${best_cost:.2f})"

    # Fallback: if no other valid bidders, keep Manek Metalcraft

    return manek_supplier, "No alternative suppliers available"


# Get all R2 landed cost columns for finding alternatives
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]

# Create a lookup dictionary for faster DataFrame access (highly optimized)
print("Creating DataFrame lookup for performance optimization...")
df_lookup = {}
for idx, row in df.iterrows():
    row_id = row.get("ROW ID #")
    if row_id is not None:
        df_lookup[row_id] = row
print(f"DataFrame lookup created with {len(df_lookup)} entries")

# Apply  Manek Metalcraft removal logic to output_data (only for specific parts)
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    row_id = row.get("ROW ID #", "")
    
    if not str(row_id) in PARTS_NOT_TO_ASSIGN_TO_MANEK:
        continue
    
    # Only apply Manek Metalcraft removal logic to specific part numbers
    if current_supplier == manek_supplier:
        # Find corresponding row in original dataframe using lookup
        row_id = row.get("ROW ID #")
        df_row = df_lookup.get(row_id)
        
        if df_row is not None:
            new_supplier, reason = find_best_alternative_to_manek(df_row, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"Manek Metalcraft avoided for red_brass: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    fob_col = f"{new_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
                    
                    
                    output_data[i]["Final quote per each FOB Port of Departure (USD)"] = df_row.get(fob_col, df_row.get('Volume-banded WAPP'))
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row[landed_pct_col] if (landed_pct_col in df_row and df_row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'pct')
                    output_data[i]["Landed Cost Savings USD"] = df_row[landed_usd_col] if (landed_usd_col in df_row and df_row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'usd')
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    
                    landed_cost_key = f"{new_supplier} - R2 - Total landed cost per UOM (USD)"
                    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

                    if new_supplier != incumbent:
                        landed_cost = df_row.get(landed_cost_key)
                        if not landed_cost:  # catches 0, None, '', False
                            landed_cost = 0
                    else:
                        landed_cost = df_row.get(incumbent_key)
                        if not landed_cost:
                            landed_cost = calculate_wapp_landed_savings(df_row, 'new_landed_cost')
                    result_extended_cost = landed_cost * df_row.get(volume_col, 0)

                    output_data[i]["Landed Extended Cost USD"] = result_extended_cost
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                
                manek_reassignments += 1

print(f"Manek Metalcraft removal complete: {manek_reassignments} parts reassigned")


# LOGIC TO Change suppliers for selected suppliers that do not supply anymore.


def find_next_best_supplier(row, current_supplier, all_suppliers):
    """Find the next best bidder among large suppliers for a given part"""

    incumbent = row.get("Normalized incumbent supplier", "")
    valid_supplier = row.get('Valid Supplier')
    if valid_supplier == 1:
        if incumbent == current_supplier:
            return current_supplier, "Has to stay with it."
        else:
            return incumbent, 'Forced to incumbent because no other bid on it.'
    else:

        part_bids = []
       
        # Get all bids for this part and sort by landed cost
        for supplier in all_suppliers:
            if supplier == current_supplier:
                continue  # Skip current supplier
                
            landed_cost_col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
            landed_cost = row.get(landed_cost_col, 0)
            
            if pd.notna(landed_cost) and landed_cost > 0:
                part_bids.append((supplier, landed_cost))
        
        # Sort by landed cost (ascending - lowest cost first)
        part_bids.sort(key=lambda x: x[1])
        
        # Find first large supplier in sorted list
        for supplier, cost in part_bids:
            return supplier, f"Lowest bidder other than {current_supplier}"

        print('Nahi hua bhai return')       
    
for i, row in enumerate(output_data):
    row_id = row.get("ROW ID #", "")
    current_supplier = row["Selected Supplier"]
    if (str(row_id) in ["11", "15", "276", "277", "4703", "4704", "4937", "9619", "9709", "11151" ] and current_supplier in ['Oston Industrial']) or ( current_supplier in ['ZHEJIANG WANDEKAI'] and str(row_id) in ["1578","1793","1794","1896","1899","3005","4377","4381","4382","4383","4406","4407","4408","4413","4414","4415","4416","4417","4421","4423","4425","4454","4455","4456","4458","4744","4749","4754","4787","4797","4800","4809","4810","4821","5853","5854","7904","8411","8412","8413","8432","8433","8434","8435","8521","8522","8539","8540","8541","9448","9695","13160","13161","13162",] ) or (current_supplier in ['Coda'] and str(row_id) in ["1163", "1164", "1165", "1166", "1167", "1173", "1176", "1177", "1178", "1179", "1180", "1181", "1182", "1183", "1184", "1185", "1186", "1187", "1188", "1190", "1213", "1277", "1288", "1289", "1290", "1305", "1306", "1308", "1309", "1310", "1311", "1312", "1318", "1319", "1320", "1321", "1322", "1323", "1327", "1328", "1333", "1335", "1341", "1342", "1346", "1347", "1348", "1352", "1358", "1359", "1360", "1361", "1362", "1364", "1365", "1366", "1367", "1368", "1369", "1370", "1372", "1374", "1379", "1386", "1387", "1388", "1389", "1390", "1393", "1394", "1395", "1396", "1397", "1398", "1399", "1400", "1405", "1406", "1407", "1408", "1409", "1410", "1411", "1412", "1413", "1414", "1415", "1416", "1417", "1418", "1425", "1429", "1430", "1439", "1441", "1445", "1446", "1448", "1489", "1490", "1498", "1499", "1510", "1511", "1516", "1520", "6813", "6815", "6825", "6838", "6839", "6844", "6851", "6852", "6864", "6866", "6890", "6893", "6909", "6910", "6911", "6912", "6917", "6918", "6919", "6927", "6928", "6929", "6930", "6932", "6933", "6934", "6939", "7060", "7071", "7072", "7073", "7076", "7089", "7090", "7102", "7111", "7117", "7119", "7125", "7126", "7136", "7145", "7185", "7186", "7187", "7188", "7189", "7197", "7207", "7209", "7210", "7211", "7212", "7213", "7214", "7215", "7254", "7256", "7300", "7301", "7306", "7331", "7332", "7826", "7919", "8742", "8772", "8915", "9994", "13613"]): 
        if current_supplier not in ['Coda', 'ZHEJIANG WANDEKAI', 'Oston Industrial']:
            continue  # Only reassign if current supplier is in the list
        # Find corresponding row in original dataframe
        row_id = row.get("ROW ID #")
        df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
        
        if df_row is not None and df_row['Valid Supplier'] >= 1:

            new_supplier, reason = find_next_best_supplier(df_row, current_supplier, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                old_supplier = current_supplier
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"{old_supplier} no longer supplies: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    fob_col = f"{new_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
                    
                    
                    output_data[i]["Final quote per each FOB Port of Departure (USD)"] = df_row.get(fob_col, df_row.get('Volume-banded WAPP'))
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row[landed_pct_col] if (landed_pct_col in df_row and df_row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'pct')
                    output_data[i]["Landed Cost Savings USD"] = df_row[landed_usd_col] if (landed_usd_col in df_row and df_row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'usd')
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    
                    landed_cost_key = f"{new_supplier} - R2 - Total landed cost per UOM (USD)"
                    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

                    if new_supplier != incumbent:
                        landed_cost = df_row.get(landed_cost_key)
                        if not landed_cost:  # catches 0, None, '', False
                            landed_cost = 0
                    else:
                        landed_cost = df_row.get(incumbent_key)
                        if not landed_cost:
                            landed_cost = calculate_wapp_landed_savings(df_row, 'new_landed_cost')
                    result_extended_cost = landed_cost * df_row.get(volume_col, 0)

                    output_data[i]["Landed Extended Cost USD"] = result_extended_cost
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"


# --- TAIL SUPPLIER RATIONALIZATION LOGIC ---

print("\nApplying tail supplier rationalization logic...\n")

# Calculate total awarded amount per supplier
supplier_awarded_amounts = {}
for row in output_data:
    supplier = row["Selected Supplier"]
    valid_supplier_count = row["valid_supplier_count"]
    # print(valid_supplier_count)
    if valid_supplier_count != 0:
        awarded_amount = row.get("Landed Extended Cost USD", 0)
        if supplier not in supplier_awarded_amounts:
            supplier_awarded_amounts[supplier] = 0
        supplier_awarded_amounts[supplier] += awarded_amount


# Update decision_rows with Binzhou Zeli changes (optimized)
output_row_lookup = {row.get("ROW ID #"): row for row in output_data}
for i, decision in enumerate(decision_rows):
    row_id = decision["row"].get("ROW ID #")
    if row_id in output_row_lookup:
        output_row = output_row_lookup[row_id]
        decision_rows[i]["new_supplier"] = output_row["Selected Supplier"]
        if "Binzhou Zeli avoided" in output_row["Reason"]:
            decision_rows[i]["reason"] = output_row["Reason"]


tail_suppliers_to_rationalize = ['Giraffe Stainless', 'Union Metal Products', 'WEFLO', 'Tianjin Outshine', 'Sichuan Y&J', 'Guangzhou Hopetrol', 'Swati Enterprise']

large_suppliers = {supplier: amount for supplier, amount in supplier_awarded_amounts.items() 
                  if amount >= 100000}

print(f"Large suppliers (≥$100k): {len(large_suppliers)}")
for supplier, amount in sorted(large_suppliers.items(), key=lambda x: x[1], reverse=True):
    print(f"  - {supplier}: ${amount:,.2f}")

print(f"\nTail suppliers (<$100k) to rationalize: {len(tail_suppliers_to_rationalize)}")
for supplier, amount in sorted([(s, supplier_awarded_amounts[s]) for s in tail_suppliers_to_rationalize], 
                              key=lambda x: x[1], reverse=True):
    print(f"  - {supplier}: ${amount:,.2f}")

# Get all R2 landed cost columns for finding next best bidders
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]

def find_next_best_large_supplier(row, current_supplier, large_suppliers, all_suppliers):
    """Find the next best bidder among large suppliers for a given part"""

    incumbent = row.get("Normalized incumbent supplier", "")
    valid_supplier = row.get('Valid Supplier')
    if valid_supplier == 1:
        if incumbent == current_supplier:
            return current_supplier, "Has to stay with it."
        else:
            return incumbent, 'Forced to incumbent because no other bid on it.'
    else:

        incumb_bid_col = f'{incumbent} - R2 - Total landed cost per UOM (USD)'
        incum_bid = row.get(incumb_bid_col, 0)
        
        part_bids = []
        # First check if incumbent is a large supplier
        if incumbent in large_suppliers:
            if pd.notna(incum_bid) and incum_bid > 0:
                part_bids.append((incumbent, incum_bid))
                return incumbent, "Rationalized to other bidder than than bidder based on logic"
        
        # Get all bids for this part and sort by landed cost
        for supplier in all_suppliers:
            if supplier == current_supplier:
                continue  # Skip current tail supplier
                
            landed_cost_col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
            landed_cost = row.get(landed_cost_col, 0)
            
            if pd.notna(landed_cost) and landed_cost > 0:
                part_bids.append((supplier, landed_cost))
        
        # Sort by landed cost (ascending - lowest cost first)
        part_bids.sort(key=lambda x: x[1])
        
        # Find first large supplier in sorted list
        for supplier, cost in part_bids:
            if supplier in large_suppliers:
                return supplier, f"Rationalized to other bidder than than bidder based on logic"
        
        # If no large supplier found, return the incumbent anyway
        return incumbent, "incumbent anyway"

# Apply rationalization

rationalization_changes = 0
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    
    if current_supplier in tail_suppliers_to_rationalize:
        # Find corresponding row in original dataframe
        row_id = row.get("ROW ID #")
        df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
        incumbent = row.get("Incumbent Supplier", "")
        if df_row is not None and df_row['Valid Supplier'] >= 1:

            new_supplier, reason = find_next_best_large_supplier(df_row, current_supplier, large_suppliers, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                old_supplier = current_supplier
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"Rationalized from {old_supplier}: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    fob_col = f"{new_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
                    
                    output_data[i]["Final quote per each FOB Port of Departure (USD)"] = df_row.get(fob_col, df_row.get('Volume-banded WAPP'))

                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row[landed_pct_col] if (landed_pct_col in df_row and df_row[landed_pct_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'pct')
                    output_data[i]["Landed Cost Savings USD"] = df_row[landed_usd_col] if (landed_usd_col in df_row and df_row[landed_usd_col] not in [0, '-']) else calculate_wapp_landed_savings(df_row, 'usd')
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    
                    landed_cost_key = f"{new_supplier} - R2 - Total landed cost per UOM (USD)"
                    incumbent_key = f"{incumbent} - R2 - Total landed cost per UOM (USD)"

                    if new_supplier != incumbent:
                        landed_cost = df_row.get(landed_cost_key)
                        if not landed_cost:  # catches 0, None, '', False
                            landed_cost = 0
                    else:
                        landed_cost = df_row.get(incumbent_key)
                        if not landed_cost:
                            landed_cost = calculate_wapp_landed_savings(df_row, 'new_landed_cost')
                    result_extended_cost = landed_cost * df_row.get(volume_col, 0)

                    output_data[i]["Landed Extended Cost USD"] = result_extended_cost
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                
                rationalization_changes += 1

print(f"Rationalization complete: {rationalization_changes} parts reassigned from tail suppliers")

# Update decision_rows with rationalized assignments
for i, decision in enumerate(decision_rows):
    row_id = decision["row"].get("ROW ID #")
    # Find corresponding output row
    for output_row in output_data:
        if output_row.get("ROW ID #") == row_id:
            decision_rows[i]["new_supplier"] = output_row["Selected Supplier"]
            if "Rationalized from" in output_row["Reason"]:
                decision_rows[i]["reason"] = output_row["Reason"]
            break

output_df = pd.DataFrame(output_data)

# --- Ensure FOB fallback for incumbent supplier rows ---
# Ensure ROW ID # is treated as string for reliable matching
for idx, row in output_df.iterrows():
    incumbent = row['Incumbent Supplier']
    selected = row['Selected Supplier']
    row_id = row['ROW ID #']

    # Only act if incumbent and selected supplier are the same
    if incumbent == selected:
        # Column to check in df
        cost_col = f'{incumbent} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)'

        # Check if row exists in df
        matching_row = df[df['ROW ID #'] == row_id]
        if not matching_row.empty:
            matching_row = matching_row.iloc[0]  # take first matching row
            # Check if cost_col exists and has valid value
            if cost_col in matching_row and pd.notna(matching_row[cost_col]) and matching_row[cost_col] not in [0, '-', '']:
                output_df.at[idx, 'Final quote per each FOB Port of Departure (USD)'] = matching_row[cost_col]
            else:
                # fallback to Volume-banded WAPP
                output_df.at[idx, 'Final quote per each FOB Port of Departure (USD)'] = matching_row['Volume-banded WAPP']



# Reset all metrics
incumbent_retained = 0
new_supplier_count = 0
net_new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0
net_new_supplier_list = set()
total_landed_savings_usd = 0
total_fob_savings_usd = 0
# parts_where_no_bids = 0

# Recalculate all metrics based on final supplier assignments
for row in output_data:
    selected_supplier = row["Selected Supplier"]
    incumbent = row["Incumbent Supplier"]
    
    # Handle parts with no bids
    if selected_supplier == "-":
        parts_where_no_bids += 1
        # Add to total cost not awarded (use original landed extended cost from output_data)
        total_cost_not_awarded += row.get("Landed Extended Cost USD", 0)
        continue
    
    # Get volume and costs for calculations
    row_id = row.get("ROW ID #")
    df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
    
    if df_row is not None:
        volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
        if pd.isna(volume):
            volume = 0
        
        landed_cost = df_row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.isna(landed_cost):
            landed_cost = 0
        
        extended_cost = landed_cost * volume
        
        # Calculate savings
        fob_usd_col = f"{selected_supplier} - Final USD savings vs baseline"
        landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
        
        try:
            fob_savings = float(df_row.get(fob_usd_col, 0)) if pd.notna(df_row.get(fob_usd_col)) else 0
        except (ValueError, TypeError):
            fob_savings = 0
        
        try:
            landed_savings = float(df_row.get(landed_usd_col, 0)) if pd.notna(df_row.get(landed_usd_col)) else 0
        except (ValueError, TypeError):
            landed_savings = 0
        
        total_fob_savings_usd += fob_savings
        total_landed_savings_usd += landed_savings
        
        # Classify supplier and update metrics
        if selected_supplier == incumbent:
            incumbent_retained += 1
            total_incumbent_volume += volume
            total_landed_cost_incumbent += extended_cost
        else:
            if selected_supplier in incumbent_suppliers:
                new_supplier_count += 1
                total_new_supplier_volume += volume
                total_landed_cost_new_suppliers += extended_cost
            else:
                net_new_supplier_count += 1
                net_new_supplier_list.add(selected_supplier)
                total_net_new_supplier_volume += volume
                total_landed_cost_completely_new_suppliers += extended_cost
        
        unique_suppliers.add(selected_supplier)


print('*************************************************')
print((output_df.loc[output_df['Selected Supplier'] == 'Manek Metalcraft', 'Final quote per each FOB Port of Departure (USD)'] * 
       output_df.loc[output_df['Selected Supplier'] == 'Manek Metalcraft', 'Annual Volume (per UOM)']).sum())

print('*************************************************')


print(f"All metrics recalculated after rationalization:")
print(f"  - Total landed savings USD: ${total_landed_savings_usd:,.2f}")
print(f"  - Total FOB savings USD: ${total_fob_savings_usd:,.2f}")
print(f"  - Total cost not awarded: ${total_cost_not_awarded:,.2f}")
print(f"  - Total landed cost incumbent: ${total_landed_cost_incumbent:,.2f}")
print(f"  - Total landed cost new suppliers: ${total_landed_cost_new_suppliers:,.2f}")
print(f"  - Total landed cost completely new suppliers: ${total_landed_cost_completely_new_suppliers:,.2f}")
print(f"  - Incumbent retained: {incumbent_retained}")
print(f"  - New suppliers: {new_supplier_count}")
print(f"  - Net new suppliers: {net_new_supplier_count}")
print(f"  - Parts with no bids: {parts_where_no_bids}")
print(f"  - Unique suppliers: {len(unique_suppliers)}")

# Final metrics recalculation after Binzhou Zeli removal
print("Final metrics recalculation after Rationalization...")

# Reset all metrics again
incumbent_retained = 0
new_supplier_count = 0
net_new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0
net_new_supplier_list = set()
total_landed_savings_usd = 0
total_fob_savings_usd = 0

# Final recalculation
for row in output_data:
    selected_supplier = row["Selected Supplier"]
    incumbent = row["Incumbent Supplier"]
    
    # Handle parts with no bids
    # if selected_supplier == "-":
    #     parts_where_no_bids += 1
    #     total_cost_not_awarded += row.get("Landed Extended Cost USD", 0)
    #     continue
    
    # Get volume and costs for calculations
    row_id = row.get("ROW ID #")
    df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
    
    if df_row is not None:
        volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
        if pd.isna(volume):
            volume = 0
        
        landed_cost = df_row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.isna(landed_cost):
            landed_cost = 0
        
        extended_cost = landed_cost * volume
        
        # Calculate savings
        fob_usd_col = f"{selected_supplier} - Final USD savings vs baseline"
        landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
        
        try:
            fob_savings = float(df_row.get(fob_usd_col, 0)) if pd.notna(df_row.get(fob_usd_col)) else 0
        except (ValueError, TypeError):
            fob_savings = 0
        
        try:
            landed_savings = float(df_row.get(landed_usd_col, 0)) if pd.notna(df_row.get(landed_usd_col)) else 0
        except (ValueError, TypeError):
            landed_savings = 0
        
        total_fob_savings_usd += fob_savings
        total_landed_savings_usd += landed_savings
        
        # Classify supplier and update metrics
        if selected_supplier == incumbent:
            incumbent_retained += 1
            total_incumbent_volume += volume
            total_landed_cost_incumbent += extended_cost
        else:
            if selected_supplier != '-':
                if selected_supplier in incumbent_suppliers:
                    new_supplier_count += 1
                    total_new_supplier_volume += volume
                    total_landed_cost_new_suppliers += extended_cost
                else:
                    net_new_supplier_count += 1
                    net_new_supplier_list.add(selected_supplier)
                    total_net_new_supplier_volume += volume
                    total_landed_cost_completely_new_suppliers += extended_cost
        
        unique_suppliers.add(selected_supplier)

print(f"Final metrics after all processing:")
print(f"  - Total annual revenue discount: ${total_annual_revenue_discount:,.2f}")
print(f"  - Total landed savings USD: ${total_landed_savings_usd:,.2f}")
print(f"  - Incumbent retained: {incumbent_retained}")
print(f"  - New suppliers: {new_supplier_count}")
print(f"  - Net new suppliers: {net_new_supplier_count}")
print(f"  - Unique suppliers: {len(unique_suppliers)}")

# In output data want to add new column Redundant Suppliers per Product Family.
'''
basically count how many unique selected suppliers are there per product family and add a column named above and add those value for each part.
'''
# Calculate redundant suppliers per product family
from collections import defaultdict

# Step 1: Map product family to all selected suppliers
family_supplier_map = defaultdict(set)
for row in output_data:
    pf = row["Product Group"]
    supplier = row["Selected Supplier"]
    if supplier != "-":
        family_supplier_map[pf].add(supplier)

# Step 2: Calculate the count of unique suppliers per family
redundancy_count_map = {pf: len(suppliers) for pf, suppliers in family_supplier_map.items()}

# Step 3: Add redundancy value to each output row
for row in output_data:
    pf = row["Product Group"]
    if pf == "No group available":
        row["Redundant Suppliers per Product Group"] = 1
    else:
        row["Redundant Suppliers per Product Group"] = redundancy_count_map.get(pf, 0)

# Step 4: Convert to DataFrame
output_df = pd.DataFrame(output_data)

# Step 5: Move column to index 14
if "Redundant Suppliers per Product Group" in output_df.columns:
    redundancy_col = output_df.pop("Redundant Suppliers per Product Group")
    output_df.insert(14, "Redundant Suppliers per Product Group", redundancy_col)

all_incumbents = set(output_df["Incumbent Supplier"].dropna().unique())

incumbent_rows = output_df[output_df["Selected Supplier"] == output_df["Incumbent Supplier"]]
incumbent_suppliers_unique = set(incumbent_rows["Selected Supplier"].dropna().unique())

new_rows = output_df[output_df["Selected Supplier"] != output_df["Incumbent Supplier"]]
new_suppliers = set(new_rows["Selected Supplier"].dropna().unique())

new_rows = output_df[output_df["Selected Supplier"] != output_df["Incumbent Supplier"]]
new_suppliers_existing = set(
    new_rows["Selected Supplier"].dropna().unique()
).intersection(all_incumbents)

net_new_suppliers = new_suppliers - all_incumbents
import numpy as np
# # --- Add country column from country_supplier_mapping.csv ---
# Helper to get supplier country from port mapping
# Load supplier port file and port-country mapping
supplier_port_file = "Supplier Port per Part table 070925.csv"
port_country_map = {
    'DALIAN': 'China', 
    'NINGBO': 'China', 
    'QINGDAO': 'China', 
    'QINGDAO2': 'China', 
    'SHANGHAI': 'China',
    'SHENZHEN': 'China', 
    'TIANJIN': 'China', 
    'XINGANG': 'China', 
    'XIAMEN': 'China',
    'AHMEDABAD': 'India', 
    'CHENNAI': 'India', 
    'DADRI': 'India', 
    'MUMBAI': 'India',
    'MUNDRA': 'India', 
    'NHAVA SHEVA': 'India',
    'SURABAYA': 'Indonesia',
    'PORT KLANG': 'Malaysia', 
    'PASIR GUDANG': 'Malaysia', 
    'TANJUNG PELAPAS': 'Malaysia',
    'BUSAN': 'South Korea',
    'KAOHSIUNG': 'Taiwan', 
    'KEELUNG': 'Taiwan', 
    'TAICHUNG': 'Taiwan', 
    'TAIPEI': 'Taiwan',
    'BANGKOK': 'Thailand',
    'LAEM CHABANG': 'Thailand',
    'HO CHI MINH CITY': 'Vietnam', 
    'VUNG TAU': 'Vietnam', 
    'HAI PHONG': 'Vietnam',
    'VIRGINIA': 'India'
}
supplier_port_df = pd.read_csv(supplier_port_file)
supplier_port_df = supplier_port_df.set_index('ROW ID #')

def get_supplier_country(row_id, supplier_name):
    try:
        port = supplier_port_df.loc[row_id, supplier_name] if supplier_name in supplier_port_df.columns else np.nan
        if pd.isna(port):
            return np.nan
        return port_country_map.get(str(port).strip().upper(), np.nan)
    except Exception:
        return np.nan


output_df['Country'] = output_df.apply(
    lambda row: get_supplier_country(row["ROW ID #"], row["Selected Supplier"]), axis=1
)

total_landed_savings_usd = pd.to_numeric(output_df['Landed Cost Savings USD'], errors='coerce').sum()


total_landed_cost_incumbent = output_df.loc[
    output_df["Incumbent Supplier"] == output_df["Selected Supplier"],
    "Landed Extended Cost USD"
].sum()

total_landed_cost_completely_new_suppliers = output_df.loc[
    ~output_df["Selected Supplier"].isin(incumbent_suppliers),
    "Landed Extended Cost USD"
].sum()

total_landed_cost_new_suppliers = output_df.loc[
    (output_df["Incumbent Supplier"] != output_df["Selected Supplier"]) &
    (output_df["Selected Supplier"].isin(incumbent_suppliers)),
    "Landed Extended Cost USD"
].sum()
summary_data = [
    # ["Total FOB Savings USD", total_fob_savings_usd],
    ["Total Landed Cost Savings USD", total_landed_savings_usd],
    # ["Total Annual Revenue Discount USD", total_annual_revenue_discount],
    # ["Total Cost of no valid suppliers", total_cost_not_awarded],

    ["Total Landed Cost where Incumbent Suppliers Retained", total_landed_cost_incumbent],
    
    ["Total Landed Cost where bid is awarded to New Suppliers", total_landed_cost_new_suppliers],
    
    ["Total Landed Cost where bid is awarded to Completely New Suppliers", total_landed_cost_completely_new_suppliers],
    
    ["Total parts where Incumbent Suppliers Retained", incumbent_retained],
    
    ["Total parts where bid is awarded to New Suppliers", new_supplier_count],
    ["Total parts where bid is awarded to Net New Suppliers", net_new_supplier_count],


    ["Parts not awarded to any supplier", parts_where_no_bids],
    ["", ""],
    ["Totally New Suppliers", len(net_new_supplier_list)],
    ["Total Unique Suppliers", len(unique_suppliers)],

]

# --- Write to Excel ---
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario: {round(incumbent_retained/14077*100, 2)}% New Supplier, {round((net_new_supplier_count+new_supplier_count)/14077*100, 2)}% Incumbent"
    header_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
    worksheet.merge_range(0, 0, 0, len(list(output_data[0].keys()))-1, scenario_header, header_format)

    # Define formats
    bold_format = workbook.add_format({'bold': True})
    # num_format = workbook.add_format({'num_format': '000,00.00'})
    usd_format = workbook.add_format({'num_format': '$000,00.00'})
    # int_format = workbook.add_format({'num_format': '000,00.00'})

    # Write summary into separate logical blocks:
    summary_row = 2
    # Grouped indices
    cost_metrics = summary_data[0:4]
    supplier_metrics = summary_data[4:15]

    # Write cost metrics: Columns A & B
    for i, item in enumerate(cost_metrics):
        worksheet.write(summary_row + i, 0, item[0], bold_format)
        worksheet.write(summary_row + i, 1, item[1], usd_format)

    # Write supplier metrics: Columns D & E
    for i, item in enumerate(supplier_metrics):
        worksheet.write(summary_row + i, 3, item[0], bold_format)
        worksheet.write(summary_row + i, 4, item[1])

    # --- Write total evaluated cost row ---
    total_label_row = summary_row + max(len(cost_metrics), len(supplier_metrics)) + 1
    worksheet.write(total_label_row, 0, "Total Landed Cost Evaluated", bold_format)

    # Formula for summing cost values (adjust B3:B8 if more/less than 6 rows of cost)

    total_cost =  total_landed_savings_usd + total_landed_cost_incumbent + total_landed_cost_new_suppliers + total_landed_cost_completely_new_suppliers
    worksheet.write(total_label_row, 1, total_cost, usd_format)

    # Write output table
    df_output = output_df
    df_output.to_excel(writer, sheet_name="Sheet1", startrow=13, index=False)

# --- Timer ---
elapsed_time = time.time() - start_time
print(f"\n✅ Done. Output written to '{output_file}'")
print(f"⏱ Time taken: {elapsed_time:.2f} seconds")