import pandas as pd
import numpy as np
from collections import defaultdict

def compare_tariffs():
    """
    Comprehensive comparison of tariff_old.csv and tariff_part_level_cleaned.csv
    showing major changes in tariffs per part, country, and metal type.
    """
    # Read both CSV files
    print("Loading tariff_old.csv...")
    df_old = pd.read_csv('tariff_old.csv')
    
    print("Loading tariff_part_level_cleaned.csv...")
    df_new = pd.read_csv('tariff_part_level_cleaned.csv')
    
    print("Merging datasets...")
    # Merge on ROW ID #, Country, and Metal Type for accurate comparison
    merged = pd.merge(
        df_old,
        df_new,
        on=['ROW ID #', 'Country', 'Metal Type'],
        suffixes=('_old', '_new'),
        how='outer',
        indicator=True
    )
    
    # Calculate changes
    merged['tariff_change'] = merged['tariff_value_new'] - merged['tariff_value_old']
    merged['tariff_change_pct'] = np.where(
        merged['tariff_value_old'] != 0,
        (merged['tariff_change'] / merged['tariff_value_old']) * 100,
        np.where(merged['tariff_value_new'] > 0, np.inf, 0)
    )
    merged['status'] = merged.apply(
        lambda x: 'New' if x['_merge'] == 'right_only' 
        else 'Removed' if x['_merge'] == 'left_only'
        else 'Increased' if x['tariff_change'] > 0
        else 'Decreased' if x['tariff_change'] < 0
        else 'No Change',
        axis=1
    )
    
    # Filter for matched rows (exist in both files)
    matched = merged[merged['_merge'] == 'both'].copy()
    
    print("\n" + "="*100)
    print("COMPREHENSIVE TARIFF CHANGE ANALYSIS REPORT")
    print("="*100)
    
    # ========== OVERALL SUMMARY ==========
    print("\n1. OVERALL SUMMARY")
    print("-"*100)
    print(f"Total rows in old file: {len(df_old):,}")
    print(f"Total rows in new file: {len(df_new):,}")
    print(f"Rows matched (exist in both): {len(matched):,}")
    print(f"New rows (only in new file): {len(merged[merged['_merge'] == 'right_only']):,}")
    print(f"Removed rows (only in old file): {len(merged[merged['_merge'] == 'left_only']):,}")
    
    print(f"\nChanges in matched rows:")
    print(f"  - Increased: {len(matched[matched['tariff_change'] > 0]):,}")
    print(f"  - Decreased: {len(matched[matched['tariff_change'] < 0]):,}")
    print(f"  - No Change: {len(matched[matched['tariff_change'] == 0]):,}")
    
    # ========== BY METAL TYPE ==========
    print("\n2. SUMMARY BY METAL TYPE")
    print("-"*100)
    metal_summary = matched.groupby('Metal Type').agg({
        'ROW ID #': 'nunique',
        'tariff_change': ['sum', 'mean', 'count'],
        'tariff_value_old': 'mean',
        'tariff_value_new': 'mean'
    }).round(4)
    metal_summary.columns = ['Unique Parts', 'Total Change', 'Avg Change', 'Rows Changed', 'Avg Old Tariff', 'Avg New Tariff']
    
    # Count increases/decreases by metal type
    metal_changes = matched.groupby(['Metal Type', 'status']).size().unstack(fill_value=0)
    metal_summary = metal_summary.join(metal_changes[['Increased', 'Decreased']] if 'Increased' in metal_changes.columns else pd.DataFrame())
    
    print(f"\n{'Metal Type':<25} {'Unique Parts':<15} {'Total Change':<15} {'Avg Change':<15} {'Increased':<12} {'Decreased':<12}")
    print("-"*100)
    for metal in metal_summary.index:
        parts = int(metal_summary.loc[metal, 'Unique Parts'])
        total_chg = f"{metal_summary.loc[metal, 'Total Change']:.4f}"
        avg_chg = f"{metal_summary.loc[metal, 'Avg Change']:.4f}"
        inc = int(metal_summary.loc[metal, 'Increased']) if 'Increased' in metal_summary.columns else 0
        dec = int(metal_summary.loc[metal, 'Decreased']) if 'Decreased' in metal_summary.columns else 0
        print(f"{metal:<25} {parts:<15} {total_chg:<15} {avg_chg:<15} {inc:<12} {dec:<12}")
    
    # ========== BY COUNTRY ==========
    print("\n3. SUMMARY BY COUNTRY")
    print("-"*100)
    country_summary = matched.groupby('Country').agg({
        'ROW ID #': 'nunique',
        'tariff_change': ['sum', 'mean', 'count'],
        'tariff_value_old': 'mean',
        'tariff_value_new': 'mean'
    }).round(4)
    country_summary.columns = ['Unique Parts', 'Total Change', 'Avg Change', 'Rows Changed', 'Avg Old Tariff', 'Avg New Tariff']
    
    country_changes = matched.groupby(['Country', 'status']).size().unstack(fill_value=0)
    country_summary = country_summary.join(country_changes[['Increased', 'Decreased']] if 'Increased' in country_changes.columns else pd.DataFrame())
    country_summary = country_summary.sort_values('Unique Parts', ascending=False)
    
    print(f"\n{'Country':<30} {'Unique Parts':<15} {'Total Change':<15} {'Avg Change':<15} {'Increased':<12} {'Decreased':<12}")
    print("-"*100)
    for country in country_summary.head(20).index:  # Show top 20
        parts = int(country_summary.loc[country, 'Unique Parts'])
        total_chg = f"{country_summary.loc[country, 'Total Change']:.4f}"
        avg_chg = f"{country_summary.loc[country, 'Avg Change']:.4f}"
        inc = int(country_summary.loc[country, 'Increased']) if 'Increased' in country_summary.columns else 0
        dec = int(country_summary.loc[country, 'Decreased']) if 'Decreased' in country_summary.columns else 0
        print(f"{country:<30} {parts:<15} {total_chg:<15} {avg_chg:<15} {inc:<12} {dec:<12}")
    if len(country_summary) > 20:
        print(f"... and {len(country_summary) - 20} more countries")
    
    # ========== BY METAL TYPE AND COUNTRY ==========
    print("\n4. SUMMARY BY METAL TYPE AND COUNTRY")
    print("-"*100)
    metal_country_summary = matched.groupby(['Metal Type', 'Country']).agg({
        'ROW ID #': 'nunique',
        'tariff_change': ['sum', 'mean'],
        'tariff_value_old': 'mean',
        'tariff_value_new': 'mean'
    }).round(4)
    metal_country_summary.columns = ['Unique Parts', 'Total Change', 'Avg Change', 'Avg Old Tariff', 'Avg New Tariff']
    metal_country_summary = metal_country_summary.sort_values('Total Change', ascending=False)
    
    print(f"\nTop 30 combinations by total change:")
    print(f"{'Metal Type':<25} {'Country':<30} {'Unique Parts':<15} {'Total Change':<15} {'Avg Change':<15}")
    print("-"*100)
    for (metal, country), row in metal_country_summary.head(30).iterrows():
        parts = int(row['Unique Parts'])
        total_chg = f"{row['Total Change']:.4f}"
        avg_chg = f"{row['Avg Change']:.4f}"
        print(f"{metal:<25} {country:<30} {parts:<15} {total_chg:<15} {avg_chg:<15}")
    
    # ========== MAJOR INCREASES ==========
    print("\n5. MAJOR TARIFF INCREASES (Top 50)")
    print("-"*100)
    increases = matched[matched['tariff_change'] > 0].copy()
    increases = increases.sort_values('tariff_change', ascending=False)
    
    print(f"\n{'ROW ID':<10} {'Metal Type':<25} {'Country':<30} {'Old Tariff':<15} {'New Tariff':<15} {'Increase':<15} {'% Change':<15}")
    print("-"*100)
    for _, row in increases.head(50).iterrows():
        pct = f"{row['tariff_change_pct']:.2f}%" if not np.isinf(row['tariff_change_pct']) else "N/A"
        print(f"{int(row['ROW ID #']):<10} {row['Metal Type']:<25} {row['Country']:<30} "
              f"{row['tariff_value_old']:<15.4f} {row['tariff_value_new']:<15.4f} "
              f"{row['tariff_change']:<15.4f} {pct:<15}")
    
    # ========== MAJOR DECREASES ==========
    print("\n6. MAJOR TARIFF DECREASES (Top 50)")
    print("-"*100)
    decreases = matched[matched['tariff_change'] < 0].copy()
    decreases = decreases.sort_values('tariff_change', ascending=True)
    
    print(f"\n{'ROW ID':<10} {'Metal Type':<25} {'Country':<30} {'Old Tariff':<15} {'New Tariff':<15} {'Decrease':<15} {'% Change':<15}")
    print("-"*100)
    for _, row in decreases.head(50).iterrows():
        pct = f"{row['tariff_change_pct']:.2f}%" if not np.isinf(row['tariff_change_pct']) else "N/A"
        print(f"{int(row['ROW ID #']):<10} {row['Metal Type']:<25} {row['Country']:<30} "
              f"{row['tariff_value_old']:<15.4f} {row['tariff_value_new']:<15.4f} "
              f"{row['tariff_change']:<15.4f} {pct:<15}")
    
    # ========== UNIQUE PARTS WITH INCREASES BY COUNTRY ==========
    print("\n7. UNIQUE PARTS WITH INCREASED TARIFFS - BY COUNTRY")
    print("-"*100)
    country_increases = defaultdict(set)
    for _, row in increases.iterrows():
        country = row['Country']
        row_id = row['ROW ID #']
        country_increases[country].add(row_id)
    
    country_counts = {country: len(row_ids) for country, row_ids in country_increases.items()}
    country_counts_sorted = dict(sorted(country_counts.items(), key=lambda x: x[1], reverse=True))
    
    print(f"\n{'Country':<30} {'Unique Parts with Increases':<30}")
    print("-"*100)
    for country, count in country_counts_sorted.items():
        print(f"{country:<30} {count:<30}")
    print(f"\nTotal unique parts with increases: {sum(country_counts_sorted.values()):,}")
    
    # ========== SAVE DETAILED REPORTS ==========
    print("\n" + "="*100)
    print("GENERATING DETAILED REPORTS...")
    print("="*100)
    
    # 1. All changes report
    all_changes = matched[[
        'ROW ID #', 'Metal Type', 'Country', 
        'tariff_value_old', 'tariff_value_new', 
        'tariff_change', 'tariff_change_pct', 'status'
    ]].copy()
    all_changes = all_changes.sort_values('tariff_change', ascending=False)
    all_changes.to_csv('report_all_changes.csv', index=False)
    print("✓ Saved: report_all_changes.csv (all tariff changes)")
    
    # 2. Increases only
    increases_report = increases[[
        'ROW ID #', 'Metal Type', 'Country',
        'tariff_value_old', 'tariff_value_new',
        'tariff_change', 'tariff_change_pct'
    ]].copy()
    increases_report.to_csv('report_increases_only.csv', index=False)
    print("✓ Saved: report_increases_only.csv (only increases)")
    
    # 3. Summary by metal type
    metal_summary.to_csv('report_by_metal_type.csv')
    print("✓ Saved: report_by_metal_type.csv (summary by metal type)")
    
    # 4. Summary by country
    country_summary.to_csv('report_by_country.csv')
    print("✓ Saved: report_by_country.csv (summary by country)")
    
    # 5. Summary by metal type and country
    metal_country_summary.to_csv('report_by_metal_country.csv')
    print("✓ Saved: report_by_metal_country.csv (summary by metal type and country)")
    
    # 6. Top increases
    top_increases = increases.head(1000)[[
        'ROW ID #', 'Metal Type', 'Country',
        'tariff_value_old', 'tariff_value_new',
        'tariff_change', 'tariff_change_pct'
    ]].copy()
    top_increases.to_csv('report_top_increases.csv', index=False)
    print("✓ Saved: report_top_increases.csv (top 1000 increases)")
    
    # 7. Top decreases
    top_decreases = decreases.head(1000)[[
        'ROW ID #', 'Metal Type', 'Country',
        'tariff_value_old', 'tariff_value_new',
        'tariff_change', 'tariff_change_pct'
    ]].copy()
    top_decreases.to_csv('report_top_decreases.csv', index=False)
    print("✓ Saved: report_top_decreases.csv (top 1000 decreases)")
    
    # 8. New and removed rows
    new_rows = merged[merged['_merge'] == 'right_only'][[
        'ROW ID #', 'Metal Type', 'Country', 'tariff_value_new'
    ]].copy()
    new_rows.columns = ['ROW ID #', 'Metal Type', 'Country', 'tariff_value']
    new_rows.to_csv('report_new_rows.csv', index=False)
    print("✓ Saved: report_new_rows.csv (rows only in new file)")
    
    removed_rows = merged[merged['_merge'] == 'left_only'][[
        'ROW ID #', 'Metal Type', 'Country', 'tariff_value_old'
    ]].copy()
    removed_rows.columns = ['ROW ID #', 'Metal Type', 'Country', 'tariff_value']
    removed_rows.to_csv('report_removed_rows.csv', index=False)
    print("✓ Saved: report_removed_rows.csv (rows only in old file)")
    
    print("\n" + "="*100)
    print("REPORT GENERATION COMPLETE!")
    print("="*100)
    
    return {
        'matched': matched,
        'increases': increases,
        'decreases': decreases,
        'metal_summary': metal_summary,
        'country_summary': country_summary,
        'metal_country_summary': metal_country_summary
    }

if __name__ == "__main__":
    try:
        results = compare_tariffs()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
