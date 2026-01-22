import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Order Analysis Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.title("📊 Order Analysis Dashboard")
st.markdown("Upload your order and product master files to generate comprehensive analytics")

# File upload section
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Orders File")
    orders_file = st.file_uploader("Upload orders.xlsx", type=['xlsx', 'xls'], key="orders")

with col2:
    st.subheader("📦 Product Master File")
    pm_file = st.file_uploader("Upload PM.xlsx", type=['xlsx', 'xls'], key="pm")

# Process button
if st.button("🚀 Generate Analysis", type="primary"):
    if orders_file is None or pm_file is None:
        st.error("⚠️ Please upload both files before processing")
    else:
        with st.spinner("Processing your files..."):
            try:
                # Load data
                Working = pd.read_excel(orders_file)
                pm = pd.read_excel(pm_file)
                
                st.success(f"✅ Loaded {len(Working)} orders and {len(pm)} products")
                
                # Data processing
                with st.expander("📝 Data Processing Steps", expanded=True):
                    st.write("**Step 1:** Converting dates...")
                    Working["date"] = pd.to_datetime(Working["purchase-date"]).dt.date
                    st.write("✓ Date conversion complete")
                    
                    st.write("**Step 2:** Cleaning column names...")
                    pm.columns = pm.columns.str.strip().str.lower()
                    Working.columns = Working.columns.str.strip().str.lower()
                    st.write("✓ Column names normalized")
                    
                    st.write("**Step 3:** Cleaning ASIN data...")
                    pm["asin"] = pm["asin"].astype(str).str.strip()
                    Working["asin"] = Working["asin"].astype(str).str.strip()
                    pm = pm[(pm["asin"] != "") & (pm["asin"].notna()) & (pm["asin"] != "nan")]
                    st.write("✓ ASIN data cleaned")
                    
                    st.write("**Step 4:** Creating VLOOKUP mappings for Brand & Brand Manager...")
                    bm_col = [c for c in pm.columns if "brand" in c and "manager" in c][0]
                    brand_col = [c for c in pm.columns if c == "brand"][0]
                    pm_unique = pm.drop_duplicates(subset="asin", keep="first")
                    
                    bm_map = pm_unique.set_index("asin")[bm_col]
                    brand_map = pm_unique.set_index("asin")[brand_col]
                    
                    Working["Brand Manager"] = Working["asin"].map(bm_map)
                    Working["Brand"] = Working["asin"].map(brand_map)
                    st.write("✓ Brand Manager and Brand mapped")

                    # STEP 5 – Vendor SKU mapping (VLOOKUP-style)
                    st.write("**Step 5:** Mapping Vendor SKU (VLOOKUP-style)...")

                    vendor_candidates = [
                        c for c in pm.columns 
                        if c in ["vendor sku", "vendor_sku", "vendor sku code", "vendor_sku_code"]
                    ]
                    if vendor_candidates:
                        vendor_col = vendor_candidates[0]
                    else:
                        # 4th column like Excel VLOOKUP(P2, A:G, 4, 0)
                        vendor_col = pm.columns[3]

                    vendor_map = pm_unique.set_index("asin")[vendor_col]
                    Working["Vendor SKU"] = Working["asin"].map(vendor_map)
                    Working["Vendor SKU"] = Working["Vendor SKU"].astype(str)
                    st.write("✓ Vendor SKU mapped")

                    # STEP 6: Cost (CP) mapping and inserting column
                    st.write("**Step 6:** Mapping Cost (CP) column and positioning it...")

                    cp_candidates = [c for c in pm.columns if c in ["cp", "cost price", "cost"]]
                    if cp_candidates:
                        cp_col = cp_candidates[0]
                    else:
                        cp_col = pm.columns[7]  # 8th column

                    cp_map = pm_unique.set_index("asin")[cp_col]
                    Working["cost"] = Working["asin"].map(cp_map)

                    Working["cost"] = pd.to_numeric(Working["cost"], errors="coerce").fillna(0)

                    if "sku" in Working.columns:
                        Working["sku"] = Working["sku"].astype(str)

                    # Place cost after item-price
                    cols = list(Working.columns)
                    if "item-price" in cols and "cost" in cols:
                        price_index = cols.index("item-price")
                        cols.insert(price_index + 1, cols.pop(cols.index("cost")))
                        Working = Working[cols]

                    # Reorder Vendor SKU between asin and item-status (Vendor SKU right after asin)
                    cols = list(Working.columns)
                    if "asin" in cols and "Vendor SKU" in cols:
                        cols.remove("Vendor SKU")
                        new_cols = []
                        for c in cols:
                            new_cols.append(c)
                            if c == "asin":
                                new_cols.append("Vendor SKU")
                        Working = Working[new_cols]

                    st.write("✓ Cost column placed and Vendor SKU positioned between asin and item-status (where available)")

                    st.write("**Step 7:** Filtering data...")
                    original_count = len(Working)
                    Working = Working[Working['quantity'] != 0]
                    Working['item-price'] = pd.to_numeric(Working['item-price'], errors='coerce')
                    Working = Working[Working['item-price'] != 0]
                    Working = Working[
                        (Working['product-name'].notna()) &
                        (Working['product-name'].astype(str).str.strip() != '-') &
                        (Working['product-name'].astype(str).str.strip() != '')
                    ]
                    Working = Working[Working['item-status'] != 'Cancelled']
                    st.write(f"✓ Filtered from {original_count} to {len(Working)} valid orders")
                
                # Prepare raw data Excel once so we can reuse
                raw_output = BytesIO()
                with pd.ExcelWriter(raw_output, engine='openpyxl') as writer:
                    Working.to_excel(writer, sheet_name='Processed Orders', index=False)
                raw_output.seek(0)

                # Display metrics
                st.markdown("---")
                st.subheader("📈 Key Metrics")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Orders", f"{len(Working):,}")
                
                with col2:
                    total_revenue = Working['item-price'].sum()
                    st.metric("Total Revenue", f"₹{total_revenue/10000000:.2f}Cr")
                
                with col3:
                    top_brand = Working.groupby('Brand')['item-price'].sum().idxmax()
                    st.metric("Top Brand", top_brand)
                
                with col4:
                    top_manager = Working.groupby('Brand Manager')['item-price'].sum().idxmax()
                    st.metric("Top Manager", top_manager)

                # Raw data download near metrics (TOP)
                st.download_button(
                    label="📥 Download Raw Data (Processed Orders)",
                    data=raw_output,
                    file_name="processed_orders_raw.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_raw_top"
                )
                
                # Generate pivot tables
                st.markdown("---")
                st.subheader("📊 Pivot Tables")
                
                tab1, tab2, tab3, tab4 = st.tabs([
                    "Brand Manager Analysis",
                    "Brand Analysis",
                    "Brand & ASIN Summary",
                    "BM / Brand / ASIN Summary"
                ])
                
                # TAB 1: Brand Manager vs Date
                with tab1:
                    st.write("**Daily Sales and Quantity by Brand Manager**")
                    pivot_Brand_Manager = pd.pivot_table(
                        Working,
                        index="Brand Manager",
                        columns="date",
                        values=["quantity", "item-price"],
                        aggfunc="sum",
                        fill_value=0,
                    )
                    pivot_Brand_Manager = pivot_Brand_Manager.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
                    pivot_Brand_Manager = pivot_Brand_Manager.rename(
                        columns={"quantity": "Sum of quantity", "item-price": "Sum of item-price"},
                        level=1
                    )

                    # Grand Total row at bottom
                    grand_total_bm = pivot_Brand_Manager.sum(numeric_only=True).to_frame().T
                    grand_total_bm.index = pd.Index(["Grand Total"], name=pivot_Brand_Manager.index.name)
                    pivot_Brand_Manager_with_total = pd.concat(
                        [pivot_Brand_Manager, grand_total_bm]
                    )

                    st.dataframe(pivot_Brand_Manager_with_total, width="stretch")
                    
                    # Download button
                    output1 = BytesIO()
                    with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                        pivot_Brand_Manager_with_total.to_excel(
                            writer, sheet_name='Brand Manager Analysis'
                        )
                    output1.seek(0)
                    st.download_button(
                        label="📥 Download Brand Manager Analysis",
                        data=output1,
                        file_name="brand_manager_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # TAB 2: Brand vs Date
                with tab2:
                    st.write("**Daily Sales and Quantity by Brand**")
                    pivot_Brand = pd.pivot_table(
                        Working,
                        index="Brand",
                        columns="date",
                        values=["quantity", "item-price"],
                        aggfunc="sum",
                        fill_value=0,
                    )
                    pivot_Brand = pivot_Brand.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
                    pivot_Brand = pivot_Brand.rename(
                        columns={"quantity": "Sum of quantity", "item-price": "Sum of item-price"},
                        level=1
                    )

                    # Grand Total row at bottom
                    grand_total_brand = pivot_Brand.sum(numeric_only=True).to_frame().T
                    grand_total_brand.index = pd.Index(["Grand Total"], name=pivot_Brand.index.name)
                    pivot_Brand_with_total = pd.concat(
                        [pivot_Brand, grand_total_brand]
                    )
                    
                    st.dataframe(pivot_Brand_with_total, width="stretch")
                    
                    # Download button
                    output2 = BytesIO()
                    with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                        pivot_Brand_with_total.to_excel(writer, sheet_name='Brand Analysis')
                    output2.seek(0)
                    st.download_button(
                        label="📥 Download Brand Analysis",
                        data=output2,
                        file_name="brand_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # TAB 3: Brand & ASIN pivot (Brand alphabetical, Grand Total last)
                with tab3:
                    st.write("**Brand & ASIN Detailed Summary**")

                    pivot_Brand_ASIN = pd.pivot_table(
                        Working,
                        index=["asin","Brand"],
                        values=["quantity", "item-price", "cost"],
                        aggfunc="sum",
                        fill_value=0
                    )

                    pivot_Brand_ASIN = pivot_Brand_ASIN.rename(
                        columns={
                            "quantity": "Sum of quantity",
                            "item-price": "Sum of item-price",
                            "cost": "Sum of cost"
                        }
                    )

                    # Sort by Brand alphabetical then ASIN
                    pivot_Brand_ASIN = pivot_Brand_ASIN.sort_index(level=["Brand", "asin"])

                    # Grand Total row at the end
                    grand_total_values = pivot_Brand_ASIN.sum(numeric_only=True).to_frame().T
                    grand_total_index = pd.MultiIndex.from_tuples(
                        [("Grand Total", "")],
                        names=pivot_Brand_ASIN.index.names
                    )
                    grand_total_values.index = grand_total_index

                    pivot_Brand_ASIN_with_total = pd.concat(
                        [pivot_Brand_ASIN, grand_total_values]
                    )

                    st.dataframe(pivot_Brand_ASIN_with_total, width="stretch")

                    # Download Brand & ASIN Summary
                    output3 = BytesIO()
                    with pd.ExcelWriter(output3, engine='openpyxl') as writer:
                        pivot_Brand_ASIN_with_total.to_excel(writer, sheet_name='Brand_ASIN_Summary')
                    output3.seek(0)

                    st.download_button(
                        label="📥 Download Brand & ASIN Summary",
                        data=output3,
                        file_name="brand_asin_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # TAB 4: BM / Brand / ASIN pivot with subtotals & grand total
                with tab4:
                    st.write("**Brand Manager / Brand / ASIN Summary (with subtotals)**")

                    agg_cols = ["quantity", "item-price", "cost"]
                    base = (
                        Working
                        .groupby(["asin","Brand Manager", "Brand"], dropna=False)[agg_cols]
                        .sum()
                        .reset_index()
                    )

                    base["order"] = 0
                    base = base.sort_values(
                        by=["Brand Manager", "Brand", "quantity"],
                        ascending=[True, True, False]
                    )

                    # Brand subtotals (e.g. Beetel Total)
                    brand_totals = (
                        base
                        .groupby(["Brand Manager", "Brand"], as_index=False)[agg_cols]
                        .sum()
                    )
                    brand_totals["asin"] = brand_totals["Brand"] + " Total"
                    brand_totals["order"] = 1

                    # Manager subtotals (e.g. Ayushi Total)
                    manager_totals = (
                        base
                        .groupby(["Brand Manager"], as_index=False)[agg_cols]
                        .sum()
                    )
                    manager_totals["Brand"] = ""
                    manager_totals["asin"] = manager_totals["Brand Manager"] + " Total"
                    manager_totals["order"] = 2

                    # Grand total row
                    grand_total = pd.DataFrame({
                        "Brand Manager": [""],
                        "Brand": [""],
                        "asin": ["Grand Total"],
                        "quantity": [base["quantity"].sum()],
                        "item-price": [base["item-price"].sum()],
                        "cost": [base["cost"].sum()],
                        "order": [3],
                    })

                    # Combine all
                    combined = pd.concat(
                        [base, brand_totals, manager_totals, grand_total],
                        ignore_index=True
                    )

                    # Flag for grand total so it always goes last
                    combined["is_grand"] = (combined["asin"] == "Grand Total").astype(int)

                    combined = combined.sort_values(
                        by=["is_grand", "Brand Manager", "Brand", "order", "quantity"],
                        ascending=[True, True, True, True, False]
                    )

                    # Rename columns like Excel pivot
                    combined = combined.rename(
                        columns={
                            "quantity": "Sum of quantity",
                            "item-price": "Sum of item-price",
                            "cost": "Sum of cost",
                        }
                    )

                    # Set index to look like multi-level pivot
                    display_df = combined.drop(columns=["order", "is_grand"]).set_index(
                        ["Brand Manager", "Brand", "asin"]
                    )

                    st.dataframe(display_df, width="stretch")

                    # Download BM / Brand / ASIN Summary
                    output4 = BytesIO()
                    with pd.ExcelWriter(output4, engine='openpyxl') as writer:
                        display_df.to_excel(
                            writer,
                            sheet_name='BM_Brand_ASIN_Summary'
                        )
                    output4.seek(0)

                    st.download_button(
                        label="📥 Download BM / Brand / ASIN Summary",
                        data=output4,
                        file_name="bm_brand_asin_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Raw processed data section
                st.markdown("---")
                st.subheader("🧾 Raw Data (Processed Orders)")
                st.dataframe(Working, width="stretch")

                # Raw data download at bottom (BOTTOM) – different key
                st.download_button(
                    label="📥 Download Raw Data (Processed Orders)",
                    data=raw_output,
                    file_name="processed_orders_raw.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_raw_bottom"
                )
                
                st.success("✅ Analysis complete! Download the Excel files above.")
                
            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
                st.exception(e)

# Instructions
with st.expander("ℹ️ How to Use", expanded=not (orders_file and pm_file)):
    st.markdown("""
    ### Instructions:
    
    1. **Upload Orders File**: Click on the first file uploader and select your `orders.xlsx` file  
       - Must contain columns: `asin`, `quantity`, `item-price`, `item-tax`, `product-name`, `item-status`, `purchase-date`
    
    2. **Upload Product Master File**: Click on the second file uploader and select your `PM.xlsx` file  
       - Must contain columns: `ASIN`, `Brand Manager`, `Brand`, `CP` (ideally in column 8) and Vendor SKU (4th column or named)
    
    3. **Generate Analysis**: Click the "Generate Analysis" button to process the data
    
    4. **View Results**: 
       - Check the key metrics dashboard  
       - Explore all four pivot table tabs  
       - View and download the processed raw data table
    
    ### Data Processing:
    - Removes cancelled orders and invalid entries  
    - Maps Brand Manager, Brand, Vendor SKU, and Cost (CP) from Product Master  
    - Creates daily and Brand/ASIN and BM/Brand/ASIN pivot tables with subtotals and grand totals  
    - Filters out zero quantities and prices  
    """)

# Footer
st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
