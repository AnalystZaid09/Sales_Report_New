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
                    
                    st.write("**Step 4:** Creating VLOOKUP mappings...")
                    bm_col = [c for c in pm.columns if "brand" in c and "manager" in c][0]
                    brand_col = [c for c in pm.columns if c == "brand"][0]
                    pm_unique = pm.drop_duplicates(subset="asin", keep="first")
                    
                    bm_map = pm_unique.set_index("asin")[bm_col]
                    brand_map = pm_unique.set_index("asin")[brand_col]
                    
                    Working["Brand Manager"] = Working["asin"].map(bm_map)
                    Working["Brand"] = Working["asin"].map(brand_map)
                    st.write("✓ Brand Manager and Brand mapped")
                    
                    st.write("**Step 5:** Filtering data...")
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
                
                # Generate pivot tables
                st.markdown("---")
                st.subheader("📊 Pivot Tables")
                
                tab1, tab2 = st.tabs(["Brand Manager Analysis", "Brand Analysis"])
                
                with tab1:
                    st.write("**Daily Sales and Quantity by Brand Manager**")
                    pivot_Brand_Manager = pd.pivot_table(
                        Working,
                        index="Brand Manager",
                        columns="date",
                        values=["quantity", "item-price"],
                        aggfunc="sum",
                        fill_value=0,
                        margins=True,
                        margins_name="Grand Total"
                    )
                    pivot_Brand_Manager = pivot_Brand_Manager.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
                    pivot_Brand_Manager = pivot_Brand_Manager.rename(
                        columns={"quantity": "Sum of quantity", "item-price": "Sum of item-price"},
                        level=1
                    )
                    
                    st.dataframe(pivot_Brand_Manager, use_container_width=True)
                    
                    # Download button
                    output1 = BytesIO()
                    with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                        pivot_Brand_Manager.to_excel(writer, sheet_name='Brand Manager Analysis')
                    output1.seek(0)
                    st.download_button(
                        label="📥 Download Brand Manager Analysis",
                        data=output1,
                        file_name="brand_manager_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with tab2:
                    st.write("**Daily Sales and Quantity by Brand**")
                    pivot_Brand = pd.pivot_table(
                        Working,
                        index="Brand",
                        columns="date",
                        values=["quantity", "item-price"],
                        aggfunc="sum",
                        fill_value=0,
                        margins=True,
                        margins_name="Grand Total"
                    )
                    pivot_Brand = pivot_Brand.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
                    pivot_Brand = pivot_Brand.rename(
                        columns={"quantity": "Sum of quantity", "item-price": "Sum of item-price"},
                        level=1
                    )
                    
                    st.dataframe(pivot_Brand, use_container_width=True)
                    
                    # Download button
                    output2 = BytesIO()
                    with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                        pivot_Brand.to_excel(writer, sheet_name='Brand Analysis')
                    output2.seek(0)
                    st.download_button(
                        label="📥 Download Brand Analysis",
                        data=output2,
                        file_name="brand_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
       - Must contain columns: `asin`, `quantity`, `item-price`, `product-name`, `item-status`, `purchase-date`
    
    2. **Upload Product Master File**: Click on the second file uploader and select your `PM.xlsx` file
       - Must contain columns: `ASIN`, `Brand Manager`, `Brand`
    
    3. **Generate Analysis**: Click the "Generate Analysis" button to process the data
    
    4. **View Results**: 
       - Check the key metrics dashboard
       - Explore the pivot tables in each tab
       - Download Excel files for further analysis
    
    ### Data Processing:
    - Removes cancelled orders and invalid entries
    - Maps Brand Manager and Brand from Product Master
    - Creates daily pivot tables with grand totals
    - Filters out zero quantities and prices
    """)

# Footer
st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")