import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# Page config
# --------------------------------------------------
st.set_page_config(page_title="Order Analysis Dashboard", layout="wide")
st.title("📊 Order Analysis Dashboard")

# --------------------------------------------------
# Upload files
# --------------------------------------------------
c1, c2 = st.columns(2)
with c1:
    orders_file = st.file_uploader("Upload Orders File", type=["xlsx"])
with c2:
    pm_file = st.file_uploader("Upload Product Master File", type=["xlsx"])

# --------------------------------------------------
# Generate button
# --------------------------------------------------
if st.button("🚀 Generate Analysis"):

    if orders_file is None or pm_file is None:
        st.error("Please upload both files")
        st.stop()

    # --------------------------------------------------
    # Load data
    # --------------------------------------------------
    Working = pd.read_excel(orders_file)
    pm = pd.read_excel(pm_file)

    # --------------------------------------------------
    # Cleaning
    # --------------------------------------------------
    Working.columns = Working.columns.str.strip().str.lower()
    pm.columns = pm.columns.str.strip().str.lower()

    Working["date"] = pd.to_datetime(Working["purchase-date"]).dt.date
    Working["asin"] = Working["asin"].astype(str).str.strip()
    pm["asin"] = pm["asin"].astype(str).str.strip()

    pm_unique = pm.drop_duplicates("asin")

    # --------------------------------------------------
    # Mapping
    # --------------------------------------------------
    bm_col = [c for c in pm.columns if "brand" in c and "manager" in c][0]

    Working["Brand"] = Working["asin"].map(pm_unique.set_index("asin")["brand"])
    Working["Brand Manager"] = Working["asin"].map(
        pm_unique.set_index("asin")[bm_col]
    )

    Working["cost"] = pd.to_numeric(
        Working["asin"].map(pm_unique.set_index("asin")["cp"]),
        errors="coerce"
    ).fillna(0)

    # --------------------------------------------------
    # Filters
    # --------------------------------------------------
    Working = Working[
        (Working["quantity"] != 0) &
        (Working["item-price"] != 0) &
        (Working["item-status"] != "Cancelled")
    ]

    # --------------------------------------------------
    # Tabs
    # --------------------------------------------------
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Brand Manager Analysis",
        "Brand Analysis",
        "Brand & ASIN Summary",
        "BM / Brand / ASIN Summary",
        "📊 Summary Pivots",
        "🧾 Raw Data"
    ])

    # ==================================================
    # TAB 1 – BRAND MANAGER ANALYSIS
    # ==================================================
    with tab1:
        pivot_bm = pd.pivot_table(
            Working,
            index="Brand Manager",
            columns="date",
            values=["item-price", "quantity"],
            aggfunc="sum",
            fill_value=0
        )

        pivot_bm = pivot_bm.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
        pivot_bm = pivot_bm.rename(
            columns={"item-price": "Sum of item-price", "quantity": "Sum of quantity"},
            level=1
        )

        pivot_bm[("Grand Total", "Total Sum of quantity")] = (
            pivot_bm.xs("Sum of quantity", level=1, axis=1).sum(axis=1)
        )
        pivot_bm[("Grand Total", "Total Sum of item-price")] = (
            pivot_bm.xs("Sum of item-price", level=1, axis=1).sum(axis=1)
        )

        grand_row = pivot_bm.sum(numeric_only=True).to_frame().T
        grand_row.index = ["Grand Total"]

        pivot_bm_final = pd.concat([pivot_bm, grand_row])

        st.dataframe(pivot_bm_final, use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            pivot_bm_final.to_excel(writer, sheet_name="Brand_Manager_Analysis")
        out.seek(0)

        st.download_button("📥 Download Brand Manager Analysis", out,
                           "brand_manager_analysis.xlsx")

    # ==================================================
    # TAB 2 – BRAND ANALYSIS
    # ==================================================
    with tab2:
        pivot_brand = pd.pivot_table(
            Working,
            index="Brand",
            columns="date",
            values=["item-price", "quantity", "cost"],
            aggfunc="sum",
            fill_value=0
        )

        pivot_brand = pivot_brand.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
        pivot_brand = pivot_brand.rename(
            columns={
                "item-price": "Sum of item-price",
                "quantity": "Sum of quantity",
                "cost": "Sum of cost",
            },
            level=1
        )

        pivot_brand[("Grand Total", "Total Sum of quantity")] = (
            pivot_brand.xs("Sum of quantity", level=1, axis=1).sum(axis=1)
        )
        pivot_brand[("Grand Total", "Total Sum of cost")] = (
            pivot_brand.xs("Sum of cost", level=1, axis=1).sum(axis=1)
        )

        grand_row = pivot_brand.sum(numeric_only=True).to_frame().T
        grand_row.index = ["Grand Total"]

        pivot_brand_final = pd.concat([pivot_brand, grand_row])

        st.dataframe(pivot_brand_final, use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            pivot_brand_final.to_excel(writer, sheet_name="Brand_Analysis")
        out.seek(0)

        st.download_button("📥 Download Brand Analysis", out,
                           "brand_analysis.xlsx")

    # ==================================================
    # TAB 3 – BRAND & ASIN SUMMARY
    # ==================================================
    with tab3:
        brand_asin = (
            Working
            .groupby(["asin", "Brand"])[["quantity", "item-price", "cost"]]
            .sum()
            .reset_index()
            .sort_values("quantity", ascending=False)   # 🔥 SORT
        )

        total_row = brand_asin[["quantity", "item-price", "cost"]].sum().to_frame().T
        total_row.insert(0, "asin", "Grand Total")
        total_row.insert(1, "Brand", "")

        brand_asin_final = pd.concat([brand_asin, total_row], ignore_index=True)

        st.dataframe(brand_asin_final, use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            brand_asin_final.to_excel(writer, index=False,
                                    sheet_name="Brand_ASIN_Summary")
        out.seek(0)

        st.download_button(
            "📥 Download Brand & ASIN Summary",
            out,
            "brand_asin_summary.xlsx"
        )


    # ==================================================
    # TAB 4 – BM / BRAND / ASIN SUMMARY
    # ==================================================
    with tab4:
        bm_brand_asin = (
            Working
            .groupby(["asin", "Brand", "Brand Manager"])[
                ["quantity", "item-price", "cost"]
            ]
            .sum()
            .reset_index()
            .sort_values("quantity", ascending=False)   # 🔥 SORT
        )

        total_row = bm_brand_asin[["quantity", "item-price", "cost"]].sum().to_frame().T
        total_row.insert(0, "asin", "Grand Total")
        total_row.insert(1, "Brand", "")
        total_row.insert(2, "Brand Manager", "")

        bm_brand_asin_final = pd.concat(
            [bm_brand_asin, total_row], ignore_index=True
        )

        st.dataframe(bm_brand_asin_final, use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            bm_brand_asin_final.to_excel(writer, index=False,
                                        sheet_name="BM_Brand_ASIN")
        out.seek(0)

        st.download_button(
            "📥 Download BM / Brand / ASIN Summary",
            out,
            "bm_brand_asin_summary.xlsx"
        )


    # ==================================================
    # TAB 5 – SUMMARY PIVOTS
    # ==================================================
    with tab5:
        st.subheader("🔹 Pivot – Brand")

        brand_summary = (
            Working
            .groupby("Brand")[["quantity", "item-price", "cost"]]
            .sum()
            .reset_index()
            .sort_values("quantity", ascending=False)   # 🔥 SORT
        )

        brand_total = brand_summary[["quantity", "item-price", "cost"]].sum().to_frame().T
        brand_total.insert(0, "Brand", "Grand Total")

        brand_summary_final = pd.concat(
            [brand_summary, brand_total], ignore_index=True
        )


        st.dataframe(brand_summary_final, use_container_width=True)

        out1 = BytesIO()
        with pd.ExcelWriter(out1, engine="openpyxl") as writer:
            brand_summary_final.to_excel(writer, index=False,
                                         sheet_name="Brand_Summary")
        out1.seek(0)

        st.download_button("📥 Download Brand Pivot", out1,
                           "brand_summary_pivot.xlsx")

        st.subheader("🔹 Pivot – Brand Manager")

        bm_summary = (
            Working
            .groupby("Brand Manager")[["quantity", "item-price", "cost"]]
            .sum()
            .reset_index()
            .sort_values("quantity", ascending=False)   # 🔥 SORT
        )

        bm_total = bm_summary[["quantity", "item-price", "cost"]].sum().to_frame().T
        bm_total.insert(0, "Brand Manager", "Grand Total")

        bm_summary_final = pd.concat(
            [bm_summary, bm_total], ignore_index=True
        )


        st.dataframe(bm_summary_final, use_container_width=True)

        out2 = BytesIO()
        with pd.ExcelWriter(out2, engine="openpyxl") as writer:
            bm_summary_final.to_excel(writer, index=False,
                                      sheet_name="BM_Summary")
        out2.seek(0)

        st.download_button("📥 Download Brand Manager Pivot", out2,
                           "brand_manager_summary_pivot.xlsx")

    # ==================================================
    # TAB 6 – RAW DATA
    # ==================================================
    with tab6:
        st.dataframe(Working, use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            Working.to_excel(writer, index=False,
                             sheet_name="Processed_Orders")
        out.seek(0)

        st.download_button("📥 Download Raw Data", out,
                           "processed_orders_raw.xlsx")

    st.success("✅ All reports generated and downloadable!")
