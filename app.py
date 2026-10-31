import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Meesho Profit Calculator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Meesho Seller Order Profit Calculator")
st.caption("Upload your Meesho Completed Payment Excel file to calculate SKU-wise Delivered Profit and Net Profit after Returns.")


def load_order_payments(uploaded_file):
    """Load and clean the 'Order Payments' sheet from a Meesho Payment Excel file."""
    xls = pd.ExcelFile(uploaded_file)
    if "Order Payments" not in xls.sheet_names:
        st.error(f"'Order Payments' sheet not found! Available sheets: {', '.join(xls.sheet_names)}")
        return None

    # Read first few rows to dynamically locate the header row containing 'Supplier SKU'
    preview = pd.read_excel(uploaded_file, sheet_name="Order Payments", header=None, nrows=10)
    header_row_idx = 1  # Default in Meesho files (Row 2 in Excel)
    for idx, row in preview.iterrows():
        row_vals = [str(v).strip() for v in row.dropna().values]
        if "Supplier SKU" in row_vals and "Final Settlement Amount" in row_vals:
            header_row_idx = idx
            break

    df = pd.read_excel(uploaded_file, sheet_name="Order Payments", header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = ["Supplier SKU", "Live Order Status", "Final Settlement Amount"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing required columns in 'Order Payments': {', '.join(missing)}")
        return None

    # Drop the Meesho formula description row and any empty rows
    df = df.dropna(subset=["Supplier SKU", "Live Order Status"], how="any").copy()
    df["Supplier SKU"] = df["Supplier SKU"].astype(str).str.strip()
    df["Live Order Status"] = df["Live Order Status"].astype(str).str.strip()

    # Filter out any header/formula row where Live Order Status is a formula code like 'U'
    valid_statuses = ["delivered", "return", "rto", "exchange", "shipped", "cancelled"]
    df = df[
        df["Live Order Status"].str.lower().isin(valid_statuses)
        | pd.to_numeric(df["Final Settlement Amount"], errors="coerce").notna()
    ].copy()

    df["Final Settlement Amount"] = pd.to_numeric(df["Final Settlement Amount"], errors="coerce").fillna(0.0)
    if "Quantity" in df.columns:
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(1.0)
    else:
        df["Quantity"] = 1.0

    if "Product Name" not in df.columns:
        df["Product Name"] = ""

    return df.reset_index(drop=True)


# 1. File Upload Section
uploaded_excel = st.file_uploader(
    "📂 Upload Meesho Payment Excel File (.xlsx)",
    type=["xlsx", "xls"]
)

if uploaded_excel is not None:
    df = load_order_payments(uploaded_excel)

    if df is not None and not df.empty:
        # Separate orders by Live Order Status (case-insensitive matching)
        status_lower = df["Live Order Status"].str.lower()
        df_delivered = df[status_lower == "delivered"].copy()
        df_return = df[status_lower == "return"].copy()
        df_rto = df[status_lower == "rto"].copy()

        # Overview Counts
        st.subheader("1️⃣ Order Status Summary")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Orders in Sheet", len(df))
        col2.metric("Delivered Orders", len(df_delivered), f"₹{df_delivered['Final Settlement Amount'].sum():,.2f}")
        col3.metric("Return Orders", len(df_return), f"₹{df_return['Final Settlement Amount'].sum():,.2f}")
        col4.metric("RTO Orders (Ignored)", len(df_rto), "₹0.00 (Ignored)")

        st.divider()

        # 2. SKU Cost Input Section (Only for Delivered Orders)
        st.subheader("2️⃣ Enter Cost per Order/Unit for Delivered Supplier SKUs")

        # Optional: Load saved SKU costs from a CSV file
        with st.expander("💾 Optional: Upload a previously saved SKU Cost CSV"):
            saved_cost_file = st.file_uploader("Upload SKU Costs CSV", type=["csv"], key="cost_csv")
            saved_costs_map = {}
            if saved_cost_file is not None:
                try:
                    cost_df_uploaded = pd.read_csv(saved_cost_file)
                    if "Supplier SKU" in cost_df_uploaded.columns and "Cost per Unit (INR)" in cost_df_uploaded.columns:
                        saved_costs_map = dict(
                            zip(
                                cost_df_uploaded["Supplier SKU"].astype(str).str.strip(),
                                pd.to_numeric(cost_df_uploaded["Cost per Unit (INR)"], errors="coerce").fillna(0.0)
                            )
                        )
                        st.success("Loaded saved SKU costs successfully!")
                except Exception as e:
                    st.warning(f"Could not read CSV: {e}")

        # Group Delivered Orders by Supplier SKU
        sku_group = (
            df_delivered.groupby("Supplier SKU", as_index=False)
            .agg(
                Product_Name=("Product Name", "first"),
                Delivered_Orders=("Supplier SKU", "count"),
                Total_Quantity=("Quantity", "sum"),
                Avg_Settlement=("Final Settlement Amount", "mean"),
                Total_Settlement=("Final Settlement Amount", "sum"),
            )
            .sort_values(by="Delivered_Orders", ascending=False)
            .reset_index(drop=True)
        )

        sku_group["Cost per Unit (INR)"] = sku_group["Supplier SKU"].map(
            lambda x: float(saved_costs_map.get(x, 0.0))
        )

        st.info("💡 Tap or click on the **Cost per Unit (INR)** column below to enter your product cost for each SKU.")

        edited_sku_table = st.data_editor(
            sku_group[[
                "Supplier SKU",
                "Product_Name",
                "Delivered_Orders",
                "Total_Quantity",
                "Avg_Settlement",
                "Total_Settlement",
                "Cost per Unit (INR)"
            ]],
            column_config={
                "Supplier SKU": st.column_config.TextColumn("Supplier SKU", disabled=True),
                "Product_Name": st.column_config.TextColumn("Product Name", disabled=True),
                "Delivered_Orders": st.column_config.NumberColumn("Delivered Orders", disabled=True),
                "Total_Quantity": st.column_config.NumberColumn("Total Qty", disabled=True),
                "Avg_Settlement": st.column_config.NumberColumn("Avg Settlement (₹)", format="₹%.2f", disabled=True),
                "Total_Settlement": st.column_config.NumberColumn("Total Settlement (₹)", format="₹%.2f", disabled=True),
                "Cost per Unit (INR)": st.column_config.NumberColumn(
                    "Cost per Unit (INR) ✏️",
                    min_value=0.0,
                    step=1.0,
                    format="₹%.2f",
                    required=True
                ),
            },
            hide_index=True,
            use_container_width=True,
            key="sku_cost_editor"
        )

        # Allow user to download their entered SKU costs for future sessions
        cost_csv_bytes = edited_sku_table[["Supplier SKU", "Cost per Unit (INR)"]].to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Save SKU Costs as CSV for Next Time",
            data=cost_csv_bytes,
            file_name="meesho_sku_costs.csv",
            mime="text/csv"
        )

        st.divider()

        # 3. Profit Calculation Section
        st.subheader("3️⃣ Profit Breakdown & Final Summary")

        # Map entered costs back to delivered orders
        cost_lookup = dict(zip(edited_sku_table["Supplier SKU"], edited_sku_table["Cost per Unit (INR)"]))
        df_delivered["Unit Cost (INR)"] = df_delivered["Supplier SKU"].map(cost_lookup).fillna(0.0)
        df_delivered["Total Order Cost (INR)"] = df_delivered["Unit Cost (INR)"] * df_delivered["Quantity"]
        df_delivered["Order Profit (INR)"] = df_delivered["Final Settlement Amount"] - df_delivered["Total Order Cost (INR)"]

        # Calculate SKU-level summary
        edited_sku_table["Total Cost (INR)"] = edited_sku_table["Cost per Unit (INR)"] * edited_sku_table["Total_Quantity"]
        edited_sku_table["Delivered Profit (INR)"] = edited_sku_table["Total_Settlement"] - edited_sku_table["Total Cost (INR)"]

        total_delivered_settlement = df_delivered["Final Settlement Amount"].sum()
        total_delivered_cost = df_delivered["Total Order Cost (INR)"].sum()
        total_delivered_profit = df_delivered["Order Profit (INR)"].sum()

        # Return settlement deduction:
        # In Meesho sheets, Return Final Settlement Amount is negative (e.g., -1548.00).
        # Using abs() ensures we always deduct the return loss magnitude from Delivered Profit.
        raw_return_settlement = df_return["Final Settlement Amount"].sum()
        return_deduction = abs(raw_return_settlement)
        net_total_profit = total_delivered_profit - return_deduction

        # Display Final KPI Cards
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Delivered Settlement", f"₹{total_delivered_settlement:,.2f}")
        k2.metric("Total Delivered Product Cost", f"₹{total_delivered_cost:,.2f}")
        k3.metric("Delivered Orders Profit", f"₹{total_delivered_profit:,.2f}")
        k4.metric(
            "Net Total Profit (After Returns)",
            f"₹{net_total_profit:,.2f}",
            delta=f"-₹{return_deduction:,.2f} Return Deduction",
            delta_color="normal" if net_total_profit >= 0 else "inverse"
        )

        # Display SKU-wise Profit Table
        st.markdown("#### 📦 SKU-Wise Delivered Profit Breakdown")
        st.dataframe(
            edited_sku_table[[
                "Supplier SKU",
                "Product_Name",
                "Delivered_Orders",
                "Total_Quantity",
                "Total_Settlement",
                "Cost per Unit (INR)",
                "Total Cost (INR)",
                "Delivered Profit (INR)"
            ]].style.format({
                "Total_Quantity": "{:.0f}",
                "Total_Settlement": "₹{:,.2f}",
                "Cost per Unit (INR)": "₹{:,.2f}",
                "Total Cost (INR)": "₹{:,.2f}",
                "Delivered Profit (INR)": "₹{:,.2f}",
            }),
            use_container_width=True,
            hide_index=True
        )

        # Expandable Order-Level Details
        with st.expander("🔍 View Individual Delivered Orders & Profit"):
            order_cols = [c for c in ["Sub Order No", "Order Date", "Supplier SKU", "Quantity", "Final Settlement Amount", "Unit Cost (INR)", "Total Order Cost (INR)", "Order Profit (INR)"] if c in df_delivered.columns]
            st.dataframe(df_delivered[order_cols], use_container_width=True, hide_index=True)

        with st.expander(f"↩️ View Return Orders ({len(df_return)} orders | Total: ₹{raw_return_settlement:,.2f})"):
            ret_cols = [c for c in ["Sub Order No", "Order Date", "Supplier SKU", "Product Name", "Live Order Status", "Final Settlement Amount"] if c in df_return.columns]
            st.dataframe(df_return[ret_cols], use_container_width=True, hide_index=True)

        # Export Full Profit Report to CSV
        report_csv = edited_sku_table.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📊 Download Full SKU Profit Report (CSV)",
            data=report_csv,
            file_name="meesho_profit_report.csv",
            mime="text/csv"
        )
