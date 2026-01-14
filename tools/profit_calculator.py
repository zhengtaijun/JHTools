from io import BytesIO
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


def render():
    st.subheader("💰 Profit Calculator")
    st.caption("All data is calculated locally · Multi-product supported · Updated by Andy Wang")

    num_products = st.number_input("Number of products", min_value=1, max_value=20, value=1)

    base_costs = []
    service_rates = []
    for i in range(num_products):
        st.markdown(f"**Product {i+1}**")
        col1, col2 = st.columns([2, 1])
        with col1:
            cost = st.number_input(
                f"Base cost (excl. GST) – Product {i+1}",
                min_value=0.0, step=0.01, format="%.2f",
                value=None, placeholder="E.g. 289.75",
                key=f"base_cost_{i}",
            ) or 0.0
        with col2:
            rate = st.radio(f"Service Fee – P{i+1}", ["15%", "5%"], horizontal=True, key=f"rate_{i}")
            rate_val = 0.15 if rate == "15%" else 0.05

        base_costs.append(cost)
        service_rates.append(rate_val)

    total_base_cost = sum(base_costs)
    service_fees = [c * r for c, r in zip(base_costs, service_rates)]
    cost_with_service = [c + s for c, s in zip(base_costs, service_fees)]
    total_cost_excl_gst = sum(cost_with_service)
    total_cost_incl_gst = total_cost_excl_gst * 1.15

    total_volume = st.number_input("Total volume (m³)", min_value=0.0, step=0.0001, format="%.3f") or 0.0
    shipping_unit_price = st.number_input("Shipping unit price (NZD/m³, excl. GST)", min_value=0.0, step=0.01, format="%.2f", value=150.0)
    shipping_cost_excl_gst = total_volume * shipping_unit_price
    shipping_cost_incl_gst = shipping_cost_excl_gst * 1.15

    sale_price = st.number_input("Input sale price (GST included, NZD)", min_value=0.0, step=0.01, format="%.2f", value=None) or 0.0

    rent = sale_price * 0.10

    total_expense = total_cost_incl_gst + shipping_cost_incl_gst + rent
    profit_with_gst = sale_price - total_expense
    profit_no_gst = profit_with_gst / 1.15 if profit_with_gst else 0.0

    def pct(n):
        return f"{(n/(sale_price or 1)*100):.2f}%" if sale_price else "-"

    result_rows = [
        ["Base Cost (Sum)", total_base_cost, pct(total_base_cost)],
        ["Service Fees", sum(service_fees), pct(sum(service_fees))],
        ["Product Total Cost (incl. GST)", total_cost_incl_gst, pct(total_cost_incl_gst)],
        ["Shipping (incl. GST)", shipping_cost_incl_gst, pct(shipping_cost_incl_gst)],
        ["Rent (10%)", rent, pct(rent)],
        ["Total Cost", total_expense, pct(total_expense)],
        ["Profit (incl. GST)", profit_with_gst, pct(profit_with_gst)],
        ["Profit (excl. GST)", profit_no_gst, ""],
    ]
    df_res = pd.DataFrame(result_rows, columns=["Item", "Amount (NZD)", "Ratio to Sale Price"])
    df_res["Amount (NZD)"] = df_res["Amount (NZD)"].map(lambda x: f"{x:.2f}")
    st.table(df_res)

    with st.expander("Calculation details"):
        st.markdown(f"""
- **Base Cost Total** = {total_base_cost:.2f} NZD
- **Service Fee Total** = {sum(service_fees):.2f} NZD
- **Cost w/ Service Fee (excl. GST)** = {total_cost_excl_gst:.2f} NZD
- **Cost w/ GST** = {total_cost_incl_gst:.2f} NZD
- **Shipping (excl. GST)** = {shipping_cost_excl_gst:.2f} NZD
- **Shipping (incl. GST)** = {shipping_cost_incl_gst:.2f} NZD
- **Rent** = {rent:.2f} NZD
- **Total Expense** = {total_expense:.2f} NZD
- **Profit (incl. GST)** = {profit_with_gst:.2f} NZD
- **Profit (excl. GST)** = {profit_no_gst:.2f} NZD
        """)

    if sale_price > 0:
        fig, ax = plt.subplots()
        ax.pie(
            [total_cost_incl_gst, shipping_cost_incl_gst, rent, max(profit_with_gst, 0)],
            labels=["Product Cost", "Shipping", "Rent", "Profit"],
            autopct="%1.1f%%",
            startangle=90,
        )
        ax.axis("equal")
        st.pyplot(fig)

    if st.button("Export results to Excel"):
        out = BytesIO()
        df_res.to_excel(out, index=False)
        out.seek(0)
        st.download_button("📥 Download Excel", out, file_name="profit_result.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
