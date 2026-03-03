# =========================================================
# STREAMLIT APP – COMBINED APPLICATION
# Nameplate Split + CPM  AND  Inflation Forecast
# =========================================================

import streamlit as st
import pandas as pd
import altair as alt
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide")

st.title("📊 Media Planning Toolkit")

# ---------------------------------------------------------
# Main Tabs
# ---------------------------------------------------------
main_tab1, main_tab2 = st.tabs(
    ["Nameplate Spend Split & CPM Tool", "2026 Inflation Forecast Calculator"]
)

# =========================================================
# TAB 1 – NAMEPLATE SPLIT + CPM
# =========================================================
with main_tab1:

    st.header("Nameplate Spend Split & CPM Impact Tool")

    EU5_COUNTRIES = ["France", "Germany", "Italy", "Spain", "UK"]
    FAMILIES = ["Pro", "Blue", "FCSD", "Model E"]

    uploaded_file = st.file_uploader(
        "Upload Nameplate Excel File",
        type=["xlsx"],
        key="nameplate_upload"
    )

    if uploaded_file:

        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip().str.lower()

        # ---------------------------------------
        # Sidebar Filters
        # ---------------------------------------
        st.sidebar.header("Market Filters")
        all_countries = sorted(df["country"].unique())
        select_all = st.sidebar.checkbox("Select All Markets", value=True)

        if select_all:
            selected_countries = all_countries
        else:
            eu5 = st.sidebar.multiselect(
                "EU5 Countries",
                [c for c in all_countries if c in EU5_COUNTRIES]
            )
            eso = st.sidebar.multiselect(
                "ESO Countries",
                [c for c in all_countries if c not in EU5_COUNTRIES]
            )
            selected_countries = eu5 + eso

        if not selected_countries:
            st.warning("Please select markets")
            st.stop()

        df = df[df["country"].isin(selected_countries)]

        # ---------------------------------------
        # CLEAN TOTAL COST
        # ---------------------------------------
        df["total_cost"] = (
            df["total_cost"]
            .astype(str)
            .str.replace(",", "", regex=False)
            .astype(float)
        )

        df = df.rename(columns={"total_cost": "previous_media_spend"})

        # ---------------------------------------
        # FAMILY TOTALS
        # ---------------------------------------
        family_totals = (
            df.groupby("family")["previous_media_spend"]
            .sum()
            .to_dict()
        )

        timestamp_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        final_outputs = []

        tabs = st.tabs(FAMILIES)

        for tab, fam in zip(tabs, FAMILIES):
            with tab:

                fam_df = df[df["family"] == fam].copy()

                if fam_df.empty:
                    st.info("No data available")
                    continue

                col_chart, col_input = st.columns([4, 1])

                with col_input:
                    revised_total = st.number_input(
                        f"{fam} Revised Spend",
                        value=int(family_totals.get(fam, 0)),
                        step=100,
                        min_value=0
                    )

                # ---------------------------------------
                # PROPORTIONAL REALLOCATION
                # ---------------------------------------
                fam_df["current_revised_spend"] = (
                    revised_total
                    * fam_df["previous_media_spend"]
                    / fam_df["previous_media_spend"].sum()
                ).round(0)

                # ---------------------------------------
                # CPM IMPACT CALCULATION
                # ---------------------------------------
                fam_df["cpm_revised_spend"] = np.where(
                    fam_df["cpm prior"] != 0,
                    (fam_df["cpm revised"] / fam_df["cpm prior"])
                    * fam_df["previous_media_spend"],
                    np.nan
                ).round(0)

                # ---------------------------------------
                # RATIO (Previous / Current Revised)
                # ---------------------------------------
                fam_df["ratio"] = np.where(
                    fam_df["current_revised_spend"] != 0,
                    fam_df["previous_media_spend"]
                    / fam_df["current_revised_spend"],
                    np.nan
                ).round(4)

                fam_df["timestamp"] = timestamp_now

                final_outputs.append(fam_df)

                # ---------------------------------------
                # CHART
                # ---------------------------------------
                chart_df = fam_df.melt(
                    id_vars="nameplate",
                    value_vars=["previous_media_spend", "current_revised_spend"],
                    var_name="Spend Type",
                    value_name="Spend"
                )

                chart = (
                    alt.Chart(chart_df)
                    .mark_bar()
                    .encode(
                        x="nameplate:N",
                        y="Spend:Q",
                        xOffset="Spend Type:N",
                        color="Spend Type:N",
                        tooltip=["nameplate", "Spend Type", "Spend"]
                    )
                    .properties(height=420)
                )

                with col_chart:
                    st.altair_chart(chart, use_container_width=True)

                    # ---------------------------------------
                    # DISPLAY IN REQUIRED ORDER
                    # ---------------------------------------
                    display_df = fam_df[[
                        "timestamp",
                        "nameplate",
                        "family",
                        "country",
                        "previous_media_spend",
                        "current_revised_spend",
                        "cpm_revised_spend",
                        "ratio"
                    ]]

                    st.dataframe(display_df, use_container_width=True)

        # ---------------------------------------
        # DOWNLOAD – 2 SHEETS
        # ---------------------------------------
        if final_outputs:

            final_df = pd.concat(final_outputs, ignore_index=True)

            # ================================
            # SHEET 1 – Revised Spend Ratio
            # ================================
            sheet1 = final_df.copy()

            sheet1["ratio"] = np.where(
                sheet1["current_revised_spend"] != 0,
                sheet1["previous_media_spend"]
                / sheet1["current_revised_spend"],
                np.nan
            ).round(4)

            sheet1 = sheet1[[
                "timestamp",
                "nameplate",
                "family",
                "country",
                "previous_media_spend",
                "current_revised_spend",
                "ratio"
            ]]

            # ================================
            # SHEET 2 – CPM Based Ratio
            # ================================
            sheet2 = final_df.copy()

            sheet2["ratio"] = np.where(
                sheet2["cpm_revised_spend"] != 0,
                sheet2["previous_media_spend"]
                / sheet2["cpm_revised_spend"],
                np.nan
            ).round(4)

            sheet2 = sheet2[[
                "timestamp",
                "nameplate",
                "family",
                "country",
                "previous_media_spend",
                "cpm_revised_spend",
                "ratio"
            ]]

            output = BytesIO()

            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                sheet1.to_excel(writer, sheet_name="Revised Spend Ratio", index=False)
                sheet2.to_excel(writer, sheet_name="CPM Based Ratio", index=False)

            output.seek(0)

            st.download_button(
                "Download Excel Output",
                data=output,
                file_name="nameplate_split_cpm.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# =========================================================
# TAB 2 – INFLATION FORECAST TOOL (UNCHANGED)
# =========================================================
# =========================================================
# TAB 2 – INFLATION FORECAST TOOL (ROW LEVEL OUTPUT)
# =========================================================
# =========================================================
# TAB 2 – 2026 INFLATION FORECAST CALCULATOR
# =========================================================
with main_tab2:

    st.header("2026 Inflation Forecast Calculator")

    uploaded_file2 = st.file_uploader(
        "Upload WPP Media 2026 Inflation Forecast.xlsx",
        type=["xlsx"],
        key="inflation_upload"
    )

    if uploaded_file2:

        try:
            inflation_df = pd.read_excel(uploaded_file2, sheet_name="Inflation")
            cost_df = pd.read_excel(uploaded_file2, sheet_name="Cost")

            inflation_df.columns = inflation_df.columns.str.strip()
            cost_df.columns = cost_df.columns.str.strip()

            # Media columns start after first 2 columns (market + family)
            media_cols = inflation_df.columns[2:]

            # ---------------------------------------------
            # Clean Inflation %
            # ---------------------------------------------
            for col in media_cols:

                if inflation_df[col].dtype == "object":
                    inflation_df[col] = (
                        inflation_df[col]
                        .astype(str)
                        .str.replace("%", "", regex=False)
                        .astype(float)
                    )

                # Convert to decimal if needed
                if inflation_df[col].max() > 1:
                    inflation_df[col] = inflation_df[col] / 100

            # ---------------------------------------------
            # Merge Cost + Inflation
            # ---------------------------------------------
            merged = pd.merge(
                cost_df,
                inflation_df,
                on="market",
                suffixes=("_cost", "_inflation")
            )

            result_df = cost_df.copy()

            # ---------------------------------------------
            # Apply Inflation (Current Values)
            # ---------------------------------------------
            for col in media_cols:
                result_df[col] = (
                    merged[f"{col}_cost"]
                    * (1 + merged[f"{col}_inflation"])
                ).round(2)

            # =============================================
            # FAMILY LEVEL TABS
            # =============================================
            families = result_df["family"].dropna().unique().tolist()
            family_tabs = st.tabs(families)

            for tab, fam in zip(family_tabs, families):

                with tab:

                    fam_df = result_df[result_df["family"] == fam].copy()

                    if fam_df.empty:
                        st.info("No data available")
                        continue

                    # -----------------------------------------
                    # FAMILY TOTAL AFTER INFLATION
                    # -----------------------------------------
                    total_current = fam_df[media_cols].sum().sum()

                    # Editable revised total
                    revised_total = st.number_input(
                        f"{fam} Revised Total Spend",
                        value=float(total_current),
                        step=1000.0,
                        min_value=0.0,
                        key=f"{fam}_revised"
                    )

                    # -----------------------------------------
                    # CREATE REVISED MEDIA COLUMNS
                    # -----------------------------------------
                    for col in media_cols:

                        media_current_total = fam_df[col].sum()

                        # Revised total for this media
                        if total_current == 0:
                            media_total_revised = 0
                        else:
                            media_total_revised = (
                                revised_total
                                * (media_current_total / total_current)
                            )

                        # Distribute proportionally to rows
                        if media_current_total == 0:
                            fam_df[f"{col}_REVISED"] = 0
                        else:
                            fam_df[f"{col}_REVISED"] = (
                                media_total_revised
                                * (fam_df[col] / media_current_total)
                            ).round(2)

                    # -----------------------------------------
                    # Add Timestamp
                    # -----------------------------------------
                    fam_df["timestamp"] = datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    # -----------------------------------------
                    # Final Column Order
                    # -----------------------------------------
                    revised_cols = [f"{col}_REVISED" for col in media_cols]

                    display_cols = (
                        ["timestamp", "nameplate", "family"]
                        + list(media_cols)
                        + revised_cols
                    )

                    st.dataframe(
                        fam_df[display_cols],
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"Error: {e}")

