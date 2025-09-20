# main.py
import streamlit as st
import pandas as pd
import numpy as np
import json
import datetime
import altair as alt
from itertools import product

st.set_page_config(layout="wide", page_title="WK What-If Premium Simulator")

# ----------------------------
# Helpers: load rate tables from excel
# ----------------------------
@st.cache_data
def load_rate_tables(excel_path="wk_rate_tables_elaborate.xlsx"):
    try:
        xls = pd.ExcelFile(excel_path)
    except Exception as e:
        st.error(f"Could not find or open '{excel_path}' in project root. Upload it and reload. Error: {e}")
        st.stop()
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
    return sheets

sheets = load_rate_tables()

# Convenience dfs
df_class = sheets.get("ClassRates", pd.DataFrame())
df_state = sheets.get("StateFactors", pd.DataFrame())
df_deductible = sheets.get("DeductibleFactors", pd.DataFrame())
df_limit = sheets.get("LimitFactors", pd.DataFrame())
df_exmods = sheets.get("ExMods", pd.DataFrame())
df_payroll = sheets.get("PayrollBands", pd.DataFrame())
df_adj = sheets.get("PolicyAdjustments", pd.DataFrame())

# Normalize some fields to expected names if needed
if "RatePer100" not in df_class.columns and "Rate" in df_class.columns:
    df_class = df_class.rename(columns={"Rate": "RatePer100"})

# ----------------------------
# Lookup functions (safe defaults)
# ----------------------------
def get_class_rate(class_code):
    try:
        row = df_class[df_class["ClassCode"].astype(str) == str(class_code)]
        if not row.empty:
            return float(row.iloc[0]["RatePer100"])
    except Exception:
        pass
    # default fallback
    if not df_class.empty:
        return float(df_class.iloc[0]["RatePer100"])
    return 1.0

def get_state_factor(state):
    try:
        row = df_state[df_state["State"].astype(str) == str(state)]
        if not row.empty:
            return float(row.iloc[0]["Factor"])
    except Exception:
        pass
    if not df_state.empty:
        return float(df_state.iloc[0]["Factor"])
    return 1.0

def get_deductible_factor(deductible):
    try:
        # find the largest deductible <= provided amount and return its factor
        df = df_deductible.sort_values("Deductible")
        df_valid = df[df["Deductible"] <= deductible]
        if df_valid.empty:
            return float(df.iloc[0]["Factor"])
        return float(df_valid.iloc[-1]["Factor"])
    except Exception:
        return 1.0

def get_limit_factor(limit_label):
    try:
        row = df_limit[df_limit["Limit"].astype(str) == str(limit_label)]
        if not row.empty:
            return float(row.iloc[0]["Factor"])
    except Exception:
        pass
    if not df_limit.empty:
        return float(df_limit.iloc[0]["Factor"])
    return 1.0

def get_exmod_multiplier(exmod_value):
    # df_exmods should define ranges or mapping; try range lookup
    try:
        # Accept rows with Min/Max or Value/Multiplier
        if {"Min","Max","Multiplier"}.issubset(set(df_exmods.columns)):
            row = df_exmods[(df_exmods["Min"] <= exmod_value) & (df_exmods["Max"] >= exmod_value)]
            if not row.empty:
                return float(row.iloc[0]["Multiplier"])
        elif {"ExMod","Multiplier"}.issubset(set(df_exmods.columns)):
            # find nearest ExMod mapping
            row = df_exmods.iloc[(df_exmods["ExMod"] - exmod_value).abs().argsort()[:1]]
            return float(row.iloc[0]["Multiplier"])
    except Exception:
        pass
    return float(1.0)

def get_payroll_band_adj(payroll):
    # Return multiplier or additive adjustment if present
    try:
        if {"MinPayroll", "MaxPayroll", "AdjMultiplier"}.issubset(set(df_payroll.columns)):
            row = df_payroll[(df_payroll["MinPayroll"] <= payroll) & (df_payroll["MaxPayroll"] >= payroll)]
            if not row.empty:
                return float(row.iloc[0]["AdjMultiplier"])
        elif {"PayrollBand","BaseAdj"}.issubset(set(df_payroll.columns)):
            # fallback: no band match
            return 1.0
    except Exception:
        pass
    return 1.0

def get_policy_adjustments():
    # Return dict of values for PolicyFee, TerrorismSurcharge, PremiumTax etc.
    adjustments = {"PolicyFee":0.0, "TerrorismSurcharge":0.0, "PremiumTax":0.0}
    try:
        for _, r in df_adj.iterrows():
            adj = str(r.get("Adjustment","")).strip()
            val = r.get("Value", 0)
            if adj:
                adjustments[adj] = float(val)
    except Exception:
        pass
    return adjustments

# ----------------------------
# Premium calc function (stepwise breakdown)
# ----------------------------
def calculate_premium(payroll, class_code, state, deductible, exmod, limit, optional_coverages):
    # Protect against None
    payroll = float(payroll or 100000)
    class_code = class_code or (df_class.iloc[0]["ClassCode"] if not df_class.empty else "")
    state = state or (df_state.iloc[0]["State"] if not df_state.empty else "")
    deductible = float(deductible or 0)
    exmod = float(exmod or 1.0)
    limit = limit or (df_limit.iloc[0]["Limit"] if not df_limit.empty else "Standard")
    optional_coverages = optional_coverages or []

    # Step 1: base premium from class rate
    class_rate = get_class_rate(class_code)
    base_premium = (payroll / 100.0) * class_rate

    # Step 2: apply state factor
    state_factor = get_state_factor(state)
    after_state = base_premium * state_factor

    # Step 3: deductible factor
    ded_factor = get_deductible_factor(deductible)
    after_ded = after_state * ded_factor

    # Step 4: exmod multiplier
    exmod_mult = get_exmod_multiplier(exmod)
    after_exmod = after_ded * exmod_mult

    # Step 5: limit factor
    limit_factor = get_limit_factor(limit)
    after_limit = after_exmod * limit_factor

    # Step 6: payroll band adjustment (multiplier)
    payroll_adj = get_payroll_band_adj(payroll)
    after_payroll = after_limit * payroll_adj

    # Step 7: coverage adjustments (optional coverages additive/minimum logic)
    coverage_df = sheets.get("Coverages", pd.DataFrame())
    coverages_total = 0.0
    if not coverage_df.empty:
        for cov in optional_coverages:
            row = coverage_df[coverage_df['Coverage'].astype(str)==str(cov)]
            if not row.empty:
                cov_factor = float(row.iloc[0].get("Factor", 0))
                min_prem = float(row.iloc[0].get("MinPremium", 0))
                cov_prem = max(after_payroll * cov_factor, min_prem)
                coverages_total += cov_prem

    # Policy adjustments
    adjustments = get_policy_adjustments()
    fee = adjustments.get("PolicyFee", 0.0)
    terrorism = adjustments.get("TerrorismSurcharge", 0.0)  # percent e.g., 0.02
    tax = adjustments.get("PremiumTax", 0.0)               # percent

    pre_fee_total = after_payroll + coverages_total
    premium_with_fee = pre_fee_total + fee
    premium_with_terror = premium_with_fee * (1.0 + terrorism)
    premium_with_tax = premium_with_terror * (1.0 + tax)

    breakdown = {
        "Base Premium": round(base_premium,2),
        "State Factor": round(after_state - base_premium,2),
        "Deductible Factor": round(after_ded - after_state,2),
        "ExMod Factor": round(after_exmod - after_ded,2),
        "Limit Factor": round(after_limit - after_exmod,2),
        "Payroll Band Adj": round(after_payroll - after_limit,2),
        "Coverages": round(coverages_total,2),
        "Policy Fee + Surcharges + Tax": round(premium_with_tax - (after_payroll + coverages_total),2)
    }

    return round(premium_with_tax,2), breakdown

# ----------------------------
# AI-style explainer
# ----------------------------
def explain_changes(breakdown_base, breakdown_new, top_n=3):
    # compute delta for each factor
    delta = {}
    for k in breakdown_base:
        delta[k] = round(breakdown_new.get(k,0) - breakdown_base.get(k,0),2)
    # sort by absolute impact
    top = sorted(delta.items(), key=lambda x: abs(x[1]), reverse=True)[:top_n]
    explanations = []
    for factor, change in top:
        if change > 0:
            explanations.append(f"Premium increased by ${change:,.2f} due to higher {factor}.")
        elif change < 0:
            explanations.append(f"Premium decreased by ${-change:,.2f} due to lower {factor}.")
    if not explanations:
        explanations = ["No significant factor changes."]
    return explanations

# ----------------------------
# AI Optimizer (heuristic search)
# ----------------------------
def optimizer_search(payroll, class_code, state, exmod, limit, optional_coverages, target_reduction_pct=0):
    # discrete options
    deductibles = sorted(df_deductible["Deductible"].unique()) if not df_deductible.empty else [0,1000,5000,10000,20000]
    limits = df_limit["Limit"].unique().tolist() if not df_limit.empty else [limit]
    # baseline premium
    base_prem, _ = calculate_premium(payroll, class_code, state, 0, exmod, limit, optional_coverages)
    target_premium = base_prem * (1 - target_reduction_pct/100.0)

    # search for best combo that yields premium <= target_premium
    best = {"deductible": None, "limit": None, "premium": base_prem, "reduction_pct":0}
    for d,l in product(deductibles, limits):
        prem, _ = calculate_premium(payroll, class_code, state, d, exmod, l, optional_coverages)
        reduction_pct = (base_prem - prem)/base_prem*100 if base_prem>0 else 0
        if prem < best["premium"]:
            best = {"deductible": int(d), "limit": l, "premium": prem, "reduction_pct": round(reduction_pct,2)}
            # early stop if we meet-or-exceed target
            if prem <= target_premium:
                break
    return best

# ----------------------------
# UI: file upload / parse
# ----------------------------
st.title("WK Policy What-If Premium Simulator")

col1, col2 = st.columns([2,3])

with col1:
    st.header("Upload / Policy")
    uploaded = st.file_uploader("Upload policy JSON (or leave blank to enter manually)", type=["json"])
    applied_defaults = []
    policy = {}
    if uploaded:
        try:
            policy = json.load(uploaded)
            st.success("Policy JSON loaded.")
        except Exception as e:
            st.error(f"Failed to parse JSON: {e}")
            policy = {}
    st.markdown("**Or fill values manually below (manual values override uploaded file):**")

    # Defaults (also used when uploaded policy missing fields)
    default_values = {
        "Payroll": 100000,
        "Deductible": 0,
        "ExMod": 1.0,
        "Coverage": "Standard",
        "Limit": df_limit.iloc[0]["Limit"] if not df_limit.empty else "Standard",
        "OptionalCoverages": []
    }

    # Read mandatory fields: ClassCode, State, Payroll
    # if uploaded policy missing, we will later use defaults or force user fill mandatory fields
    uploaded_payroll = policy.get("Payroll")
    uploaded_class = policy.get("ClassCode")
    uploaded_state = policy.get("State")
    uploaded_deductible = policy.get("Deductible")
    uploaded_exmod = policy.get("ExMod")
    uploaded_coverage = policy.get("Coverage")
    uploaded_limit = policy.get("Limit")
    uploaded_optional = [c.get("Coverage") for c in policy.get("OptionalCoverages", []) if c.get("Selected", True)] if policy.get("OptionalCoverages") else policy.get("OptionalCoverages", [])

    # Manual input (use uploaded if exists else defaults)
    payroll_input = st.number_input("Payroll *", min_value=0, value=int(uploaded_payroll) if uploaded_payroll is not None else default_values["Payroll"])
    class_options = df_class["ClassCode"].astype(str).tolist() if not df_class.empty else ["1001"]
    default_class_idx = class_options.index(str(uploaded_class)) if uploaded_class and str(uploaded_class) in class_options else 0
    class_input = st.selectbox("Class Code *", class_options, index=default_class_idx)
    state_options = df_state["State"].astype(str).tolist() if not df_state.empty else ["CA"]
    default_state_idx = state_options.index(str(uploaded_state)) if uploaded_state and str(uploaded_state) in state_options else 0
    state_input = st.selectbox("State *", state_options, index=default_state_idx)

    deductible_input = st.select_slider("Deductible", options=sorted(df_deductible["Deductible"].unique().tolist()) if not df_deductible.empty else [0,1000,5000,10000], value=int(uploaded_deductible) if uploaded_deductible is not None else default_values["Deductible"])
    exmod_input = st.number_input("Experience Modifier (ExMod)", min_value=0.5, max_value=3.0, value=float(uploaded_exmod) if uploaded_exmod is not None else default_values["ExMod"], step=0.01)
    limit_options = df_limit["Limit"].astype(str).tolist() if not df_limit.empty else [default_values["Limit"]]
    default_limit_idx = limit_options.index(str(uploaded_limit)) if uploaded_limit and str(uploaded_limit) in limit_options else 0
    limit_input = st.selectbox("Limit", limit_options, index=default_limit_idx)

    # coverages
    coverage_options = sheets.get("Coverages", pd.DataFrame())
    if not coverage_options.empty:
        cov_list = coverage_options["Coverage"].astype(str).tolist()
    else:
        cov_list = []
    default_cov = uploaded_optional if uploaded_optional else []
    coverages_input = st.multiselect("Optional Coverages", cov_list, default=default_cov)

    # Check which defaults applied (if uploaded policy had missing fields)
    applied_defaults = []
    if uploaded:
        if uploaded_payroll is None:
            applied_defaults.append(f"Payroll={default_values['Payroll']}")
        if uploaded_class is None:
            applied_defaults.append(f"ClassCode={class_options[0]}")
        if uploaded_state is None:
            applied_defaults.append(f"State={state_options[0]}")
        if uploaded_deductible is None:
            applied_defaults.append(f"Deductible={default_values['Deductible']}")
        if uploaded_exmod is None:
            applied_defaults.append(f"ExMod={default_values['ExMod']}")
        if uploaded_limit is None:
            applied_defaults.append(f"Limit={default_values['Limit']}")
        if not uploaded.get("OptionalCoverages"):
            # if missing entire optional coverages
            applied_defaults.append("OptionalCoverages=[]")

    if applied_defaults:
        st.info("Some fields were missing in uploaded policy. Defaults applied: " + ", ".join(applied_defaults))

with col2:
    st.header("What-If Controls & Results")
    st.markdown("Adjust the What-If inputs to see premium impact instantly.")
    # Mirror What-If controls (start with Base values)
    st.subheader("What-If Inputs")
    payroll_new = st.number_input("What-If Payroll", min_value=0, value=int(payroll_input))
    class_new = st.selectbox("What-If Class Code", class_options, index=class_options.index(class_input))
    state_new = st.selectbox("What-If State", state_options, index=state_options.index(state_input))
    deductible_new = st.select_slider("What-If Deductible", options=sorted(df_deductible["Deductible"].unique().tolist()) if not df_deductible.empty else [0,1000,5000,10000], value=int(deductible_input))
    exmod_new = st.number_input("What-If ExMod", min_value=0.5, max_value=3.0, value=float(exmod_input), step=0.01)
    limit_new = st.selectbox("What-If Limit", limit_options, index=limit_options.index(limit_input))
    coverages_new = st.multiselect("What-If Optional Coverages", cov_list, default=coverages_input)

    # Calculate base & what-if
    premium_base, breakdown_base = calculate_premium(payroll_input, class_input, state_input, deductible_input, exmod_input, limit_input, coverages_input)
    premium_new, breakdown_new = calculate_premium(payroll_new, class_new, state_new, deductible_new, exmod_new, limit_new, coverages_new)

    # Summary
    st.subheader("Premium Comparison")
    delta_prem = premium_new - premium_base
    st.metric("Base Premium", f"${premium_base:,.2f}", delta=None)
    st.metric("What-If Premium", f"${premium_new:,.2f}", delta=f"${delta_prem:,.2f}")

    # Stepwise breakdown tables
    st.subheader("Stepwise Breakdown")
    left, right = st.columns(2)
    with left:
        st.write("Base Breakdown")
        st.table(pd.DataFrame.from_dict(breakdown_base, orient="index", columns=["Amount"]))
    with right:
        st.write("What-If Breakdown")
        st.table(pd.DataFrame.from_dict(breakdown_new, orient="index", columns=["Amount"]))

    # Delta highlight
    st.subheader("Delta Highlight")
    delta = {}
    for k in breakdown_base:
        diff = round(breakdown_new.get(k,0) - breakdown_base.get(k,0),2)
        if diff != 0:
            delta[k] = diff
    if delta:
        delta_df = pd.DataFrame(list(delta.items()), columns=["Factor", "Change ($)"])
        delta_df["Impact"] = delta_df["Change ($)"].apply(lambda x: "Increase" if x>0 else "Decrease")
        st.table(delta_df)
    else:
        st.write("No factor changes between Base and What-If.")

    # AI Explainer
    st.subheader("AI Explainer (Heuristic)")
    if st.button("AI Explainer"):
        explanations = explain_changes(breakdown_base, breakdown_new)
        for e in explanations:
            st.info(e)

    # AI Optimizer
    st.subheader("AI Optimizer (Heuristic search)")
    tgt = st.slider("Target premium reduction (%)", 0, 50, 10)
    if st.button("Run AI Optimizer"):
        best = optimizer_search(payroll_input, class_input, state_input, exmod_input, limit_input, coverages_input, target_reduction_pct=tgt)
        st.success(f"Suggested Deductible: ${best['deductible']}, Suggested Limit: {best['limit']}")
        st.write(f"Estimated Premium under suggestion: ${best['premium']:,.2f} (reduction {best['reduction_pct']}%)")

    # Factor contribution chart
    st.subheader("Factor Contribution Chart")
    dfb = pd.DataFrame.from_dict(breakdown_base, orient="index", columns=["Base"]).reset_index().rename(columns={"index":"Factor"})
    dfn = pd.DataFrame.from_dict(breakdown_new, orient="index", columns=["What-If"]).reset_index().rename(columns={"index":"Factor"})
    dfc = pd.merge(dfb, dfn, on="Factor", how="outer").fillna(0)
    df_melt = dfc.melt(id_vars="Factor", var_name="Scenario", value_name="Contribution")
    df_melt["Color"] = df_melt["Contribution"].apply(lambda x: "Increase" if x>0 else "Decrease")
    chart = alt.Chart(df_melt).mark_bar().encode(
        x=alt.X("Factor:N", sort=None),
        y=alt.Y("Contribution:Q"),
        color=alt.Color("Color:N", scale=alt.Scale(domain=["Increase","Decrease"], range=["red","green"])),
        column="Scenario:N"
    ).properties(width=200, height=300)
    st.altair_chart(chart, use_container_width=True)

st.markdown("---")
st.caption("Built for hackathon: use wk_rate_tables_elaborate.xlsx in project root. Uses only open-source libs.")
