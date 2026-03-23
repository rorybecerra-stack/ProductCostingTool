import streamlit as st
import pandas as pd
import json
from io import BytesIO

# --- 1. CORE DATA ENGINES ---

def process_sap_bom(raw_df, mapping_dict):
    flat_bom = []
    parent_stack = {}
    for _, row in raw_df.iterrows():
        lvl = int(row[mapping_dict['Level']])
        part_id = str(row[mapping_dict['Comp Material']])
        if lvl == 1:
            parent_id = str(row[mapping_dict['Material']])
            parent_stack[lvl] = part_id
        else:
            parent_id = parent_stack.get(lvl - 1)
            parent_stack[lvl] = part_id

        qty_raw = row[mapping_dict['Req. Qty/1000']]
        qty_per = float(qty_raw) / 1000 if pd.notnull(qty_raw) else 0

        flat_bom.append({
            "Parent_ID": parent_id,
            "Child_ID": part_id,
            "Qty_Per": qty_per,
            "UOM": row.get(mapping_dict['Component UOM'], 'EA'),
            "Fixed_Qty_Flag": str(row.get(mapping_dict['Fixed Qty'], '')).upper() == 'X'
        })
    return pd.DataFrame(flat_bom)

def get_purchase_price_with_moq(part_id, required_qty, purchase_df, cny_rate, purchase_moq):
    actual_buy_qty = max(required_qty, purchase_moq)
    part_data = purchase_df[purchase_df['Part_ID'] == part_id]
    
    if part_data.empty: return 0
        
    mask = part_data['Min_Qty'] <= actual_buy_qty
    valid_prices = part_data[mask]
    
    if valid_prices.empty:
        row = part_data.sort_values('Min_Qty', ascending=True).iloc[0]
    else:
        row = valid_prices.sort_values('Min_Qty', ascending=False).iloc[0]
        
    base_price = row['Unit_Price']
    if str(row.get('Currency')).upper() == "CNY":
        base_price = base_price / cny_rate
        
    return (base_price * actual_buy_qty) / required_qty

# --- UPDATED: RETURNS COST BUCKETS ---
def calculate_master(part_id, lot_size, m_rate, a_rate, log_pct, parts_dict, bom_df, routing_dict, purchase_df, overrides, cny_rate, i_tar_rate, efficiency, fixed_qty_flag=False):
    cost_buckets = {"material": 0.0, "labor": 0.0, "logistics": 0.0, "tariff": 0.0, "adders": 0.0, "total": 0.0}
    
    if overrides.get(part_id, {}).get("ignored") or fixed_qty_flag: 
        return cost_buckets

    part = parts_dict.get(part_id, {})
    proc = part.get("Procurement_Type", "F")
    
    if overrides.get(part_id, {}).get("use_std_cost"):
        cost_buckets["material"] = (part.get("Standard_Cost", 0) / part.get("Price_Unit", 1))
        cost_buckets["total"] = cost_buckets["material"]
        return cost_buckets

    if proc == "F":
        u_price = get_purchase_price_with_moq(part_id, lot_size, purchase_df, cny_rate, part.get("Purchase_MOQ", 0))
        scrap_factor = 1 - part.get("Buy_Scrap", 0)
        
        cost_buckets["material"] += u_price / scrap_factor
        cost_buckets["logistics"] += (u_price * log_pct) / scrap_factor
        base_cost = cost_buckets["material"] + cost_buckets["logistics"]

    else: 
        yield_m = 1 - part.get("Make_Scrap", 0)
        if part_id in bom_df['Parent_ID'].values:
            for _, row in bom_df[bom_df['Parent_ID'] == part_id].iterrows():
                req_q = (lot_size * row['Qty_Per']) / yield_m
                child_costs = calculate_master(row['Child_ID'], req_q, m_rate, a_rate, log_pct, parts_dict, bom_df, routing_dict, purchase_df, overrides, cny_rate, i_tar_rate, efficiency, row['Fixed_Qty_Flag'])
                
                # Roll up child costs proportionally
                qty_mult = row['Qty_Per'] / yield_m
                cost_buckets["material"] += child_costs["material"] * qty_mult
                cost_buckets["labor"] += child_costs["labor"] * qty_mult
                cost_buckets["logistics"] += child_costs["logistics"] * qty_mult
                cost_buckets["tariff"] += child_costs["tariff"] * qty_mult
                cost_buckets["adders"] += child_costs["adders"] * qty_mult

        routing = routing_dict.get(part_id, {"Setup_Hrs": 0, "Run_Hrs_Per_Unit": 0})
        current_hrs = (routing["Setup_Hrs"] + (routing["Run_Hrs_Per_Unit"] * lot_size)) / (yield_m * efficiency)
        rate = m_rate if overrides.get(part_id, {}).get("labor_type") == "Machine Shop" else a_rate
        
        cost_buckets["labor"] += ((current_hrs * rate) / lot_size)
        base_cost = cost_buckets["material"] + cost_buckets["labor"] + cost_buckets["logistics"] + cost_buckets["tariff"] + cost_buckets["adders"]

    if overrides.get(part_id, {}).get("apply_tariff"): 
        t_val = base_cost * i_tar_rate
        cost_buckets["tariff"] += t_val
        
    test_fee = overrides.get(part_id, {}).get("test_charge", 0)
    osp_fee = overrides.get(part_id, {}).get("osp_charge", 0)
    cost_buckets["adders"] += (test_fee + osp_fee) / lot_size

    cost_buckets["total"] = cost_buckets["material"] + cost_buckets["labor"] + cost_buckets["logistics"] + cost_buckets["tariff"] + cost_buckets["adders"]
    return cost_buckets

# --- 2. STREAMLIT UI SETUP ---

st.set_page_config(page_title="SAP Cost Dashboard", layout="wide", initial_sidebar_state="expanded")

if 'test_codes_df' not in st.session_state:
    st.session_state.test_codes_df = pd.DataFrame({"Test_Code": ["None", "FAT-01", "HYDRO-01", "XRAY-01"], "Cost": [0.0, 500.0, 250.0, 400.0]})

if 'osp_codes_df' not in st.session_state:
    st.session_state.osp_codes_df = pd.DataFrame({"OSP_Code": ["None", "PAINT-01", "HEAT-01", "ANODIZE-01"], "Cost": [0.0, 150.0, 300.0, 125.0]})

if 'fg_data' not in st.session_state:
    st.session_state.fg_data = pd.DataFrame({"Part_ID": ["FIN-01"], "Lot_Sizes": ["10, 50, 100, 500"]})

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 1.4rem; color: #1f77b4; }
    .stExpander { border: 1px solid #f0f2f6; border-radius: 8px; margin-bottom: 5px; }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("⚙️ Global Rates")
    
    with st.expander("💾 Session Management", expanded=True):
        st.info("Save or load your custom Lot Sizes, Test Codes, and OSP settings.")
        
        uploaded_json = st.file_uploader("Load Saved Session (.json)", type="json")
        if uploaded_json is not None:
            try:
                loaded_data = json.load(uploaded_json)
                st.session_state.fg_data = pd.DataFrame(loaded_data["fg_data"])
                st.session_state.test_codes_df = pd.DataFrame(loaded_data["test_codes"])
                st.session_state.osp_codes_df = pd.DataFrame(loaded_data["osp_codes"])
                st.success("Session loaded! (Clear file to reset)")
            except Exception as e:
                st.error(f"Error loading session: {e}")

        session_export = {
            "fg_data": st.session_state.fg_data.to_dict(orient="records"),
            "test_codes": st.session_state.test_codes_df.to_dict(orient="records"),
            "osp_codes": st.session_state.osp_codes_df.to_dict(orient="records")
        }
        json_string = json.dumps(session_export, indent=4)
        st.download_button(
            label="📥 Save Current Session",
            data=json_string,
            file_name="SAP_Costing_Session.json",
            mime="application/json",
            use_container_width=True
        )
    
    with st.expander("Labor & Shop Rates", expanded=False):
        m_rate = st.number_input("Machine Shop ($/hr)", value=75.0)
        a_rate = st.number_input("Assembly ($/hr)", value=45.0)
        eff = st.slider("Efficiency (%)", 50, 100, 100) / 100
    
    with st.expander("Logistics & Tariffs", expanded=False):
        log_pct = st.slider("Logistics Load (%)", 0, 30, 8) / 100
        cny = st.number_input("USD/CNY Rate", value=7.25)
        i_tar_rate = st.number_input("Duty Rate (%)", value=25.0) / 100

test_dict = dict(zip(st.session_state.test_codes_df["Test_Code"], st.session_state.test_codes_df["Cost"]))
osp_dict = dict(zip(st.session_state.osp_codes_df["OSP_Code"], st.session_state.osp_codes_df["Cost"]))

up_file = st.file_uploader("📂 Upload Excel Master", type="xlsx")

if up_file:
    p_df = pd.read_excel(up_file, sheet_name='Part_Master')
    raw_b = pd.read_excel(up_file, sheet_name='BOM_Structure')
    r_df = pd.read_excel(up_file, sheet_name='Labor_Routing')
    pur_df = pd.read_excel(up_file, sheet_name='Purchase_Matrix')
    
    p_df['Part_ID'] = p_df['Part_ID'].astype(str)
    r_df['Part_ID'] = r_df['Part_ID'].astype(str)
    pur_df['Part_ID'] = pur_df['Part_ID'].astype(str)
    
    parts_dict = p_df.set_index('Part_ID').to_dict('index')
    routing_dict = r_df.set_index('Part_ID').to_dict('index')
    b_df = process_sap_bom(raw_b, {'Material': 'Material', 'Level': 'Level', 'Comp Material': 'Comp Material', 'Req. Qty/1000': 'Req. Qty/1000', 'Component UOM': 'Component UOM', 'Fixed Qty': 'Fixed Qty'})

    c1, c2 = st.columns([1, 2])
    with c1:
        st.subheader("🎯 Target Selection")
        fg_df = st.data_editor(
            st.session_state.fg_data, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={"Part_ID": st.column_config.TextColumn("Finished Good", required=True), "Lot_Sizes": st.column_config.TextColumn("Target Lot Sizes", required=True)},
            key="fg_editor"
        )
        st.session_state.fg_data = fg_df 
        
        fg_targets = {}
        for _, row in fg_df.iterrows():
            pid = str(row['Part_ID']).strip()
            if pid and pid != "None":
                lots = [int(x.strip()) for x in str(row['Lot_Sizes']).split(',') if x.strip().isdigit()]
                if lots:
                    fg_targets[pid] = lots

    demand_by_fg = {fg: {p: {q: 0 for q in lots} for p in parts_dict.keys()} for fg, lots in fg_targets.items()}
    
    for fg, fg_lots in fg_targets.items():
        def explode_demand_fg(parent_id, qty_per_top):
            for q in fg_lots:
                demand_by_fg[fg][parent_id][q] += qty_per_top * q
            children = b_df[b_df['Parent_ID'] == parent_id]
            for _, row in children.iterrows():
                explode_demand_fg(row['Child_ID'], qty_per_top * row['Qty_Per'])
        
        if fg in parts_dict:
            explode_demand_fg(fg, 1.0)

    user_overrides = {fg: {} for fg in fg_targets.keys()}

    with c2:
        st.subheader("🛠️ Component Management")
        search_query = st.text_input("🔍 Global Filter...", placeholder="Type Part ID or Keyword to filter across all tabs...")
        live_calc_active = st.toggle("Enable Real-Time Cost Preview", value=True)

        if fg_targets:
            fg_tabs = st.tabs(list(fg_targets.keys()))
            
            for idx, (fg, fg_lots) in enumerate(fg_targets.items()):
                with fg_tabs[idx]:
                    
                    def get_bom_tree(parent_id, b_df, tree_set):
                        tree_set.add(parent_id)
                        children = b_df[b_df['Parent_ID'] == parent_id]['Child_ID'].unique()
                        for child in children: get_bom_tree(child, b_df, tree_set)
                        return tree_set
                        
                    bom_components = get_bom_tree(fg, b_df, set())
                    filtered_parts = [p for p in bom_components if p in parts_dict and (search_query.lower() in str(p).lower() or search_query.lower() in str(parts_dict[p].get('Description', '')).lower())]

                    fg_part = [p for p in filtered_parts if p == fg]
                    child_parts = [p for p in filtered_parts if p != fg]
                    sorted_parts = fg_part + child_parts

                    for p_id in sorted_parts:
                        data = parts_dict[p_id]
                        is_fg = (p_id == fg)
                        
                        max_demand = max([demand_by_fg[fg][p_id].get(q, 0) for q in fg_lots]) if fg_lots else 0
                        routing_ref = routing_dict.get(p_id, {"Setup_Hrs": 0.0, "Run_Hrs_Per_Unit": 0.0})

                        if is_fg:
                            st.info(f"### 👑 FINISHED GOOD: {p_id} \n **{data.get('Description','')}**")
                            ui_block = st.container()
                        else:
                            ui_block = st.expander(f"📦 {p_id} | {data.get('Description','')[:60]}")

                        with ui_block:
                            m1, m2, m3, m4, m5, m6 = st.columns(6)
                            m1.metric("Source", data.get('Procurement_Type', 'F'))
                            m2.metric("Stock", f"{data.get('Total_Stock', 0):,.0f}")
                            m3.metric("Std Cost", f"${data.get('Standard_Cost', 0):,.2f}")
                            m4.metric("SAP Lot", f"{data.get('Min. Lot Size', 0):,.0f}")
                            m5.metric("Setup (Hrs)", f"{routing_ref.get('Setup_Hrs', 0):,.2f}")
                            m6.metric("Run (Hrs/U)", f"{routing_ref.get('Run_Hrs_Per_Unit', 0):,.3f}")
                            
                            if data.get('Total_Stock', 0) < max_demand and max_demand > 0:
                                m2.caption("⚠️ Stock < Max Demand")
                            
                            ctl1, ctl2, ctl3, ctl4, ctl5 = st.columns(5)
                            with ctl1: ign = st.toggle("Ignore Part", key=f"i_{fg}_{p_id}")
                            with ctl2: std_o = st.checkbox("Use Std Cost", value=(data.get('Total_Stock', 0) >= max_demand and max_demand > 0), key=f"s_{fg}_{p_id}")
                            with ctl3: t_code = st.selectbox("Test Code", list(test_dict.keys()), key=f"t_{fg}_{p_id}")
                            with ctl4: o_code = st.selectbox("OSP Code", list(osp_dict.keys()), key=f"o_{fg}_{p_id}")
                            with ctl5: tar_inc = st.toggle("Duty Appli.", value=False, key=f"tar_{fg}_{p_id}")

                            l_type = st.radio("Primary Workcenter", ["Machine Shop", "Assembly"], key=f"l_{fg}_{p_id}", horizontal=True) if data.get("Procurement_Type") == "E" else "Machine Shop"

                            user_overrides[fg][p_id] = {
                                "ignored": ign or t_code != "None", 
                                "use_std_cost": std_o, 
                                "labor_type": l_type, 
                                "test_code": t_code, 
                                "test_charge": test_dict.get(t_code, 0.0), 
                                "osp_code": o_code, 
                                "osp_charge": osp_dict.get(o_code, 0.0),
                                "apply_tariff": tar_inc
                            }

                            if live_calc_active:
                                preview_data = {"Metric": ["Demand Qty", "Total Unit Cost"]}
                                for q in fg_lots:
                                    cost_res = calculate_master(p_id, q, m_rate, a_rate, log_pct, parts_dict, b_df, routing_dict, pur_df, user_overrides[fg], cny, i_tar_rate, eff)
                                    req_qty = demand_by_fg[fg][p_id].get(q, 0)
                                    preview_data[f"Lot {q}"] = [f"{req_qty:,.1f}", f"${cost_res['total']:,.2f}"]
                                st.dataframe(pd.DataFrame(preview_data), hide_index=True, use_container_width=True)

                        if is_fg:
                            st.divider()
                            st.markdown("#### 📦 Bill of Materials (Sub-Components)")

    st.divider()

    # --- FINAL RESULTS TABS ---
    st.header("3. Output Analysis")
    t1, t2, t3 = st.tabs(["📊 Cost Matrix", "🚩 Deep Scan Audit", "🧪 Master Services Dictionary"])
    
    with t1:
        if st.button("🚀 Run Batch Calculation", use_container_width=True, type="primary"):
            final_res = []
            for fg, fg_lots in fg_targets.items():
                if fg not in parts_dict: continue
                for q in fg_lots:
                    res_dict = calculate_master(fg, q, m_rate, a_rate, log_pct, parts_dict, b_df, routing_dict, pur_df, user_overrides[fg], cny, i_tar_rate, eff)
                    
                    # --- NEW: DETAILED COST BUCKETS ---
                    final_res.append({
                        "Part": fg, 
                        "Qty": q, 
                        "Material_Cost": round(res_dict["material"], 2),
                        "Setup_&_Labor": round(res_dict["labor"], 2),
                        "Logistics_&_Tariffs": round(res_dict["logistics"] + res_dict["tariff"], 2),
                        "OSP_&_Tests": round(res_dict["adders"], 2),
                        "Total_Unit_Cost": round(res_dict["total"], 2)
                    })
            st.session_state['res_df'] = pd.DataFrame(final_res)

        if 'res_df' in st.session_state:
            res_df = st.session_state['res_df']
            
            # Format display table to look like currency
            format_dict = {col: "${:,.2f}" for col in res_df.columns if "Cost" in col or "Labor" in col or "Tariff" in col or "OSP" in col}
            st.dataframe(res_df.style.format(format_dict), use_container_width=True, hide_index=True)
            
            # Pivot table for Excel export, retaining all detailed buckets
            df_pivot = res_df.pivot(index='Part', columns='Qty', values=['Material_Cost', 'Setup_&_Labor', 'Logistics_&_Tariffs', 'OSP_&_Tests', 'Total_Unit_Cost'])
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_pivot.to_excel(writer, sheet_name='Cost_Matrix')
            st.download_button("📥 Download Excel (With Breakdowns)", output.getvalue(), "Detailed_Cost_Audit.xlsx", use_container_width=True)

    with t2:
        st.subheader("🚩 Missing Data Deep Scan")
        if st.button("🔍 Scan All BOMs"):
            gaps, checked = [], set()
            def scan(p_id, current_fg):
                check_key = f"{current_fg}_{p_id}"
                if check_key in checked: return
                checked.add(check_key)
                
                ov = user_overrides[current_fg].get(p_id, {})
                if ov.get("ignored") or ov.get("use_std_cost"): return
                
                p_data = parts_dict.get(p_id, {})
                if p_data.get("Procurement_Type") == "E":
                    if p_id not in routing_dict: gaps.append({"Finished_Good": current_fg, "Part": p_id, "Type": "E", "Missing": "Router"})
                    for c in b_df[b_df['Parent_ID'] == p_id]['Child_ID'].unique(): scan(c, current_fg)
                else:
                    if p_id not in pur_df['Part_ID'].values: gaps.append({"Finished_Good": current_fg, "Part": p_id, "Type": "F", "Missing": "Price"})

            for fg in fg_targets.keys():
                if fg in parts_dict: scan(fg, fg)
                
            if gaps: st.table(pd.DataFrame(gaps).drop_duplicates())
            else: st.success("✅ No Data Gaps Found Across Any Finished Good")

    with t3:
        st.subheader("🧪 Manage Tests and OSP Charges")
        st.info("Changes made here will instantly update the selection menus across all tabs.")
        col_t, col_o = st.columns(2)
        
        with col_t:
            st.markdown("**Test/Service Charges (Flat Lot Fee)**")
            edited_tests = st.data_editor(st.session_state.test_codes_df, num_rows="dynamic", hide_index=True, column_config={"Test_Code": st.column_config.TextColumn("Test Code", required=True), "Cost": st.column_config.NumberColumn("Fee ($)", min_value=0.0, format="$%.2f")}, use_container_width=True, key="editor_test")
            st.session_state.test_codes_df = edited_tests

        with col_o:
            st.markdown("**Outside Processing (Flat Lot Fee)**")
            edited_osp = st.data_editor(st.session_state.osp_codes_df, num_rows="dynamic", hide_index=True, column_config={"OSP_Code": st.column_config.TextColumn("OSP Code", required=True), "Cost": st.column_config.NumberColumn("Fee ($)", min_value=0.0, format="$%.2f")}, use_container_width=True, key="editor_osp")
            st.session_state.osp_codes_df = edited_osp