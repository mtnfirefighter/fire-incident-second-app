
import os
import re
from datetime import datetime
from typing import Dict, List
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident Reports (Clean App)", page_icon="🧪", layout="wide")

# === Files ===
DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
PRIMARY_KEY = "IncidentNumber"

# === Schemas ===
PERSONNEL_SCHEMA = ["PersonnelID","Rank","FirstName","LastName","Name"]
APPARATUS_SCHEMA = ["UnitNumber","CallSign","Name"]
CHILD_TABLES = {
    "Incident_Times": ["IncidentNumber","Alarm","Enroute","Arrival","Clear"],
    "Incident_Personnel": ["IncidentNumber","Name","Role","Hours","RespondedIn","Notes","PersonnelID"],
    "Incident_Apparatus": ["IncidentNumber","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame({c: pd.Series(dtype="object") for c in cols})
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df[cols + [c for c in df.columns if c not in cols]]

def load_or_init(path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(path):
        # initialize empty workbook with all sheets
        data = {
            "Personnel": ensure_columns(pd.DataFrame(), PERSONNEL_SCHEMA),
            "Apparatus": ensure_columns(pd.DataFrame(), APPARATUS_SCHEMA),
        }
        for k, cols in CHILD_TABLES.items():
            data[k] = ensure_columns(pd.DataFrame(), cols)
        save_to_path(data, path)
        return data
    # load
    xls = pd.ExcelFile(path)
    data = {}
    for name in ["Personnel","Apparatus"] + list(CHILD_TABLES.keys()):
        try:
            df = pd.read_excel(path, sheet_name=name)
        except Exception:
            df = pd.DataFrame()
        if name in ["Personnel","Apparatus"]:
            schema = PERSONNEL_SCHEMA if name=="Personnel" else APPARATUS_SCHEMA
            data[name] = ensure_columns(df, schema)
        else:
            data[name] = ensure_columns(df, CHILD_TABLES[name])
    return data

def save_to_path(data: Dict[str, pd.DataFrame], path: str):
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        for sheet, df in data.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

# === Label builders ===
def build_person_options(df: pd.DataFrame) -> list:
    """Clean labels for the member picker (Rank + First + Last, fall back to Name/FullName)."""
    if df is None or df.empty:
        return []
    cols = {c.lower(): c for c in df.columns}
    first = cols.get("firstname") or cols.get("first") or cols.get("first_name") or cols.get("given")
    last  = cols.get("lastname")  or cols.get("last")  or cols.get("last_name")  or cols.get("family")
    rank  = cols.get("rank") or cols.get("title")
    name  = cols.get("name") or cols.get("fullname") or cols.get("full_name") or cols.get("display")

    def clean(s):
        s = "" if s is None else str(s).strip()
        s = re.sub(r"\['([^]]+?)'\]", r"\1", s).replace("][", " ").replace("[", "").replace("]", "").replace("'", "")
        return re.sub(r"\s+", " ", s).strip()

    labels = []
    if first and last and first in df and last in df:
        f = df[first].map(clean)
        l = df[last].map(clean)
        if rank and rank in df:
            r = df[rank].map(clean)
            labels = (r + " " + f + " " + l).str.strip().replace("", pd.NA).dropna().tolist()
        else:
            labels = (f + " " + l).str.strip().replace("", pd.NA).dropna().tolist()
    if not labels and name and name in df:
        labels = df[name].map(clean).replace("", pd.NA).dropna().tolist()
    return sorted(dict.fromkeys(labels))

def build_unit_options(df: pd.DataFrame) -> list:
    """Labels like '12 – Engine 12 – Whitmer VFD' using any available columns."""
    if df is None or df.empty:
        return []
    cols = [c for c in ["UnitNumber","CallSign","Name","Unit","Label"] if c in df.columns]
    if not cols:
        return []
    def nz(x):
        s = "" if x is None else str(x).strip()
        return "" if s.lower() in ("nan","none") else s
    labels = []
    for _, row in df[cols].fillna("").iterrows():
        parts, seen = [], set()
        for c in cols:
            v = nz(row[c])
            if v and v not in seen:
                parts.append(v); seen.add(v)
        if parts:
            labels.append(" – ".join(parts))
    return sorted(dict.fromkeys([p for p in labels if p]))

# === UI ===
st.title("Fire Incident Reports — Clean Test App")
st.caption("This clean version isolates personnel + apparatus logic to verify saving Responded In.")

file_path = st.text_input("Excel path", value=DEFAULT_FILE, help="Uses/creates fire_incident_db.xlsx next to this app by default.")
data = load_or_init(file_path)

# Incident selector
inc_num = st.text_input("Incident Number", value="1", help="Enter the incident number to work on.")

with st.container(border=True):
    st.subheader("All Members on Scene")
    people_df = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    app_df_all = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    person_opts = build_person_options(people_df)
    unit_opts_all = build_unit_options(app_df_all)

    picked_people = st.multiselect("Pick members", options=person_opts, key="w_pick_people_auth")
    roles = ["OIC","Driver","Firefighter","Officer","EMT"]
    cc = st.columns(4)
    role_default = cc[0].selectbox("Default Role", options=roles, index=0 if roles else None, key="w_role_default_auth")
    hours_default = cc[1].number_input("Default Hours", value=0.0, min_value=0.0, step=0.5, key="w_hours_default_auth")
    responded_in_default = cc[2].selectbox("Responded In (optional)", options=[""]+unit_opts_all, index=0, key="w_resp_in_default_auth")

    if cc[3].button("Add Selected Members", key="w_add_people_btn_auth"):
        if not inc_num or str(inc_num).strip() == "":
            st.error("Enter **IncidentNumber** before adding members.")
        else:
            inc_key = str(inc_num).strip()
            df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
            resp_value = responded_in_default if isinstance(responded_in_default, str) and responded_in_default.strip() != "" else None
            new = []
            for n in picked_people:
                new.append({
                    PRIMARY_KEY: inc_key,
                    "Name": str(n).strip(),
                    "Role": role_default,
                    "Hours": hours_default,
                    "RespondedIn": resp_value,
                    "Notes": None,
                    "PersonnelID": None,
                })
            if new:
                data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                save_to_path(data, file_path)
                st.success(f"Added {len(new)} member(s) to incident {inc_key}. Responded In saved: {resp_value or '— none —'}")
            else:
                st.warning("No members selected.")

    # Show personnel for this incident
    cur_per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    this_per = cur_per[cur_per[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else "__none__")].copy()
    if not this_per.empty and "Delete" not in this_per.columns:
        this_per["Delete"] = False
    st.write(f"**Total Personnel on Scene:** {0 if this_per.empty else len(this_per)}")

    # Reorder columns
    display_cols = [PRIMARY_KEY, "Name", "RespondedIn", "Role", "Hours", "Notes", "PersonnelID"]
    for c in display_cols:
        if c not in this_per.columns:
            this_per[c] = pd.NA
    other_cols = [c for c in this_per.columns if c not in display_cols]
    this_per = this_per[display_cols + other_cols]

    this_per_edit = st.data_editor(this_per, num_rows="dynamic", use_container_width=True, key="editor_incident_personnel")
    cdel = st.columns(2)
    if cdel[0].button("Save Personnel Grid", key="btn_save_incident_personnel"):
        base = cur_per[cur_per[PRIMARY_KEY].astype(str) != (str(inc_num).strip() if inc_num else "__none__")]
        if "Delete" in this_per_edit.columns:
            this_per_edit = this_per_edit[this_per_edit["Delete"] != True].drop(columns=["Delete"], errors="ignore")
        data["Incident_Personnel"] = pd.concat([base, this_per_edit], ignore_index=True)
        save_to_path(data, file_path)
        st.success("Incident personnel updated (removals applied if any).")

with st.container(border=True):
    st.subheader("Apparatus on Scene")
    app_df = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    unit_opts = build_unit_options(app_df)
    picked_units = st.multiselect("Pick apparatus units", options=unit_opts, key="w_pick_units_auth")
    unit_type_options = ["Engine","Tanker","Rescue","Mini Pumper"]
    cc2 = st.columns(4)
    unit_type = cc2[0].selectbox("UnitType", options=[""]+unit_type_options, index=0, key="w_unit_type_auth")
    unit_role = cc2[1].selectbox("Role", options=["Primary","Support","Staging","Water Supply"], index=0, key="w_unit_role_auth")
    unit_actions = cc2[2].text_input("Actions (e.g., 'Directing traffic')", key="w_unit_actions_auth")

    if cc2[3].button("Add Units", key="w_add_units_btn_auth"):
        if not inc_num or str(inc_num).strip() == "":
            st.error("Enter **IncidentNumber** before adding units.")
        else:
            inc_key = str(inc_num).strip()
            df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
            new = [{
                PRIMARY_KEY: inc_key,
                "Unit": u,
                "UnitType": (unit_type if unit_type else None),
                "Role": unit_role,
                "Actions": unit_actions or ""
            } for u in picked_units]
            if new:
                data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                save_to_path(data, file_path)
                st.success(f"Added {len(new)} unit(s) to incident {inc_key}.")
            else:
                st.warning("No units selected.")

    # Show apparatus for this incident
    cur_app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    this_app = cur_app[cur_app[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else "__none__")].copy()
    if not this_app.empty and "Delete" not in this_app.columns:
        this_app["Delete"] = False
    st.write(f"**Total Apparatus on Scene:** {0 if this_app.empty else len(this_app)}")
    this_app_edit = st.data_editor(this_app, num_rows="dynamic", use_container_width=True, key="editor_incident_apparatus")
    cdel2 = st.columns(2)
    if cdel2[0].button("Save Apparatus Grid", key="btn_save_incident_apparatus"):
        base = cur_app[cur_app[PRIMARY_KEY].astype(str) != (str(inc_num).strip() if inc_num else "__none__")]
        if "Delete" in this_app_edit.columns:
            this_app_edit = this_app_edit[this_app_edit["Delete"] != True].drop(columns=["Delete"], errors="ignore")
        data["Incident_Apparatus"] = pd.concat([base, this_app_edit], ignore_index=True)
        save_to_path(data, file_path)
        st.success("Incident apparatus updated (removals applied if any).")

with st.container(border=True):
    st.subheader("Diagnostics")
    st.write(f"**Excel path:** {file_path}  |  Exists: {'✅' if os.path.exists(file_path) else '❌'}")
    try:
        xls = pd.ExcelFile(file_path); st.write("**Sheets:**", xls.sheet_names)
    except Exception as e:
        st.error(f"Open failed: {e}")
    st.write("**Personnel Top 10:**")
    st.dataframe(data['Personnel'].head(10), use_container_width=True)
    st.write("**Apparatus Top 10:**")
    st.dataframe(data['Apparatus'].head(10), use_container_width=True)
