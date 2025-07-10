#!/usr/bin/env python
"""Subregional emission factor analysis for Ontario.

This script downloads generator, zonal demand and trade flow data from IESO
and computes hourly supply- and demand-based emission factors for ten
Ontario subregions. The approach mirrors the Ontario-wide calculations but
includes an additional linear programming step to allocate flows between
subregions.
"""

import os
import pandas as pd
import numpy as np
import requests
from scipy.optimize import linprog

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def download_file(url: str, save_path: str) -> str:
    """Download a file if it does not exist locally."""
    if os.path.exists(save_path):
        print(f"File already exists: {save_path}")
        return save_path

    response = requests.get(url, timeout=30)
    response.raise_for_status()
    with open(save_path, "wb") as fh:
        fh.write(response.content)
    print(f"Downloaded: {save_path}")
    return save_path


def parse_and_clean_generator_month(file_path: str) -> pd.DataFrame:
    cleaned = []
    with open(file_path, "r") as fh:
        lines = fh.readlines()[3:]  # skip headers
        for line in lines:
            line = line.strip().rstrip(",")
            fields = line.split(",")
            if len(fields) == 28:
                cleaned.append(fields)
    cols = [
        "Delivery Date",
        "Generator",
        "Fuel Type",
        "Measurement",
    ] + [f"Hour {i}" for i in range(1, 25)]
    return pd.DataFrame(cleaned, columns=cols)


def aggregate_generator_data(year: int) -> pd.DataFrame:
    base_url = "https://reports-public.ieso.ca/public/GenOutputCapabilityMonth/"
    data_dir = os.path.join("data", "IESO", str(year), "Generator")
    os.makedirs(data_dir, exist_ok=True)

    frames = []
    for month in range(1, 13):
        month_str = f"{month:02d}"
        name = f"PUB_GenOutputCapabilityMonth_{year}{month_str}.csv"
        url = f"{base_url}{name}"
        path = os.path.join(data_dir, name)
        download_file(url, path)
        df = parse_and_clean_generator_month(path)
        header = [
            "Delivery Date",
            "Generator",
            "Fuel Type",
            "Measurement",
        ] + [f"Hour {i}" for i in range(1, 25)]
        df = df[~(df == header).all(axis=1)]
        if frames:
            df = df[1:]
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def transform_generator_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df[df["Measurement"] == "Output"].drop(columns=["Measurement"])
    df["Fuel-Generator"] = df["Fuel Type"] + " - " + df["Generator"]
    df = df.drop(columns=["Generator", "Fuel Type"])
    melted = df.melt(
        id_vars=["Delivery Date", "Fuel-Generator"],
        var_name="Hour",
        value_name="Value",
    )
    melted["Hour"] = melted["Hour"].str.extract(r"(\d+)").astype(int)
    pivot = (
        melted.pivot(index=["Delivery Date", "Hour"], columns="Fuel-Generator", values="Value")
        .reset_index()
    )
    pivot.columns.name = None
    return pivot


def parse_and_clean_demand(file_path: str) -> pd.DataFrame:
    with open(file_path, "r") as fh:
        lines = fh.readlines()
    skip_rows = [i for i, line in enumerate(lines) if "\\" in line]
    return pd.read_csv(file_path, skiprows=skip_rows)


def parse_and_clean_trade_flow(file_path: str) -> pd.DataFrame:
    with open(file_path, "r") as fh:
        lines = fh.readlines()
    skip_rows = [i for i, line in enumerate(lines) if "\\" in line]
    raw_df = pd.read_csv(file_path, skiprows=skip_rows, header=None)
    raw_headers = raw_df.iloc[:2].fillna("")
    headers = [f"{str(raw_headers.iloc[0, i]).strip()} {str(raw_headers.iloc[1, i]).strip()}" for i in range(len(raw_df.columns))]
    headers[0] = "Date"
    headers[1] = "Hour"
    raw_df.columns = headers
    raw_df = raw_df[2:]
    return raw_df


def transform_trade_flow(trade_df: pd.DataFrame) -> pd.DataFrame:
    result = trade_df[["Date", "Hour"]].copy()
    flow_cols = [c for c in trade_df.columns if c.endswith("Flow") and not c.startswith("Total")]
    filtered = trade_df[flow_cols].copy()
    man_cols = [c for c in filtered.columns if c.startswith("MANITOBA")]
    result["MANITOBA Total Flow"] = filtered[man_cols].astype(float).sum(axis=1)
    filtered = filtered.drop(columns=man_cols, errors="ignore")
    qc_cols = [c for c in filtered.columns if c.startswith("PQ")]
    result["QUEBEC Total Flow"] = filtered[qc_cols].astype(float).sum(axis=1)
    filtered = filtered.drop(columns=qc_cols, errors="ignore")
    result = pd.concat([result, filtered], axis=1)
    return result


def get_emission_rates() -> dict:
    path = os.path.join("data", "emission_rates.csv")
    df = pd.read_csv(path)
    return {row["Technology"]: row["Emission Rate (t CO2e/GWh)"] / 1000 for _, row in df.iterrows()}


def get_neighboring_emission_factors() -> dict:
    path = os.path.join("data", "neighboring_emission_factors.csv")
    df = pd.read_csv(path)
    return {row["Region"].upper()[:3]: row["Emission Factor (t CO2e/GWh)"] / 1000 for _, row in df.iterrows()}


def load_generator_list() -> pd.DataFrame:
    path = os.path.join("data", "Generator_List.csv")
    return pd.read_csv(path)


def setup_year_data(year: int):
    year_dir = os.path.join("data", "IESO", str(year))
    demand_dir = os.path.join(year_dir, "Demand")
    trade_dir = os.path.join(year_dir, "Trade")
    zonal_dir = os.path.join(year_dir, "DemandZonal")
    os.makedirs(demand_dir, exist_ok=True)
    os.makedirs(trade_dir, exist_ok=True)
    os.makedirs(zonal_dir, exist_ok=True)

    gen_df = aggregate_generator_data(year)

    demand_url = f"https://reports-public.ieso.ca/public/Demand/PUB_Demand_{year}.csv"
    demand_path = os.path.join(demand_dir, f"PUB_Demand_{year}.csv")
    download_file(demand_url, demand_path)
    demand_df = parse_and_clean_demand(demand_path)

    zonal_url = f"https://reports-public.ieso.ca/public/DemandZonal/PUB_DemandZonal_{year}.csv"
    zonal_path = os.path.join(zonal_dir, f"PUB_DemandZonal_{year}.csv")
    download_file(zonal_url, zonal_path)
    zonal_df = parse_and_clean_demand(zonal_path)

    trade_url = f"https://reports-public.ieso.ca/public/IntertieScheduleFlowYear/PUB_IntertieScheduleFlowYear_{year}.csv"
    trade_path = os.path.join(trade_dir, f"PUB_IntertieScheduleFlowYear_{year}.csv")
    download_file(trade_url, trade_path)
    trade_df = parse_and_clean_trade_flow(trade_path)
    trade_df = transform_trade_flow(trade_df)

    return gen_df, demand_df, zonal_df, trade_df

# ---------------------------------------------------------------------------
# Core calculations
# ---------------------------------------------------------------------------

REGIONS = [
    "Northwest",
    "Northeast",
    "Ottawa",
    "East",
    "Toronto",
    "Essa",
    "Bruce",
    "Southwest",
    "Niagara",
    "West",
]
TECHS = ["Biofuel", "Hydro", "Natural Gas", "Nuclear", "Solar", "Wind"]


def compute_generation_by_region(gen_data: pd.DataFrame, gen_list: pd.DataFrame, emission_rates: dict):
    region_idx = {r: i for i, r in enumerate(REGIONS)}
    tech_idx = {t: i for i, t in enumerate(TECHS)}

    mapping = {}
    for _, row in gen_list.iterrows():
        mapping[row["Generator List"].strip().upper()] = (
            region_idx[row["Region"].strip()],
            tech_idx[row["Technology"].replace("Gas", "Natural Gas").strip()],
        )

    gen_cols = [c for c in gen_data.columns if c not in ["Delivery Date", "Hour"]]
    n_steps = len(gen_data)
    outputs = np.zeros((n_steps, len(REGIONS)))
    ef = np.zeros((n_steps, len(REGIONS)))

    for col in gen_cols:
        gen_name = col.split(" - ", 1)[1].strip().upper()
        if gen_name not in mapping:
            continue
        r_idx, t_idx = mapping[gen_name]
        rate = emission_rates[TECHS[t_idx]]
        values = pd.to_numeric(gen_data[col], errors="coerce").fillna(0).to_numpy()
        outputs[:, r_idx] += values
        ef[:, r_idx] += values * rate

    ef = np.divide(ef, outputs, out=np.zeros_like(ef), where=outputs != 0)
    return outputs, ef


def calculate_supply_based_ef(outputs: np.ndarray, ef: np.ndarray):
    ont_output = outputs.sum(axis=1)
    ont_ef = np.divide((ef * outputs).sum(axis=1), ont_output, out=np.zeros_like(ont_output), where=ont_output != 0)
    return ont_ef


def calculate_new_ontario_ef(ont_ef, demand_df, trade_df, neighboring_factors):
    ef_trade = np.array([neighboring_factors[r[:3].upper()] for r in ["Manitoba", "Michigan", "Minnesota", "New York", "Quebec"]])
    n_steps = len(ont_ef)
    new_ef = np.zeros(n_steps)

    manitoba = trade_df["MANITOBA Total Flow"].astype(float).to_numpy()
    michigan = trade_df[[c for c in trade_df.columns if c.startswith("MICHIGAN")][0]].astype(float).to_numpy()
    minnesota = trade_df[[c for c in trade_df.columns if c.startswith("MINNESOTA")][0]].astype(float).to_numpy()
    newyork = trade_df[[c for c in trade_df.columns if c.startswith("NEW YORK")][0]].astype(float).to_numpy()
    quebec = trade_df["QUEBEC Total Flow"].astype(float).to_numpy()

    ont_demand = demand_df["Ontario Demand"].astype(float).to_numpy()

    trades = np.vstack([manitoba, michigan, minnesota, newyork, quebec])

    for t in range(n_steps):
        trade = np.where(trades[:, t] < 0, -trades[:, t], 0)
        self_supplied = ont_demand[t] - trade.sum()
        total = (ef_trade * trade).sum() + ont_ef[t] * self_supplied
        new_ef[t] = total / ont_demand[t] if ont_demand[t] != 0 else 0
    return new_ef


def build_lp_matrices():
    Aeq = np.zeros((18, 50))
    Aineq = np.zeros((12, 50))
    bineq = np.array([325,350,2100,2900,2000,1500,2000,7500,5000,3000,1990,1800], dtype=float)

    Aineq[0,0] = 1
    Aineq[1,3] = 1
    Aineq[2,4] = 1
    Aineq[3,7] = 1
    Aineq[4,11] = 1
    Aineq[5,12] = 1
    Aineq[6,13] = 1
    Aineq[7,14] = 1
    Aineq[8,15:17] = 1
    Aineq[9,17] = 1
    Aineq[10,18] = 1
    Aineq[11,20] = 1

    cost = np.ones(50)
    cost[30:50] = 1000
    cost[[36,46]] = 10
    lb = np.zeros(50)

    # equality constraint definitions
    Aeq[0, [0,1,2]] = 1
    Aeq[0, [3,22,24]] = -1

    Aeq[1, [3,4,5]] = 1
    Aeq[1, [0,12,27]] = -1

    Aeq[2, [6]] = 1
    Aeq[2, [7,28]] = -1

    Aeq[3, 7:10] = 1
    Aeq[3, [10,25,29]] = -1

    Aeq[4, 10:12] = 1
    Aeq[4, [13,15]] = -1

    Aeq[5, 12:14] = 1
    Aeq[5, [4,11,16]] = -1

    Aeq[6, 14] = 1

    Aeq[7, 15:18] = 1
    Aeq[7, [14,18,20]] = -1

    Aeq[8, 18:20] = 1
    Aeq[8, 26] = -1

    Aeq[9, 20:22] = 1
    Aeq[9, [17,23]] = -1

    Aeq[10,1] = 1
    Aeq[10,22] = -1

    Aeq[11,21] = 1
    Aeq[11,23] = -1

    Aeq[12,2] = 1
    Aeq[12,24] = -1

    Aeq[13,[8,19]] = 1
    Aeq[13,25:27] = -1

    Aeq[14,5] = 1
    Aeq[14,27] = -1

    Aeq[15,6] = 1
    Aeq[15,28] = -1

    Aeq[16,9] = 1
    Aeq[16,29] = -1

    Aeq[17,30:40] = 1
    Aeq[17,40:50] = -1

    for i in range(10):
        Aeq[i, i+30] = 1
        Aeq[i, i+40] = -1

    return Aeq, Aineq, bineq, cost, lb


def subregion_lp(gen_output, zonal_demand, trade_df):
    n_steps = gen_output.shape[0]
    results = []
    Aeq, Aineq, bineq, cost, lb = build_lp_matrices()

    manitoba = trade_df["MANITOBA Total Flow"].astype(float).to_numpy()
    michigan = trade_df[[c for c in trade_df.columns if c.startswith("MICHIGAN")][0]].astype(float).to_numpy()
    minnesota = trade_df[[c for c in trade_df.columns if c.startswith("MINNESOTA")][0]].astype(float).to_numpy()
    newyork = trade_df[[c for c in trade_df.columns if c.startswith("NEW YORK")][0]].astype(float).to_numpy()
    qcne = trade_df[[c for c in trade_df.columns if "QUEBEC" in c and "NE" in c]][0].astype(float).to_numpy() if any("NE" in c for c in trade_df.columns if "QUEBEC" in c) else np.zeros(n_steps)
    qcott = trade_df[[c for c in trade_df.columns if "QUEBEC" in c and "OTT" in c]][0].astype(float).to_numpy() if any("OTT" in c for c in trade_df.columns if "QUEBEC" in c) else np.zeros(n_steps)
    qce = trade_df["QUEBEC Total Flow"].astype(float).to_numpy()

    for t in range(n_steps):
        beq = np.zeros(18)
        for r in range(10):
            beq[r] = gen_output[t, r] - zonal_demand.iloc[t, r+2]
        beq[10] = manitoba[t]
        beq[11] = michigan[t]
        beq[12] = minnesota[t]
        beq[13] = newyork[t]
        beq[14] = qcne[t]
        beq[15] = qcott[t]
        beq[16] = qce[t]
        beq[17] = beq[:10].sum() - beq[10:17].sum()

        res = linprog(cost, A_ub=Aineq, b_ub=bineq, A_eq=Aeq, b_eq=beq, bounds=list(zip(lb, [None]*50)), method="highs")
        if res.success:
            results.append(res.x)
        else:
            results.append(np.zeros(50))
    return np.array(results)


def compute_subregion_ef(gen_output, ef, zonal_demand, lp_results, neighboring_factors):
    n_steps = gen_output.shape[0]
    new_ef = np.zeros((n_steps, 10))
    ef_trade = np.array([neighboring_factors[r[:3].upper()] for r in ["Manitoba", "Michigan", "Minnesota", "New York", "Quebec"]])

    for t in range(n_steps):
        x = lp_results[t]
        P = np.zeros(15)
        P[:10] = gen_output[t]
        P[10] = x[22]
        P[11] = x[23]
        P[12] = x[24]
        P[13] = x[25] + x[26]
        P[14] = x[27] + x[28] + x[29]

        D = np.zeros(15)
        D[:10] = zonal_demand.iloc[t, 2:12]
        D[10] = x[1]
        D[11] = x[21]
        D[12] = x[2]
        D[13] = x[8] + x[19]
        D[14] = x[5] + x[6] + x[9]

        T = np.zeros((15,15))
        T[0,1] = x[0]
        T[0,10] = x[1]
        T[0,12] = x[2]
        T[1,0] = x[3]
        T[1,5] = x[4]
        T[1,14] = x[5]
        T[2,14] = x[6]
        T[3,2] = x[7]
        T[3,13] = x[8]
        T[3,14] = x[9]
        T[4,3] = x[10]
        T[4,5] = x[11]
        T[5,1] = x[12]
        T[5,4] = x[13]
        T[6,7] = x[14]
        T[7,4] = x[15]
        T[7,5] = x[16]
        T[7,9] = x[17]
        T[8,7] = x[18]
        T[8,13] = x[19]
        T[9,7] = x[20]
        T[9,11] = x[21]
        T[10,0] = x[22]
        T[11,9] = x[23]
        T[12,0] = x[24]
        T[13,3] = x[25]
        T[13,8] = x[26]
        T[14,1] = x[27]
        T[14,2] = x[28]
        T[14,3] = x[29]

        F = np.zeros(15)
        F[:10] = ef[t]
        F[10:] = ef_trade

        X = np.zeros(15)
        for i in range(10):
            X[i] = D[i] + T[i].sum()
        X[10] = x[1] + x[22]
        X[11] = x[21] + x[23]
        X[12] = x[2] + x[24]
        X[13] = x[8] + x[19] + x[25] + x[26]
        X[14] = x[5] + x[6] + x[9] + x[27] + x[28] + x[29]

        invX = np.linalg.pinv(np.diag(X))
        B = invX @ T
        G = np.linalg.inv(np.eye(15) - B)
        H = G @ np.diag(D) @ invX
        Eg = F * P
        Ec = np.diag(Eg) @ H
        Ec2 = np.ones((1,15)) @ Ec
        invD = np.linalg.pinv(np.diag(D))
        new_ef[t] = (Ec2 @ invD)[:,:10]
    return new_ef


def main():
    year = int(input("Enter the year for analysis (valid years are 2020-2024): ") or 2023)
    emission_rates = get_emission_rates()
    neighbor_factors = get_neighboring_emission_factors()

    gen_df, demand_df, zonal_df, trade_df = setup_year_data(year)
    gen_transformed = transform_generator_data(gen_df).fillna(0)
    gen_list = load_generator_list()

    outputs, ef = compute_generation_by_region(gen_transformed, gen_list, emission_rates)
    ont_ef = calculate_supply_based_ef(outputs, ef)
    new_ont_ef = calculate_new_ontario_ef(ont_ef, demand_df, trade_df, neighbor_factors)

    lp_results = subregion_lp(outputs, zonal_df, trade_df)
    subregion_ef = compute_subregion_ef(outputs, ef, zonal_df, lp_results, neighbor_factors)

    supply_df = pd.DataFrame(ef, columns=REGIONS)
    supply_df.insert(0, "Hour", gen_transformed["Hour"])
    supply_df.insert(0, "Delivery Date", gen_transformed["Delivery Date"])
    supply_df.insert(1, "Ontario", ont_ef * 1000)

    demand_df_out = pd.DataFrame(subregion_ef, columns=REGIONS)
    demand_df_out.insert(0, "Hour", gen_transformed["Hour"])
    demand_df_out.insert(0, "Delivery Date", gen_transformed["Delivery Date"])
    demand_df_out.insert(1, "Ontario", new_ont_ef * 1000)

    out_dir = os.path.join("data", "output")
    os.makedirs(out_dir, exist_ok=True)
    supply_path = os.path.join(out_dir, f"SubOntario_Supply_EF_{year}.csv")
    demand_path = os.path.join(out_dir, f"SubOntario_Demand_EF_{year}.csv")
    supply_df.to_csv(supply_path, index=False)
    demand_df_out.to_csv(demand_path, index=False)
    print(f"Supply-based EF saved to: {supply_path}")
    print(f"Demand-based EF saved to: {demand_path}")


if __name__ == "__main__":
    main()


