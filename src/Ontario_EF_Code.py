#!/usr/bin/env python
# coding: utf-8



## 1. Import Necessary Libraries
import pandas as pd
import numpy as np
import os
import requests
import time



## 2. Helper Functions
### 2.1 Fetch data from URLs
def download_file(url, save_path):
    """Download ``url`` to ``save_path`` if it does not already exist."""
    if os.path.exists(save_path):
        print(f"File already exists: {save_path}")
        return save_path

    response = requests.get(url, timeout=30)
    response.raise_for_status()

    with open(save_path, "wb") as f:
        f.write(response.content)
    print(f"Downloaded: {save_path}")

    return save_path

### 2.2 Parse and Clean Generator Data
def parse_and_clean_generator_month(file_path):
    # Read the file line-by-line to preprocess and fix inconsistencies
    cleaned_rows = []
    with open(file_path, 'r') as file:
        lines = file.readlines()
        for line in lines[3:]:  # Skip the first three rows (headers)
            line = line.strip().rstrip(',')  # Remove trailing spaces and commas
            fields = line.split(',')  # Split the line into fields
            if len(fields) == 28:  # Only keep rows with the correct number of fields
                cleaned_rows.append(fields)
            else:
                print(f"Skipped row with {len(fields)} fields: {line}")

    # Convert the cleaned rows into a DataFrame
    raw_df = pd.DataFrame(cleaned_rows, columns=[
        'Delivery Date', 'Generator', 'Fuel Type', 'Measurement',
        'Hour 1', 'Hour 2', 'Hour 3', 'Hour 4', 'Hour 5', 'Hour 6',
        'Hour 7', 'Hour 8', 'Hour 9', 'Hour 10', 'Hour 11', 'Hour 12',
        'Hour 13', 'Hour 14', 'Hour 15', 'Hour 16', 'Hour 17', 'Hour 18',
        'Hour 19', 'Hour 20', 'Hour 21', 'Hour 22', 'Hour 23', 'Hour 24'
    ])

    return raw_df

### 2.3 Aggregate Generator Data for a Year
def aggregate_generator_data(year):
    base_url = "https://reports-public.ieso.ca/public/GenOutputCapabilityMonth/"
    data_dir = os.path.join("data", "IESO", str(year), "Generator")
    os.makedirs(data_dir, exist_ok=True)

    # Prepare list for monthly data
    monthly_data = []

    for month in range(1, 13):
        month_str = f"{month:02d}"
        file_name = f"PUB_GenOutputCapabilityMonth_{year}{month_str}.csv"
        file_url = f"{base_url}{file_name}"
        local_path = os.path.join(data_dir, file_name)

        # Download the file
        download_file(file_url, local_path)

        # Parse and clean the monthly data
        month_data = parse_and_clean_generator_month(local_path)

        # Remove any rows matching the header row
        header_row = [
            'Delivery Date', 'Generator', 'Fuel Type', 'Measurement',
            'Hour 1', 'Hour 2', 'Hour 3', 'Hour 4', 'Hour 5', 'Hour 6',
            'Hour 7', 'Hour 8', 'Hour 9', 'Hour 10', 'Hour 11', 'Hour 12',
            'Hour 13', 'Hour 14', 'Hour 15', 'Hour 16', 'Hour 17', 'Hour 18',
            'Hour 19', 'Hour 20', 'Hour 21', 'Hour 22', 'Hour 23', 'Hour 24'
        ]
        month_data = month_data[~(month_data == header_row).all(axis=1)]

        # Remove the header row from subsequent monthly files
        if monthly_data:
            month_data = month_data[1:]

        monthly_data.append(month_data)

    # Concatenate all monthly data into a single yearly DataFrame
    yearly_data = pd.concat(monthly_data, ignore_index=True)

    return yearly_data

### 2.4 Transform Generator Data to Match Demand/Trade Flow Format
def transform_generator_data(gen_output_df):
    # Step 1: Keep only rows where Measurement is "Output"
    gen_output_df = gen_output_df[gen_output_df['Measurement'] == 'Output']

    # Check if there are any rows left after filtering
    if gen_output_df.empty:
        raise ValueError("No rows with 'Output' in the 'Measurement' column.")

    # Step 2: Drop the Measurement column
    gen_output_df = gen_output_df.drop(columns=['Measurement'])

    # Step 3: Combine Generator and Fuel Type into a single column
    gen_output_df['Fuel-Generator'] = gen_output_df['Fuel Type'] + ' - ' + gen_output_df['Generator']
    gen_output_df = gen_output_df.drop(columns=['Generator', 'Fuel Type'])

    ## Debug: Confirm creation of 'Fuel-Generator'
    #print("\nAfter Creating 'Fuel-Generator':")
    #print(gen_output_df[['Fuel-Generator']].head())

    # Step 4: Reshape the data to have Hours as rows and Fuel-Generators as columns
    melted = gen_output_df.melt(
        id_vars=['Delivery Date', 'Fuel-Generator'],  # Include 'Fuel-Generator' here
        var_name='Hour',
        value_name='Value'
    )

    ## Debug: Verify the melted DataFrame
    #print("\nAfter Melting Data:")
    #print(melted.head())

    # Extract numeric hour values, handling NaN values
    melted['Hour'] = melted['Hour'].str.extract(r'(\d+)')
    melted = melted.dropna(subset=['Hour'])  # Drop rows where Hour extraction failed
    melted['Hour'] = melted['Hour'].astype(int)

    # Pivot the data to have one column per Fuel-Generator
    reshaped_data = melted.pivot(
        index=['Delivery Date', 'Hour'],
        columns='Fuel-Generator',
        values='Value'
    ).reset_index()

    # Ensure columns are well-aligned
    reshaped_data.columns.name = None  # Remove the multi-index column name

    ## Debug: Final reshaped DataFrame
    #print("\nFinal Reshaped Data:")
    #print(reshaped_data.head())

    return reshaped_data


### 2.5 Load Emission Rates from File
def get_emission_rates():
    """Load generator emission rates from ``data/emission_rates.csv``."""
    path = os.path.join("data", "emission_rates.csv")
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Required emission rate file not found at {path}. "
            "Create this CSV with columns 'Technology' and 'Emission Rate (t CO2e/GWh)'."
        )

    df = pd.read_csv(path)
    expected_cols = {"Technology", "Emission Rate (t CO2e/GWh)"}
    if not expected_cols.issubset(df.columns):
        raise ValueError(
            f"Emission rate file must contain columns {expected_cols}."
        )

    print("\nLoaded emission rates from:", path)
    print(df.to_string(index=False))
    input("\nIf you wish to change these values, edit the file above and then press Enter to continue...")

    # Convert to t CO2e/MWh for calculations
    return {row['Technology']: row['Emission Rate (t CO2e/GWh)'] / 1000 for _, row in df.iterrows()}

### 2.6 Load Emission Factors of Neighboring Regions from File
def get_neighboring_emission_factors():
    """Load neighboring region emission factors from ``data/neighboring_emission_factors.csv``."""
    path = os.path.join("data", "neighboring_emission_factors.csv")
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Required neighboring emission factors file not found at {path}. "
            "Create this CSV with columns 'Region' and 'Emission Factor (t CO2e/GWh)'."
        )

    df = pd.read_csv(path)
    expected_cols = {"Region", "Emission Factor (t CO2e/GWh)"}
    if not expected_cols.issubset(df.columns):
        raise ValueError(
            f"Neighboring emission factor file must contain columns {expected_cols}."
        )

    print("\nLoaded emission factors for neighboring regions from:", path)
    print(df.to_string(index=False))
    input("\nIf you wish to change these values, edit the file above and then press Enter to continue...")

    # Convert to t CO2e/MWh for calculations
    return {row['Region']: row['Emission Factor (t CO2e/GWh)'] / 1000 for _, row in df.iterrows()}

### 2.7 Parse and Clean Demand Data
def parse_and_clean_demand(file_path):
    # Read the file once to find rows to skip
    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Identify rows to skip based on the presence of double backslashes
    skip_rows = [i for i, line in enumerate(lines) if '\\' in line]

    # Read CSV, skipping identified rows
    raw_df = pd.read_csv(file_path, skiprows=skip_rows)

    # Reset column headers if necessary
    raw_df.columns = [str(col).strip() for col in raw_df.columns]

    return raw_df

### 2.8 Parse and Clean Trade Flow Data
def parse_and_clean_trade_flow(file_path):
    # Read the file once to find rows to skip
    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Identify rows to skip based on the presence of backslashes
    skip_rows = [i for i, line in enumerate(lines) if '\\' in line]

    # Read CSV, skipping identified rows
    raw_df = pd.read_csv(file_path, skiprows=skip_rows, header=None)

    # Combine first two rows to form column names
    raw_headers = raw_df.iloc[:2].fillna('')
    headers = [f"{str(raw_headers.iloc[0, i]).strip()} {str(raw_headers.iloc[1, i]).strip()}" for i in range(len(raw_df.columns))]
    headers[0] = "Date"  # Rename the first column to Date
    headers[1] = "Hour"  # Rename the second column to Hour

    # Ensure the number of headers matches the number of columns
    if len(headers) != len(raw_df.columns):
        raise ValueError("Header length mismatch with columns in the data.")

    raw_df.columns = headers
    raw_df = raw_df[2:]  # Drop the first two rows used for headers

    return raw_df

### 2.9 Transform Trade Flow Data
def transform_trade_flow(trade_flow_df):
    """
    Transforms the trade flow DataFrame by keeping relevant columns:
    - Retains "Date" and "Hour" columns.
    - Removes columns starting with "Total".
    - Removes columns ending with "Imp" or "Exp".
    - Keeps only columns ending with "Flow".
    - Sums columns starting with "MANITOBA" into "MANITOBA Total Flow".
    - Sums columns starting with "PQ" into "QUEBEC Total Flow".

    Args:
        trade_flow_df (pd.DataFrame): The original trade flow DataFrame.

    Returns:
        pd.DataFrame: The transformed trade flow DataFrame.
    """
    # Step 1: Keep only "Date" and "Hour" columns
    transformed_df = trade_flow_df[["Date", "Hour"]].copy()

    # Step 2: Filter columns ending with "Flow", excluding those starting with "Total"
    flow_columns = [
        col for col in trade_flow_df.columns
        if col.endswith("Flow") and not col.startswith("Total")
    ]

    # Append the filtered flow columns to the result
    filtered_df = trade_flow_df[flow_columns].copy()

    # Step 3: Sum values of columns starting with "MANITOBA" into "MANITOBA Total Flow"
    manitoba_columns = [col for col in filtered_df.columns if col.startswith("MANITOBA")]
    transformed_df.loc[:, "MANITOBA Total Flow"] = filtered_df[manitoba_columns].astype(float).sum(axis=1)

    # Remove the added up MANITOBA columns
    filtered_df = filtered_df.drop(columns=manitoba_columns, errors="ignore")

    # Step 4: Sum values of columns starting with "PQ" into "QUEBEC Total Flow"
    quebec_columns = [col for col in filtered_df.columns if col.startswith("PQ")]
    transformed_df.loc[:, "QUEBEC Total Flow"] = filtered_df[quebec_columns].astype(float).sum(axis=1)

    # Remove the added up PQ columns
    filtered_df = filtered_df.drop(columns=quebec_columns, errors="ignore")

    # Append the remaining flow columns
    transformed_df = pd.concat([transformed_df, filtered_df], axis=1)

    return transformed_df

### 2.10 Calculate supply-based emission factors for each timestep
def calculate_supply_based_ef(transformed_gen_data, emission_rates):

    # Initialize a list for the output data
    ef_rows = []
    total_output_list = []


    # Group columns by technology prefix
    tech_prefix_map = {
        "B": "Biofuel",
        "H": "Hydro",
        "G": "Natural Gas",
        "N": "Nuclear",
        "S": "Solar",
        "W": "Wind"
    }

    # Ensure all cells are properly cleaned and non-numeric entries are handled
    transformed_gen_data = transformed_gen_data.apply(lambda col: pd.to_numeric(col, errors='coerce').fillna(0) if col.name not in ["Delivery Date", "Hour"] else col)

    # Loop through each row to calculate EF
    for _, row in transformed_gen_data.iterrows():
        # Extract date and hour
        delivery_date = row["Delivery Date"]
        hour = row["Hour"]

        # Sum outputs by technology
        tech_sums = {tech: 0 for tech in tech_prefix_map.values()}
        for col in transformed_gen_data.columns[2:]:
            prefix = col[0]
            if prefix in tech_prefix_map:
                tech = tech_prefix_map[prefix]
                value = row[col]
                tech_sums[tech] += value

        # Calculate total emissions and total output
        total_emissions = sum(tech_sums[tech] * emission_rates[tech] for tech in tech_sums)
        total_output = sum(tech_sums.values())

         # Store total output for this timestep
        total_output_list.append({
            "Delivery Date": delivery_date,
            "Hour": hour,
            "Total Output": total_output
        })

        # Calculate EF in t CO2e/MWh
        ef_t_co2e_mwh = total_emissions / total_output if total_output > 0 else 0

        # Convert EF to g CO2e/kWh for publishing
        ef_g_co2e_kwh = ef_t_co2e_mwh * 1000

        # Append results to the list
        ef_rows.append({
            "Delivery Date": delivery_date,
            "Hour": hour,
            "Supply-based EF (g CO2e/kWh)": ef_g_co2e_kwh
        })

    # Convert the results to a DataFrame
    supplybased_ef = pd.DataFrame(ef_rows)
    total_output_df = pd.DataFrame(total_output_list)


    return supplybased_ef, total_output_df

### 2.11 Calculate consumption-based emission factors for each timestep.

def calculate_consumption_based_ef(supplybased_ef, demand_df, transformed_trade_flow, neighboring_emission_factors, total_output_df):
    """
    Calculate the consumption-based emission factors for each timestep.

    Args:
        supplybased_ef (pd.DataFrame): Supply-based emission factors.
        demand_df (pd.DataFrame): Demand data.
        transformed_trade_flow (pd.DataFrame): Preprocessed trade flow data.
        neighboring_emission_factors (dict): Emission factors for neighboring regions.
        total_output_df (pd.DataFrame): Total output data.

    Returns:
        pd.DataFrame: Consumption-based emission factors.
        pd.DataFrame: Spot-check data for debugging.
    """
    # Check if all inputs have the same length
    if not (len(supplybased_ef) == len(demand_df) == len(transformed_trade_flow)):
        raise ValueError("Input DataFrames (supplybased_ef, demand_df, transformed_trade_flow) must have the same length.")

    # Normalize neighboring_emission_factors keys
    neighboring_emission_factors = {key.upper()[:3]: value for key, value in neighboring_emission_factors.items()}

    # Initialize lists for results and debugging
    consumption_rows = []
    spot_check_data = []

    total_steps = len(supplybased_ef)
    progress_interval = max(1, total_steps // 10)
    start_time = time.time()

    # Iterate over each timestep
    for idx in range(total_steps):
        # Extract rows for the current timestep
        supply_ef_row = supplybased_ef.iloc[idx]
        demand_row = demand_df.iloc[idx]
        trade_flow_row = transformed_trade_flow.iloc[idx]
        total_output_row = total_output_df.iloc[idx]

        # Supply-based EF in g CO2e/kWh to t CO2e/MWh
        supply_based_ef_t_co2e_mwh = supply_ef_row["Supply-based EF (g CO2e/kWh)"] / 1000

        # Ontario demand and total output
        ontario_demand_mwh = demand_row["Ontario Demand"]
        total_output_mwh = total_output_row["Total Output"]

        # Initialize variables for exports and imports
        total_exports_mwh = 0
        total_imports_mwh = 0
        total_imports_emissions_t_co2e = 0

        # Process trade flow data dynamically
        for col in transformed_trade_flow.columns:
            if col in ["Date", "Hour"]:
                continue

            # Extract the region name using the first 3 characters
            region = col.split(" ")[0].upper()[:3]

            # Get the flow value, handling NaN gracefully
            flow_mwh = trade_flow_row.get(col, 0)
            flow_mwh = float(flow_mwh) if pd.notna(flow_mwh) else 0

            # Lookup emission factor
            applied_factor = neighboring_emission_factors.get(region, 0)
            if applied_factor == 0:
                print(f"Warning: No emission factor found for region '{region}'. Defaulting to 0.")

            # Determine whether the flow is import or export
            if flow_mwh > 0:  # Net export
                total_exports_mwh += flow_mwh
            elif flow_mwh < 0:  # Net import
                import_mwh = abs(flow_mwh)
                total_imports_mwh += import_mwh
                total_imports_emissions_t_co2e += import_mwh * applied_factor

        # Spot-check data for the timestep
        spot_check_data.append({
            "Delivery Date": supply_ef_row["Delivery Date"],
            "Hour": supply_ef_row["Hour"],
            "Total Exports (MWh)": total_exports_mwh,
            "Total Imports (MWh)": total_imports_mwh,
            "Total Import Emissions (t CO2e)": total_imports_emissions_t_co2e
        })

        # Calculate balance difference
        net_balance = total_output_mwh - total_exports_mwh + total_imports_mwh
        balance_difference_mwh = net_balance - ontario_demand_mwh

        # Remove emissions associated with the balance difference
        adjusted_emissions_t_co2e = balance_difference_mwh * supply_based_ef_t_co2e_mwh if balance_difference_mwh > 0 else 0

        # Calculate consumption-based EF
        consumption_ef_t_co2e_mwh = (
            (supply_based_ef_t_co2e_mwh * total_output_mwh) -
            (supply_based_ef_t_co2e_mwh * total_exports_mwh) +
            total_imports_emissions_t_co2e -
            adjusted_emissions_t_co2e
        ) / ontario_demand_mwh if ontario_demand_mwh > 0 else 0

        # Convert EF to g CO2e/kWh for publishing
        consumption_ef_g_co2e_kwh = consumption_ef_t_co2e_mwh * 1000

        # Append the result
        consumption_rows.append({
            "Delivery Date": supply_ef_row["Delivery Date"],
            "Hour": supply_ef_row["Hour"],
            "Consumption-based EF (g CO2e/kWh)": consumption_ef_g_co2e_kwh
        })

        # Progress update every 10% completed
        if (idx + 1) % progress_interval == 0 or (idx + 1) == total_steps:
            elapsed_time = time.time() - start_time
            print(f"Progress: {idx + 1}/{total_steps} ({(idx + 1) / total_steps * 100:.1f}%) steps completed. Elapsed time: {elapsed_time:.2f} seconds.")

    # Create DataFrames for results and spot-check data
    consumption_based_ef = pd.DataFrame(consumption_rows)
    spot_check_df = pd.DataFrame(spot_check_data)

    return consumption_based_ef, spot_check_df





## 3. Setup Data for a Specific Year
### 3.1 Example setup for a single year (2020)
def setup_year_data(year):
    if year < 2020 or year > 2024:
        raise ValueError("Valid years for generator data are 2020-2024.")

    year_dir = os.path.join("data", "IESO", str(year))
    demand_dir = os.path.join(year_dir, "Demand")
    trade_dir = os.path.join(year_dir, "Trade")
    os.makedirs(demand_dir, exist_ok=True)
    os.makedirs(trade_dir, exist_ok=True)

    # Aggregate generator data
    gen_output_df = aggregate_generator_data(year)

    # Parse and clean demand data
    demand_url = f"https://reports-public.ieso.ca/public/Demand/PUB_Demand_{year}.csv"
    demand_path = os.path.join(demand_dir, f"PUB_Demand_{year}.csv")
    download_file(demand_url, demand_path)
    demand_df = parse_and_clean_demand(demand_path)

    # Parse and clean trade flow data
    trade_flow_url = f"https://reports-public.ieso.ca/public/IntertieScheduleFlowYear/PUB_IntertieScheduleFlowYear_{year}.csv"
    trade_flow_path = os.path.join(trade_dir, f"PUB_IntertieScheduleFlowYear_{year}.csv")
    download_file(trade_flow_url, trade_flow_path)
    trade_flow_df = parse_and_clean_trade_flow(trade_flow_path)

    return gen_output_df, demand_df, trade_flow_df




## 4. Main Execution
### 4.1 Main entry point

### Notes: Downloaded generator data, demand data and flow data is all in MW

def main():

    # Simple runtime check for required packages
    import importlib
    for pkg in ["pandas", "numpy", "requests", "openpyxl"]:
        importlib.import_module(pkg)

    # Get emission rates
    emission_rates = get_emission_rates()

    # Get emission factors for neighboring regions
    neighboring_emission_factors = get_neighboring_emission_factors()

    # Download Data
    year = int(input("Enter the year for analysis (valid years are 2020-2024): ") or 2020)

    output_dir = os.path.join("data", "output")
    os.makedirs(output_dir, exist_ok=True)
    consumption_path = os.path.join(output_dir, f"Consumption-based_EF_{year}.csv")

    if os.path.exists(consumption_path):
        print(f"Consumption-based EF for {year} already exists at {consumption_path}.")
        print("Skipping analysis.")
        return

    gen_output_df, demand_df, trade_flow_df = setup_year_data(year)



    # Transform generator data to match new format
    transformed_gen_data = transform_generator_data(gen_output_df)
    # Ensure empty cells are filled with 0
    transformed_gen_data = transformed_gen_data.fillna(0)

    # Transform the trade flow DataFrame
    transformed_trade_flow = transform_trade_flow(trade_flow_df)


    ## Output cleaned data for verification
    #print("\nCleaned Generator Data Preview:")
    #print(gen_output_df.head())
    #print("\nTransformed Generator Data Preview:")
    #print(transformed_gen_data.head())
    #print("\nCleaned Demand Data Preview:")
    #print(demand_df.head())
    #print("\nCleaned Trade Flow Data Preview:")
    #print(trade_flow_df.head())

    ## Save the cleaned data to CSV files
    #output_dir = "data/cleaned_data"
    #os.makedirs(output_dir, exist_ok=True)

    #gen_output_df.to_csv(f"{output_dir}/cleaned_generator_data_{year}.csv", index=False)
    #print(f"Cleaned generator data saved to {output_dir}/cleaned_generator_data_{year}.csv")

    #transformed_gen_data.to_csv(f"{output_dir}/transformed_generator_data_{year}.csv", index=False)
    #print(f"Transformed generator data saved to {output_dir}/transformed_generator_data_{year}.csv")

    #demand_df.to_csv(f"{output_dir}/cleaned_demand_data_{year}.csv", index=False)
    #print(f"Cleaned demand data saved to {output_dir}/cleaned_demand_data_{year}.csv")

    #trade_flow_df.to_csv(f"{output_dir}/cleaned_trade_flow_data_{year}.csv", index=False)
    #print(f"Cleaned trade flow data saved to {output_dir}/cleaned_trade_flow_data_{year}.csv")

    #output_path = f"{output_dir}/transformed_trade_flow_data_{year}.csv"
    #transformed_trade_flow.to_csv(output_path, index=False)
    #print(f"Transformed trade flow data saved to {output_path}")



    # Calculate supply-based EF and total output
    supplybased_ef, total_output_df = calculate_supply_based_ef(transformed_gen_data, emission_rates)


    # Calculate consumption-based EF and retrieve spot-check data
    consumption_based_ef, spot_check_df = calculate_consumption_based_ef(
      supplybased_ef, demand_df, transformed_trade_flow, neighboring_emission_factors, total_output_df
      )

    # Save the consumption-based EF to CSV
    consumption_based_ef.to_csv(consumption_path, index=False)
    print(f"Consumption-based EF data saved to: {consumption_path}")

if __name__ == "__main__":
    main()
