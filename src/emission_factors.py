"""Process IESO data for emission factor calculations.

This script is a direct conversion of the original Google Colab notebook. It no
longer mounts Google Drive and instead expects the input data to live under
``data/IESO_Data`` relative to the repository root.
"""

import os

import numpy as np
import pandas as pd

# Base path relative to this file where the data is located
BASE_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "IESO_Data")

# Define file path templates for the different data types
supply_template = 'Supply/GOC-{}.xlsx'
demand_template = 'Demand/PUB_DemandZonal_{}.csv'
generator_list_template = 'Generator_List.xlsx'
transmission_template = 'Transmission/PUB_IntertieScheduleFlowYear_{}.csv'


def main() -> None:
    """Load and preprocess data files."""

    # Initialize an empty dictionary to store data for each year
    data = {}

    # Loop through each year from 2016 to 2023
    for year in range(2016, 2024):
        # Generate file paths based on the year using the relative base path
        supply_file = os.path.join(BASE_PATH, supply_template.format(year))
        demand_file = os.path.join(BASE_PATH, demand_template.format(year))
        generator_list_file = os.path.join(BASE_PATH, generator_list_template)
        transmission_file = os.path.join(BASE_PATH, transmission_template.format(year))

        # Load files and handle missing data gracefully
        try:
            if os.path.exists(supply_file) and os.path.exists(demand_file) and os.path.exists(
                generator_list_file
            ) and os.path.exists(transmission_file):
                supply_data = pd.read_excel(supply_file)
                demand_data = pd.read_csv(demand_file)
                generator_list = pd.read_excel(generator_list_file)
                transmission_data = pd.read_csv(transmission_file)

                data[year] = {
                    "supply": supply_data,
                    "demand": demand_data,
                    "generator_list": generator_list,
                    "transmission": transmission_data,
                }
            else:
                print(f"Files for {year} are missing. Skipping this year.")

        except Exception as e:
            print(f"Error loading data for {year}: {e}")

    def clean_data(df):
        """Replace NaN values with zero."""
        return df.fillna(0)

    for year in data:
        data[year]["supply"] = clean_data(data[year]["supply"])
        data[year]["demand"] = clean_data(data[year]["demand"])
        data[year]["transmission"] = clean_data(data[year]["transmission"])

    for year in data:
        try:
            time_data = pd.to_datetime(data[year]["supply"].iloc[:, 0], errors="coerce")
            hours_data = pd.to_numeric(data[year]["supply"].iloc[:, 1], errors="coerce")
            if hours_data.notnull().all():
                time_data = time_data + pd.to_timedelta(hours_data - 1, unit="h")
            data[year]["time"] = time_data
        except Exception as e:
            print(f"Error processing time data for {year}: {e}")

    regions = [
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
    technologies = ["Biofuel", "Hydro", "Natural Gas", "Nuclear", "Solar", "Wind"]

    demand_data = data[year]["demand"]
    demand_by_region = {}
    for i in range(1, 11):
        demand_by_region[i] = demand_data.iloc[:, 2 + i].to_numpy()

    gen = {
        region: {tech: np.zeros(len(data[year]["time"])) for tech in technologies}
        for region in regions
    }

    gencount = 0
    for col_idx in range(3, len(data[year]["supply"].columns)):
        generator_name = data[year]["supply"].columns[col_idx]

        for _, row in data[year]["generator_list"].iterrows():
            if generator_name == row[0]:
                region_name = row[2]
                if region_name in regions:
                    region = region_name

                tech_name = row[1]
                if tech_name in technologies:
                    technology = tech_name

                gen[region][technology] = np.column_stack(
                    (gen[region][technology], data[year]["supply"].iloc[:, col_idx])
                )
                gen[region][technology][np.isnan(gen[region][technology])] = 0
                gencount += 1
                break

    print(f"Processed year {year} with {gencount} generators")

if __name__ == '__main__':
    main()

