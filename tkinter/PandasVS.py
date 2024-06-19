import pandas as pd   
import sys

# Check if the file path is provided as a command-line argument
if len(sys.argv) != 3:
    print("Usage: python PandasVS.py <input_file_path> <output_file_path>")
    sys.exit(1)

# Get the file path from the command-line argument
input_file_path = sys.argv[1]
output_file_path = sys.argv[2]

# Load the Excel file into a DataFrame
df = pd.read_excel(input_file_path)

#filepath = input('Enter filepath name: ')
#df = pd.read_csv(filepath)

# Load the CSV file into a DataFrame without using any column as the index
#df = pd.read_csv('Clarivate Data Delivery 202405 Hospital Only 2017 thru 2024 Q1(sales_hovertech_202405).csv', index_col=None)

# Define the specific COT_GROUP value to filter
specific_cot = ['HOSPITAL/HEALTH SYSTEM']

# Filter the DataFrame based on the specific COT_GROUP
df_filtered = df[df['COT_GROUP'].isin(specific_cot)].copy()

# List of columns in the order you want them
cols = ['COT_GROUP', 'MANF_DESC', 'SKU', 'PROD_DESC']

# Reorder the columns
df = df_filtered[cols]

# Combine the YEAR and QUARTER columns into a single column named 'year-quarter' with a dash
df_filtered['year-quarter'] = df_filtered['YEAR'].astype(str) + '-Q' + df_filtered['QUARTER'].astype(str)

# Pivot the DataFrame so that 'year-quarter' are the columns and 'REVENUE' and 'UNITS' are the values
pivot_revenue = df_filtered.pivot_table(index=['COT_GROUP', 'MANF_DESC', 'SKU', 'PROD_DESC'],
                                        columns='year-quarter',
                                        values='REVENUE',
                                        aggfunc='sum')

pivot_units = df_filtered.pivot_table(index=['COT_GROUP', 'MANF_DESC', 'SKU', 'PROD_DESC'],
                                      columns='year-quarter',
                                      values='UNITS',
                                      aggfunc='sum')

# Sort the columns based on the year and quarter values
pivot_revenue = pivot_revenue.reindex(sorted(pivot_revenue.columns), axis=1)
pivot_units = pivot_units.reindex(sorted(pivot_units.columns), axis=1)

# Rename the columns to include 'Revenue' and 'Units'
pivot_revenue.columns = [f"{col} Revenue" for col in pivot_revenue.columns]
pivot_units.columns = [f"{col} Units" for col in pivot_units.columns]

# Interleave the columns of revenue and units
revenue_cols = pivot_revenue.columns.tolist()
units_cols = pivot_units.columns.tolist()

# Ensure both have the same length and correct pairing
interleaved_columns = []
for rev_col, unit_col in zip(revenue_cols, units_cols):
    interleaved_columns.append(rev_col)
    interleaved_columns.append(unit_col)

# Combine the DataFrames
combined = pd.concat([pivot_revenue, pivot_units], axis=1)
combined = combined[interleaved_columns]


# Reset index to flatten the DataFrame
combined.reset_index(inplace=True, drop=False)

# Replace NaN values with blank
#combined = combined.fillna('')

# Copy and modify the original DataFrames
pivot_revenue_copy = pivot_revenue.copy() * 0.98
pivot_units_copy = pivot_units.copy()

# Rename the columns in the copied DataFrames
pivot_revenue_copy.columns = [f'{col} SUM' for col in pivot_revenue_copy.columns]
pivot_units_copy.columns = [f'{col} SUM' for col in pivot_units_copy.columns]

# Create a list of DataFrames in the order you want them interleaved
dfs = [pivot_revenue, pivot_units, pivot_revenue_copy, pivot_units_copy]

# Use pd.concat with keys to create a multi-index DataFrame, then swap levels and sort to interleave
combined = pd.concat(dfs, axis=1, keys=range(len(dfs))).swaplevel(0, 1, axis=1).sort_index(axis=1)
# Concatenate the DataFrames without keys
# Flatten the column MultiIndex
combined.columns = combined.columns.get_level_values(0)


# Reset index to flatten the DataFrame
combined.reset_index(inplace=True)

# Replace NaN values with blank
#combined = combined.fillna('')

# Interleave the original columns
original_columns = [col for pair in zip(pivot_revenue.columns, pivot_units.columns) for col in pair]

# Interleave the copied columns
copied_columns = [col for pair in zip(pivot_revenue_copy.columns, pivot_units_copy.columns) for col in pair]

# Concatenate the interleaved original and copied columns
columns = cols + original_columns + copied_columns

# Assign the interleaved columns to the combined DataFrame
combined = combined[columns]

# Extract the years from the column names
years = set(col.split('-')[0] for col in pivot_revenue.columns)

# Initialize a dictionary to store the total revenue and total units for each year
totals_dict = {year: {'revenue': [], 'units': []} for year in years}

# Iterate over the DataFrame columns once
for col in combined.columns:
    # Split the column name into components
    components = col.split()

    # Check if the column name has the correct format
    if len(components) < 3:
        continue

    year = components[0].split('-')[0]
    type_ = components[1]

    if type_ in ["Revenue", "Units"]:
        totals_dict[year][type_.lower()].append(col)



# Initialize a new DataFrame to store the total revenue and total units for each year
totals = pd.DataFrame(index=combined.index)

# For each year, create a new column in the totals DataFrame with the total revenue and total units for that year
for year in sorted(totals_dict.keys()):
    totals[f'{year} TOTAL REVENUE'] = combined[totals_dict[year]['revenue']].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    totals[f'{year} TOTAL UNITS'] = combined[totals_dict[year]['units']].apply(pd.to_numeric, errors='coerce').sum(axis=1)

years = set(col.split('-')[0] for col in pivot_revenue.columns)

# Initialize a dictionary to store the total revenue and total units for each year
ASP_dict = {year: {'revenue': [], 'units': []} for year in years}

# Iterate over the DataFrame columns once
for col in combined.columns:
    # Split the column name into components
    components = col.split()

    # Check if the column name has the correct format
    if len(components) < 3:
        continue

    year = components[0].split('-')[0]
    type_ = components[1]

    if type_ in ["Revenue", "Units"]:
        ASP_dict[year][type_.lower()].append(col)

# Initialize a new DataFrame to store the total revenue and total units for each year
ASPs = pd.DataFrame(index=combined.index)

# For each year, create a new column in the totals DataFrame with the total revenue and total units for that year
for year in sorted(ASP_dict.keys()):
    revenue = combined[ASP_dict[year]['revenue']].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    units = combined[ASP_dict[year]['units']].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    

    ASPs[f'{year} ASP'] = revenue / units

    ###
# Extract the quarters from the column names
quarters = set(col.split()[0] for col in pivot_revenue.columns)

# Initialize a dictionary to store the total revenue and total units for each quarter
ASP_6_dict = {quarter: {'revenue': [], 'units': []} for quarter in quarters}

# Iterate over the DataFrame columns once
for col in combined.columns:
    # Split the column name into components
    components = col.split()

    # Check if the column name has the correct format
    if len(components) < 3:
        continue

    quarter = components[0]
    type_ = components[1]

    if type_ in ["Revenue", "Units"]:
        ASP_6_dict[quarter][type_.lower()].append(col)

# Get the last 6 quarters
last_6_quarters = sorted(ASP_6_dict.keys())[-6:]

# Initialize a new DataFrame to store the ASP for the last 6 quarters
ASPs_6_quarters = pd.DataFrame(index=combined.index)

# For each of the last 6 quarters, create a new column in the ASP DataFrame with the ASP for that quarter
for quarter in last_6_quarters:
    revenue = combined[ASP_6_dict[quarter]['revenue']].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    units = combined[ASP_6_dict[quarter]['units']].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    ASPs_6_quarters[f'{quarter} ASP'] = revenue / units
    ###


# If indices do not match, reset them
df.reset_index(drop=True, inplace=True)
combined.reset_index(drop=True, inplace=True)
totals.reset_index(drop=True, inplace=True)
ASPs.reset_index(drop=True, inplace=True)
ASPs_6_quarters.reset_index(drop=True, inplace=True)

# Try concatenating again
combined_final = pd.concat([combined, totals, ASPs, ASPs_6_quarters], axis=1, ignore_index=False)

# Round the revenue and units columns to the nearest hundredths
def round_columns(combined_final):
    for col in combined_final.columns:
        if 'Revenue' in col or 'Units' in col or 'TOTAL' in col or 'ASP' in col:
            combined_final[col] = pd.to_numeric(combined_final[col], errors='coerce').round(2)
    return combined_final

# Apply the function to your DataFrame
combined_final = round_columns(combined_final)

# Continue with your code
#combined_final = combined_final.replace(0, '') 
#combined_final = combined_final.fillna('')

#print(combined_final.head().to_string(index=False))

# Save to CSV (if needed)
combined_final.to_excel(output_file_path, index=False)
