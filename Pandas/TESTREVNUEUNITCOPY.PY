import pandas as pd

# Load the CSV file into a DataFrame without using any column as the index
df = pd.read_csv('Clarivate Data Delivery 202405 Hospital Only 2017 thru 2024 Q1(sales_hovertech_202405).csv', index_col=None)

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
combined = combined.fillna('')

# Copy and modify the original DataFrames
pivot_revenue_copy = (pivot_revenue.copy() * 0.98).round(2)
pivot_units_copy = pivot_units.copy().round(2)

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

# Round the revenue and units columns to the nearest hundredths
for col in combined.columns:
    if 'Revenue' in col or 'Units' in col:
        combined[col] = combined[col].round(2)

# Reset index to flatten the DataFrame
combined.reset_index(inplace=True)

# Replace NaN values with blank
combined = combined.fillna('')

# Interleave the original columns
original_columns = [col for pair in zip(pivot_revenue.columns, pivot_units.columns) for col in pair]

# Interleave the copied columns
copied_columns = [col for pair in zip(pivot_revenue_copy.columns, pivot_units_copy.columns) for col in pair]

# Concatenate the interleaved original and copied columns
columns = original_columns + copied_columns

# Assign the interleaved columns to the combined DataFrame
combined = combined[columns]

# Display the combined DataFrame without index
print(combined.to_string(index=False))

combined.to_csv('testingtest2.csv', index=False)


