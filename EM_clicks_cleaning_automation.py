import pandas as pd
import re

## Read in raw EM data file
raw_EM_clicks_df = pd.read_excel("EM_Clicks_T3.xlsx")

    ## Create new dataframe for the cleaned data with the correct column names
    ## Define column names
columns = ["Campaign", "Deploy Date", "Delivery Label (Treatment)", "Touch", "OS", "Cohort", "SL Testing Variant", "Other Testing Variant", "Audience Details 1", "Audience Details 2", "Audience Details 3", "CTA", "Position (Module #)", "CTA Offer Details", "CTA Category", "Link Style", "Device Category", "Device Type", "Deliveries", "Total Clicks", "CTR", "Click Share"]
clean_EM_clicks_df = pd.DataFrame(columns=columns)

## Define mapping: {new column : old column}
mapping = {
    "Deploy Date": "Deploy Date",
    "Delivery Label (Treatment)": ["Delivery Label", "DeliveryLabel", "Label"],
    "Sends": ["Sent", "Processed"],
    "Deliveries": ["Deliveries", "Received", "Success"],
    "Unique Opens": "Unique Opens",
    "Unique Clicks": "Unique Clicks",
    "CTA": "CTA",
    "Total Clicks": ["Clicks", "Total Clicks", "CTA Clicks"]
}

# Define all regex patterns
regex_map = {
    "Touch": r"(T1|T2|T3|T4)",
    "OS": r"(IOS|AND)",
    "Cohort": r"(Growth|Churn)",
    "SL Testing Variant": r"(SLA|SLB)",
    "SMS Testing Variant": r"(|)",
    "Other Testing Variant": r"(|)",
    "Audience Details 1": r"(k12|college)",
    "Audience Details 2": r"(|)",
    "Audience Details 3": r"(|)"
}

## Function to perform EM clicks data cleaning steps: Matching 1:1 columns, pulling info from delivery label, calculating CTR & share of clicks
def transform_data_clicks(raw_EM_clicks_df, mapping, source_col, regex_map):
    ## Pull in metrics from raw data
    for new_col, old_cols in mapping.items():
        # Ensure old_cols is always a list
        if not isinstance(old_cols, list):
            old_cols = [old_cols]

        # Find the first matching column in raw data
        found_col = next((col for col in old_cols if col in raw_EM_clicks_df.columns), None)

        if found_col is not None:
            col_data = raw_EM_clicks_df[found_col].fillna("")
            # Only round if the column is numeric
            if pd.api.types.is_numeric_dtype(col_data):
                col_data = col_data.round(0)
            clean_EM_clicks_df[new_col] = col_data

    ## Pull out info from delivery label
    for new_col, pattern in regex_map.items():
        clean_EM_clicks_df[new_col] =clean_EM_clicks_df[source_col].str.extract(pattern, expand=False, flags=re.IGNORECASE).fillna("")

    ## Drop columns that are not specific to this EM data sheet brought in from the delivery label regex sheet (specifically the SMS Testing column)
    clean_EM_clicks_df.drop(columns=['SMS Testing Variant'], inplace=True)

    ## Calculate CTR
    clean_EM_clicks_df["Deliveries"] = pd.to_numeric(clean_EM_clicks_df["Deliveries"], errors='coerce').fillna(0)
    clean_EM_clicks_df["Total Clicks"] = pd.to_numeric(clean_EM_clicks_df["Total Clicks"], errors='coerce').fillna(0)
    clean_EM_clicks_df["CTR"] = clean_EM_clicks_df.apply(
                lambda row: round(row["Total Clicks"] / row["Deliveries"], 5) if row["Deliveries"] != 0 else 0.00,
                axis=1
            )
    
    ## Calculate click share
    sum_clicks_per_label = {}
    sum_clicks_per_label = clean_EM_clicks_df.groupby('Delivery Label (Treatment)')['Total Clicks'].transform('sum')
    clean_EM_clicks_df["Click Share"] = clean_EM_clicks_df.apply(
                    lambda row: round(row["Total Clicks"] / sum_clicks_per_label[row.name], 5),
                    axis=1
                )
    
    return clean_EM_clicks_df


## Run the function
final_EM_clicks_df = transform_data_clicks(raw_EM_clicks_df, mapping, "Delivery Label (Treatment)", regex_map)

## Save the clean EM click data as a csv file
final_EM_clicks_df.to_excel('clean_EM_clicks.xlsx', index=False)
