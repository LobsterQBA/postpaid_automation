import pandas as pd
import streamlit as st
import re
from openpyxl import load_workbook
import io
from io import BytesIO

## Set configurations and title of the page
st.set_page_config(page_title="Postpaid Email Heat Maps", layout="wide")
st.title("Postpaid Email Heat Maps")

# --------------------
# Data Input Section
# --------------------
st.subheader("Input the Clean EM Clicks File")
with st.expander("📧 Email Data Inputs", expanded=False):
   EM_clicks_uploaded_file = st.file_uploader(
       "Drop or upload the clean EM clicks data file", 
       type='xlsx', 
       key="EM_clicks"
   )

# Read in the data file as a dataframe in pandas
if EM_clicks_uploaded_file is not None:
   EM_clicks_df = pd.read_excel(EM_clicks_uploaded_file)
   st.write(EM_clicks_df)

## Define mapping: {new column : old column}
mapping = {
   "Module": "Position (Module #)",
   "CTA": "CTA",
   "Offer Details": "CTA Offer Details",
}

## Create the heat map tables
def build_heat_maps(EM_clicks_df, mapping, delivery_label_col="Delivery Label (Treatment)"):

   # Ensure the delivery label column exists
   if delivery_label_col not in EM_clicks_df.columns:
       st.error(f"Error: The '{delivery_label_col}' column was not found in the uploaded data. Please ensure your Excel file has this column or adjust the 'delivery_label_col' variable in the code.")
       return None # Return None if the column is missing
  
   # This will store heat maps for each label
   unique_delivery_labels = EM_clicks_df[delivery_label_col].unique()
   all_heat_maps = {}

   # Filter the dataframe for the current delivery label
   for label in unique_delivery_labels:
       df_for_label = EM_clicks_df[EM_clicks_df[delivery_label_col] == label].copy()

       # Create a blank dataframe for the heat maps
       columns = ["Module", "CTA", "Offer Details", "CTR"]
       heat_map_df = pd.DataFrame(columns=columns)

       # Fill the df in with the columns from the data source
       for new_col, old_cols in mapping.items():
           # Ensure old_cols is always a list
           if not isinstance(old_cols, list):
               old_cols = [old_cols]

           # Find the first matching column in data (should be in df_for_label)
           found_col = next((col for col in old_cols if col in df_for_label.columns), None)

           if found_col is not None:
               col_data = df_for_label[found_col].fillna("") # Use df_for_label here
               # Only round if the column is numeric
               if pd.api.types.is_numeric_dtype(col_data):
                   col_data = col_data.round(0)
               heat_map_df[new_col] = col_data
           else:
              # If a mapped column is not found, fill with NaN or empty string
              heat_map_df[new_col] = "" # Or pd.NA

       ## Calculate CTR
       # Ensure 'Deliveries' and 'Clicks' are numeric in df_for_label
       df_for_label["Deliveries"] = pd.to_numeric(df_for_label["Deliveries"], errors='coerce').fillna(0)
       df_for_label["Clicks"] = pd.to_numeric(df_for_label["Clicks"], errors='coerce').fillna(0)
       
       # Apply CTR calculation using df_for_label
       heat_map_df["CTR"] = df_for_label.apply(
                   lambda row: round(row["Clicks"] / row["Deliveries"], 5) if row["Deliveries"] != 0 else 0.00,
                   axis=1
               )
       
       all_heat_maps[label] = heat_map_df # Store the heat map for this label
   
   # FIX: This return statement is now correctly unindented to the function's base level
   return all_heat_maps

# Button to run the heat map generating function
st.subheader("Build the heat maps")
if st.button("Build the heat maps!"):

   # EM data cleaning
   if EM_clicks_uploaded_file is not None:
       all_heat_maps = build_heat_maps(EM_clicks_df, mapping, delivery_label_col="Delivery Label (Treatment)")

       # Now all_heat_maps is a dictionary, so `if all_heat_maps:` correctly checks if it's not empty
       if all_heat_maps: # Only proceed if heat maps were successfully generated (i.e., dictionary is not empty or None)
          st.session_state["all_heat_maps"] = all_heat_maps
       
       # # Apply campaign metadata
       # # if campaign_name_input:
       # #     final_EM_df["Campaign"] = campaign_name_input
       # # if deploy_date_input:
       # #     final_EM_df["Deploy Date"] = pd.to_datetime(deploy_date_input)

# Show download buttons if data exists in session_state
if "all_heat_maps" in st.session_state and st.session_state["all_heat_maps"]:
  st.subheader("Generated Heat Maps")

  # Display each heat map individually
  for label, df in st.session_state["all_heat_maps"].items():
      st.write(f"### Heat Map for: {label}")
      st.write(df)
      st.markdown("---") # Separator for better readability

  buffer = io.BytesIO()
  with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
      for label, df in st.session_state["all_heat_maps"].items():
          # Excel sheet names have a max length of 31 characters and cannot contain certain chars
          # Added .replace(' ', '_') for robustness
          sheet_name = str(label)[:31].replace('[', '').replace(']', '').replace(':', '').replace('*', '').replace('?', '').replace('/', '').replace('\\','').replace(' ', '_')
          df.to_excel(writer, sheet_name=sheet_name, index=False)
  buffer.seek(0)
  st.download_button(
      label="Download All Heat Maps (Excel)",
      data=buffer,
      file_name="all_heat_maps_by_delivery_label.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )