"""
Leo's Automated Weekly Report Generator
Reads cleaned data from EM/SMS automation scripts and generates PowerPoint report
"""

import functions as DS
from pptx import Presentation
import pandas as pd
import os


def load_data_file(filepath):
    """Load Excel file if exists, return None if not found"""
    if os.path.exists(filepath):
        df = pd.read_excel(filepath)
        print(f"Loaded {filepath} ({len(df)} rows)")
        return df
    else:
        print(f"File not found: {filepath}")
        return None


def format_number(x):
    """Format large numbers as K/M"""
    if pd.isna(x) or x == "":
        return ""
    try:
        x = float(x)
        if x >= 1_000_000:
            return f"{x/1_000_000:.2f}M"
        elif x >= 1_000:
            return f"{x/1_000:.1f}K"
        else:
            return str(int(x))
    except:
        return str(x)


def format_percent(x):
    """Format as percentage"""
    if pd.isna(x) or x == "":
        return ""
    try:
        x = float(x)
        if x < 1:
            return f"{x*100:.2f}%"
        else:
            return f"{x:.2f}%"
    except:
        return str(x)


def main():
    # ========== 1. SETUP & LOAD TEMPLATE ==========
    pptx_file = "template.pptx"
    prs = Presentation(pptx_file)
    navigation_screen = DS.create_navigation_screen(prs)
    
    # ========== 2. LOAD CLEANED DATA FILES ==========
    print("\n--- Loading Data Files ---")
    em_data = load_data_file("clean_EM.xlsx")
    em_clicks_data = load_data_file("clean_EM_clicks.xlsx")
    sms_data = load_data_file("clean_SMS.xlsx")
    
    # ========== 3. TITLE SLIDE ==========
    prs = DS.duplicate_slide(prs, navigation_screen['titleslide'], verbose=False)
    title_slide = prs.slides[-1]
    DS.find_replace_text(title_slide, 'PLACE_TEXT_TITLE', "Leo's Automated Weekly Report", verbose=False)
    
    # ========== 4. EMAIL PERFORMANCE ==========
    if em_data is not None:
        prs = DS.duplicate_slide(prs, navigation_screen['subtitleslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'Email Performance', verbose=False)
        
        em_table = em_data[['Touch', 'Cohort', 'Subject Line', 'Deliveries', 'Unique Opens']].head(4).copy()
        em_table.columns = ['Touch', 'Audience', 'Creative', 'Volume', 'CTR']
        em_table['Volume'] = em_table['Volume'].apply(format_number)
        em_table['CTR'] = em_table['CTR'].apply(format_percent)
        
        prs = DS.duplicate_slide(prs, navigation_screen['smsdataslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'Email Performance Overview', verbose=False)
        DS.add_data_table_new(prs.slides[-1], 'Touch', em_table)
    
    # ========== 5. CREATIVE ENGAGEMENT ==========
    if em_clicks_data is not None:
        prs = DS.duplicate_slide(prs, navigation_screen['subtitleslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'Creative Engagement Analysis', verbose=False)
        
        clicks_table = em_clicks_data[['CTA', 'Cohort', 'Delivery Label (Treatment)', 'Total Clicks', 'CTR']].head(4).copy()
        clicks_table.columns = ['Touch', 'Audience', 'Creative', 'Volume', 'CTR']
        clicks_table['Volume'] = clicks_table['Volume'].apply(format_number)
        clicks_table['CTR'] = clicks_table['CTR'].apply(format_percent)
        
        prs = DS.duplicate_slide(prs, navigation_screen['smsdataslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'Creative Engagement: Email', verbose=False)
        DS.add_data_table_new(prs.slides[-1], 'Touch', clicks_table)
    
    # ========== 6. SMS PERFORMANCE ==========
    if sms_data is not None:
        prs = DS.duplicate_slide(prs, navigation_screen['subtitleslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'SMS/RBM Performance', verbose=False)
        
        sms_table = sms_data[['Touch', 'Cohort', 'Creative', 'Deliveries', 'CTR']].head(4).copy()
        sms_table.columns = ['Touch', 'Audience', 'Creative', 'Volume', 'CTR']
        sms_table['Volume'] = sms_table['Volume'].apply(format_number)
        sms_table['CTR'] = sms_table['CTR'].apply(format_percent)
        
        prs = DS.duplicate_slide(prs, navigation_screen['smsdataslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'T1: SMS/RBM Performance', verbose=False)
        DS.add_data_table_new(prs.slides[-1], 'Touch', sms_table)
    
    # ========== 7. FALLBACK: Mock data if no files ==========
    if em_data is None and em_clicks_data is None and sms_data is None:
        print("\nNo data files found, using mock data for demo...")
        
        prs = DS.duplicate_slide(prs, navigation_screen['smsdataslide'], verbose=False)
        DS.find_replace_text(prs.slides[-1], 'PLACE_TEXT_TITLE', 'Performance Overview (Demo)', verbose=False)
        
        mock_data = pd.DataFrame({
            'Touch': ['Launch', 'Preorder', 'Announce', 'Reminder'],
            'Audience': ['Growth_AAL', 'Churn_Upgrade', 'New_Users', 'Engaged'],
            'Creative': ['Spring Campaign', 'Early Access', 'Product Reveal', 'Follow-up'],
            'Volume': ['1.2M', '850K', '2.5M', '500K'],
            'CTR': ['2.5%', '3.8%', '1.9%', '4.2%']
        })
        DS.add_data_table_new(prs.slides[-1], 'Touch', mock_data)
    
    # ========== 8. CLEANUP ==========
    for i in sorted(navigation_screen.values(), reverse=True):
        rId = prs.slides._sldIdLst[i]
        prs.slides._sldIdLst.remove(rId)
        print(f"Deleted template slide at index {i}")
    
    # ========== 9. SAVE ==========
    output_file = 'output_leo_report.pptx'
    prs.save(output_file)
    print(f"\nReport saved to: {output_file}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
