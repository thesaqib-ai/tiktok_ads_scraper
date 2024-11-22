import requests
import pandas as pd
import json
import re
import streamlit as st

def combine_excel_sheets(input_file, output_file):
    """
    Reads all sheets from an Excel file and combines them into a single sheet.
    
    Parameters:
    input_file (str): Path to the input Excel file.
    output_file (str): Path to save the combined Excel file.
    """
    excel_data = pd.read_excel(input_file, sheet_name=None)
    
    combined_data = pd.concat(excel_data.values(), ignore_index=True)

    combined_data.to_excel(output_file, index=False, sheet_name='Combined_Data')
    st.success(f"Data combined and saved to {output_file}")

def get_industry_name(industry_id, json_data):
    industry_id = industry_id.replace('label_', '')

    for category in json_data:
        if category["id"] == industry_id:
            return category["name"]
        
        for sub_category in category.get("sub_industry", []):
            if sub_category["id"] == industry_id:
                return sub_category["name"]
    
    return "-"

def sanitize_string(value):
    if isinstance(value, str):
        return re.sub(r'[\x00-\x1F\x7F]', '', value)
    return value

def getTikTokAds():
    st.title('TikTok Ads Scraper')
    if st.button("Start Scraping Ads"):
        with st.spinner("Fetching TikTok Ads..."):
            file_path = "categories.json"
            with open(file_path) as file:
                data = json.load(file)

            industry_ids = [
                "22102000000", "22101000000", 
                # "22107000000", "22108000000", "22109000000", "22106000000", "22999000000", "22112000000",
                # "22105000000", "22113000000", "22110000000", "22111000000", "16105000000", "16104000000", "16100000000", "11102000000",
                # "20108000000", "12104000000", "12108000000", "12107000000", "12109000000", "12999000000", "12106000000", "14105000000",
                # "14104000000", "14107000000", "14106000000", "14101000000", "14100000000", "14103000000", "14102000000", "24103000000",
                # "24109000000", "24117000000", "24112000000", "24999000000", "24113000000", "24100000000", "24102000000", "30100000000",
                # "30101000000", "10000000000", "10103000000", "23114000000", "23122000000", "23102000000", "23111000000", "23116000000",
                # "23104000000", "23107000000", "23124000000", "19000000000", "19999000000", "19103000000", "19101000000", "19102000000",
                # "19105000000", "19106000000", "19104000000", "19100000000", "28100000000", "28101000000", "15100000000", "15105000000",
                # "15103000000", "15101000000", "15102000000", "15104000000", "15106000000", "15107000000", "15999000000", "11111000000",
                # "11101000000", "11103000000"
            ]
            x_rapidapi_key = st.secrets["X-RAPIDAPI-KEY"]
            url = "https://tiktok-api23.p.rapidapi.com/api/trending/ads"
            headers = {
                "x-rapidapi-key": x_rapidapi_key,
                "x-rapidapi-host": "tiktok-api23.p.rapidapi.com"
            }

            main_excel_file = "tiktok_ads_data.xlsx"
            secondary_excel_file = "top_ads_data.xlsx"
            with pd.ExcelWriter(main_excel_file, engine='openpyxl') as writer_main, \
             pd.ExcelWriter(secondary_excel_file, engine='openpyxl') as writer_secondary:
                for industry_id in industry_ids:
                    industry_name = get_industry_name(industry_id, data)
                    sheet_name = industry_name if industry_name != "-" else f"Industry_{industry_id}"
                    all_ad_data = []
                    filtered_ad_data = []
                    for page in range(1, 11):
                        querystring = {
                            "page": str(page),
                            "period": "7",
                            "limit": "10",
                            "country": "US",
                            "order_by": "ctr",
                            "like": "1",
                            "ad_format": "2",
                            "industry": industry_id,
                            "ad_language": "en"
                        }

                        try:
                            response = requests.get(url, headers=headers, params=querystring)
                            response.raise_for_status()
                            ads = response.json().get('data', {}).get('materials', [])

                            for ad in ads:
                                try:
                                    details_url = "https://tiktok-api23.p.rapidapi.com/api/trending/ads/detail"
                                    details_querystring = {"ads_id": ad.get('id')}
                                    ads_response = requests.get(details_url, headers=headers, params=details_querystring)
                                    ads_response.raise_for_status()
                                    detailed_ads = ads_response.json().get('data', {})

                                    all_ad_data.append({
                                        "Ad ID": ad.get('id'),
                                        "Brand Name": sanitize_string(ad.get('brand_name')),
                                        "Ad Industry": sanitize_string(get_industry_name(ad.get('industry_key'), data)),
                                        "CTR": ad.get('ctr'),
                                        "Ad Objective": sanitize_string(ad.get('objective_key')),
                                        "Total Likes": ad.get('like'),
                                        "Total Comments": detailed_ads.get('comment'),
                                        "Total Shares": detailed_ads.get('share'),
                                        "Video Url": ad.get('video_info', {}).get('video_url', {}).get('720p'),
                                        "Video Cover URL": ad.get('video_info', {}).get('cover', {}),
                                        "Video Duration": ad.get('video_info', {}).get('duration'),
                                        "Landing Page": detailed_ads.get('landing_page'),
                                        "Ad Description": sanitize_string(ad.get('ad_title'))
                                    })

                                    if float(ad.get('ctr', 0)) >= 0.05 and int(ad.get('like', 0)) >= 2000 and int(detailed_ads.get('comment', 0)) >= 150:
                                        filtered_ad_data.append({
                                            "Ad ID": ad.get('id'),
                                            "Brand Name": sanitize_string(ad.get('brand_name')),
                                            "Ad Industry": sanitize_string(get_industry_name(ad.get('industry_key'), data)),
                                            "CTR": ad.get('ctr'),
                                            "Ad Objective": sanitize_string(ad.get('objective_key')),
                                            "Total Likes": ad.get('like'),
                                            "Total Comments": detailed_ads.get('comment'),
                                            "Total Shares": detailed_ads.get('share'),
                                            "Video Url": ad.get('video_info', {}).get('video_url', {}).get('720p'),
                                            "Video Cover URL": ad.get('video_info', {}).get('cover', {}),
                                            "Video Duration": ad.get('video_info', {}).get('duration'),
                                            "Landing Page": detailed_ads.get('landing_page'),
                                            "Ad Description": sanitize_string(ad.get('ad_title'))
                                        })
                                except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                                    st.error(f"Error processing ad {ad.get('id')}: {e}")
                        except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                            st.error(f"Error retrieving data for industry {industry_id} on page {page}: {e}")

                    all_data_df = pd.DataFrame(all_ad_data)
                    top_data_df = pd.DataFrame(filtered_ad_data)

                    all_data_df.to_excel(writer_main, index=False, sheet_name=sheet_name)
                    top_data_df.to_excel(writer_secondary, index=False, sheet_name=sheet_name)
                    
                    st.success(f"Ads for {sheet_name} appended to {main_excel_file} and Top Ads appended to {secondary_excel_file}")

            combine_excel_sheets('top_ads_data.xlsx', 'combined_top_ads_data.xlsx')

            # Provide download options for both Excel files
            st.download_button("Download TikTok Ads Data", main_excel_file, file_name=main_excel_file)
            st.download_button("Download Combined Data", "combined_top_ads_data.xlsx", file_name="combined_top_ads_data.xlsx")

if __name__ == "__main__":
    getTikTokAds()
