import requests
import pandas as pd
import json
import re
import streamlit as st
from io import BytesIO

def combine_excel_sheets(input_file, output_file):
    """
    Combines all sheets from an Excel file into a single sheet and returns it as a BytesIO object.
    
    Parameters:
    input_file (BytesIO): The input Excel file as a BytesIO object.
    
    Returns:
    BytesIO: The combined Excel file as a BytesIO object.
    """
    input_file.seek(0)
    excel_data = pd.read_excel(input_file, sheet_name=None)

    combined_data = pd.concat(excel_data.values(), ignore_index=True)

    output_stream = BytesIO()
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        combined_data.to_excel(writer, index=False, sheet_name='Combined_Data')

    output_stream.seek(0)
    return output_stream

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
            # Load category data from a local file or replace with Streamlit secrets if needed.
            data = st.secrets["CATEGORIES_JSON"]
            json_data = json.loads(data)

            industry_ids = [
                "22102000000", "22101000000", "22107000000", "22108000000", "22109000000", "22106000000", "22999000000", "22112000000"
                # Add more IDs as required...
            ]
            x_rapidapi_key = st.secrets["X-RAPIDAPI-KEY"]
            url = "https://tiktok-api23.p.rapidapi.com/api/trending/ads"
            headers = {
                "x-rapidapi-key": x_rapidapi_key,
                "x-rapidapi-host": "tiktok-api23.p.rapidapi.com"
            }

            main_excel_stream = BytesIO()
            secondary_excel_stream = BytesIO()

            with pd.ExcelWriter(main_excel_stream, engine='openpyxl') as writer_main, \
                 pd.ExcelWriter(secondary_excel_stream, engine='openpyxl') as writer_secondary:
                for industry_id in industry_ids:
                    industry_name = get_industry_name(industry_id, json_data)
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

                                    ad_data = {
                                        "Ad ID": ad.get('id'),
                                        "Brand Name": sanitize_string(ad.get('brand_name')),
                                        "Ad Industry": sanitize_string(get_industry_name(ad.get('industry_key'), json_data)),
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
                                    }
                                    all_ad_data.append(ad_data)

                                    if float(ad.get('ctr', 0)) >= 0.05 and int(ad.get('like', 0)) >= 2000 and int(detailed_ads.get('comment', 0)) >= 150:
                                        filtered_ad_data.append(ad_data)
                                except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                                    st.error(f"Error processing ad {ad.get('id')}: {e}")
                        except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                            st.error(f"Error retrieving data for industry {industry_id} on page {page}: {e}")

                    pd.DataFrame(all_ad_data).to_excel(writer_main, index=False, sheet_name=sheet_name)
                    pd.DataFrame(filtered_ad_data).to_excel(writer_secondary, index=False, sheet_name=sheet_name)

            st.success("Ads data processed successfully!")

            # Create a downloadable combined Excel file.
            combined_data_stream = combine_excel_sheets(secondary_excel_stream, None)

            st.download_button(
                "Download Combined Top Ads Data",
                combined_data_stream,
                file_name="combined_top_tiktok_ads_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    getTikTokAds()
