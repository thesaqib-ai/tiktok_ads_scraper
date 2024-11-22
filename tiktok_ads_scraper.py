import requests
import pandas as pd
import json
import re
import streamlit as st
from io import BytesIO

def combine_excel_sheets(input_file):
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
    # Custom CSS styling
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');

            body {
                background-color: #f0f2f6;
                font-family: 'Montserrat', sans-serif;
            }
            .stApp {
                background-image: url('https://images.unsplash.com/photo-1564750566895-e0a4857846f8?ixlib=rb-4.0.3&auto=format&fit=crop&w=1350&q=80');
                background-size: cover;
                background-position: center;
            }
            h1 {
                color: #ffffff;
                text-align: center;
                font-weight: 700;
                font-size: 3em;
                margin-top: 0.5em;
                text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            }
            .stButton>button {
                background-color: #1DA1F2;
                color: #ffffff;
                border: none;
                padding: 0.75em 1.5em;
                font-size: 1em;
                font-weight: 700;
                border-radius: 8px;
                margin: 0 auto;
                display: block;
                box-shadow: 2px 2px 4px rgba(0,0,0,0.2);
            }
            .stButton>button:hover {
                background-color: #0d95e8;
                cursor: pointer;
            }
            .css-1kyxreq {
                background-color: rgba(255, 255, 255, 0.85);
                padding: 2em;
                border-radius: 15px;
            }
            .download-section {
                text-align: center;
                margin-top: 2em;
                background-color: rgba(255, 255, 255, 0.85);
                padding: 2em;
                border-radius: 15px;
            }
            .download-button {
                background-color: #28a745;
                color: #ffffff;
                border: none;
                padding: 0.75em 1.5em;
                font-size: 1em;
                font-weight: 700;
                border-radius: 8px;
                margin: 1em;
                box-shadow: 2px 2px 4px rgba(0,0,0,0.2);
            }
            .download-button:hover {
                background-color: #218838;
                cursor: pointer;
            }
            .stSpinner>div>div {
                border-top-color: #1DA1F2;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title('TikTok Ads Scraper')

    # Initialize session state to store Excel files
    if 'main_excel_stream' not in st.session_state:
        st.session_state.main_excel_stream = None
    if 'secondary_excel_stream' not in st.session_state:
        st.session_state.secondary_excel_stream = None

    if st.button("Start Scraping Ads"):
        with st.spinner("Fetching TikTok Ads..."):
            try:
                # Load category data from Streamlit secrets
                data = st.secrets["CATEGORIES_JSON"]
                json_data = json.loads(data)

                industry_ids = [
                    "22102000000", "22101000000"
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

                        # Write all ads data to the main Excel file
                        pd.DataFrame(all_ad_data).to_excel(writer_main, index=False, sheet_name=sheet_name)
                        # Write filtered ads data to the secondary Excel file
                        pd.DataFrame(filtered_ad_data).to_excel(writer_secondary, index=False, sheet_name=sheet_name)

                # Store the Excel streams in session state
                st.session_state.main_excel_stream = main_excel_stream
                st.session_state.secondary_excel_stream = secondary_excel_stream

                st.success("Ads data processed successfully!")

            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

    # Display download buttons only if the Excel files are available
    if st.session_state.secondary_excel_stream and st.session_state.main_excel_stream:
        combined_data_stream = combine_excel_sheets(st.session_state.secondary_excel_stream)
        
        st.markdown("<div class='download-section'><h3>Download Files</h3>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="Download Top Ads Data",
                data=combined_data_stream,
                file_name="top_tiktok_ads_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='download-top',
            )

        with col2:
            st.download_button(
                label="Download All Ads Data",
                data=st.session_state.main_excel_stream,
                file_name="tiktok_ads_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='download-all',
            )
        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    getTikTokAds()
