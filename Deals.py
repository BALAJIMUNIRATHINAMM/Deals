import pandas as pd
import streamlit as st
from worksheet.draup import DashboardFormatter
import os
print(os.listdir("/"))

# Page Configuration
st.set_page_config(
    page_title="Outsourcing Deals Data Formatter",
    layout="wide",
    initial_sidebar_state="auto",
    menu_items={
        'About': 'Outsourcing Deals Data Formatter streamlines data requests and improves collaboration among stakeholders.'
    }
)

# Header & Subheader
st.markdown("<h2 style='text-align: center; color: #4e4e4e;'>Outsourcing Deals Formatter</h2>", unsafe_allow_html=True)
st.header('Braindesk Data Project Optimization Pullouts', divider='rainbow')

# Upload Section
st.header("📤 Upload Your Files")
col1, col2 = st.columns(2)
col3, = st.columns(1)

with col1:
    deals_file = st.file_uploader("Deals CSV File", type=['csv'], key="deals")
with col2:
    headcount_file = st.file_uploader("Headcount CSV File", type=['csv'], key="headcount")
with col3:
    formatter_file = st.file_uploader("Formatter Excel Template", type=['xlsx'], key="formatter")

custom_title = st.text_input("Enter Custom Title for Formatted File", "Requested Accounts")

# Cached Function to Clean & Format Data
@st.cache_data
def clean_and_format_data(deals, headcount):
    required_columns = ['Description', 'Deal Start Date', 'Deal End Date', 'Headcount', 
                        'Linkedin Link', 'Gmail Id', 'Client Name', 'Provider Name', 
                        'Provider MSA', 'Client MSA']

    # Ensure all required columns are present
    missing_cols = [col for col in required_columns if col not in deals.columns]
    if missing_cols:
        raise ValueError(f"Missing columns in Deals file: {', '.join(missing_cols)}")

    deals['Description'] = deals['Description'].str.replace('[-=]', '', regex=True)
    deals = deals[~deals['Description'].str.contains("#NAME?", na=False)]

    # Format Dates
    deals['Deal Start Date'] = pd.to_datetime(deals['Deal Start Date'], errors='coerce')
    deals['Deal End Date'] = pd.to_datetime(deals['Deal End Date'], errors='coerce')

    deals['Formatted Start Date'] = deals['Deal Start Date'].apply(
        lambda x: f"Q{int((x.month + 2) / 3)} {x.year}" if pd.notnull(x) else '-'
    )
    deals['Formatted End Date'] = deals['Deal End Date'].apply(
        lambda x: f"Q{int((x.month + 2) / 3)} {x.year}" if pd.notnull(x) else '-'
    )

    # Headcount Range Mapping
    headcount_mapping = headcount.set_index('Main')['Range']
    deals['Headcount Range'] = deals['Headcount'].map(headcount_mapping).fillna('-')

    # LinkedIn & Gmail Filtering
    deals['Linkedin Link_clean'] = deals['Linkedin Link'].apply(
        lambda x: x if 'linkedin.com' in str(x) else None
    )
    deals['Gmail Id_clean'] = deals['Gmail Id'].apply(
        lambda x: x if '@gmail.com' in str(x) else None
    )
    deals['LinkedIn_URL_CVID'] = deals.apply(
        lambda row: row['Linkedin Link_clean'] or row['Gmail Id_clean'], axis=1
    ).fillna('Secondary Research')

    # Remove Duplicates
    deals['Duplication check'] = (
        deals['Client Name'] + deals['Provider Name'] + deals['Description'] +
        deals['Provider MSA'] + deals['Formatted Start Date']
    )
    deals = deals.drop_duplicates(subset='Duplication check')
    deals.fillna('-', inplace=True)

    # Deliverables
    internal = deals[[
        'Deal Id', 'Client Name', 'Draup Verticals', 'Provider Name', 'Client MSA', 'Provider MSA',
        'Headcount Range', 'Description', 'Formatted Start Date', 'Business Function',
        'Functional Workload', 'Digital Product', 'Skills', 'Digital Technology Evidence', 
        'Headcount', 'Deal Start Date'
    ]]

    client = deals[[
        'Deal Id', 'Client Name', 'Draup Verticals', 'Provider Name', 'Provider MSA', 'Description',
        'Formatted Start Date', 'Functional Workload', 'Digital Product', 'Skills',
        'Digital Technology Evidence'
    ]]

    zinnov = deals[[
        'Deal Id', 'Client Name', 'Draup Verticals', 'Provider Name', 'Client MSA', 'Provider MSA',
        'Headcount Range', 'Description', 'Formatted Start Date', 'Formatted End Date',
        'Business Function', 'Functional Workload', 'Digital Product', 'Skills',
        'Digital Technology Evidence', 'LinkedIn_URL_CVID', 'Headcount', 'Client Subsidiary'
    ]]

    return internal, client, zinnov

# Main Logic
if deals_file and headcount_file and formatter_file:
    with st.spinner("🔄 Processing your files..."):
        try:
            deals = pd.read_csv(deals_file)
            headcount = pd.read_csv(headcount_file)
            internal, client, zinnov = clean_and_format_data(deals, headcount)

            deliverable_option = st.selectbox("Select Deliverable Type", ["Internal", "Client", "Zinnov"])

            deliverable_map = {
                "Internal": internal,
                "Client": client,
                "Zinnov": zinnov
            }

            selected_data = deliverable_map[deliverable_option]

            st.subheader(f"📋 Preview of {deliverable_option} Deliverable")
            st.dataframe(selected_data, use_container_width=True)

            # Download CSV
            st.subheader("📥 Download Processed Files")
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                st.download_button(
                    label=f"Download {deliverable_option} CSV",
                    data=selected_data.to_csv(index=False),
                    file_name=f"{deliverable_option.lower()}_deliverable.csv",
                    mime="text/csv"
                )

            # Formatter & Excel Export
            with col_dl2:
                st.subheader("\U0001F3A8 Formatted Excel Output")
                formatter = DashboardFormatter(formatter_file, title=custom_title)

                formatter_func = {
                    "Internal": formatter.internal_deals,
                    "Client": formatter.client_deals,
                    "Zinnov": formatter.client_deals  # assuming Zinnov uses same format as Client
                }

                formatter_func[deliverable_option](selected_data)

                output_name = f"deals_{deliverable_option.lower()}_formatted.xlsx"
                formatter.save(output_name)

                with open(output_name, "rb") as file:
                    st.download_button(
                        label=f"Download {deliverable_option} Formatted Excel",
                        data=file,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            st.success("✅ All files are ready for download!")

        except Exception as e:
            st.error(f"🚨 An error occurred while processing your files:\n\n{e}")
else:
    st.warning("⚠️ Please upload all three required files to proceed.")

# Footer
st.markdown(
    """
    <style>
    .footer {position: fixed; left: 0; bottom: -17px; width: 100%;
             background-color: #b1b1b5; color: black; text-align: center;}
    </style>
    <div class="footer"><p>© 2025 BDDRDP | Powered by Draup | All Rights Reserved</p></div>
    """, 
    unsafe_allow_html=True
)
