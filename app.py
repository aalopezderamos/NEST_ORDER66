import streamlit as st
import pandas as pd
import pyautogui
import pygetwindow as gw
import time

# --- Page Settings ---
st.set_page_config(page_title="NEST Order 66", layout="wide")

# --- Logo in top-right corner ---
st.markdown(
    """
    <div style='text-align: right'>
        <img src='https://d1ynl4hb5mx7r8.cloudfront.net/wp-content/uploads/2020/02/19180100/logo.png' width='280'>
    </div>
    """,
    unsafe_allow_html=True
)

# --- App Title ---
st.title("🧠 SEB NEST Order 66")

# --- Upload Excel File ---
uploaded_file = st.file_uploader("📤 Upload your 'NEST_Order66.xlsx' file", type=["xlsx"])

# --- GIF placeholders ---
gif_placeholder = st.empty()

if uploaded_file:
    # Load Excel data
    df = pd.read_excel(uploaded_file)
    df = df[['sku_id', 'yhat', 'Week number']]  # Ensure correct columns

    # Get unique SKU list
    sku_list = sorted(df['sku_id'].unique())

    # Select SKU
    selected_sku = st.selectbox("🔍 Select SKU to input:", sku_list)

    if st.button("🚀 Execute Automation"):
        filtered_df = df[df['sku_id'] == selected_sku].sort_values(by='Week number')

        # Display Order 66 GIF
        gif_placeholder.markdown(
            "<div style='text-align: center'>"
            "<img src='https://media.giphy.com/media/xTiIzrRyvrFijaEtY4/giphy.gif' width='300'>"
            "</div>",
            unsafe_allow_html=True
        )

        # Try focusing the browser
        browser_window = None
        for w in gw.getWindowsWithTitle(''):
            if any(browser in w.title for browser in ['Chrome', 'Edge']):
                browser_window = w
                break

        if browser_window:
            browser_window.activate()
            time.sleep(1)
            pyautogui.click()
            st.success("✅ Browser window focused.")
            time.sleep(2)
        else:
            st.warning("⚠️ Could not auto-focus browser. Click into Week 1 field manually.")
            time.sleep(5)

        st.info(f"Typing values for SKU: **{selected_sku}**...")
        for value in filtered_df['yhat']:
            pyautogui.write(str(round(value, 2)))
            pyautogui.press('tab')
            time.sleep(0.2)

        # Replace GIF with Darth Vader
        gif_placeholder.markdown(
            "<div style='text-align: center'>"
            "<img src='https://media.giphy.com/media/12qFOaBbu9TZny/giphy.gif' width='300'>"
            "</div>",
            unsafe_allow_html=True
        )

        st.success("✅ Input complete!")

else:
    st.info("👆 Upload your Excel file to begin.")
