import streamlit as st
import pandas as pd
import pyautogui
import pygetwindow as gw
import time

@st.cache_data
def load_data():
    return pd.read_excel('NEST_Order66.xlsx')

df = load_data()
sku_list = sorted(df['sku_id'].unique())
selected_sku = st.selectbox("Select SKU to input", sku_list)

if st.button("Execute Automation"):

    filtered_df = df[df['sku_id'] == selected_sku].sort_values(by='Week number')

    browser_window = None
    for w in gw.getWindowsWithTitle(''):
        if any(browser in w.title for browser in ['Chrome', 'Edge']):
            browser_window = w
            break

    if browser_window:
        browser_window.activate()
        time.sleep(1)
        pyautogui.click()
        st.success("Browser window focused. Starting input...")
        time.sleep(2)
    else:
        st.warning("Browser window not found. Please click it manually.")
        time.sleep(5)

    for value in filtered_df['yhat']:
        pyautogui.write(str(round(value, 2)))
        pyautogui.press('tab')
        time.sleep(0.2)

    st.success("Done inputting values for " + selected_sku)
