import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

st.set_page_config(page_title="Core Sandbox Engine", layout="wide")

# Precision UI Masking: Removes fork and toolbar while preserving sidebar mechanics
st.markdown("""
    <style>
        /* Target and remove ONLY the right-side toolbar and options dropdown */
        div[data-testid="stToolbar"] {
            display: none !important;
            visibility: hidden !important;
        }
        
        /* Clear the upper colorful accent decoration strip */
        div[data-testid="stDecoration"] {
            display: none !important;
            visibility: hidden !important;
        }
        
        /* Keep the platform status indicators hidden */
        div[data-testid="stStatusWidget"] {
            visibility: hidden !important;
            display: none !important;
        }
        
        /* Adjust the block container top padding to keep elements high without breaking headers */
        .block-container {
            padding-top: 2rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Custom CSS to enhance aesthetics (sleek cards and colors)
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .section-card {
        background-color: #F9FAFB;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1.5rem;
    }
    .result-box {
        background-color: #ECFDF5;
        border: 1px solid #10B981;
        padding: 1rem;
        border-radius: 0.375rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Main Title & Subtitle
st.markdown('<div class="main-header">📊 Mini-Excel: Accounting & MIS Lab</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">An interactive portal to learn spreadsheet operations, ledger flow, and basic system logic.</div>', unsafe_allow_html=True)

# System Architecture Labeling for MIS Students
st.info("💻 **System Flow Overview:** Data Input (Ledger) ➡️ Processing (Formula Engine) ➡️ Output (Visual Analytics)")

# Formula Library - easily extensible dictionary with math, description, and concept definitions
# Only financial formulas (SUM, AVERAGE, MIN, MAX) are supported
FORMULA_LIBRARY = {
    "SUM": {
        "func": lambda s: s.sum(),
        "description": "Adds all values. Negative numbers represent expenses or outflows, so SUM calculates the net cash flow or net profit/loss.",
        "concept": "Summing positive revenue and negative expenses nets out the total balance. A negative result means the company had a net loss.",
        "math": r"\text{SUM} = x_1 + x_2 + \dots + x_n"
    },
    "AVERAGE": {
        "func": lambda s: s.mean() if len(s) > 0 else 0.0,
        "description": "Calculates the arithmetic mean. It averages out both inflows (positives) and outflows (negatives) to find the net typical transaction size.",
        "concept": "Calculates the average balance impact per transaction, accounting for both positive additions and negative deductions.",
        "math": r"\text{AVERAGE} = \frac{\sum_{i=1}^{n} x_i}{n}"
    },
    "MAX": {
        "func": lambda s: s.max() if len(s) > 0 else 0.0,
        "description": "Finds the highest value. In business, this identifies the largest revenue or cash inflow (the most positive number).",
        "concept": "Extracts the maximum positive inflow. If all transactions are negative, it highlights the smallest expense.",
        "math": r"\text{MAX} = \max(x_1, x_2, \dots, x_n)"
    },
    "MIN": {
        "func": lambda s: s.min() if len(s) > 0 else 0.0,
        "description": "Finds the lowest value. In business, this identifies the largest expense or cash outflow (the most negative number).",
        "concept": "Extracts the minimum transaction value. Heavily negative values pinpoint the single largest cash payment.",
        "math": r"\text{MIN} = \min(x_1, x_2, \dots, x_n)"
    }
}

st.sidebar.markdown("""
### 💡 How to use:
Explore pre-loaded accounting transactions below. Learn how basic spreadsheet calculations (`SUM`, `AVERAGE`, `MIN`, `MAX`) relate to accounting metrics.
""")

# Sidebar Formula Cheat Sheet Reference Library
with st.sidebar.expander("📖 Formula Cheat Sheet", expanded=True):
    st.markdown("""
    Keep this cheat sheet open for reference during calculations and tasks!
    
    ### 🧮 Mathematical & Statistical
    * **SUM**: Adds a range of cells.
    * **AVERAGE**: Calculates the arithmetic mean.
    * **MIN**: Finds the smallest value in a range.
    * **MAX**: Finds the largest value in a range.
    * **MEDIAN**: Finds the middle value in a range.
    * **COUNT**: Counts cells containing numbers.
    * **COUNTA**: Counts non-empty cells in a range.
    * **ROUND**: Rounds a number to specified decimal places.
    * **ABS**: Returns the absolute value of a number.
    * **SUMIF**: Sums cells that meet a specified condition.

    ### ⚖️ Logical
    * **IF**: Checks a condition; returns one value if TRUE, another if FALSE.
    * **AND**: Returns TRUE if all arguments are true.
    * **OR**: Returns TRUE if any argument is true.
    * **NOT**: Reverses the logical value of its argument.
    * **IFERROR**: Returns a specified value if a formula errors out.

    ### 🔤 Text Operations
    * **CONCATENATE**: Joins two or more text strings.
    * **UPPER**: Converts a text string to all uppercase.
    * **LOWER**: Converts a text string to all lowercase.
    * **LEFT**: Extracts a specific number of characters from the start of a string.
    * **RIGHT**: Extracts a specific number of characters from the end of a string.
    * **LEN**: Returns the number of characters in a text string.
    * **TRIM**: Removes all spaces from text except single spaces between words.

    ### 📅 Date & Time
    * **TODAY**: Returns the current date.
    * **NOW**: Returns the current date and time.
    * **YEAR**: Extracts the year from a given date.
    * **MONTH**: Extracts the month from a given date.
    * **DAY**: Extracts the day from a given date.

    ### 🔍 Lookup & Reference
    * **VLOOKUP**: Looks up a value in the first column of a table to return a value in another column.
    * **HLOOKUP**: Looks up a value in the top row of a table to return a value in another row.
    * **INDEX**: Returns a value or reference from within a table or range.
    * **MATCH**: Locates the relative position of a lookup value in a range.
    """)

# Utility function to get column letter in Excel style (1-based)
def get_column_letter(index):
    return chr(65 + index) if index < 26 else f"Col{index}"

# Define sample data with 1-based index and Excel Column Letters
default_data = pd.DataFrame([
    {"(A) Item": "Sales Revenue", "(B) Amount": 5000.0},
    {"(A) Item": "Office Rent", "(B) Amount": -1200.0},
    {"(A) Item": "Software License", "(B) Amount": -150.0},
    {"(A) Item": "Consulting Income", "(B) Amount": 2500.0},
    {"(A) Item": "Utility Bills", "(B) Amount": -300.0},
    {"(A) Item": "Advertising Spend", "(B) Amount": -800.0},
])
default_data.index = range(1, len(default_data) + 1)
default_data.index.name = "Row"

# Helper callbacks to sync selectboxes with formula bars
def update_demo_formula():
    sel = st.session_state.demo_formula
    num_rows = len(st.session_state.get("demo_editor", default_data))
    # 'Amount' is column B (index 1)
    st.session_state.demo_formula_bar = f"={sel}(B1:B{num_rows})"

# Main Guided Ledger Layout
st.subheader("📚 Guided Demo: Understanding Basic Ledger Calculations")

# Split UI into layout columns: Table (Left), Actions & Charts (Right)
col1, col2 = st.columns([3, 2], gap="large")

with col1:
    st.markdown("### 📥 Transaction Ledger")
    st.write("Below is a standard ledger. Note the column letters (**A, B**) and row numbers (**1 to 6**). Double-click values to edit them.")
    
    # Display sample accounting data in a data editor
    edited_df = st.data_editor(
        default_data,
        num_rows="dynamic",
        use_container_width=True,
        key="demo_editor"
    )
    
    # Ensure 1-based index and named index persists
    if not edited_df.empty:
        edited_df.index = range(1, len(edited_df) + 1)
        edited_df.index.name = "Row"
        
    # Export Button at the bottom of the ledger table section
    demo_csv = edited_df.to_csv(index=True)
    st.download_button(
        label="📥 Export Ledger to CSV",
        data=demo_csv,
        file_name="guided_demo_ledger.csv",
        mime="text/csv",
        help="Exporting converts your data into a universal format, bridging our classroom tool with industry-standard spreadsheet software.",
        key="demo_export_csv"
    )
    st.success("Data successfully exported! You can now open this file in Excel to perform advanced reporting.")
    
    st.markdown("---")
    st.markdown("### 🧮 Formula Bar & Calculation")
    
    # Initialize selected formula key to ensure we don't cause lookup error on load
    current_formula = st.session_state.get("demo_formula", list(FORMULA_LIBRARY.keys())[0])
    tooltip_text = FORMULA_LIBRARY[current_formula]["description"]
    
    # Selection of the formula pulling from formula library
    formula_selection = st.selectbox(
        "Select formula:",
        options=list(FORMULA_LIBRARY.keys()),
        key="demo_formula",
        help=tooltip_text,
        on_change=update_demo_formula
    )
    
    # Display contextual explanations immediately below selection
    st.info(f"**{formula_selection}**: {FORMULA_LIBRARY[formula_selection]['description']}")
    
    # Find index of the target column '(B) Amount'
    num_rows = len(edited_df)
    amount_col_idx = 1
    for idx, col_name in enumerate(edited_df.columns):
        if "Amount" in col_name:
            amount_col_idx = idx
            break
    
    col_letter = get_column_letter(amount_col_idx)
    default_formula_text = f"={formula_selection}({col_letter}1:{col_letter}{num_rows})"
    
    # Initialize session state for formula bar if not set
    if "demo_formula_bar" not in st.session_state:
        st.session_state.demo_formula_bar = default_formula_text
        
    # Formula Bar (fx) - clearly labeled as Pseudo-code
    formula_bar = st.text_input("Formula Bar (fx) - Pseudo-code:", key="demo_formula_bar")
    
    calculate_clicked = st.button("Calculate", key="demo_calculate_btn")
    
    # Calculate logic with input validation
    target_col = edited_df.columns[amount_col_idx] if len(edited_df.columns) > amount_col_idx else None
    valid_input = True
    try:
        if target_col and not edited_df.empty:
            # Clean invalid entries and coerce to numeric (keeps positive and negative numbers)
            amounts = pd.to_numeric(edited_df[target_col], errors='coerce').dropna()
            if len(amounts) == 0:
                valid_input = False
        else:
            valid_input = False
    except Exception:
        valid_input = False
        
    if "demo_result" not in st.session_state:
        st.session_state.demo_result = None
        st.session_state.demo_formula_run = None
        st.session_state.demo_formula_bar_text = None
        
    if calculate_clicked:
        if not valid_input:
            st.session_state.demo_result = "error_sentinel"
        else:
            func_to_apply = FORMULA_LIBRARY[formula_selection]["func"]
            try:
                st.session_state.demo_result = func_to_apply(amounts)
                st.session_state.demo_formula_run = formula_selection
                st.session_state.demo_formula_bar_text = formula_bar
            except Exception:
                st.session_state.demo_result = "error_sentinel"

    # Result box display with validation warning check
    if st.session_state.demo_result is not None:
        if st.session_state.demo_result == "error_sentinel":
            st.warning("Please enter valid data. Make sure amount records contain valid numbers.")
        else:
            formula_used = st.session_state.demo_formula_bar_text or formula_bar
            formatted_val = f"${st.session_state.demo_result:,.2f}"
            st.markdown(f"""
            <div class="result-box">
                <strong>Result for {formula_used}:</strong>
                <span style="font-size: 1.25rem; font-weight: bold; color: #065F46; margin-left: 10px;">
                    {formatted_val}
                </span>
            </div>
            """, unsafe_allow_html=True)
            
    # Permanent "How it works" expander directly under controls (linked to currently selected formula)
    lib_entry = FORMULA_LIBRARY.get(formula_selection)
    if lib_entry:
        with st.expander("📖 How it works", expanded=True):
            st.markdown(f"""
            **Formula Definition ({formula_selection}):**
            {lib_entry["description"]}
            
            **Mathematical Expression:**
            $${lib_entry["math"]}$$
            
            **Accounting / Practical Concept:**
            {lib_entry["concept"]}
            """)
        
with col2:
    st.markdown("### 📈 Visual Analytics Dashboard")
    
    # Visualization Gallery Selector
    chart_type = st.radio(
        "Choose Visualization:",
        ["Comparison (Bar)", "Trend (Area)", "Composition (Donut)"],
        horizontal=True,
        key="demo_chart_selector"
    )
    
    # Auto-updating chart based on edited table data
    item_col = edited_df.columns[0] if len(edited_df.columns) > 0 else None
    amount_col = edited_df.columns[1] if len(edited_df.columns) > 1 else None
    
    if not edited_df.empty and item_col and amount_col:
        chart_data = edited_df.dropna(subset=[item_col, amount_col]).copy()
        chart_data[amount_col] = pd.to_numeric(chart_data[amount_col], errors='coerce')
        chart_data = chart_data.dropna(subset=[item_col, amount_col])
        
        if len(chart_data) > 0:
            if chart_type == "Comparison (Bar)":
                chart_data["Type"] = chart_data[amount_col].apply(lambda x: "Revenue/Inflow" if x >= 0 else "Expense/Outflow")
                st.bar_chart(
                    data=chart_data,
                    x=item_col,
                    y=amount_col,
                    color="Type",
                    use_container_width=True
                )
                st.caption("ℹ️ **Bar Chart:** Bar charts are ideal for comparing distinct categories of transaction items.")
                
            elif chart_type == "Trend (Area)":
                # Cumulative balance trend
                chart_data["Cumulative Balance"] = chart_data[amount_col].cumsum()
                st.area_chart(
                    data=chart_data,
                    x=item_col,
                    y="Cumulative Balance",
                    use_container_width=True
                )
                st.caption("ℹ️ **Area Chart:** Area charts are excellent for showing cumulative trends and net balance changes sequentially.")
                
            elif chart_type == "Composition (Donut)":
                # Absolute amounts for relative composition ratios
                chart_data_donut = chart_data.copy()
                chart_data_donut["Abs Amount"] = chart_data_donut[amount_col].abs()
                
                donut = alt.Chart(chart_data_donut).mark_arc(innerRadius=50).encode(
                    theta=alt.Theta(field="Abs Amount", type="quantitative"),
                    color=alt.Color(field=item_col, type="nominal"),
                    tooltip=[item_col, amount_col]
                ).properties(height=280)
                
                st.altair_chart(donut, use_container_width=True)
                st.caption("ℹ️ **Donut Chart:** Donut charts display composition, showing the relative size of each item (by absolute value) compared to the total budget.")
        else:
            st.info("Waiting for data...")
    else:
        st.info("Waiting for data...")

# ---------------------------------------------------------
# New UI Section: Formula Reference Manual
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📚 Formula Reference Manual")
with st.expander("View Formula Reference Guide", expanded=False):
    st.markdown("""
    | Formula Name | Definition/Use Case |
    | :--- | :--- |
    | **SUM** | Adds all numerical values in a range of cells. |
    | **AVERAGE** | Calculates the arithmetic mean of a range of cells. |
    | **MIN** | Finds the lowest numerical value in a range of cells. |
    | **MAX** | Finds the highest numerical value in a range of cells. |
    | **MEDIAN** | Finds the middle value in a range of cells. |
    | **COUNT** | Counts cells that contain numerical values. |
    | **COUNTA** | Counts non-empty cells in a range. |
    | **IF** | Evaluates a logical condition and returns one value if true and another if false. |
    | **SUMIF** | Adds values in a range that meet a specified single criteria. |
    | **COUNTIF** | Counts the number of cells in a range that meet a specified single criteria. |
    | **VLOOKUP** | Looks up a value in the first column of a vertical range and returns a value in the same row from a specified column. |
    | **HLOOKUP** | Looks up a value in the first row of a horizontal range and returns a value in the same column from a specified row. |
    | **CONCATENATE** | Joins two or more text strings together. |
    | **LEFT** | Extracts a specified number of characters starting from the left of a text string. |
    | **RIGHT** | Extracts a specified number of characters starting from the right of a text string. |
    | **MID** | Extracts characters from the middle of a text string, starting at a specified position. |
    | **UPPER** | Converts all characters in a text string to uppercase. |
    | **LOWER** | Converts all characters in a text string to lowercase. |
    | **TRIM** | Removes extra spaces from text, leaving only single spaces between words. |
    | **NOW** | Returns the current system date and time. |
    | **TODAY** | Returns the current system date. |
    | **RANK** | Returns the statistical rank of a number relative to a list of numbers. |
    | **IFERROR** | Catches formulas that result in an error and displays a custom message/fallback value. |
    | **ROUND** | Rounds a number to a specified number of decimal places. |
    | **POWER** | Raises a number to a specified mathematical power. |
    """)
