import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(page_title="NeuroSinQ NPS Adjuster", layout="centered")

st.image("nslogo.png", width=200)
st.title("NeuroSinQ NPS Adjuster")

st.markdown("""
### INSTRUCTIONS:
1. Please upload a file with **ONE worksheet** named `"Data"`.
2. Enter the correct row numbers for:
   - **Last Data Record**
   - **Desired NPS Row**
3. Click the button below to apply adjustments.
""")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
individual_score_end = st.number_input("Enter row number of LAST data record (e.g., 1503)", min_value=1, step=1)-1
desired_nps_row = st.number_input("Enter row number for Desired NPS (e.g., 1523)", min_value=1, step=1)-1

can_run = uploaded_file is not None and individual_score_end > 0 and desired_nps_row > 0

if st.button("Make Adjustment", disabled=not can_run):

    def calculate_nps(scores):
        scores = np.array(scores)
        total = len(scores)
        if total == 0:
            return 0
        promoters = np.sum((scores == 4) | (scores == 5))
        detractors = np.sum((scores == 1) | (scores == 2))
        return (promoters - detractors) / total * 100

    def adjust_nps(df, individual_score_start, individual_score_end, desired_nps_row):
        last_column_index = df.shape[1] - 1
        total_columns = last_column_index + 1
        progress = st.progress(0)
        status = st.empty()

        for col in range(0, last_column_index + 1):
            try:
                desired_nps = float(df.at[desired_nps_row, col])
            except (ValueError, TypeError):
                continue  # skip if not a valid desired NPS

            original_scores = df.loc[individual_score_start:individual_score_end, col].copy()
            scores = pd.to_numeric(original_scores, errors='coerce').dropna()

            if len(scores) < 10:
                #status.warning(f"Skipping column {col} due to low valid responses ({len(scores)})")
                progress.progress((col + 1) / total_columns)
                continue

            neutral_cap = np.random.uniform(3, 12)
            #status.info(f"Processing column {col}: Desired NPS = {desired_nps} | Neutral cap = {round(neutral_cap, 2)}%")

            current_nps = calculate_nps(scores)
            changes = 0
            max_changes = 400

            while abs(current_nps - desired_nps) > 0.1 and changes < max_changes:
                neutral_percent = (scores == 3).sum() / len(scores) * 100

                if current_nps > desired_nps:
                    candidates = scores[(scores == 5) | (scores == 4)].index
                    if len(candidates):
                        idx = np.random.choice(candidates)
                        scores.loc[idx] = 4 if scores.loc[idx] == 5 else 3
                        changes += 1
                    elif neutral_percent > 0:
                        candidates = scores[scores == 3].index
                        if len(candidates):
                            idx = np.random.choice(candidates)
                            scores.loc[idx] = 2
                            changes += 1
                        else:
                            break
                    else:
                        break
                else:
                    if neutral_percent < neutral_cap:
                        candidates = scores[scores == 2].index
                        if len(candidates):
                            idx = np.random.choice(candidates)
                            scores.loc[idx] = 3
                            changes += 1
                        else:
                            candidates = scores[scores == 3].index
                            if len(candidates):
                                idx = np.random.choice(candidates)
                                scores.loc[idx] = 4
                                changes += 1
                            else:
                                break
                    else:
                        candidates = scores[scores == 3].index
                        if len(candidates):
                            idx = np.random.choice(candidates)
                            scores.loc[idx] = 4
                            changes += 1
                        else:
                            break

                current_nps = calculate_nps(scores)

            df.loc[individual_score_start:individual_score_end, col] = scores
            progress.progress((col + 1) / total_columns)

        progress.progress(1.0)  # Ensure progress bar stays at 100%
        status.success("✅ Adjustment complete. Processing the results.You can now download the adjusted file once the results are ready.")
        return df

    try:
        df = pd.read_excel(uploaded_file, sheet_name='Data', header=None)
        adjusted_df = adjust_nps(df.copy(), individual_score_start=3, individual_score_end=individual_score_end, desired_nps_row=desired_nps_row)

        output_filename = uploaded_file.name.replace(".xlsx", "_Adjusted.xlsx")
        adjusted_df.to_excel(output_filename, index=False, header=False)

        with open(output_filename, "rb") as f:
            st.download_button(label="📥 Download Adjusted File", data=f, file_name=output_filename)

    except Exception as e:
        st.error(f"❌ Error: {e}")