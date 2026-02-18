import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Set style
sns.set_theme(style="whitegrid")

# Paths
data_path = 'data/Grad Program Exit Survey Data 2024.xlsx'
output_dir = 'outputs'
output_file = os.path.join(output_dir, 'rank_order.png')

# Create output dir if not exists
os.makedirs(output_dir, exist_ok=True)

# Load data
# We read header=None first to get the question text in row 1 (index 1)
df_raw = pd.read_excel(data_path, header=None)
# Row 0: Header Codes
# Row 1: Question Text
question_texts = df_raw.iloc[1].tolist()
header_codes = df_raw.iloc[0].tolist()

# Map codes to questions
code_to_question = dict(zip(header_codes, question_texts))

# Define columns of interest
core_columns = ['Q35_1', 'Q35_5', 'Q35_2', 'Q35_4', 'Q35_3', 'Q35_8', 'Q35_9', 'Q35_10']
elective_columns = ['Q76_1', 'Q77_2', 'Q78_3', 'Q83_4', 'Q82_5', 'Q80_6', 'Q81_9', 'Q79_7']

# Load actual data
# header=0 means first row (codes) is columns.
# Rows 1 and 2 in the file (0-indexed) are subheader/question text and ImportId.
# Row 3 is data start.
df = pd.read_excel(data_path, header=0)
# df index 0 is Question Text
# df index 1 is ImportId
# df index 2 is Data start
df_data = df.iloc[2:].copy() # Skip metadata rows

# Helper to extract course name
def extract_course_name(question_text):
    if pd.isna(question_text):
        return "Unknown"
    text = str(question_text)
    if ' - ' in text:
        return text.rsplit(' - ', 1)[-1].strip()
    return text

# Process Core Courses (Rank)
core_data = []
for col in core_columns:
    if col in df.columns:
        q_text = code_to_question.get(col, "Unknown")
        course_name = extract_course_name(q_text)

        # Convert to numeric, errors='coerce' turns non-numeric to NaN
        series = pd.to_numeric(df_data[col], errors='coerce')
        mean_rank = series.mean()
        core_data.append({'Course': course_name, 'Mean Rank': mean_rank})
    else:
        print(f"Warning: Column {col} not found in dataframe.")

df_core = pd.DataFrame(core_data).sort_values('Mean Rank') # Ascending for rank (1 is best)

# Process Elective Courses (Rating)
elective_data = []
for col in elective_columns:
    if col in df.columns:
        q_text = code_to_question.get(col, "Unknown")
        course_name = extract_course_name(q_text)

        series = pd.to_numeric(df_data[col], errors='coerce')
        mean_rating = series.mean()
        elective_data.append({'Course': course_name, 'Mean Rating': mean_rating})
    else:
        print(f"Warning: Column {col} not found in dataframe.")

df_elective = pd.DataFrame(elective_data).sort_values('Mean Rating', ascending=False) # Descending for rating (5 is best)

# Plotting
fig, axes = plt.subplots(2, 1, figsize=(12, 14))

# Plot Core
if not df_core.empty:
    sns.barplot(x='Mean Rank', y='Course', data=df_core, ax=axes[0], hue='Course', palette='viridis', legend=False)
    axes[0].set_title('Core Courses Ranking (Lower Rank is Better)')
    axes[0].set_xlabel('Mean Rank')
    axes[0].set_ylabel('')
    # Add values to bars
    for i, v in enumerate(df_core['Mean Rank']):
        axes[0].text(v + 0.05, i, f'{v:.2f}', va='center')

# Plot Elective
if not df_elective.empty:
    sns.barplot(x='Mean Rating', y='Course', data=df_elective, ax=axes[1], hue='Course', palette='magma', legend=False)
    axes[1].set_title('Elective Courses Rating (Higher Rating is Better)')
    axes[1].set_xlabel('Mean Rating (1-5)')
    axes[1].set_ylabel('')
    # Add values to bars
    for i, v in enumerate(df_elective['Mean Rating']):
        axes[1].text(v + 0.05, i, f'{v:.2f}', va='center')

plt.tight_layout()
plt.savefig(output_file)
print(f"Figure saved to {output_file}")
