import pandas as pd
import re
import argparse
import xlsxwriter

#get CSV from command line
parser = argparse.ArgumentParser(description="Validate AP exam item data")
parser.add_argument('file_path', help='Path to the CSV file')
args = parser.parse_args()
file_path = args.file_path

#load and clean csv
df = pd.read_csv(file_path)
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('\n', '_')

df['item_status'] = df['item_status'].replace(
    {'Accepted without modifications' : 'Finalized'}
)

allowed_statuses = [
    'Finalized',
    'Accepted with modifications',
    'Accepted with minor modifications'
]


def colnum_to_excel_col(n):
    result = ''
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


report = {}

#1. check for items per section per intended_form with totals
items_per_form_section = df.groupby(['intended_form', 'section']).size().unstack(fill_value=0)
items_with_totals = items_per_form_section.copy()
items_with_totals.loc['Total'] = items_per_form_section.sum()
report['items_per_form_section'] = items_with_totals.reset_index()
equal_items_check = items_per_form_section.nunique().max() == 1

#2. check for duplicate items
duplicates = df[df.duplicated('item_sequence', keep=False)].copy()
duplicates['row_number'] = duplicates.index + 2
report['duplicate_items'] = duplicates[['row_number', 'item_sequence', 'intended_form', 'section']]
duplicate_item_sequences = duplicates['item_sequence'].unique()
#df['item_ready'] = True
#df.loc[df['item_sequence'].isin(duplicate_item_sequences), 'item_ready'] = False
#these aren't working right- come back to later

#3 check topic, skill, and difficulty balance
topic_distribution = df.groupby('intended_form')['topic_(label)'].value_counts(normalize=True).unstack(
    fill_value=0)
skill_distribution = df.groupby('intended_form')['skill'].value_counts(normalize=True).unstack(
    fill_value=0)
complexity_distribution = df.groupby('intended_form')['complexity'].value_counts(normalize=True).unstack(
    fill_value=0)

topic_distribution = topic_distribution.reset_index().rename(
    {'intended_form': 'Form'}, axis=1)  # Add Form Number
skill_distribution = skill_distribution.reset_index().rename(
    {'intended_form': 'Form'}, axis=1)  # Add Form Number
complexity_distribution = complexity_distribution.reset_index().rename(
    {'intended_form': 'Form'}, axis=1)  # Add Form Number

report['topic_distribution'] = topic_distribution
report['skill_distribution'] = skill_distribution
report['complexity_distribution'] = complexity_distribution

#4 check that skills and subskills match
def extract_skill_number(val):
    match = re.search(r'\d+', str(val))
    return int(match.group()) if match else None

df['skill_number'] = df['skill'].apply(extract_skill_number)
df['subskill_number'] = df['subskill'].apply(extract_skill_number)
mismatch_rows =df[df['skill_number'] != df['subskill_number']].copy()
mismatch_rows['row_number'] = mismatch_rows.index + 2
report['skill_subskill_mismatch'] = mismatch_rows[
    ['row_number', 'item_sequence', 'intended_form', 'section', 'skill', 'subskill']]

#5 items needing > 30 revision time
needs_revision = df[~df['item_status'].isin(allowed_statuses)].copy()
needs_revision['row_number'] = needs_revision.index + 2
report['items_needing_too_much_revision'] = needs_revision[['row_number', 'item_sequence', 'intended_form', 'item_status']]

#6 check for missing graphics sign off
df['graphics_status'] = df['graphics_status'].fillna('').str.strip()
graphics_issues = df[df['graphics_status'].str.lower() != 'graphics lead'].copy()
graphics_issues['row_number'] = graphics_issues.index + 2
report['missing_graphics_sign_off'] = graphics_issues[['row_number', 'item_sequence', 'intended_form', 'graphics_status']]

#7 overall form readiness
df['item_ready'] = True
df.loc[df['item_sequence'].isin(duplicate_item_sequences), 'item_ready'] = False  # Duplicates
df.loc[~df['item_status'].isin(allowed_statuses), 'item_ready'] = False  # Revision needed
df.loc[df['graphics_status'].str.lower() != 'graphics lead', 'item_ready'] = False  # Missing graphics
df.loc[df['skill_number'] != df['subskill_number'], 'item_ready'] = False # Skill/Subskill mismatch

form_readiness = df.groupby('intended_form')['item_ready'].mean() * 100
form_readiness = form_readiness.reset_index().rename(columns={'intended_form': 'Form', 'item_ready': 'Readiness (%)'})
form_readiness['Readiness (%)'] = form_readiness['Readiness (%)'].round(2)
report['form_readiness_percent'] = form_readiness


#8 not ready items summary
not_ready = df[~df['item_ready']].copy()
not_ready['row_number'] = not_ready.index + 2


def describe_issues(row):
    issues = []
    if row['item_status'] not in allowed_statuses:
        issues.append('Needs >30 min revision')
    if (row['graphics_status']).strip().lower() != 'graphics lead':
        issues.append('Graphics status issue')
    if row['item_sequence'] in duplicate_item_sequences:
        issues.append('Duplicate item')
    if row['skill_number'] != row['subskill_number']:
        issues.append('Skill/Subskill mismatch')
    return '; '.join(issues)


not_ready['issues'] = not_ready.apply(describe_issues, axis=1)
report['not_ready_items'] = not_ready[
    ['row_number', 'item_sequence', 'intended_form', 'item_status', 'graphics_status', 'issues']]

print("Number of items in not_ready_items:", len(report['not_ready_items'])) #debug

#summary data
report['summary'] = pd.DataFrame({
    'Check' : [
        'Equal number of items per section',
        'Duplicate item sequences found',
        'Items needing >30 min revision',
        'Items missing graphics sign-off',
        'Skill/Subskill mismatches'
    ],
    'Results' : [
        'PASS' if equal_items_check else 'FAIL',
        'YES' if not duplicates.empty else 'NO',
        len(needs_revision),
        len(graphics_issues),
        len(mismatch_rows)
    ]
})

report['summary'].loc[report['summary']['Check'] == 'Duplicate item sequences found', 'Results'] = 'YES' if not duplicates.empty else 'NO'

#9 export to excel
with pd.ExcelWriter('validation_report.xlsx', engine='xlsxwriter') as writer:
    workbook = writer.book
    
    #autocolors still not working right- need more time on xlsx writer
    bold = writer.book.add_format({"bold" : True})
    green_fill = workbook.add_format({'bg_color': "#70F682"})
    yellow_fill = workbook.add_format({'bg_color': '#F6F070'})

    #write the all_data sheet
    df_export = df.copy()
    df_export['original_excel_row'] = df_export.index + 2
    df_export['item_ready_for_excel'] = df_export['item_ready'].map(
        {True: 'TRUE', False: 'FALSE'})  # create new column with string
    print("Data type of 'item_ready_for_excel' before writing to Excel:",
        df_export['item_ready_for_excel'].dtype)  # check datatype
    df_export.to_excel(writer, sheet_name='all_data', index=False, header=True)
    worksheet_all = writer.sheets['all_data']
    worksheet_all.set_column("A:AA", 10)
    num_rows_all, num_cols_all = df_export.shape
    last_col_letter_all = colnum_to_excel_col(num_cols_all - 1)
    highlight_range_all = f'A2:{last_col_letter_all}{num_rows_all + 1}'
    print(f"Highlight range: {highlight_range_all}")  # print range

    #TRUE
    worksheet_all.conditional_format(
        highlight_range_all,
        {'type': 'formula', 'criteria': '=AA2="TRUE"', 'format': green_fill}
    )

    #FALSE
    worksheet_all.conditional_format(
        highlight_range_all,
        {'type': 'formula', 'criteria': '=$AA2="FALSE"', 'format': yellow_fill}
    )
   
    #write other dataframes
    for sheet_name, data in report.items():
        if not data.empty:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:Z', 15)

    #format summary sheet
    if 'summary' in report:
        summary_worksheet = writer.sheets['summary']
        summary_worksheet.set_column('A:B', 10)
        #bold header
        header_format = workbook.add_format({'bold': True})
        for col_num, value in enumerate(report['summary'].columns.values):
            summary_worksheet.write(0, col_num, value, header_format)


#print console summary
print("== FORM VALIDATION SUMMARY ==")
print("✔ Same number of items per section:", equal_items_check)
print("✖ Duplicate items across forms:", not duplicates.empty)
print("✖ Items needing >30 min revision:", len(needs_revision))
print("✖ Items missing graphics sign-off:", len(graphics_issues))
print("✖ Skill/subskill mismatches:", len(mismatch_rows))
print("\n% of items ready per form:")
print(form_readiness)
