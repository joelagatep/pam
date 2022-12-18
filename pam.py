# Moving to ipynb
import pandas as pd
import warnings
import altair as alt

fn_bfore = 'PSG20220713'
fn_after = 'PSG20220822'

print('Starting ...')

if fn_bfore >= fn_after:
    raise Exception("The fn_bfore must contain the name of the BEFORE file.  Please correct before restarting.")

pd.set_option('display.max_columns', None)

output_xls = r"pam_review.xlsx"

# fn_bfore = 'PSG20220209'
# fn_after = 'PSG20220308'

with warnings.catch_warnings(record=True):
    warnings.simplefilter('always')
    after_df = pd.read_excel(fn_after + '.xlsx', 'Privileged SGs, Worker Accounts', header=0, engine='openpyxl')

after_df = after_df.fillna('<blank>')
after_df.columns = [c.replace(' ', '_') for c in after_df.columns]
after_df.columns = [c.replace(':', '_') for c in after_df.columns]
after_df.columns = [c.replace('-', '_') for c in after_df.columns]
    
after_df = after_df.sort_values(by=['Members__Members', \
                                    'Security_Group', \
                                    'Members__Position', \
                                    'Members__Business_Title', \
                                    'Members__Supervisory_Organization'])
after_df = after_df[after_df[['Members__Workday_Account']].notnull().all(1)]  # drops nans
after_df = after_df.loc[(after_df.User_Based == 'YES') & \
                        (~after_df.Members__Workday_Account.str.contains('ISU_')) & \
                        (~after_df.Members__Workday_Account.str.contains('wd-')) & \
                        (~after_df.Members__Workday_Account.str.contains('-impl')) & \
                        (after_df.Members__Account_Inactive == 0) & \
                        (after_df.Active_Worker == 'Active')]

after_df_1 = after_df.loc[:, ['Members__Workday_Account', \
                              'Members__Members', \
                              'Security_Group', \
                              'Members__Position', \
                              'Members__Business_Title', \
                              'Members__Supervisory_Organization']]

df_distinct_members_wd_acct = after_df_1['Members__Workday_Account'].unique()
    
# Check if rows in after_df_1 is NOT ISC - if true then add this to df_non_isc dataframe
# June 2022 do no include any that contains:  *OCIO | Program Delivery*, *OCIO | Integration Enablement Center*,
#                                             *Risk Assurance*, *Operational Excellence*
df_all_non_isc = after_df_1.loc[(~after_df_1.Members__Supervisory_Organization.str.contains('Integrated Service Centre')) & \
                                (~after_df_1.Members__Supervisory_Organization.str.contains('OCIO | Program Delivery')) & \
                                (~after_df_1.Members__Supervisory_Organization.str.contains('OCIO | Integration Enablement Center')) & \
                                (~after_df_1.Members__Supervisory_Organization.str.contains('Risk Assurance')) & \
                                (~after_df_1.Members__Supervisory_Organization.str.contains('Operational Excellence')) & \
                                (~after_df_1.Members__Business_Title.str.contains('ISC'))]

with warnings.catch_warnings(record=True):
    warnings.simplefilter('always')
    bfore_df = pd.read_excel(fn_bfore + '.xlsx', 'Privileged SGs, Worker Accounts', header=0, engine='openpyxl')

bfore_df = bfore_df.fillna('<blank>')
bfore_df.columns = [c.replace(' ', '_') for c in bfore_df.columns]
bfore_df.columns = [c.replace(':', '_') for c in bfore_df.columns]
bfore_df.columns = [c.replace('-', '_') for c in bfore_df.columns]
bfore_df = bfore_df.sort_values(by=['Members__Members', \
                                    'Security_Group', \
                                    'Members__Position', \
                                    'Members__Business_Title', \
                                    'Members__Supervisory_Organization'])
bfore_df = bfore_df[bfore_df[['Members__Workday_Account']].notnull().all(1)]  # drops nans
bfore_df = bfore_df.loc[(bfore_df.User_Based == 'YES') & \
                        (~bfore_df.Members__Workday_Account.str.contains('ISU_')) & \
                        (~bfore_df.Members__Workday_Account.str.contains('wd-')) & \
                        (~bfore_df.Members__Workday_Account.str.contains('-impl')) & \
                        (bfore_df.Members__Account_Inactive == 0) & \
                        (bfore_df.Active_Worker == 'Active')]
bfore_df_1 = bfore_df.loc[:, ['Members__Workday_Account', \
                              'Members__Members', \
                              'Security_Group', \
                              'Members__Position', \
                              'Members__Business_Title', \
                              'Members__Supervisory_Organization']]

#df_new_add = pd.DataFrame()
df_diff = pd.DataFrame()
df_diff_old = pd.DataFrame()
df_diff_new = pd.DataFrame()

for key in df_distinct_members_wd_acct:
    df_compare = pd.DataFrame()
    after_df_2 = after_df_1.loc[(after_df_1.Members__Workday_Account == key)]
    after_df_2 = after_df_2.reset_index(drop=True)
    bfore_df_2 = bfore_df_1.loc[(bfore_df_1.Members__Workday_Account == key)]
    bfore_df_2 = bfore_df_2.reset_index(drop=True)
        
    # Search for after rows that are not in, or different from bfore
    df_compare = after_df_2 # We will delete rows that have no differences
    for inda in after_df_2.index:
        Found = False
        for indb in bfore_df_2.index:
            if ((bfore_df_2['Security_Group'][indb] == after_df_2['Security_Group'][inda]) & \
                (bfore_df_2['Members__Position'][indb] == after_df_2['Members__Position'][inda]) & \
                (bfore_df_2['Members__Business_Title'][indb] == after_df_2['Members__Business_Title'][inda])):
                Found = True
                break
        if Found:
            # Exact matching row found - remove from diff
            df_compare = df_compare.drop(axis=1, index=inda)
    if not df_compare.empty:
        # This means there are deltas for this person between before and after
        df_compare = df_compare.reset_index(drop=True)
        df_compare = df_compare.sort_values(by=['Members__Members', \
                            'Security_Group', \
                            'Members__Position', \
                            'Members__Business_Title', \
                            'Members__Supervisory_Organization'])
        df_diff = df_diff.append(df_compare)
        df_diff_old = df_diff_old.append(bfore_df_2)
        df_diff_new = df_diff_new.append(after_df_2)

if not df_diff.empty:
    df_diff = df_diff.drop(labels='Members__Workday_Account', axis=1)
    df_diff.rename(columns={'Members__Members':'Member', \
                  'Security_Group':'Security Group', \
                  'Members__Position':'Position (New)', \
                  'Members__Business_Title':'Business Title (New)', \
                  'Members__Supervisory_Organization':'Sup Org (New)'}, \
                  inplace=True)
if not df_diff_old.empty:
    df_diff_old = df_diff_old.drop(labels='Members__Workday_Account', axis=1)
    df_diff_old.rename(columns={'Members__Members':'Member', \
                  'Security_Group':'Security Group', \
                  'Members__Position':'Position', \
                  'Members__Business_Title':'Business Title', \
                  'Members__Supervisory_Organization':'Sup Org'}, \
                  inplace=True)
if not df_diff_new.empty:
    df_diff_new = df_diff_new.drop(labels='Members__Workday_Account', axis=1)
    df_diff_new.rename(columns={'Members__Members':'Member', \
                  'Security_Group':'Security Group', \
                  'Members__Position':'Position', \
                  'Members__Business_Title':'Business Title', \
                  'Members__Supervisory_Organization':'Sup Org'}, \
                  inplace=True)
if not df_all_non_isc.empty:
    df_all_non_isc = df_all_non_isc.drop(labels='Members__Workday_Account', axis=1)
    df_all_non_isc.rename(columns={'Members__Members':'Member', \
                  'Security_Group':'Security Group', \
                  'Members__Position':'Position', \
                  'Members__Business_Title':'Business Title', \
                  'Members__Supervisory_Organization':'Sup Org'}, \
                  inplace=True)

writer = pd.ExcelWriter(output_xls, engine = 'xlsxwriter')

def prep_sheet(df1, sheet1):
    df1.to_excel(writer, sheet_name = sheet1, index=False, header=False, startrow=1)
    for column in df1:
        column_width = max(df1[column].astype(str).map(len).max(), len(column))
        col_idx = df1.columns.get_loc(column)
        writer.sheets[sheet1].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df1.columns]
    (max_row, max_col) = df1.shape
    worksheet1 = writer.sheets[sheet1]
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)
    
prep_sheet(df_diff, 'Deltas')
prep_sheet(df_diff_old, fn_bfore)
prep_sheet(df_diff_new, fn_after)
prep_sheet(df_all_non_isc, 'All Non-ISC')

writer.save()

print('PAM Review Report Generated - Please open pam_review.xlsx')

# Prepare tiny bar chart data

print('Summary chart below...')
    
df_tc = pd.DataFrame()
    
row_count = bfore_df[bfore_df.columns[1]].count()
row_list = [[fn_bfore, row_count]]
df_tc = df_tc.append(pd.DataFrame(row_list, columns=['Data','Rows']),ignore_index=True)

row_count = after_df[after_df.columns[1]].count()
row_list = [[fn_after, row_count]]
df_tc = df_tc.append(pd.DataFrame(row_list, columns=['Data','Rows']),ignore_index=True)

row_count = df_diff[df_diff.columns[1]].count()
row_list = [['Deltas', row_count]]
df_tc = df_tc.append(pd.DataFrame(row_list, columns=['Data','Rows']),ignore_index=True)
    
row_count = df_all_non_isc[df_all_non_isc.columns[1]].count()
row_list = [['All Non-ISC', row_count]]
df_tc = df_tc.append(pd.DataFrame(row_list, columns=['Data','Rows']),ignore_index=True)
        
# Generate tiny bar chart

bars = alt.Chart(df_tc).mark_bar(color='lightgrey').encode(
    x=alt.X('Rows', type='quantitative', sort=None),
    y=alt.Y('Data', type='nominal', sort=None)
).properties(height=100, width=600)

text = bars.mark_text(
    align='left',
    baseline='middle',
    dx=3  # Nudges text to right so it doesn't appear on top of the bar
).encode(
    text='Rows'
)
    
bars + text
