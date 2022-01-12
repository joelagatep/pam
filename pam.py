import pandas as pd
import warnings

def main():
    
    pd.set_option('display.max_columns', None)

    output_xls = r"pam_review.xlsx"

    fn_after = "PSG20211209"
    fn_bfore = "PSG20211117"

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        df_after = pd.read_excel(fn_after + ".xlsx", "Privileged SGs, Worker Accounts", header=0, engine="openpyxl")

    df_after.columns = [c.replace(' ', '_') for c in df_after.columns]
    df_after.columns = [c.replace(':', '_') for c in df_after.columns]
    df_after.columns = [c.replace('-', '_') for c in df_after.columns]

    # for colname in df_after.columns:
    #     print(colname)
    # Security Group
    # Members: Members
    # Members: Position
    # Members: Business Title
    # Members: Supervisory Organization
    # Members: Workday ID
    # Members: Workday Account
    # Members: Account Inactive
    # Assignable Role
    # User-Based
    # Manage Roles
    # Maintain Assign Role
    # Active Worker

    # Security_Group
    # Members__Members
    # Members__Position
    # Members__Business_Title
    # Members__Supervisory_Organization
    # Members__Workday_ID
    # Members__Workday_Account
    # Members__Account_Inactive
    # Assignable_Role
    # User_Based
    # Manage_Roles
    # Maintain_Assign_Role
    # Active_Worker
    
    df_after = df_after.sort_values(by=["Members__Members", \
                                        "Security_Group", \
                                        "Members__Position", \
                                        "Members__Business_Title", \
                                        "Members__Supervisory_Organization"])
    df_after = df_after[df_after[['Members__Workday_Account']].notnull().all(1)]  # drops nans
    df_after = df_after.loc[(df_after.User_Based == "YES") & \
                            #(df_after.Members__Members == "Hari Mailvaganam") & \
                            #(df_after.Members__Members == "Aarif Khan") | \
                            #(df_after.Members__Members == "Aaron Boley") & \
                            (~df_after.Members__Workday_Account.str.contains("ISU_")) & \
                            (~df_after.Members__Workday_Account.str.contains("wd-")) & \
                            (~df_after.Members__Workday_Account.str.contains("-impl")) & \
                            (df_after.Members__Account_Inactive == 0) & \
                            (df_after.Active_Worker == "Active")]

    df_after_1 = df_after.loc[:, ['Members__Workday_Account', \
                                  'Members__Members', \
                                  'Security_Group', \
                                  'Members__Position', \
                                  'Members__Business_Title', \
                                  'Members__Supervisory_Organization']]

    df_distinct_members_wd_acct = df_after_1["Members__Workday_Account"].unique()
    
    # Check if rows in df_after_1 is NOT ISC - if true then add this to df_non_isc dataframe
    df_all_non_isc = df_after_1.loc[(~df_after_1.Members__Supervisory_Organization.str.contains("Integrated Service Centre")) & \
                                    #(~df_after_1.Members__Supervisory_Organization.str.contains("Integrated Serviec Centre")) & \
                                    (~df_after_1.Members__Business_Title.str.contains("ISC"))]

    ########

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        df_bfore = pd.read_excel(fn_bfore + ".xlsx", "Privileged SGs, Worker Accounts", header=0, engine="openpyxl")

    df_bfore.columns = [c.replace(' ', '_') for c in df_bfore.columns]
    df_bfore.columns = [c.replace(':', '_') for c in df_bfore.columns]
    df_bfore.columns = [c.replace('-', '_') for c in df_bfore.columns]
    df_bfore = df_bfore.sort_values(by=["Members__Members", \
                                        "Security_Group", \
                                        "Members__Position", \
                                        "Members__Business_Title", \
                                        "Members__Supervisory_Organization"])
    df_bfore = df_bfore[df_bfore[['Members__Workday_Account']].notnull().all(1)]  # drops nans
    df_bfore = df_bfore.loc[(df_bfore.User_Based == "YES") & \
                            #(df_bfore.Members__Members == "Hari Mailvaganam") & \
                            #(df_bfore.Members__Members == "Aarif Khan") | \
                            #(df_bfore.Members__Members == "Aaron Boley") & \
                            (~df_bfore.Members__Workday_Account.str.contains("ISU_")) & \
                            (~df_bfore.Members__Workday_Account.str.contains("wd-")) & \
                            (~df_bfore.Members__Workday_Account.str.contains("-impl")) & \
                            (df_bfore.Members__Account_Inactive == 0) & \
                            (df_bfore.Active_Worker == "Active")]
    df_bfore_1 = df_bfore.loc[:, ['Members__Workday_Account', \
                                  'Members__Members', \
                                  'Security_Group', \
                                  'Members__Position', \
                                  'Members__Business_Title', \
                                  'Members__Supervisory_Organization']]

    #df_new_add = pd.DataFrame()
    df_diff = pd.DataFrame()
    df_diff_old = pd.DataFrame()
    df_diff_new = pd.DataFrame()

    #orow = 0
    for key in df_distinct_members_wd_acct:
        #print(key)
        #ws_out.write(orow, 0, key, style0)
        #orow += 1
        df_compare = pd.DataFrame()
        df_after_2 = df_after_1.loc[(df_after_1.Members__Workday_Account == key)]
        df_after_2 = df_after_2.reset_index(drop=True)
        df_bfore_2 = df_bfore_1.loc[(df_bfore_1.Members__Workday_Account == key)]
        df_bfore_2 = df_bfore_2.reset_index(drop=True)
        # Check if df_bfore is empty.  If it is then it means this is a new employee
        #if df_bfore_2.empty:
        #    df_new_add = df_new_add.append(df_after_2)
        #else:
        #df_compare = df_after_2[df_bfore_2.ne(df_after_2).any(axis=1)]
        
        # Search for after rows that are not in, or different from bfore
        df_compare = df_after_2 # We will delete rows that have no differences
        for inda in df_after_2.index:
            Found = False
            for indb in df_bfore_2.index:
                if ((df_bfore_2['Security_Group'][indb] == df_after_2['Security_Group'][inda]) & \
                    (df_bfore_2['Members__Position'][indb] == df_after_2['Members__Position'][inda]) & \
                    (df_bfore_2['Members__Business_Title'][indb] == df_after_2['Members__Business_Title'][inda])):
                    #(df_bfore_2['Members__Supervisory_Organization'][indb] == df_after_2['Members__Supervisory_Organization'][inda])):
                    Found = True
                    break
            if Found:
                # Exact matching row found - remove from diff
                df_compare = df_compare.drop(axis=1, index=inda)
        if not df_compare.empty:
            # This means there are deltas for this person between before and after
            df_compare = df_compare.reset_index(drop=True)
            df_compare = df_compare.sort_values(by=["Members__Members", \
                                "Security_Group", \
                                "Members__Position", \
                                "Members__Business_Title", \
                                "Members__Supervisory_Organization"])
            df_diff = df_diff.append(df_compare)
            df_diff_old = df_diff_old.append(df_bfore_2)
            df_diff_new = df_diff_new.append(df_after_2)

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
    #if not df_new_add.empty:
    #    df_new_add = df_new_add.drop(labels='Members__Workday_Account', axis=1)
    #    df_new_add.rename(columns={'Members__Members':'Member', \
    #                  'Security_Group':'Security Group', \
    #                  'Members__Position':'Position', \
    #                  'Members__Business_Title':'Business Title', \
    #                  'Members__Supervisory_Organization':'Sup Org'}, \
    #                  inplace=True)
    if not df_all_non_isc.empty:
        df_all_non_isc = df_all_non_isc.drop(labels='Members__Workday_Account', axis=1)
        df_all_non_isc.rename(columns={'Members__Members':'Member', \
                      'Security_Group':'Security Group', \
                      'Members__Position':'Position', \
                      'Members__Business_Title':'Business Title', \
                      'Members__Supervisory_Organization':'Sup Org'}, \
                      inplace=True)

    #df_final.to_excel('pam_output.xls', index=False, header=False)
    writer = pd.ExcelWriter(output_xls, engine = 'xlsxwriter')

    df_diff.to_excel(writer, sheet_name = 'Deltas', index=False, header=False, startrow=1)
    for column in df_diff:
        column_width = max(df_diff[column].astype(str).map(len).max(), len(column))
        col_idx = df_diff.columns.get_loc(column)
        writer.sheets['Deltas'].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_diff.columns]
    (max_row, max_col) = df_diff.shape
    worksheet1 = writer.sheets['Deltas']
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)

    df_diff_old.to_excel(writer, sheet_name = fn_bfore, index=False, header=False, startrow=1)
    for column in df_diff_old:
        column_width = max(df_diff_old[column].astype(str).map(len).max(), len(column))
        col_idx = df_diff_new.columns.get_loc(column)
        writer.sheets[fn_bfore].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_diff_old.columns]
    (max_row, max_col) = df_diff_old.shape
    worksheet1 = writer.sheets[fn_bfore]
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)

    df_diff_new.to_excel(writer, sheet_name = fn_after, index=False, header=False, startrow=1)
    for column in df_diff_new:
        column_width = max(df_diff_new[column].astype(str).map(len).max(), len(column))
        col_idx = df_diff_old.columns.get_loc(column)
        writer.sheets[fn_after].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_diff_new.columns]
    (max_row, max_col) = df_diff_new.shape
    worksheet1 = writer.sheets[fn_after]
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)

    #df_new_add.to_excel(writer, sheet_name = 'New Additions', index=False, header=False, startrow=1)
    #for column in df_new_add:
    #    column_width = max(df_new_add[column].astype(str).map(len).max(), len(column))
    #    col_idx = df_new_add.columns.get_loc(column)
    #    writer.sheets['New Additions'].set_column(col_idx, col_idx, column_width)    
    #column_settings2 = [{'header': column} for column in df_new_add.columns]
    #(max_row, max_col) = df_new_add.shape
    #worksheet2 = writer.sheets['New Additions']
    #worksheet2.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings2})
    #worksheet2.freeze_panes(1, 1)
    
    df_all_non_isc.to_excel(writer, sheet_name = 'All Non-ISC', index=False, header=False, startrow=1)
    for column in df_all_non_isc:
        column_width = max(df_all_non_isc[column].astype(str).map(len).max(), len(column))
        col_idx = df_all_non_isc.columns.get_loc(column)
        writer.sheets['All Non-ISC'].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_all_non_isc.columns]
    (max_row, max_col) = df_all_non_isc.shape
    worksheet1 = writer.sheets['All Non-ISC']
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)

    writer.save()
    #writer.close()


if __name__ == '__main__':
    main()


