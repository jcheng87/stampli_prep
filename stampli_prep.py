import pandas as pd
import os
import numpy as np

dir_path = os.getcwd()
work_folder_path = os.path.join(dir_path,'work_folder')

'''
Docs required in work_folder:
1) Team COA
2) stampli file

function preps the report for distribution
'''
def prep_stampli_file():
    #column order file
    column_order = pd.read_csv(os.path.join(work_folder_path,'column_order.csv'))

    #team COA File
    team_COA_file_name = 'FP&A & GL team COA (1).xlsx'

    #Department Sheet
    dept_df = pd.read_excel(os.path.join(work_folder_path,team_COA_file_name),sheet_name='Dept', dtype={'Department Number':'object','Accrual Account':'object','Prepaid Account':'object'})
    dept_df.set_index('Department Number', inplace=True)

    #COA sheet
    account_df = pd.read_excel(os.path.join(work_folder_path,team_COA_file_name),sheet_name='COA',dtype={'Account':'object'})
    account_df.set_index('Account', inplace=True)

    #stampli file
    stampli_df = pd.read_csv(os.path.join(work_folder_path,'stampli_sampl.csv'))
    stampli_df = stampli_df.drop(0) #drop subtotal row
    stampli_df =stampli_df.drop('Number of Records', axis = 1) #drop redundant columns

    # additional columns
    column_ord_list = np.array(column_order['column_name'])
    stampli_col_list = np.array(stampli_df.columns)
    mask = np.isin(column_ord_list,stampli_col_list, invert=True)
    add_cols = column_ord_list[mask]
    for col in add_cols:
        stampli_df[col] = ''

    stampli_df.fillna('', inplace=True)

    # #fill cells with blank '' value for match to work
    # stampli_df['ACM PO Subaccount'].fillna('', inplace=True)
    # stampli_df['ACM Vendor Department'].fillna('', inplace=True)

    #this functions returns either the [8:12] char of ['ACM PO Subaccount'] or [0:4] char of the ['ACN Vendor Department']
    def dept_look(acm_po_dept, acm_vendor_dept):
        if acm_po_dept != '':
            return acm_po_dept[8:12]
        elif acm_vendor_dept != '' :
            return acm_vendor_dept[0:4]  

    #depending on type, looks up value against certain file docs
    def lookup_func(value, type):
        try:
            if type == 'GL Owner':
                return dept_df['GL Owner'].loc[int(value)]
            elif type == 'PO Dept Name':
                return dept_df['Department Name'].loc[int(value[8:12])]
            elif type == 'PO Account Name':
                return account_df['Description'].loc[int(value[0:6])]
        except:
            return ''
        
    #look ups
    stampli_df['dept_lookup'] = stampli_df.apply(lambda x: dept_look(x['ACM PO Subaccount'], x['ACM Vendor Department']), axis=1)
    stampli_df['GL Owner'] = stampli_df['dept_lookup'].apply(lambda x: lookup_func(x,'GL Owner'),)
    stampli_df['Dept Name per PO/PR'] = stampli_df['ACM PO Subaccount'].apply(lambda x: lookup_func(x,'PO Dept Name'))
    stampli_df['Account Description per PO/PR'] = stampli_df['ACM PO Account'].apply(lambda x: lookup_func(x,'PO Account Name'))

    stampli_df = stampli_df[column_ord_list]

    return stampli_df


def stampli_to_je(stampli_df):
    stampli_df['Transaction Description'] = (stampli_df['Line-Item Description'] + '//' +
                                    stampli_df['PO/PR #'] + '//' +
                                    stampli_df['Invoice #'] + '//'  + 
                                    'ACRL//' + 
                                    stampli_df['Service Period/Ship Date'] + '//' +
                                    stampli_df['Vendor'] + '//' + 
                                    'stampli:' + stampli_df['PK'])

    stampli_df['char_cnt: Transaction Description'] = stampli_df['Transaction Description'].apply(len)
                                    
    je_column = ['Account', 'Account Description', 'Subaccount','Debit Amount','Credit Amount','Transaction Description','char_cnt: Transaction Description', 'Link', 'Currency',
                'Line-Item Description','PO/PR #', 'Invoice #','Service Period/Ship Date','Vendor', 'PK']

    return stampli_df[je_column]