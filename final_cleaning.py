import os
import pandas as pd



def process_files(directory_path):
    def P11_clean_cashbook(filepath):
        # Check the file extension
        if filepath.lower().endswith('.xls'):
            cashbook = pd.read_html(filepath)[1]
            #first_row = cashbook[cashbook[22].str.contains('Line Description', na=False)].index[0]
            #cashbook.columns = cashbook.iloc[first_row]
            cashbook.rename(columns={'Document Date':'Date'},inplace=True)
            cashbook=cashbook.rename(columns={'Reference Number':'VoucherN°'})
            cashbook=cashbook.rename(columns={'Narration':'Description'})
            cashbook=cashbook.rename(columns={'Debits':'Un credited Items'})
            cashbook=cashbook.rename(columns={'Credits':'Un Paid Items'})
            cashbook=cashbook.rename(columns={'Balance amount':'Running Total'})

            cashbook['Cheque N°']=''
            #cashbook = cashbook[first_row+1:]
            #cashbook.columns.name = None
            cashbook.reset_index(drop=True, inplace=True)
            cashbook.dropna(subset='Date',inplace=True)

        else:
            cashbook = pd.read_excel(filepath)
            cashbook.rename(columns={'Document Date':'Date'},inplace=True)
            cashbook=cashbook.rename(columns={'Reference Number':'VoucherN°'})
            cashbook=cashbook.rename(columns={'Narration':'Description'})
            cashbook=cashbook.rename(columns={'Debits':'Un credited Items'})
            cashbook=cashbook.rename(columns={'Credits':'Un Paid Items'})
            #cashbook = cashbook[first_row+1:]
            cashbook['Cheque N°']=''
            cashbook=cashbook.rename(columns={'Balance amount':'Running Total'})
            cashbook.dropna(subset='Date',inplace=True)


        return cashbook
    
    def clean_cashbook(filepath):
        # Check the file extension
        if filepath.lower().endswith('.xls'):
            cashbook = pd.read_html(filepath)[1]
            first_row = cashbook[cashbook[22].str.contains('Line Description', na=False)].index[0]
            cashbook.columns = cashbook.iloc[first_row]
            cashbook = cashbook[first_row+1:]
            cashbook.columns.name = None
            cashbook.reset_index(drop=True, inplace=True)
        else:
            cashbook = pd.read_excel(filepath)
            first_row = cashbook[cashbook["Unnamed: 22"].str.contains('Line Description', na=False)].index[0]
            cashbook.columns = cashbook.iloc[first_row]
            cashbook = cashbook[first_row+1:]
            cashbook.reset_index(drop=True, inplace=True)
            cashbook.columns.name = None

        return cashbook

    def clean_dfcu_bank_statement(filepath):
        df_bank = pd.read_excel(filepath)
        first_row = df_bank[df_bank['Unnamed: 3'].str.contains('Description', na=False)].index[0]
        df_bank.columns = df_bank.iloc[first_row]
        df_bank = df_bank[first_row+1:]
        df_bank.reset_index(drop=True, inplace=True)
        df_bank.columns.name = None
        df_bank.rename(columns={'Transaction date':'Transaction Date'},inplace=True)
        df_bank.rename(columns={'Debit Value':'Debits'},inplace=True)
        df_bank['Debits'] = pd.to_numeric(df_bank['Debits'], errors='coerce').abs()
        df_bank.rename(columns={'Credit Value':'Credits'},inplace=True)
        df_bank.rename(columns={'Balance':'Running Balance'},inplace=True)
        df_bank.rename(columns={'Description':'Transaction Details'},inplace=True)
        df_bank.rename(columns={'[DATALIST:Custom]':'Transaction Type'},inplace=True)
        df_bank.rename(columns={'Trans. Date]':'Value Date'},inplace=True)
        return df_bank
    
    def clean_absa_bank_statement(filepath):
        df_bank = pd.read_excel(filepath)
        first_row = df_bank[df_bank["Unnamed: 1"].str.contains('Value date', na=False)].index[0]
        df_bank.columns = df_bank.iloc[first_row]
        df_bank = df_bank[first_row+1:]
        df_bank.reset_index(drop=True, inplace=True)
        df_bank.columns.name = None
        df_bank.rename(columns={'Transaction date':'Transaction Date'},inplace=True)
        df_bank.rename(columns={'Cheque number':'Cheque Number'},inplace=True)
        df_bank.rename(columns={'Value date':'Value Date'},inplace=True)
        df_bank.rename(columns={'Debit amount':'Debits'},inplace=True)           
        df_bank.rename(columns={'Credit amount':'Credits'},inplace=True)
        df_bank.rename(columns={'Running balance':'Running Balance'},inplace=True)
        df_bank.rename(columns={'Customer reference':'Transaction Details'},inplace=True)
        df_bank.rename(columns={'Transaction Reference Number':'Reference'},inplace=True)
        df_bank.rename(columns={'Description':'Transaction Type'},inplace=True)
        return df_bank

    def clean_scb_bank_statement(filepath):
        df_bank = pd.read_excel(filepath)
        first_row = df_bank[df_bank["Unnamed: 8"].str.contains('Date', na=False)].index[0]
        df_bank.columns = df_bank.iloc[first_row]
        df_bank = df_bank[first_row+1:]
        df_bank.reset_index(drop=True, inplace=True)
        df_bank.columns.name = None
        df_bank.columns = df_bank.columns.str.strip()
        df_bank['Transaction Date']=''
        df_bank['Cheque Number']=''
        #df_bank.rename(columns={'Transaction date':'Transaction Date'},inplace=True)
        #df_bank.rename(columns={'Cheque number':'Cheque Number'},inplace=True)

        df_bank.rename(columns={'Date':'Value Date'},inplace=True)
        df_bank.rename(columns={'Withdrawal':'Debits'},inplace=True)           
        df_bank.rename(columns={'Deposit':'Credits'},inplace=True)
        df_bank.rename(columns={'Balance':'Running Balance'},inplace=True)
        df_bank.rename(columns={'Account Name':'Transaction Details'},inplace=True)
        df_bank.rename(columns={'Account Number':'Reference'},inplace=True)

        df_bank.rename(columns={'Description':'Transaction Type'},inplace=True)
        #df_bank.drop(columns=['NaN'], inplace=True)
        df_bank=df_bank[['Transaction Date','Reference','Transaction Details','Address','Currency','Value Date', 'Transaction Type', 'Debits', 'Credits', 'Running Balance', 'Transaction Date', 'Cheque Number']]
        return df_bank
    
    def clean_kcb_bank_statement(filepath):
        df_bank = pd.read_excel(filepath)
        first_row = df_bank[df_bank["Unnamed: 1"].str.contains('Value Date', na=False)].index[0]
        df_bank.columns = df_bank.iloc[first_row]
        df_bank = df_bank[first_row+1:]
        df_bank.reset_index(drop=True, inplace=True)
        df_bank.columns.name = None
        df_bank.columns = df_bank.columns.str.strip()
        df_bank['Cheque Number']=''
        #df_bank.rename(columns={'Transaction date':'Transaction Date'},inplace=True)
        #df_bank.rename(columns={'Cheque number':'Cheque Number'},inplace=True)

        #df_bank.rename(columns={'Date':'Value Date'},inplace=True)
        df_bank.rename(columns={'Money Out':'Debits'},inplace=True)           
        df_bank.rename(columns={'Money In':'Credits'},inplace=True)
        df_bank.rename(columns={'Ledger Balance':'Running Balance'},inplace=True)
        #df_bank.rename(columns={'Account Name':'Transaction Details'},inplace=True)
        df_bank.rename(columns={'Bank Reference Number':'Reference'},inplace=True)
        df_bank['Transaction Type']=''

        #df_bank.rename(columns={'Description':'Transaction Type'},inplace=True)
        return df_bank
    
    def clean_ncba_bank_statement(filepath):
        df_bank = pd.read_excel(filepath)
        first_row = df_bank[df_bank["Unnamed: 1"].str.contains('Value Date', na=False)].index[0]
        df_bank.columns = df_bank.iloc[first_row]
        df_bank = df_bank[first_row+1:]
        df_bank.reset_index(drop=True, inplace=True)
        df_bank.columns.name = None
        df_bank.columns = df_bank.columns.str.strip()
        df_bank['Cheque Number']=''
        df_bank.rename(columns={'Transaction date':'Transaction Date'},inplace=True)
        #df_bank.rename(columns={'Cheque number':'Cheque Number'},inplace=True)

        #df_bank.rename(columns={'Date':'Value Date'},inplace=True)
        df_bank.rename(columns={'Debit':'Debits'},inplace=True)           
        df_bank.rename(columns={'Credit':'Credits'},inplace=True)
        df_bank.rename(columns={'Balance':'Running Balance'},inplace=True)
        df_bank.rename(columns={'Description':'Transaction Details'},inplace=True)
        df_bank.rename(columns={'Reference Number':'Reference'},inplace=True)
        #df_bank['Transaction Type']=''

        #df_bank.rename(columns={'Description':'Transaction Type'},inplace=True)

        return df_bank


    def clean_previous(filepath):
        earlier_workings=pd.read_excel(filepath)
        first_row = earlier_workings[earlier_workings["Unnamed: 8"].str.contains('Matching', na=False)].index[0]
        earlier_workings.columns = earlier_workings.iloc[first_row]
        earlier_workings = earlier_workings[first_row+1:]
        earlier_workings.reset_index(drop=True, inplace=True)
        earlier_workings.columns.name = None
        earlier_workings.columns=['Date','VoucherN°','Cheque N°','Description',' Direct Debits','Un receipted Items','Un credited Items','Un Paid Items','Matching']         
        return earlier_workings






    try:
        # Create a new directory for the clean files
        clean_directory = os.path.join(directory_path, 'clean')
        os.makedirs(clean_directory, exist_ok=True)

        # Get the list of files in the directory
        files = [f for f in os.listdir(directory_path) if f.lower().endswith(('.xlsx', '.xls'))]

        # Iterate over the files
        for file in files:
            if 'Cashbook' in file:
                try:
                    # Clean cashbook
                    cashbook = clean_cashbook(os.path.join(directory_path, file))

                    # Save the clean cashbook in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    cashbook.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing cashbook file:", file)
                    print("Error message:", str(e))


            ###P11
            elif 'P11' in file:
                try:
                    # Clean cashbook
                    cashbook = P11_clean_cashbook(os.path.join(directory_path, file))

                    # Save the clean cashbook in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    cashbook.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing cashbook file:", file)
                    print("Error message:", str(e))



            ###P11

            elif 'DFCU' in file and 'Bank' in file:
                try:
                    # Clean DFCU bank statement
                    df_bank = clean_dfcu_bank_statement(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))

            elif 'ABSA' in file and 'Bank' in file:
                try:
                    # Clean DFCU bank statement
                    df_bank = clean_absa_bank_statement(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))
                    
            elif 'KCB' in file and 'Bank' in file:
                try:
                    # Clean DFCU bank statement
                    df_bank = clean_kcb_bank_statement(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))
                    
            elif 'NCBA' in file and 'Bank' in file:
                try:
                    # Clean DFCU bank statement
                    df_bank = clean_ncba_bank_statement(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))
                    
            elif 'SCB' in file and 'Bank' in file:
                try:
                    # Clean DFCU bank statement
                    df_bank = clean_scb_bank_statement(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))
            

            elif 'Previous' in file:
                try:
                    earlier_workings = clean_previous(os.path.join(directory_path, file))

                    # Save the clean DFCU bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    earlier_workings.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU bank statement file:", file)
                    print("Error message:", str(e))
                    
            elif 'DFCU' in file and 'Previous' in file and 'Bank' not in file:
                try:
                    # Clean DFCU UGX Previous
                    earlier_workings = clean_previous(os.path.join(directory_path, file))

                    # Save the clean DFCU UGX Previous in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    earlier_workings.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing DFCU UGX Previous file:", file)
                    print("Error message:", str(e))

            else:
                # If the file does not contain 'Cashbook' and 'DFCU', clean the bank statement
                try:
                    df_bank = pd.read_excel(os.path.join(directory_path, file))
                    first_row = df_bank[df_bank["Unnamed: 1"].str.contains('Value Date', na=False)].index[0]
                    df_bank.columns = df_bank.iloc[first_row]
                    df_bank = df_bank[first_row+1:]
                    df_bank.reset_index(drop=True, inplace=True)
                    df_bank.columns.name = None

                    if 'stanbic' in file.lower() and 'cashbook' not in file.lower():
                                df_bank.rename(columns={'Debit':'Debits'},inplace=True)           
                                df_bank.rename(columns={'Credit':'Credits'},inplace=True)
                                df_bank.rename(columns={'Balance':'Running Balance'},inplace=True)
                                df_bank.rename(columns={'Transaction Description':'Transaction Details'},inplace=True)
                                df_bank.rename(columns={'Type':'Transaction Type'},inplace=True)

                        

                    # Save the clean bank statement in the clean directory with xlsx extension
                    clean_filepath = os.path.join(clean_directory, os.path.splitext(file)[0] + '.xlsx')
                    df_bank.to_excel(clean_filepath, index=False)

                except Exception as e:
                    print("An error occurred while processing bank statement file:", file)
                    print("Error message:", str(e))

    except Exception as e:
        print("An error occurred:", str(e))

# Example usage of the process_files function
directory_path =r"C:\Users\cynthia.mutisya\Downloads\Data (1)\Data"
process_files(directory_path)
