# Standard library imports
from datetime import date, datetime, timedelta
import os
import re
import shutil

# Third-party imports
from dateutil.relativedelta import relativedelta
import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# Configuration
pd.options.mode.chained_assignment = None

# Constants
US_CAL = USFederalHolidayCalendar()
TODAY = datetime(2023, 11, 1)
TEMPLATE_FILE = r'\\CIBG-SRV-TOR08\dpss\ged_applications\beta\Structured Notes\Template\Blank calc file.xlsx'
INPUT_DIRECTORY = r'\\CIBG-SRV-TOR08\dpss\ged_applications\beta\Structured Notes\Input'
OUTPUT_DIRECTORY = r'\\CIBG-SRV-TOR08\dpss\ged_applications\beta\Structured Notes\Output'
INPUT_FILE = os.path.join(INPUT_DIRECTORY, 'ValuationDateExtract_2023-11-01.xlsx')
USHOLIDAYS = US_CAL.holidays(start=f'{TODAY.year}-01-01', end=f'{TODAY.year}-12-31').to_pydatetime()
CURRENT_MONTH_START = datetime(year=TODAY.year, month=TODAY.month, day=1)
NEXT_MONTH_START = CURRENT_MONTH_START + relativedelta(months=1)
CURRENT_MONTH_LIST = pd.date_range(CURRENT_MONTH_START, NEXT_MONTH_START - timedelta(days=1)).to_pydatetime().tolist()

def process_dataframe(df, tickerlist):
    # ... [same as your original function]

def process_worksheet(ws, df, refassetlist, vallist, settlelist, list_, notional, memory, inv, ident, auto, tpayingint, interest):
    # ... [same as your original function]

def main():
    print('Obtaining Structures Data')
    main_df = pd.read_excel(INPUT_FILE, sheet_name='Structures')
    print('Obtaining Notional Data')
    notional_df = pd.read_excel(INPUT_FILE, sheet_name='Positional')
    print('Obtaining Ticker Data')
    ticker_df = pd.read_excel(os.path.join(INPUT_DIRECTORY, 'tickers.xlsx'))
    combined_df = main_df.join(notional_df.set_index('Package Code'), on='Package Code')
    combined_df.drop_duplicates(subset=['Package Code', 'Observation Date', 'Settlement Date', 'Notional'], inplace=True)
    ticker_list = ticker_df['ticker'].str.rsplit('.', n=1, expand=True)[0]

    test_counter = 0
    for day in CURRENT_MONTH_LIST:
        
        # print(day)
        if (day not in usholidays) & (day.isoweekday() != 6) & (day.isoweekday() !=7):
            fileloc = outputdirectory+rf'\{datetime.strftime(day,"%b %#d")}.xlsx'
            if testcounter == 40:
                break
            try:
                os.remove( outputdirectory+rf'\{datetime.strftime(day,"%b %#d")}.xlsx')
    
                os.remove( outputdirectory+rf'\cusip_draft_'+rf'{datetime.strftime(day,"%b %#d")}.csv')
                os.remove( outputdirectory+rf'\isin_draft_'+rf'{datetime.strftime(day,"%b %#d")}.csv')
                os.remove( outputdirectory+rf'\df_'+rf'{datetime.strftime(day,"%b %#d")}.csv')
            except:
                pass
            
            if os.path.exists(fileloc) == False:
                shutil.copy(templatefile,fileloc)
                print(day,'file created.')
            else:
                print(day,'file exists.')   
                
               
            daydf = combineddf.loc[combineddf['Observation Date']==datetime.strftime(day,'%Y-%m-%d')]
            if daydf.empty == False:
                daydf['Long Name'] = daydf['Long Name'].fillna('None')
                        # Adding ISM Code to CUSIP if Structure is Callable
                daydf['CUSIP']=daydf[['CUSIP','ISM Code']].add('\n').sum(axis=1).where(daydf['Long Name'].str.contains('Issuer',case=False),other=daydf['CUSIP']).str.rstrip()
    
                
                daydf["Autocall Field"] = ""
                # daydf["Paying Interest"] = "=IF(CLOSING VALUE>=BARRIER VALUE,TRUE,FALSE)"
                daydf.loc[daydf['Settlement Date']==daydf['Structure Maturity'],'Autocall Field'] = 'Maturity'
                daydf.loc[daydf['downstrikePayout']==100,'Autocall Field'] = '=IF(CLOSING VALUE >= CALL THRESHOLD VALUE, TRUE FALSE)'
                daydf.loc[((daydf['downstrikePayout']== 0) | (daydf['Long Name'].str.contains('Blackrock', case=False))),'Autocall Field'] = 'N/A'
                daydf.loc[daydf['Long Name'].str.contains('issuer', case=False),'Autocall Field'] = 'Valid IC Date'
                daydf["Interest"] = ""        
            
                cusipdf = daydf[daydf['CUSIP'].str[:2] == '89']
                cusipdf = cusipdf.reset_index(drop=True)
                isindf = daydf[daydf['CUSIP'].str[:2] != '89']
                isindf = isindf.reset_index(drop=True)
                
                
            def process_dataframe(df, tickerlist):
                refassetlist = []
                finaldaylist = []
                vallist = []
                settlelist = []
                df_list = []  # Generic name to handle both cusip and isin
                notional = []
                memory = []
                inv = []
                ident = []
                auto = []
                tpayingint = []
                interest = []
                longname = []
    
                tickers = df['Long Name'].apply(lambda x: set.intersection(set(re.split('[ /]',x)), set(tickerlist)))
                tickcounter = 0
            
                for i in range(df.shape[0]):
                    df.at[i, 'Paying Interest'] = f"=IF(AA{i+1}>=Y{i+1},TRUE,FALSE)"
            
                for tickers_set in tickers:
                    tickers_list = list(tickers_set)
                    if len(tickers_list) == 0:
                        refassetlist.append('NoTickerFound')
                        vallist.append(datetime.strftime(df['Observation Date'][tickcounter], '%m/%d/%Y'))
                        settlelist.append(datetime.strftime(df['Settlement Date'][tickcounter], '%m/%d/%Y'))
                        df_list.append(df['CUSIP'][tickcounter])
                        notional.append(df['Notional'][tickcounter])
                        inv.append(df['inventoryName'][tickcounter])
                        ident.append(df['ident'][tickcounter])
                        finaldaylist.append([datetime.strftime(df['Observation Date'][tickcounter], '%m/%d/%Y'),
                                             datetime.strftime(df['Settlement Date'][tickcounter], '%m/%d/%Y'),
                                             df['CUSIP'][tickcounter], 'NoTickerFound'])
                        auto.append(df['Autocall Field'][tickcounter])
                        tpayingint.append(df['Paying Interest'][tickcounter])
                        interest.append(df['Interest'][tickcounter])
                        longname.append(df['Long Name'][tickcounter])
            
                        memory_val = 'Refer to previous month' if 'memory' in df['Long Name'][tickcounter].lower() else ''
                        memory.append(memory_val)
                    else:
                        for ticker in tickers_list:
                            if ticker != 'TD':
                                refassetlist.append(ticker)
                                vallist.append(datetime.strftime(df['Observation Date'][tickcounter], '%m/%d/%Y'))
                                settlelist.append(datetime.strftime(df['Settlement Date'][tickcounter], '%m/%d/%Y'))
                                df_list.append(df['CUSIP'][tickcounter])
                                notional.append(df['Notional'][tickcounter])
                                inv.append(df['inventoryName'][tickcounter])
                                ident.append(df['ident'][tickcounter])
                                finaldaylist.append([datetime.strftime(df['Observation Date'][tickcounter], '%m/%d/%Y'),
                                                     datetime.strftime(df['Settlement Date'][tickcounter], '%m/%d/%Y'),
                                                     df['CUSIP'][tickcounter], ticker])
                                auto.append(df['Autocall Field'][tickcounter])
                                tpayingint.append(df['Paying Interest'][tickcounter])
                                interest.append(df['Interest'][tickcounter])
                                longname.append(df['Long Name'][tickcounter])
            
                                memory_val = 'Refer to previous month' if 'memory' in df['Long Name'][tickcounter].lower() else ''
                                memory.append(memory_val)
            
                    tickcounter += 1
            
                return refassetlist, finaldaylist, vallist, settlelist, df_list, notional, memory, inv, ident, auto, tpayingint, interest, longname
            
            if not cusipdf.empty:
                cusiprefassetlist, cusipfinaldaylist, cusipvallist, cusipsettlelist, cusiplist, cusipnotional, cusipmemory, cusipinv, cusipident, cusipauto, cusiptpayingint, cusipinterest, cusiplongname = process_dataframe(cusipdf, tickerlist)
            
            if not isindf.empty:
                isinrefassetlist, isinfinaldaylist, isinvallist, isinsettlelist, isinlist, isinnotional, isinmemory, isininv, isinident, isinauto, isintpayingint, isininterest, isinlongname = process_dataframe(isindf, tickerlist)
    
    
                def process_worksheet(ws, df, refassetlist, vallist, settlelist, list_, notional, memory, inv, ident, auto, tpayingint, interest):
                    """
                    Process the worksheet based on the provided data.
                
                    Parameters:
                    - ws: The worksheet to process.
                    - df: The dataframe to check if it's empty.
                    - refassetlist, vallist, ... : Lists containing the data to be processed.
                    """
                
                    if df.empty == False:
                        # Initialize columns
                        valdatecol = paydatecol = cusipcol = cusipcolletter = refassetcol = principalcol = None
                        notecol = memorycol = autocallcol = autocolletter = invcol = identcol = identcolletter = None
                        payingintcol = interestcol = intcolletter = None
                
                        def cellmaker(colletter, rownum, value):
                            """Helper function to set cell value and alignment."""
                            ws[rf'{colletter}{rownum}'] = value
                            ws[rf'{colletter}{rownum}'].alignment = Alignment(horizontal='center', vertical='center')
                
                        for i in range(len(refassetlist)):
                            for cell in ws[1]:
                                if cell.value:
                                    cell_value = cell.value.strip()
                                    if cell_value == 'Valuation Date':
                                        valdatecol = cell.column if valdatecol is None else valdatecol
                                        cellmaker(cell.column_letter, i+2, vallist[i])
                                    elif cell_value == 'Payment Date':
                                        paydatecol = cell.column if paydatecol is None else paydatecol
                                        cellmaker(cell.column_letter, i+2, settlelist[i])
                                    elif cell_value == 'ISIN / CUSIP':
                                        cusipcol = cell.column if cusipcol is None else cusipcol
                                        cusipcolletter = cell.column_letter if cusipcolletter is None else cusipcolletter
                                        cellmaker(cell.column_letter, i+2, list_[i])
                                    elif cell_value == 'Reference Asset':
                                        refassetcol = cell.column if refassetcol is None else refassetcol
                                        cellmaker(cell.column_letter, i+2, refassetlist[i])
                                    elif cell_value == 'Principal':
                                        principalcol = cell.column if principalcol is None else principalcol
                                        cellmaker(cell.column_letter, i+2, abs(notional[i]))
                                    elif cell_value == 'No. of Notes':
                                        notecol = cell.column if notecol is None else notecol
                                        cellmaker(cell.column_letter, i+2, abs(notional[i]/1000))
                                    elif cell_value == 'Memory':
                                        memorycol = cell.column if memorycol is None else memorycol
                                        cellmaker(cell.column_letter, i+2, memory[i])
                                    elif cell_value == 'Inventory':
                                        invcol = cell.column if invcol is None else invcol
                                        cellmaker(cell.column_letter, i+2, inv[i])
                                    elif cell_value == 'Ident':
                                        identcol = cell.column if identcol is None else identcol
                                        identcolletter = cell.column_letter if identcolletter is None else identcolletter
                                        cellmaker(cell.column_letter, i+2, ident[i])
                                    elif cell_value == 'Autocalled':
                                        autocallcol = cell.column if autocallcol is None else autocallcol
                                        autocolletter = cell.column_letter if autocolletter is None else autocolletter
                                        cellmaker(cell.column_letter, i+2, auto[i])
                                    elif cell_value == 'Paying interest?':
                                        payingintcol = cell.column if payingintcol is None else payingintcol
                                        cellmaker(cell.column_letter, i+2, tpayingint[i])
                                    elif cell_value == 'Interest':
                                        interestcol = cell.column if interestcol is None else interestcol
                                        intcolletter = cell.column_letter if intcolletter is None else intcolletter
                                        # cellmaker(cell.column_letter, i+2, interest[i])
    
                        int_col = ref_asset_col = ident_col = None
                        for col in range(1, ws.max_column+1):
                            cell_value = ws.cell(row=1, column=col).value
                            if cell_value == 'Interest':
                                int_col = col
                            elif cell_value == 'Reference Asset':
                                ref_asset_col = col
                            elif cell_value == 'Ident':
                                ident_col = col
                
                        # fill out 'Interest' Column            
                        for i in range(1,ws.max_row):
                            ws.cell(row=i+1, column=int_col).value = f"=AC{i+1}*D{i+1}/12"
        
                                                        
                        # Merging cells
                        rows_iterate = []
                        rowcounter = 2
                        startmergerow = None
                        for cell in ws[cusipcolletter]: # Only iterates through cusip column
                            if cell.row == 1:
                                continue
                            elif (cell.value is not None) & (cell.value != 1):
                                if ws[rf'{cell.column_letter}{rowcounter}'].value == ws[rf'{cell.column_letter}{rowcounter+1}'].value:
                                    if startmergerow is not None:
                                        rowcounter += 1
                                        continue
                                    else:
                                        startmergerow = ws[rf'{cell.column_letter}{rowcounter}'].row
                                        rowcounter += 1
                                        continue
                                else:
                                    if startmergerow is None:
                                        startmergerow = ws[rf'{cell.column_letter}{rowcounter}'].row
                                    endmergerow = ws[rf'{cell.column_letter}{rowcounter}'].row
                                    ws.merge_cells(start_row=startmergerow,start_column=cusipcol,end_row=endmergerow,end_column=cusipcol)
                                    ws.merge_cells(start_row=startmergerow,start_column=valdatecol,end_row=endmergerow,end_column=valdatecol)
                                    ws.merge_cells(start_row=startmergerow,start_column=paydatecol,end_row=endmergerow,end_column=paydatecol)
                                    ws.merge_cells(start_row=startmergerow,start_column=principalcol,end_row=endmergerow,end_column=principalcol)
                                    ws.merge_cells(start_row=startmergerow,start_column=notecol,end_row=endmergerow,end_column=notecol)
                                    ws.merge_cells(start_row=startmergerow,start_column=memorycol,end_row=endmergerow,end_column=memorycol)
                                    ws.merge_cells(start_row=startmergerow,start_column=invcol,end_row=endmergerow,end_column=invcol)
                                    ws.merge_cells(start_row=startmergerow,start_column=int_col,end_row=endmergerow,end_column=int_col)
                                    ws.merge_cells(start_row=startmergerow,start_column=autocallcol,end_row=endmergerow,end_column=autocallcol)
                                    rows_iterate.append(startmergerow)
                                    
                                    # Sort cols     
                                    RefAsset_Ident = []
                                    
                                    for row in range(startmergerow, endmergerow+1):
                                        RefAsset_Ident.append([(ws.cell(row=row, column =ref_asset_col).value),(ws.cell(row=row, column =ident_col).value)])
        
                                    # Sort the list first by "Ident" (the int value) then by "Reference Asset" (alphabetically)
                                    RefAsset_Ident.sort(key=lambda x: (x[1],x[0]))
                                    
                                    for i in range(startmergerow, endmergerow+1):
                                        ws.cell(row=i, column=ref_asset_col).value = RefAsset_Ident[i-startmergerow][0]
                                        ws.cell(row=i, column=ident_col).value = RefAsset_Ident[i-startmergerow][1]
        
                                    startmergerow = None
                                    rowcounter += 1
                                    
                        # Merge "Ident" Column
                        ident_iterate = []
                        ident_counter = 2
                        ident_start_merge_row = None
                        for cell in ws[identcolletter]:
                            if cell.row == 1:
                                continue
                            elif (cell.value is not None) & (cell.value != 1):
                                if ws[rf'{cell.column_letter}{ident_counter}'].value == ws[rf'{cell.column_letter}{ident_counter+1}'].value:
                                    if ident_start_merge_row is not None:
                                        ident_counter += 1
                                        continue
                                    else:
                                        ident_start_merge_row = ws[rf'{cell.column_letter}{ident_counter}'].row
                                        ident_counter += 1                            
                                        continue
                                else:
                                    if ident_start_merge_row is None:
                                        ident_start_merge_row = ws[rf'{cell.column_letter}{ident_counter}'].row
                                    ident_end_merge_row = ws[rf'{cell.column_letter}{ident_counter}'].row
                                    ws.merge_cells(start_row=ident_start_merge_row,start_column=identcol,end_row=ident_end_merge_row,end_column=identcol)
                                    rows_iterate.append(ident_start_merge_row)
                                    ident_start_merge_row = None
                                    ident_counter += 1
                
                        # Change the font color to blue or red
                        red = Font(color='FF0000', bold=True)
                        blue = Font(color='0000FF')
                        for d, h in zip(ws['AF'], ws['C']):
                            if d.row in rows_iterate:
                                if d.value == 'Valid IC Date':
                                    d.font = red
                                    h.font = blue
                
                wb = load_workbook(filename=fileloc)
                process_worksheet(wb['CUSIP'], cusipdf, cusiprefassetlist, cusipvallist, cusipsettlelist, cusiplist, cusipnotional, cusipmemory, cusipinv, cusipident, cusipauto, cusiptpayingint, cusipinterest)
                process_worksheet(wb['ISIN'], isindf, isinrefassetlist, isinvallist, isinsettlelist, isinlist, isinnotional, isinmemory, isininv, isinident, isinauto, isintpayingint, isininterest)
    
                wb.save(fileloc)
                wb.close()
        
        testcounter += 1

    # ... [rest of your code]

if __name__ == "__main__":
    main()

    
