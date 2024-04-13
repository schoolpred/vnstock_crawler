import vnstock as vn
import pandas as pd
import yaml


# Read YAML file
with open("config.yaml", 'r') as file:
    data = yaml.safe_load(file)

list_stock = data['list_stock']

stock_price_df = pd.DataFrame()
financial_report_df = pd.DataFrame()
financial_ratio_df = pd.DataFrame()
company_overview_df = pd.DataFrame()
balance_sheet_df = pd.DataFrame()
cash_flow_df = pd.DataFrame()
for stock in list_stock:
    df_price =  vn.stock_historical_data(symbol=stock, 
                                start_date="2015-01-01", 
                                end_date='2024-01-02', resolution='1D', type='stock', beautify=True, decor=False, source='dnse')
    
    #report
    df_report = vn.financial_report(symbol=stock, report_type='IncomeStatement', frequency='Quarterly', periods=32, latest_year=None)
    df_bs = vn.financial_report(symbol=stock, report_type='BalanceSheet', frequency='Quarterly', periods=32, latest_year=None)
    df_cf = vn.financial_report(symbol=stock, report_type='CashFlow', frequency='Quarterly', periods=32, latest_year=None)

    df_report = pd.melt(df_report, id_vars='CHỈ TIÊU', value_vars=df_report.columns, var_name='Quarter',value_name='Value')
    df_bs = pd.melt(df_bs, id_vars='CHỈ TIÊU', value_vars=df_bs.columns, var_name='Quarter',value_name='Value')
    df_cf = pd.melt(df_cf, id_vars='Unnamed: 0', value_vars=df_cf.columns, var_name='Quarter',value_name='Value')

    df_report['ticker'] = stock
    df_bs['ticker'] = stock
    df_cf['ticker'] = stock


    #financial
    df_ratio = vn.financial_ratio(symbol=stock, report_range='quarterly', is_all=True)
    df_ratio.drop(index=['ticker', 'quarter', 'year'], inplace=True)
    df_ratio.reset_index(inplace=True)
    df_ratio = pd.melt(df_ratio, id_vars='index', value_vars=df_ratio.columns, var_name='Quarter',value_name='Value')
    df_ratio['ticker'] = stock


    #company overview
    df_company_overview = vn.company_overview(stock).T
    df_company_overview.drop(index=['ticker'], inplace=True)
    df_company_overview.reset_index(inplace=True)
    df_company_overview.rename({0: 'value'}, axis='columns', inplace=True)
    df_company_overview['ticker'] = stock
    

    stock_price_df = pd.concat([stock_price_df, df_price], ignore_index=True)
    print('price')
    financial_report_df = pd.concat([financial_report_df, df_report], ignore_index=True)
    print('report')
    financial_ratio_df = pd.concat([financial_ratio_df, df_ratio], ignore_index=True)
    print('ratio')
    company_overview_df = pd.concat([company_overview_df, df_company_overview], ignore_index=True)
    print('overview')
    balance_sheet_df = pd.concat([balance_sheet_df, df_bs], ignore_index=True)
    print('overview')
    cash_flow_df = pd.concat([cash_flow_df, df_cf], ignore_index=True)
    print('overview')


    print(stock)

try:
    stock_price_df.to_excel('../result/price.xlsx', index = False)
    financial_report_df.to_excel('../result/financial_report.xlsx', index = False)
    financial_ratio_df.to_excel('../result/financial_ratio.xlsx', index = False)
    company_overview_df.to_excel('../result/company_overview.xlsx', index = False)
    balance_sheet_df.to_excel('../result/balance_sheet.xlsx', index = False)
    cash_flow_df.to_excel('../result/cash_flow.xlsx', index = False)
    print('brilliant')
except:
    print('it bug, you idiot!')
    stock_price_df.to_excel('price.xlsx', index = False)
    financial_report_df.to_excel('financial_report.xlsx', index = False)
    financial_ratio_df.to_excel('financial_ratio.xlsx', index = False)
    company_overview_df.to_excel('company_overview.xlsx', index = False)
    balance_sheet_df.to_excel('balance_sheet.xlsx', index = False)
    cash_flow_df.to_excel('cash_flow.xlsx', index = False)



#calculate price group by quarter
print('calculate price group by quarter')
df = pd.read_excel('../result/price.xlsx')

df['time'] = pd.to_datetime(df['time']) 
# Set 'time' as index
df.set_index('time', inplace=True)
# Group by 'ticker' and 'quarter', and aggregate
agg_df = df.groupby(['ticker', pd.Grouper(freq='QE')]).agg({
    'open': 'first',
    'high': 'max',
    'low': 'min',
    'close': 'last',
    'volume': 'sum'
}).reset_index()

agg_df['time'] = agg_df['time'].dt.to_period('Q').astype(str)

agg_df.to_excel('../result/price_by_quarter.xlsx', index = False)

print('you did it :D')