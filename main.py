import pandas as pd
import pprint, calendar
from decouple import config
from dateutil import parser

file_path = 'sample2.xlsx'

df = pd.read_excel(file_path)

def until_none(string, column):
    column_data = df.iloc[string:, column]
    values = 0
    for value in column_data:
        if pd.isna(value):  
            break
        values += 1 
    return values

def convert_russian_date(date_str):
    russian_to_english_month = {
        'января': 'January',
        'февраля': 'February',
        'марта': 'March',
        'апреля': 'April',
        'мая': 'May',
        'июня': 'June',
        'июля': 'July',
        'августа': 'August',
        'сентября': 'September',
        'октября': 'October',
        'ноября': 'November',
        'декабря': 'December'
    }

    for ru_month, en_month in russian_to_english_month.items():
        if ru_month in date_str:
            date_str = date_str.replace(ru_month, en_month)
            break
    date_str = date_str.strip()[:-2]
    date = parser.parse(date_str)
    return date.strftime('%Y-%m-%d')

def auto_fill(start_y, end_y, x, file):
    column = file.iloc[start_y:end_y + 1, x]
    column.fillna(method='ffill', inplace=True)
    file.iloc[start_y:end_y + 1, x] = column
    df.to_excel(file_path, index=False)

def can_be_integer(s):
    try:
        int(s)  
        return True 
    except TypeError:
        return False
    except ValueError:
        return False
    
def can_be_bool(s):
    try:
        bool(s)  
        return True 
    except TypeError:
        return False
    except ValueError:
        return False

start_reports = 9
end_reports = df[df.iloc[:, 0].astype(str).str.contains("Всего:", na=False)]
end_reports_int = end_reports.index[0]+2

Sources_info = df.iloc[end_reports_int:end_reports_int+4, 3].tolist()
Sources_data = df.iloc[end_reports_int:end_reports_int+4, 5].tolist()
Sources = dict(zip(Sources_info,Sources_data))

Export_sources_end = until_none(end_reports_int+4, 4)
Export_sources_info = df.iloc[end_reports_int+4:end_reports_int+4+Export_sources_end, 4].tolist()
Export_sources_data = df.iloc[end_reports_int+4:end_reports_int+4+Export_sources_end, 5].tolist()
Export_sources = dict(zip(Export_sources_info,Export_sources_data))

Services_end = until_none(end_reports_int, 0)
Services_info = df.iloc[end_reports_int:end_reports_int+Services_end, 0].tolist()
Services_data = df.iloc[end_reports_int:end_reports_int+Services_end, 1].tolist()
Services = dict(zip(Services_info,Services_data))

Applications_end = until_none(end_reports_int, 8)
Applications_info = df.iloc[end_reports_int:end_reports_int+Applications_end, 8].tolist()

Reports_info = df.iloc[:7, 0].tolist()
Reports_data = df.iloc[:7, 2].tolist()
Booker_info = df.iloc[end_reports_int+Applications_end+1, 7]
Booker_data = df.iloc[end_reports_int+Applications_end+1, 9]
additionall_info = df.iloc[:4, 9].to_list()
additionall_info2 = df.iloc[:4, 10].to_list()
Reports_info.append(Booker_info)
Reports_info.extend(additionall_info)
Reports_data.append(Booker_data)
Reports_data.extend(additionall_info2)


Title = df.iloc[start_reports-2, 0].split(' ')
Quarter = Title[-4]
Year = Title[-2]

# WHERE company_name =  {df.iloc[1, 9]})''', 
Reports = {
    'counter' : 2,
    'submission_status_id': 4,  
    'payment_status_id': 3,
    'submission_date': "'" + str(convert_russian_date(df.iloc[3, 9])) + "'", 
    'resident_id': f'''
    (SELECT "Id"
    FROM public."Residents"
    WHERE "company_name" = 'Walmart')''', 
    'quarter_id': f'''
    (SELECT "Id"
    FROM public."Quarters"
    WHERE "quarter" = {Quarter}
    AND "year_id" IN (
    SELECT "Id"
    FROM public."Years"
    WHERE "year" = {Year}))''', 
    'CreatedAt': 'NOW()',  
    'UpdatedAt': 'NOW()',
    'DeletedAt': 'NULL',  
    'IsDeactivated': 'false',  
    'original_report_id': 'NULL',  
    'comment': 'NULL'
}


result = 'INSERT INTO public."Reports" \nVALUES ('
for key, value in Reports.items():
    if can_be_integer(value):
        result += str(value) + ', '
    elif can_be_bool(value):
        result += value + ', '
    else:
        result += "'" + str(value) + "', "
result = result[:-2]
result += ');'
print(result)


for i in range(13):
    auto_fill(start_reports, end_reports_int-2, i, df)

Report_data = {}
Report_data['amount'] = df.iloc[start_reports:end_reports_int-2, 5].tolist()
Report_data['amount_date'] = df.iloc[start_reports:end_reports_int-2, 2].tolist()
Report_data['payment'] = df.iloc[start_reports:end_reports_int-2, 6].tolist()
Report_data['customer_id'] = df.iloc[start_reports:end_reports_int-2, 4].tolist()
Report_data['currency_id'] = ((end_reports_int - 2) - start_reports) * [1]
Report_data['service_type_id'] = df.iloc[start_reports:end_reports_int-2, 3].tolist()



for i in range(len(Report_data['amount'])):
    result = 'INSERT INTO public."Report_transactions"\nVALUES ('
    for key, value in Report_data.items():
        if isinstance(value[i], pd.Timestamp):
            result += 'NOW()' + ', '
        elif can_be_integer(value[i]):
            result += str(value[i]) + ', '
        else:
            result += "'" + str(value[i]) + "', "
    result = result[:-2]
    result += ');'
    # print(result)
    result = 'INSERT INTO public.Report_Transactions '