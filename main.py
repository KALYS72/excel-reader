import pandas as pd
import pprint
from dateutil import parser
import json
import os

sample2 = 'sample2.xlsx'
sample1 = 'sample1.xlsx'
sample3 = 'sample3.xlsx'
counter_file = 'counter.json'

df = pd.read_excel(sample1)
df2 = pd.read_excel(sample2)
df3 = pd.read_excel(sample3)

def until_none(string, column):
    column_data = df.iloc[string:, column]
    values = 0
    for value in column_data:
        if pd.isna(value):  
            break
        values += 1 
    return values

def convert_russian_date(date_str):
    try:
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
        date_stri = date_str.replace('"', '').replace("'", '').split()
        date_str2 = date_stri[:2]
        date_str2.append(date_str[2][:4])
        date_str3 = ' '.join(date_str2)
        date = parser.parse(date_str3)
        return date.strftime('%Y-%m-%d')
    except:
        return date_str

def load_counter():
    if os.path.exists(counter_file):
        with open(counter_file, 'r') as file:
            data = json.load(file)
            return data.get('counter', 0)
    return 0

def save_counter(counter):
    with open(counter_file, 'w') as file:
        json.dump({'counter': counter}, file)

def auto_fill(start_y, end_y, x, file, path_file):
    column = file.iloc[start_y:end_y + 1, x]
    column.fillna(method='ffill', inplace=True)
    file.iloc[start_y:end_y + 1, x] = column
    df.to_excel(path_file, index=False)

def can_be_integer(s):
    try:
        int(s)  
        return True 
    except TypeError:
        return False
    except ValueError:
        return False
    
def find_quarter(Title):
    Quarter = None
    Year = None
    number_count = 0
    while len(Title) != 0:
        if can_be_integer(Title[-1]):
            number_count += 1
            if Year == None:
                Year = Title[-1]
            elif Quarter == None:
                Quarter = Title[-1]
        Title.pop()
    return Quarter, Year
    
    
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

# Reports_info = df.iloc[:7, 0].tolist()
# Reports_data = df.iloc[:7, 2].tolist()
# Booker_info = df.iloc[end_reports_int+Applications_end+1, 7]
# Booker_data = df.iloc[end_reports_int+Applications_end+1, 9]
# additionall_info = df.iloc[:4, 9].to_list()
# additionall_info2 = df.iloc[:4, 10].to_list()
# Reports_info.append(Booker_info)
# Reports_info.extend(additionall_info)
# Reports_data.append(Booker_data)
# Reports_data.extend(additionall_info2)

def get_Reports(df):
    Title = df.iloc[start_reports-2, 0].split(' ')
    Quarter, Year = find_quarter(Title)
    counter = load_counter()
    # WHERE company_name = '{df.columns[2]}')""", 
    Reports = {
        'counter' : counter,
        'submission_status_id': 4,  
        'payment_status_id': 3,
        'submission_date': "'" + str(convert_russian_date(df.iloc[3, 9])) + "'", 
        'resident_id': f"""
        (SELECT "Id"
        FROM public."Residents"
        WHERE company_name = 'Walmart')""", 
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
    for value in Reports.values():
        if can_be_integer(value):
            result += str(value) + ', '
        elif can_be_bool(value):
            result += value + ', '
        else:
            result += "'" + str(value) + "', "
    result = result[:-2]
    result += ');'
    counter += 1
    save_counter(counter)
    print(result)
    return counter-1


# for i in range(13):
#     auto_fill(start_reports, end_reports_int-2, i, file=df, path_file=sample1)

Report_data = {}
Report_data = {
        'amount': df.iloc[start_reports:end_reports_int-2, 5].tolist(),
        'amount_date': df.iloc[start_reports:end_reports_int-2, 2].tolist(),
        'payment': df.iloc[start_reports:end_reports_int-2, 6].tolist(),
        'customer_id': df.iloc[start_reports:end_reports_int-2, 4].tolist(),
        'currency_id': ((end_reports_int - 2) - start_reports) * [1], 
        'CreatedAt': 'NOW()',  
        'UpdatedAt': 'NOW()',
        'DeletedAt': 'NULL',  
        'report_id': 'NULL',  
        'IsDeactivated': 'false',  
        'service_type_id': df.iloc[start_reports:end_reports_int-2, 3].tolist(),
        'RateToKGS': 'NULL', 
        'original_currency_amount': 'NULL' 
    }

for i in range(len(Report_data['amount'])):
    result = 'INSERT INTO public."Report_transactions" VALUES ('
    for key, value in Report_data.items():
        if key == 'customer_id':
            result += f"""(SELECT "Id" FROM public."Customers" WHERE customer_name = '{value[i]}'),"""
        # elif isinstance(value[i], pd.Timestamp):
        #     result += 'NOW()' + ','
        # elif can_be_integer(value[i]):
        #     result += str(value[i]) + ','
        else:
            result += "'" + str(value[i]) + "',"
    result = result[:-1]
    result += ');'
    pprint.pprint(result)

# get_Reports(df3)
