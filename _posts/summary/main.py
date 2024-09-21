import pandas as pd

def preprocess(data_execl):
    data_execl.insert(data_execl.shape[1], 'date', 0)
    for idx, row in data_execl.iterrows():
        if (idx == 0):
            data_execl.loc[idx, 'date'] = int(data_execl.loc[idx, '演出日期'])
            continue

        isnull = row.isnull()
        for col in data_execl.columns:
            if isnull[col] and col != '备注':
                data_execl.loc[idx, col] = data_execl.loc[idx - 1, col]

        data_execl.loc[idx, 'date'] = int(data_execl.loc[idx, '演出日期'])

def process_data_format(data):
    data = data.drop(columns=['date'])

    for idx, row in data.iterrows():
        date_str = data.loc[idx, '演出日期']
        str_dot = date_str[:4] + '.' + date_str[4:6] + '.' + date_str[6:]
        data.loc[idx, '演出日期'] = str_dot

    return data

def year_html_to_md(year, year_total, year_html, md_str) :
    md_str += '\n<tr>\n<td colspan="4" style="text-align: left;" class="font-weight-bold text-danger">'
    md_str += str(year) + '年(共' + str(year_total) + '场)' + '</td>\n<tr>\n\n'

    lines = year_html.split('\n')
    for line in lines :
        if '<table' in line or '</table>' in line or \
           '<tbody' in line or '</tbody>' in line :
            continue
        md_str += line.strip()
        md_str +='\n'
        if '</tr>' in line:
            md_str += '\n'
    return md_str

def add_oneyear_to_md(md_str, data_execl, year):
    year_data = pd.DataFrame()
    year_total = 0
    for idx, row in data_execl.iterrows():
        if data_execl.loc[idx, 'date'] > year * 10000 \
           and data_execl.loc[idx, 'date'] < (year + 1) * 10000 :
            year_data = pd.concat([year_data, pd.DataFrame([row])], ignore_index=True)
            year_total += 1

    year_data = year_data.sort_values(by='date', ascending=False)
    year_data = process_data_format(year_data)
    year_html = year_data.to_html(index=False, header=False, na_rep='/')
    md_str = year_html_to_md(year, year_total, year_html, md_str)
    return md_str


years =[2023, 2022, 2021, 2020, 2019, 2018, 2017, 2016, 2015, 2014, 2013]
if __name__ == '__main__':

    execl_file = '李云霄2013-2023十年演出记录.xlsx'
    md_file = 'generated.md'

    md_str = ''
    data_execl = pd.read_excel(execl_file, header=[2], dtype=str)
    preprocess(data_execl)
    for year in years:
        md_str = add_oneyear_to_md(md_str, data_execl, year)

    with open(md_file, 'w', encoding='utf-8') as f:
        f.write(md_str)



