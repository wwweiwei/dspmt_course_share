import re
import argparse
import pandas as pd
import numpy as np
from ast import literal_eval
import warnings
warnings.filterwarnings("ignore")

def get_student_curriculum(foldername: str, filename: str):
    data = pd.read_excel(foldername + filename)

    info = data.columns[0]
    data = data.set_axis(['學年', '學期', '科號', '科目名稱'], axis=1, inplace=False).drop([0])
    enrollment_year = min(data['學年'])

    def delete_last_two_digits(num):
        if re.search(r'\d+', num) is not None:
            integer_parts = re.search(r'\d+', num).group()
            if len(integer_parts) > 4:
                return num[:-2]
        return num
    def delete_space(num):
        return num.replace(" ", "")
    def shorten_course_name(text):
        return text.split('--', 1)[0]
    
    ## 科號移除後兩碼和空格, 科目名稱去除--之後的文字
    data['科號'] = data['科號'].apply(delete_space).apply(delete_last_two_digits)
    data['科目名稱'] = data['科目名稱'].apply(shorten_course_name)

    ## 刪除服學(Z)和體育(PE)
    for i, value in enumerate(data['科號']):
        if str(value.strip()[0]) == 'Z' or str(value.strip()[0:2]) == 'PE':
            data = data.drop(index = [i+1])

    data = data.reset_index() ## some index had been delete

    return data, info, enrollment_year

def find_specialty(info):
    first_specialty, _ , second_specialty = info[info.find('第一專長 ：')+6:].partition('\u3000')
    second_specialty, _ , _ = second_specialty[second_specialty.find('第二專長：')+5:].partition('\u3000')
    return first_specialty, second_specialty

def get_output(data, rules, first_specialty, second_specialty):
    output = pd.DataFrame(np.full((data.shape[0], 5), np.nan), columns = ['編號', '課名', '學年度', '學期', '類別'])

    count = 0
    for i, row in data.iterrows():
        c_year = row['學年']
        c_semester = row['學期']
        c_num = row['科號'].strip()
        c_name = row['科目名稱']
        
        for j, rule in rules.iterrows():
            if count > i: ## find the match source
                break
                
            r_num = eval(rule['科號'])
            r_cat = rule['類別']
            r_specialty = rule['專長名稱']

            for _, num in enumerate(r_num):
                if str(c_num) == str(num.replace(' ', '')):
                    output['課名'].iloc[i] = c_name
                    output['學年度'].iloc[i] = c_year
                    output['學期'].iloc[i] = c_semester
                
                    if str(r_cat) == '基礎必修':
                        count += 1
                        output['類別'].iloc[i] = r_cat
                        break

                    elif str(r_cat) == '第一專長':
                        if r_specialty == first_specialty: ## 一專課程
                            count += 1
                            output['類別'].iloc[i] = r_cat
                            break
                            
                    elif str(r_cat) == '第二專長':
                        if r_specialty == second_specialty: ## 二專課程
                            count += 1
                            output['類別'].iloc[i] = r_cat
                            break
                
            if j == rules.shape[0]-1: ## 其他課程
                count += 1
                output['課名'].iloc[i] = c_name
                output['學年度'].iloc[i] = c_year
                output['類別'].iloc[i] = '其他'
                output['學期'].iloc[i] = c_semester

    return output

def formatting(output, enrollment_year):
    output['學年度'] = output['學年度'].astype(int)
    output['學期'] = output['學期'].astype(int)

    match = {
        '基礎必修' : 1 ,
        '第一專長' : 2 ,
        '第二專長' : 3 ,
        '其他' : 4 
    }

    output['編號'] = output['類別'].map(match)

    for i in range(output.shape[0]):
        if output['學年度'].iloc[i] == enrollment_year:
            if output['學期'].iloc[i] == 10:
                output['編號'].iloc[i] = output['編號'].iloc[i]
            else: ## '20' or others
                output['編號'].iloc[i] += 4
        elif output['學年度'].iloc[i] == enrollment_year+1:
            if output['學期'].iloc[i] == 10:
                output['編號'].iloc[i] += 8
            else: ## '20' or others
                output['編號'].iloc[i] += 12
        elif output['學年度'].iloc[i] == enrollment_year+2:
            if output['學期'].iloc[i] == 10:
                output['編號'].iloc[i] += 16
            else: ## '20' or others
                output['編號'].iloc[i] += 20
        elif output['學年度'].iloc[i] == enrollment_year+3:
            if output['學期'].iloc[i] == 10:
                output['編號'].iloc[i] += 24
            else: ## '20' or others
                output['編號'].iloc[i] += 28
        else:
            if output['學期'].iloc[i] == 10:
                output['編號'].iloc[i] += 32
            else: ## '20' or others
                output['編號'].iloc[i] += 36

    output = output.drop(columns=['學年度', '學期', '類別'])
    return output

def get_argument():
    opt = argparse.ArgumentParser()
    opt.add_argument('--folder_name',
                        type = str,
                        default = './data/',
                        help = 'folder name')
    opt.add_argument('--student_filename',
                        type = str,
                        default = '範例課表.xlsx',
                        help = 'filename of original curriculum')
    config = vars(opt.parse_args())
    return config

if __name__ == '__main__':
    config = get_argument()
    curriculum, info, enrollment_year = get_student_curriculum(config['folder_name'], config['student_filename'])
    first_specialty, second_specialty = find_specialty(info)

    rules = pd.read_csv('rule.csv')
    output = get_output(curriculum, rules, first_specialty, second_specialty)
    output = formatting(output, enrollment_year)

    ## check
    print('學生姓名: ', info[3:6])
    print('學生入學年度:', enrollment_year)
    print('第一專長:', first_specialty)
    print('第二專長:', second_specialty)
    print(output.tail(10))

    output.to_excel(info[3:6]+'.xlsx', index = False)