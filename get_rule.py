import re
import argparse
import pandas as pd

def getRule(foldername: str, filename: str):
    def delete_last_two_digits(num):
        if re.search(r'\d+', num) is not None:
            integer_parts = re.search(r'\d+', num).group()
            if len(integer_parts) > 4:
                return num[:-2]
        return num
    def delete_space(num):
        return num.replace(" ", "")
    def multiple_num(nums):
        return nums.split(',')
    def delete_extra_digits(integer_list):
        return list(map(delete_last_two_digits, integer_list))

    rule = pd.read_excel(foldername + filename)
    rule['科號'] = rule['科號'].apply(delete_last_two_digits).apply(delete_space).apply(multiple_num).apply(delete_extra_digits)

    return rule

def get_argument():
    opt = argparse.ArgumentParser()
    opt.add_argument('--folder_name',
                        type = str,
                        default = './data/',
                        help = 'folder name')
    opt.add_argument('--rule_filename',
                        type = str,
                        default = '專長課程代號.xlsx',
                        help = 'filename of original rule')
    config = vars(opt.parse_args())
    return config

if __name__ == '__main__':
    config = get_argument()
    rule = getRule(config['folder_name'], config['rule_filename'])
    rule['學年度':'科號'].to_csv('rule.csv', index = False)