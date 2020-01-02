# coding=utf-8

import time
import configparser
import kindee_data.core_other_funcs as funcs
import kindee_data.core_db_connect as db_connect
import kindee_data.tool_balance as balance
import kindee_data.tool_report as report
import kindee_data.tool_detail as detail

print('\n金蝶 K3/WISE 数据导出工具 V2.0\n')

print('检查配置文件...')
funcs.check_file()

try:
    print('读取配置文件...')
    fyear = funcs.get_config('data', 'year')
    fperiod = funcs.get_config('data', 'end_period')
    flevel = funcs.get_config('data', 'level')
    fnumber_begin = funcs.convert_str(funcs.get_config('data', 'fnumber_begin'))
    fnumber_end = funcs.convert_str(funcs.get_config('data', 'fnumber_end'))
    split_mon = funcs.convert_str(funcs.get_config('data', 'split_mon'))
    split_file = funcs.convert_str(funcs.get_config('data', 'split_file'))
except configparser.NoSectionError as es:
    print('配置文件中存在错误:')
    print(es)
    funcs.exit_with_anykey()
except configparser.NoOptionError as eo:
    print('配置文件中存在错误:')
    print(eo)
    funcs.exit_with_anykey()
    

print('测试数据库连接...')
msg = db_connect.connect_db()

print('获取账套列表...')
db_name, company_name = db_connect.get_all_db_company(msg)  # 获取数据库名

db_list = []
print('\n可获取数据的账套如下：')
for i in range(len(db_name)):

    print('['+str(i+1)+'] '+company_name[i])
    db_list.append(str(i+1))


while True:
    input_db_num = input('\n输入要获取数据的账套号：')
    li_needed = sorted(input_db_num.split(','), key=funcs.emb_numbers)
    if input_db_num == '':
        li_needed = db_list
        break
    elif len(li_needed) != len(set(li_needed)):
        print('禁止输入重复值！')
    elif set(li_needed).issubset(set(db_list)):
        break
    else:
        print('输入错误！')

db_name_needed = []
company_name_needed = []
for i in range(len(li_needed)):
    db_name_needed.append(db_name[int(li_needed[i])-1])
    company_name_needed.append(company_name[int(li_needed[i])-1])
    
funcs.change_dir(funcs.mkdir('财务数据_'+fyear+'年_'+funcs.get_current_time()))

sheet_type = input('\n输入要导出的报表类型：\n[1] 科目余额表\n[2] 科目明细表\n[3] 财务报表\n')
t0 = time.time()
while True:
    if sheet_type == '1':
        balance.get_balance(db_name_needed, company_name_needed, msg,
                            fyear, fperiod, fnumber_begin, fnumber_end, flevel)
        break
    elif sheet_type == '2':
        detail.get_detail()
        break
    elif sheet_type == '3':
        report.get_report(db_name_needed, company_name_needed, msg,
                          fyear, fperiod)
        break
    else:
        print('输入错误！')

# TODO 不导出没有数据的账套
# TODO

t1 = time.time()
t = t1-t0
print('导出完成！\n用时：%.2f秒' % t)
funcs.exit_with_anykey()
