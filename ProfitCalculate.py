# encoding:utf-8
import xlrd
import openpyxl
import os
import shutil
import time
from common import copy_sheet, change_sheet_name, output_info_redirect


def core_method(dst_dir, err_dir, wt_logfile):
    calc_err_dir = u'E:\报表处理\利润表\不平的报表'
    dept_list = os.listdir(dst_dir)
    dept_list.remove(u'利润表合并')
    combine_dir = os.path.join(dst_dir, u'城市合并版')
    merge_dir = os.path.join(dst_dir, u'利润表合并')

    os.makedirs(combine_dir)
    # prime_dept_name = [u'北京诺互', u'成都诺互', u'广州景沪', u'广州诺互', u'海口诺互', u'昆明百当诺互',
    #                    u'南京诺互', u'青岛楷模', u'上海昌保', u'深圳深南', u'沈阳诺互', u'天津诺互',
    #                    u'武汉云景', u'西安来护']
    if os.path.exists(calc_err_dir):
        shutil.rmtree(calc_err_dir)
    os.makedirs(calc_err_dir)

    if u'上海诺互利润表.xlsx' in os.listdir(merge_dir):
        merge_file = os.path.join(merge_dir, u'上海诺互利润表.xlsx')
        merge_wb = openpyxl.load_workbook(filename=merge_file)
        merge_wb_shts = merge_wb.sheetnames
    else:
        text = u'\n发现错误!\n' \
               u'没有找到 利润表合并的Excel - 上海诺互利润表.xlsx\n' \
               u'系统将退出! 请按回车键退出!\n'
        output_info_redirect(text, wt_logfile)
        raw_input()
        return

    merge_data = {}  # 存放门店的利润表数据

    text = u'开始处理利润表...\n'
    output_info_redirect(text, wt_logfile)

    for dept in dept_list:
        dept_dir = os.path.join(dst_dir, dept)
        shops_list = os.listdir(dept_dir)
        try:
            cp_to_file_tmp = [x for x in shops_list if x.startswith(dept)][0]
        except IndexError:
            text = u'#########\n发现错误!\n' \
                   u'%s 里, 没有以 %s 开头的利润汇总表\n' \
                   u'#########' % (dept_dir, dept)
            output_info_redirect(text, wt_logfile)
            continue
        shops_list.remove(cp_to_file_tmp)
        cp_to_file = os.path.join(dept_dir, cp_to_file_tmp)
        rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
        total_profit = rb.sheet_by_index(0).cell(17, 1).value

        text = u'\n' + u'###' * 10 + u'\n' + \
               u'%s 在 汇总表中的净利润是: %s' % (dept, total_profit)
        output_info_redirect(text, wt_logfile)

        date_part = cp_to_file_tmp.split('-')[-1]
        combine_sht_name = dept + u'-利润表-' + date_part.split('.xls')[0]
        change_sheet_name(cp_to_file, combine_sht_name, sht_num=0)

        cp_to_sht_num = 2
        from_sht_num = 1

        if len(shops_list) > 0:
            shops_profit = 0.0
        else:
            shops_profit = total_profit
        merge_shops_name = []
        month = int(date_part.split('.')[1])
        merge_col = month * 3 - 1
        for shop in shops_list:
            from_file = os.path.join(dept_dir, shop)
            shop_name_tmp = shop.split(u'利润表')[0]
            shop_name = shop_name_tmp + u'店' if not shop_name_tmp.endswith(u'店') else shop_name_tmp
            cp_to_sht_name = shop_name + u'-利润表-' + date_part.split('.xls')[0]

            # 获取当月的实际净利润
            shop_rb = xlrd.open_workbook(from_file, formatting_info=True)
            my_profit = shop_rb.sheet_by_index(0).cell(17, 1).value
            shops_profit += my_profit

            text = u'\n开始编辑 %s\n' \
                   u'加上 %s 的 %s 后, %s 的实际净利润是: %s\n' % (shop, shop_name, my_profit, dept, shops_profit)
            output_info_redirect(text, wt_logfile)

            # 获取当月的其他数据 写入字典merge_data
            shop_rs = shop_rb.sheet_by_index(0)
            merge_data[shop_name] = shop_rs.col_values(1, 1, 21)
            merge_shops_name.append(shop_name)

            # 调用common方法中的 - 复制sheet表函数
            copy_sheet(from_file, cp_to_file, from_sht_num, cp_to_sht_num, cp_to_sht_name, wt_logfile=wt_logfile)

        total_profit = float(str(total_profit).decode('utf-8'))
        shops_profit = float(str(shops_profit).decode('utf-8'))

        text = u'现在 %s 的总实际净利润是: %s\n' % (dept, shops_profit) + \
               u'开始计算 %s 利润表B18单元格(净利润) 是否准确...\n' \
               u'汇总数额 %s VS 实际数额 %s\n' % (dept, total_profit, shops_profit)
        output_info_redirect(text, wt_logfile)

        # 如果total数和实际数一致, 那么归档报表, 并把merge_data写入利润表合并
        if total_profit != shops_profit:
            text = u'%s B18(净利润) 有错误, 将把该文件夹移动到 [不平的报表] 中...' % dept
            output_info_redirect(text, wt_logfile)

            move_to_dir = os.path.join(calc_err_dir, dept)
            shutil.move(dept_dir, move_to_dir)

            text = u'移动完毕!\n'
            output_info_redirect(text, wt_logfile)
        else:
            # 汇总 [利润表 - 城市合并版]到同一个文件夹
            name_part = dept + u'利润表'
            combine_file_name = '-'.join((name_part, date_part))
            combine_file = os.path.join(combine_dir, combine_file_name)

            text = u'准确无误!\n' \
                   u'准备将 汇总表 %s\n' \
                   u'复制到 城市合并版 %s\n' % (cp_to_file, combine_file)
            output_info_redirect(text, wt_logfile)

            shutil.copy(cp_to_file, combine_file)

            text = u'复制完成!\n\n' \
                   u'开始合并利润表...'
            output_info_redirect(text, wt_logfile)

            # 整理 利润表合并
            for merge_shop in merge_shops_name:
                merge_sht_num = map((lambda y: merge_shop in y), merge_wb_shts).index(True)
                merge_ws = merge_wb[merge_wb_shts[merge_sht_num]]
                data = merge_data[merge_shop]

                for i in range(len(data)):
                    fill_num = data[i]
                    if i == 0 and fill_num != 0:
                        merge_ws.cell(3, merge_col, fill_num)
                    elif i not in [10, 14, 16] and fill_num != 0:
                        merge_row = i + 5
                        merge_ws.cell(merge_row, merge_col, fill_num)
                merge_wb.save(merge_file)

                text = u'%s 利润表合并完毕' % merge_shop
                output_info_redirect(text, wt_logfile)

    merge_wb.close()

    text = u'###' * 20 + \
           u'\n利润表 Excel文件处理完毕.\n' \
           u'请在 %s 中获取处理后的Excel文件 \n' \
           u'在 %s 中获取 报表不平的门店\n' \
           u'在 %s 中获取 有错误的原始报表\n' % (dst_dir, calc_err_dir, err_dir) + \
           u'###' * 20 + \
           u'\n请按 [回车键] 退出...'
    output_info_redirect(text, wt_logfile)

    raw_input()
    time.sleep(1)
    return
