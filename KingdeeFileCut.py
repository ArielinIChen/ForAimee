# encoding:utf-8
import xlrd
import win32com.client
import os
import time

from common import copy_sheet, output_info_redirect


def core_method(dst_dir, wt_logfile):
    from_sht_num = 1
    cp_to_sht_num = 2
    cp_to_sht_name = 'Page1_Copy'
    for part_name in os.listdir(dst_dir):
        from_file = cp_to_file = os.path.join(dst_dir, part_name)
        copy_sheet(from_file, cp_to_file, from_sht_num, cp_to_sht_num, cp_to_sht_name, wt_logfile=wt_logfile)
        # 获取需要处理的sheet, 以及需要处理的col
        ori_rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
        ori_sht = ori_rb.sheet_by_name('Page1_Copy')
        sht_name = 'Page1_Copy'
        row_title = ori_sht.row_values(0)

        text = u'正在获取 科目代码、贷方 和 凭证摘要 所在的列...'
        output_info_redirect(text, wt_logfile)

        col_summary_num = col_subject_num = col_credit_num = -1
        for i in row_title:
            if i == u'凭证摘要':
                col_summary_num = row_title.index(i)
            elif i == u'科目代码':
                col_subject_num = row_title.index(i)
            elif i == u'贷方':
                col_credit_num = row_title.index(i)

        xlapp = win32com.client.Dispatch('Excel.Application')
        if col_summary_num >= 0 and col_subject_num >= 0 and col_credit_num >= 0:
            text = u'获取成功! 科目代码、贷方、凭证摘要 分别在 %s %s %s 列\n\n' \
                   % (col_subject_num, col_credit_num, col_summary_num) + \
                   u'开始删除 凭证摘要为: 结转本期损益 以及 之后连续空单元格 所在的行...'
            output_info_redirect(text, wt_logfile)

            # 每进行一次删除行循环, 都要重新加载一次excel文件
            # 删除'凭证摘要'列中, 单元格内容为: '结转本期损益' 和 之后的连续空单元格 的行
            while True:
                ori_rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
                ori_sht = ori_rb.sheet_by_name('Page1_Copy')
                col_summary = ori_sht.col_values(col_summary_num)
                selected = [x for x in range(len(col_summary)) if col_summary[x] == u'结转本期损益']
                if len(selected) == 0:
                    text = u'这张表格中 没有 凭证摘要为: 结转本期损益\n'
                    output_info_redirect(text, wt_logfile)
                    break
                else:
                    selected = selected[0]
                    text = u'已找到 结转本期损益, 当前所在行号为: %s\n' % (selected + 1) + \
                           u'开始计算 连续空单元格 的行号...'
                    output_info_redirect(text, wt_logfile)

                    void_col = [x for x in range(len(col_summary)) if col_summary[x] == '']
                    to_del = [selected, ]
                    tmp_num = selected + 1
                    if tmp_num in void_col:
                        for j in range(void_col.index(tmp_num), len(void_col)):
                            if void_col[j] == tmp_num:
                                to_del.append(void_col[j])
                                tmp_num += 1
                            else:
                                break
                    text = u'计算完毕, 当前需要删除的行号起始位置是: %s\n' \
                           u'需要删除连续的 %s 行' % (tmp_num, len(to_del))
                    output_info_redirect(text, wt_logfile)

                    if len(to_del) > 0:
                        xlbook = xlapp.Workbooks.Open(cp_to_file)
                        xlsht = xlbook.Worksheets(sht_name)
                        del_line = selected + 1
                        for i in range(len(to_del)):
                            xlsht.Rows(del_line).Delete()
                        # xlbook.Save()
                        xlbook.Close(SaveChanges=True)
                    text = u'删除完毕！继续检查... \n'
                    output_info_redirect(text, wt_logfile)

            # 删除 '贷方' 不等于0的行
            text = u'开始删除 贷方 不等于0的行'
            output_info_redirect(text, wt_logfile)

            ori_rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
            ori_sht = ori_rb.sheet_by_name('Page1_Copy')
            col_credit = ori_sht.col_values(col_credit_num)
            i = 1
            xlbook = xlapp.Workbooks.Open(cp_to_file)
            xlsht = xlbook.Worksheets(sht_name)
            while i < len(col_credit):
                del_line = i + 1
                if col_credit[i] != 0:
                    col_credit.remove(col_credit[i])
                    xlsht.Rows(del_line).Delete()
                    i -= 1
                i += 1
            # xlbook.Save()
            xlbook.Close(SaveChanges=True)
            text = u'删除完毕！\n'
            output_info_redirect(text, wt_logfile)

            # 删除 '科目代码' 属于del_list的行
            del_list = ['1001', '1002.001', '1002.002', '1002.003', '1002.004 ',
                        '1122.01', '1122.02', '1122.03', '1122.05',
                        '1221.04.001', '1221.04.002', '1221.04.003', '1405.01.01', '1405.01.02',
                        '2221.01.01.01', '2221.01.01.02', '2221.01.01.03',
                        '2241.04.001', '2241.04.002', '2203.01.001', '2203.01.003',
                        '1301.02.001.0001', '1301.02.001.0002']
            text = u'开始删除 科目代码 在列表中的行\n' \
                   u'该列表为:\n' \
                   u'%s' % del_list
            output_info_redirect(text, wt_logfile)

            ori_rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
            ori_sht = ori_rb.sheet_by_name('Page1_Copy')
            col_subject = ori_sht.col_values(col_subject_num)
            i = 1
            xlbook = xlapp.Workbooks.Open(cp_to_file)
            xlsht = xlbook.Worksheets(sht_name)

            while i < len(col_subject):
                del_line = i + 1
                if str(col_subject[i]) in del_list:
                    col_subject.remove(col_subject[i])
                    xlsht.Rows(del_line).Delete()
                    i -= 1
                i += 1
            # xlbook.Save()
            xlbook.Close(SaveChanges=True)
            text = u'删除完毕！\n'
            output_info_redirect(text, wt_logfile)

    text = u'金蝶文件切割 Excel文件处理完毕, 请在 %s 中获取处理后的Excel文件 \n' \
           u'请按 [回车键] 退出...' % dst_dir
    output_info_redirect(text, wt_logfile)

    raw_input()
    time.sleep(1)
    return
