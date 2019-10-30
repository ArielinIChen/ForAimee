# encoding:utf-8
import os
import sys
import time

from openpyxl import Workbook


def output_info_redirect(text, wt_logfile):
    print (text)
    text = text.encode('utf-8')
    # wt_logfile.write(text + '\n')
    with open(wt_logfile, 'a+') as write_file:
        write_file.write(text + '\n')
    # f = open(filename, 'a+')
    # # print >> f, text
    # f.write(text + '\n')
    # f.close()
    return


def create_logfile(base_dir):
    log_path = os.sep.join((base_dir, u'Script-Log'))

    if not os.path.exists(log_path):
        os.makedirs(log_path)

    logfile_list = os.listdir(log_path)
    for _file in logfile_list:
        _filepath = os.path.join(log_path, _file)
        if os.path.isdir(_filepath):
            pass
        else:
            os.remove(_filepath)

    now_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
    logfile_name = now_time + '.txt'
    _logfile_path = os.path.join(log_path, logfile_name)
    return _logfile_path


def common_check_folder(path):
    if not os.path.exists(path):
        return 'not found'
    else:
        if os.path.isdir(path):
            return 'dir'
        elif os.path.isfile(path):
            return 'file'
        else:
            return 'special'


def check_work_folder(path_list, logfile):
    for path_and_type in path_list:
        path = path_and_type[0]
        path_type = path_and_type[1]
        text = u'正在检查 %s \n' % path
        output_info_redirect(text, logfile)
        resp = common_check_folder(path)
        if resp == 'not found' or resp != path_type:
            if resp == 'not found':
                text = u'检查不到 %s 这个文件或文件夹! 请检查... \n' % path
            else:
                if resp == 'dir':
                    text = u'%s 应该是一个文件, 但检测到的是一个文件夹, 请检查 \n' % path
                elif resp == 'file':
                    text = u'%s 应该是一个文件夹, 但检测到的是一个文件, 请检查 \n' % path
                else:
                    text = u'%s 是一个特殊文件 \n' % path
            text += u'系统将在3秒后退出... \n'
            output_info_redirect(text, logfile)
            time.sleep(3)
            sys.exit(1)
        else:
            text = u'检查完成...OK \n'
            output_info_redirect(text, logfile)
    return


def create_xl_file(file_path, data_list):
    wb = Workbook()
    ws = wb.active

    for my_row in range(len(data_list)):
        row_data = data_list[my_row]
        my_row += 1
        for my_col in range(len(row_data)):
            insert_data = row_data[my_col]
            my_col += 1
            ws.cell(my_row, my_col, insert_data)

    wb.save(file_path)
    wb.close()
