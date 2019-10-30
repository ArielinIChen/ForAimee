# encoding:utf-8
import os
import sys
import time
from common import chk_folder, copy_files, output_info_redirect
import KingdeeFileCut
import ProfitCalculate
from cost_proportion import chk_proportion_file, merge_proportion_result


def create_logfile(pj_name):
    log_dir = u'E:\报表处理\日志'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    logfile_list = sorted(os.listdir(log_dir))
    i = 0
    while i < len(logfile_list):
        if not logfile_list[i].startswith('20') and not logfile_list[i].endswith('.txt'):
            logfile_list.remove(logfile_list[i])
            i -= 1
        i += 1

    while True:
        if len(logfile_list) < 5:
            break
        else:
            filename = logfile_list[0]
            file_path = os.path.join(log_dir, filename)
            os.remove(file_path)
            logfile_list.remove(logfile_list[0])

    now_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
    logfile_name = now_time + '_' + pj_name + '.txt'
    logfile_tmp = os.path.join(log_dir, logfile_name)
    return logfile_tmp


while True:
    print (u'脚本启动! \n'
           u'请选择需要操作的项目: \n'
           u'1: 处理金蝶导出文件 \n'
           u'2: 利润表 \n'
           u'3: 费用占比 \n'
           u'4: 直接退出 \n')
    choice = raw_input()

    if choice == '1':
        wt_logfile = create_logfile('kingdee')

        text = u'选择了1\n' + \
               u'##' * 14 + '\n' + \
               u'#  开始运行 金蝶报表切割脚本  #\n' + \
               u'##' * 14
        output_info_redirect(text, wt_logfile)

        src_dir = u'E:\报表处理\金蝶报表切割\原始报表'
        dst_dir = u'E:\报表处理\金蝶报表切割\处理后报表'
        err_dir = u'E:\报表处理\金蝶报表切割\\z有错误的报表'

        chk_folder(src_dir, wt_logfile, dst_dir, err_dir)
        copy_files(src_dir, dst_dir, tag='file', wt_logfile=wt_logfile)
        KingdeeFileCut.core_method(dst_dir, err_dir, wt_logfile=wt_logfile)
        break

    elif choice == '2':
        wt_logfile = create_logfile('profit')

        text = u'选择了2\n' + \
               u'##' * 14 + '\n' + \
               u'#  开始运行  利润表统计脚本  #\n' + \
               u'##' * 14
        output_info_redirect(text, wt_logfile)

        src_dir = u'E:\报表处理\利润表\原始报表'
        dst_dir = u'E:\报表处理\利润表\处理后报表'
        err_dir = u'E:\报表处理\利润表\\z有错误的报表'

        chk_folder(src_dir, wt_logfile, dst_dir, err_dir)
        copy_files(src_dir, dst_dir, tag='tree', wt_logfile=wt_logfile)
        ProfitCalculate.core_method(dst_dir, err_dir, wt_logfile=wt_logfile)
        break

    elif choice == '3':
        wt_logfile = create_logfile('proportion')

        text = u'选择了3\n' + \
               u'##' * 14 + '\n' + \
               u'# 开始运行 费用占比统计脚本 #\n' + \
               u'##' * 14
        output_info_redirect(text, wt_logfile)

        dst_path = u'E:\报表处理\费用占比'
        file_path = chk_proportion_file(dst_path, wt_logfile)
        src_dir = u'E:\报表处理\费用占比\原始报表'
        chk_folder(src_dir, wt_logfile)
        merge_proportion_result(file_path, src_dir, wt_logfile)

        break

    elif choice == '4':
        print (u'选择了4, 将直接退出')
        time.sleep(1)
        break
    else:
        print (u'选择错误, 请重新选择!')

sys.exit(0)
