# encoding:utf-8
import os
import sys
import time
from common import chk_folder, copy_files, output_info_redirect
import KingdeeFileCut
import ProfitCalculate


def create_logfile(pj_name):
    log_dir = u'E:\报表处理\日志'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    now_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
    logfile_name = now_time + '_' + pj_name + '.txt'
    logfile_tmp = os.path.join(log_dir, logfile_name)
    return logfile_tmp


while True:
    print (u'脚本启动! \n'
           u'请选择需要操作的项目: \n'
           u'1: 处理金蝶导出文件 \n'
           u'2: 利润表 \n'
           u'3: 直接退出 \n')
    choice = raw_input()

    if choice == '1':
        logfile = create_logfile('kingdee')
        wt_logfile = open(logfile, 'a+')

        text = u'选择了1\n' + \
               u'##' * 14 + '\n' + \
               u'#  开始运行 金蝶报表切割脚本  #\n' + \
               u'##' * 14
        output_info_redirect(text, wt_logfile)

        src_dir = u'E:\报表处理\金蝶报表切割\原始报表'
        dst_dir = u'E:\报表处理\金蝶报表切割\处理后报表'
        err_dir = u'E:\报表处理\金蝶报表切割\错误的原始报表'

        chk_folder(src_dir, dst_dir, err_dir, wt_logfile=wt_logfile)
        copy_files(src_dir, dst_dir, tag='file', wt_logfile=wt_logfile)
        KingdeeFileCut.core_method(dst_dir, wt_logfile=wt_logfile)

        wt_logfile.close()
        break
    elif choice == '2':
        logfile = create_logfile('profit')
        wt_logfile = open(logfile, 'a+')

        text = u'选择了2\n' + \
               u'##' * 14 + '\n' + \
               u'#  开始运行  利润表统计脚本  #\n' + \
               u'##' * 14
        output_info_redirect(text, wt_logfile)

        src_dir = u'E:\报表处理\利润表\原始报表'
        dst_dir = u'E:\报表处理\利润表\处理后报表'
        err_dir = u'E:\报表处理\利润表\错误的原始报表'

        chk_folder(src_dir, dst_dir, err_dir, wt_logfile=wt_logfile)
        copy_files(src_dir, dst_dir, tag='tree', wt_logfile=wt_logfile)
        ProfitCalculate.core_method(dst_dir, err_dir, wt_logfile=wt_logfile)

        wt_logfile.close()
        break
    elif choice == '3':
        print (u'选择了3, 将直接退出')
        time.sleep(1)
        break
    else:
        print (u'选择错误, 请重新选择!')

sys.exit(0)
