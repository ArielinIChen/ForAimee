# encoding:utf-8
import xlrd
from xlutils.copy import copy
from openpyxl import load_workbook
from win32com.client import Dispatch
import os
import sys
import shutil
import time


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


def change_sheet_name(filename, sht_name, sht_num):
    file_type = filename.split('.')[-1]
    if file_type == 'xls':
        rb = xlrd.open_workbook(filename, formatting_info=True)
        wb = copy(rb)
        print (u'正在把sheet名修改为 %s ...' % sht_name)
        if sht_name is not None:
            ws = wb.get_sheet(sht_num)
            ws.set_name(sht_name)
        if sht_name != 'Page1_Copy':
            print (u'正在把默认显示的sheet表修改为第一个sheet')
            wb.set_active_sheet(0)
        wb.save(filename)
    elif file_type == 'xlsx':
        wb = load_workbook(filename)
        print (u'将要把sheet名修改为 %s ...' % sht_name)
        if sht_name is not None:
            wb.worksheets[sht_num].title = sht_name
        if sht_num != 'Page_Copy':
            print (u'正在把默认显示的sheet表修改为第一个sheet...')
            wb.active = 0
        wb.save(filename)
        wb.close()
    return


def _del_path(path, wt_logfile, re_create=1):
    # 判断path是文件or文件夹, 如果是文件, 则彻底删除; 如果是文件夹, 则清空文件夹下内容
    # 传递一个文件夹作为参数path, 如果文件夹不为空, 则彻底删除path这个文件夹
    # 如果re_create == 1, 则在相同位置, 重新创建一个同名文件夹(不创建子文件/子文件夹)
    # if os.path.exists(path) and os.path.isdir(path):
    #     if os.listdir(path):
    #         # for i in os.listdir(path):
    #         #     path_file = os.path.join(path, i)
    #         #     if os.path.isfile(path_file):
    #         #         os.remove(path_file)
    #         #     if os.path.isdir(path_file):
    #         #         if len(os.listdir(path_file)) != 0:
    #         #             _del_path(path_file)
    #         shutil.rmtree(path)
    #         if re_create == 1:
    #             os.makedirs(path)
    if os.path.isfile(path):
        text = u'%s 是一个已存在的文件, 准备删除...' % path
        output_info_redirect(text, wt_logfile)
        os.remove(path)
    elif os.path.isdir(path):
        text = u'%s 是一个已存在的文件夹, 准备清理...' % path
        output_info_redirect(text, wt_logfile)
        shutil.rmtree(path)
    else:
        text = u'%s 原本就不存在, 无需删除/清理...' % path
        output_info_redirect(text, wt_logfile)

    text = u'检查完毕! 准备重新创建文件夹 %s ...' % path
    output_info_redirect(text, wt_logfile)
    if re_create == 1:
        while True:
            os.makedirs(path)
            if os.path.exists(path):
                text = u'重新创建完毕!'
                output_info_redirect(text, wt_logfile)
                break
    else:
        text = u'检查到: 无需重新创建文件夹. 将跳过该步骤...'
        output_info_redirect(text, wt_logfile)
    return


def _move_err_sub_file(filename, to_dir, from_dir, tag, wt_logfile):
    rd_cp_file = os.path.join(from_dir, filename)
    cp_to_file = os.path.join(to_dir, filename)
    if tag.startswith('chk_kingdee'):
        text = u'发现错误!\n' \
               u'%s 不是Excel文件 也 不是一个文件, 将把它移动到 [有错误的原始报表] 中...\n' % filename
        output_info_redirect(text, wt_logfile)

        rd_cp_file = os.path.join(from_dir, filename)
        cp_to_file = os.path.join(to_dir, filename)

    elif tag == 'chk_profit_src_isfile':
        text = u'发现错误!\n' \
               u'%s 目录下只能存放文件夹, %s 是一个文件\n' \
               u'将把它移动到 [有错误的原始报表] 中...\n' % (from_dir, filename)
        output_info_redirect(text, wt_logfile)

        rd_cp_file = filename
        sub_name = filename.split(os.path.sep)[-1]
        cp_to_file = os.path.join(to_dir, sub_name)

    elif tag == 'chk_profit_not_excel' or tag == 'chk_profit_empty_subdir':
        text = u'发现错误!\n' \
               u'%s 里有不是Excel的文件, 或它是个空文件夹\n' \
               u'将把它移动到 [有错误的原始报表] 中...\n' % filename
        output_info_redirect(text, wt_logfile)

        rd_cp_file = filename
        sub_name = filename.split(os.path.sep)[-1]
        cp_to_file = os.path.join(to_dir, sub_name)

    shutil.move(rd_cp_file, cp_to_file)
    text = u'移动完成!\n'
    output_info_redirect(text, wt_logfile)
    return


def _chk_kingdee_src(src_dir, err_src_dir, wt_logfile):
    text = u'正在检查 %s 是否符合 金蝶报表切割 的原始文件规则...' % src_dir
    output_info_redirect(text, wt_logfile)

    for chk_file in os.listdir(src_dir):
        if os.path.isdir(chk_file):
            tag = 'chk_kingdee_isdir'
            os.path.exists(err_src_dir) and 1 or os.mkdir(err_src_dir)
            _move_err_sub_file(filename=chk_file,
                               to_dir=err_src_dir,
                               from_dir=src_dir,
                               tag=tag, wt_logfile=wt_logfile)
        elif chk_file.split('.')[-1] != 'xls' and chk_file.split('.')[-1] != 'xlsx':
            tag = 'chk_kingdee_not_excel'
            os.path.exists(err_src_dir) and 1 or os.mkdir(err_src_dir)
            _move_err_sub_file(filename=chk_file,
                               to_dir=err_src_dir,
                               from_dir=src_dir,
                               tag=tag, wt_logfile=wt_logfile)

    text = u'检查完毕, OK!'
    output_info_redirect(text, wt_logfile)
    return


def _chk_profit_src(src_dir, err_src_dir, wt_logfile):
    text = u'正在检查 %s 是否符合 利润表 的原始文件规则...' % src_dir
    output_info_redirect(text, wt_logfile)

    tag = 'chk_profit_not_excel'
    for sub_dir in os.listdir(src_dir):
        sub_dir = os.path.join(src_dir, sub_dir)
        if os.path.isdir(sub_dir):
            if len(os.listdir(sub_dir)) != 0:
                for chk_file in os.listdir(sub_dir):
                    if chk_file.split('.')[-1] != 'xls' and chk_file.split('.')[-1] != 'xlsx':
                        os.path.exists(err_src_dir) and 1 or os.mkdir(err_src_dir)
                        _move_err_sub_file(filename=sub_dir,
                                           to_dir=err_src_dir,
                                           from_dir=src_dir,
                                           tag=tag,
                                           wt_logfile=wt_logfile)
                        continue
            else:
                tag = 'chk_profit_empty_subdir'
                os.path.exists(err_src_dir) and 1 or os.mkdir(err_src_dir)
                _move_err_sub_file(filename=sub_dir,
                                   to_dir=err_src_dir,
                                   from_dir=src_dir,
                                   tag=tag,
                                   wt_logfile=wt_logfile)
        else:
            tag = 'chk_profit_src_isfile'
            os.path.exists(err_src_dir) and 1 or os.mkdir(err_src_dir)
            _move_err_sub_file(filename=sub_dir,
                               to_dir=err_src_dir,
                               from_dir=src_dir,
                               tag=tag,
                               wt_logfile=wt_logfile)

    text = u'检查完毕, OK!'
    output_info_redirect(text, wt_logfile)
    return


def chk_folder(src_dir, dst_dir, err_dir, wt_logfile):
    # err_dir 为[z有错误的报表] (文件夹)
    #   根据错误类型, 向下细分文件/文件夹
    #   初始化时 清空/删除 && 重建文件夹
    # src_dir 为[原始报表] (文件夹)
    #   初始必须存在, 且不为空. 否则就 报错&&退出
    # dst_dir 为[处理后报表] (文件夹)
    #   初始不存在则创建文件夹
    #   已存在 且 为文件时: 加上 backup_at_时间戳 标记, 然后后自动备份到同级目录, 再创建文件夹
    #   已存在 且 为空文件夹时: pass
    #   已存在 且 不为空文件夹时: 选择是否需要(自动)备份, 自动备份文件夹为: 处理后报表-历史备份
    # 开始检查err_dir
    text = u'开始检查 %s ...\n' % err_dir
    output_info_redirect(text, wt_logfile)

    _del_path(err_dir, wt_logfile, re_create=1)

    # 开始检查src_dir
    text = u'\n检查完毕!\n\n' \
           u'开始检查 %s...' % src_dir
    output_info_redirect(text, wt_logfile)
    if not os.path.isdir(src_dir):
        # 没有 src_dir 或 src_dir 不为文件夹时
        text = u'发现错误...\n' \
               u'%s\n' \
               u'必须存在, 且必须是 [文件夹] , 请确认!\n' \
               u'请按任意键退出...' % src_dir
        output_info_redirect(text, wt_logfile)
        raw_input()
        time.sleep(1)
        sys.exit(1)
    elif len(os.listdir(src_dir)) == 0:
        # src_dir 为空文件夹时的处理
        text = u'发现错误...\n' \
               u'原始文件夹 %s 不能为空文件夹\n' \
               u'请按任意键退出...' % src_dir
        output_info_redirect(text, wt_logfile)
        raw_input()
        time.sleep(1)
        sys.exit(1)

    # 开始检查dst_dir
    text = u'\n检查完毕!\n\n' \
           u'开始检查 %s...' % dst_dir
    output_info_redirect(text, wt_logfile)
    if not os.path.exists(dst_dir):
        os.mkdir(dst_dir)
    elif os.path.isfile(dst_dir):
        now_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
        dst_dir_file_back = dst_dir + '_backup_at_' + now_time
        shutil.move(dst_dir, dst_dir_file_back)
        os.mkdir(dst_dir)
    elif len(os.listdir(dst_dir)) != 0:
        # dst_dir 不为空文件夹时的处理
        while True:
            print (u'警告! 警告! 警告!\n'
                   u'警告! 警告! 警告!\n'
                   u'%s 不是空文件夹, 需要进行一下哪项操作:\n'
                   u'1: 自动备份到同级的 [处理后报表-历史备份] 目录(该目录原有文件将被删除)\n'
                   u'2: 退出脚本, 手动备份\n'
                   u'3: 直接删除该文件夹下内容\n' % dst_dir)
            get_var = raw_input()

            if get_var == '1':
                text = u'\n选择了1(自动备份), 系统将在清空 [处理后报表-历史备份] 文件夹后, 再进行备份...\n'
                output_info_redirect(text, wt_logfile)

                mypath = os.path.dirname(dst_dir)
                cp_to_dir = os.path.join(mypath, u'处理后报表-历史备份')
                _del_path(cp_to_dir, wt_logfile, re_create=0)
                text = u'清空完毕, 开始备份...'
                output_info_redirect(text, wt_logfile)

                shutil.move(dst_dir, cp_to_dir)
                os.makedirs(dst_dir)
                text = u'备份完毕!'
                output_info_redirect(text, wt_logfile)
                break
            elif get_var == '2':
                text = u'\n选择了2(手动备份), 系统将在2秒后退出\n'
                output_info_redirect(text, wt_logfile)
                time.sleep(2)
                sys.exit(0)
            elif get_var == '3':
                text = u'\n选择了3(直接删除), 系统将直接删除 %s 文件内容' % dst_dir
                output_info_redirect(text, wt_logfile)

                _del_path(dst_dir, wt_logfile, re_create=1)
                text = u'删除完毕!'
                output_info_redirect(text, wt_logfile)
                break

    text = u'\n开始检查子文件/子文件夹...'
    output_info_redirect(text, wt_logfile)

    pjt_name = os.path.dirname(src_dir).split(os.path.sep)[-1]
    err_src_dir = os.path.join(err_dir, u'有错误的原始报表')
    if pjt_name == u'金蝶报表切割':
        _chk_kingdee_src(src_dir, err_src_dir, wt_logfile)
    elif pjt_name == u'利润表':
        _chk_profit_src(src_dir, err_src_dir, wt_logfile)
    return


def copy_files(src_dir, dst_dir, tag, wt_logfile):
    text = u'\n开始复制文件!\n' \
           u'正在将原始文件/文件夹 从 %s\n' \
           u'复制到 %s' % (src_dir, dst_dir)
    output_info_redirect(text, wt_logfile)

    for filename in os.listdir(src_dir):
        rd_cp_file = os.path.join(src_dir, filename)
        cp_to_file = os.path.join(dst_dir, filename)
        if tag == 'tree':
            shutil.copytree(rd_cp_file, cp_to_file)
        elif tag == 'file':
            shutil.copyfile(rd_cp_file, cp_to_file)
    text = u'文件复制完毕\n'
    output_info_redirect(text, wt_logfile)
    return


def copy_sheet(from_file, cp_to_file, from_sht_num, cp_to_sht_num=2, cp_to_sht_name=None, wt_logfile=None):
    # from_file 和 cp_to_file 是待处理文件的完整路径
    if from_sht_num < 1 or cp_to_sht_num < 2:
        text = u'无效的sheet表编号\n'
        output_info_redirect(text, wt_logfile)
        return
    if from_file != cp_to_file:
        text = u'开始复制sheet表\n' \
               u'将 %s 的 sheet表%s\n' \
               u'复制到 %s 的 sheet表%s' % (from_file, from_sht_num, cp_to_file, cp_to_sht_num)
        output_info_redirect(text, wt_logfile)

        xlapp = Dispatch('Excel.Application')
        from_book = xlapp.Workbooks.Open(Filename=from_file)
        from_shts = from_book.Worksheets
        cp_to_book = xlapp.Workbooks.Open(Filename=cp_to_file)
        cp_to_shts = cp_to_book.Worksheets
        from_shts(from_sht_num).Copy(None, cp_to_shts(cp_to_sht_num - 1))
        # from_book.Save()
        # cp_to_book.Save()
        from_book.Close(SaveChanges=True)
        cp_to_book.Close(SaveChanges=True)

    else:
        text = u'开始复制sheet表\n' \
               u'将 %s 的 sheet表%s\n' \
               u'复制到 sheet表%s' % (cp_to_file, from_sht_num, cp_to_sht_num)
        output_info_redirect(text, wt_logfile)

        xlapp = Dispatch('Excel.Application')
        xlbook = xlapp.Workbooks.Open(Filename=cp_to_file)
        xlshts = xlbook.Worksheets
        xlshts(from_sht_num).Copy(None, xlshts(cp_to_sht_num - 1))
        xlbook.Save()
        xlbook.Close(SaveChanges=True)
    text = u'\nsheet表复制完成, 开始修改sheet名称...\n'
    output_info_redirect(text, wt_logfile)

    # 修改sheet名称
    change_sheet_name(cp_to_file, cp_to_sht_name, sht_num=1)

    if cp_to_sht_name is not None:
        text = u'已将复制后的sheet名称修改为 %s' % cp_to_sht_name
        output_info_redirect(text, wt_logfile)
    else:
        text = u'检测到无需修改sheet名称. 将跳过本步骤...'
        output_info_redirect(text, wt_logfile)

    text = u'sheet工作表 复制完成!\n'
    output_info_redirect(text, wt_logfile)
    return
