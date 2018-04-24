# encoding:utf-8
import xlrd
# import xlwt, xlutils, xlsxwriter
from xlutils.copy import copy
from openpyxl import load_workbook
# import pandas
from win32com.client import Dispatch
import win32com.client
import os
import shutil
import time


def _del_path(path, re_create=0):
    # 传递一个文件夹作为参数path, 如果文件夹不为空, 则彻底删除path这个文件夹
    # 如果re_create == 1, 则在相同位置, 重新创建一个同名文件夹(不创建子文件/子文件夹)
    if os.path.exists(path) and os.path.isdir(path):
        if os.listdir(path):
            # for i in os.listdir(path):
            #     path_file = os.path.join(path, i)
            #     if os.path.isfile(path_file):
            #         os.remove(path_file)
            #     if os.path.isdir(path_file):
            #         if len(os.listdir(path_file)) != 0:
            #             _del_path(path_file)
            shutil.rmtree(path)
            if re_create == 1:
                os.makedirs(path)
    return


def _move_err_sub_file(sub_file, err_dir, src_dir, tag):
    rd_cp_file = os.path.join(src_dir, sub_file)
    cp_to_file = os.path.join(err_dir, sub_file)
    if tag == 'chk_profit_src':
        print (u'发现错误!\n'
               u'%s 目录下只能存放文件夹, %s 是一个文件\n'
               u'将把它移动到 [错误的原始报表] 中...' % (src_dir, sub_file))
        rd_cp_file = sub_file
        sub_name = sub_file.split(os.path.sep)[-1]
        cp_to_file = os.path.join(err_dir, sub_name)
    elif tag == 'chk_excel_profit':
        print (u'发现错误!\n'
               u'%s 里有不是Excel的文件, 或它是个空文件夹\n'
               u'将把它移动到 [错误的原始报表] 中...' % sub_file)
        rd_cp_file = sub_file
        sub_name = sub_file.split(os.path.sep)[-1]
        cp_to_file = os.path.join(err_dir, sub_name)

    elif tag == 'chk_excel_kingdee':
        print (u'发现错误!\n'
               u'%s 不是Excel文件, 将把它移动到 [错误的原始报表] 中...' % sub_file)
        rd_cp_file = os.path.join(src_dir, sub_file)
        cp_to_file = os.path.join(err_dir, sub_file)

    shutil.move(rd_cp_file, cp_to_file)
    print (u'移动完成!\n')
    return


def _chk_kingdee_src(src_dir, err_dir):
    print (u'正在检查 %s 是否符合 金蝶报表切割 的原始文件规则...' % src_dir)
    tag = 'chk_excel_kingdee'
    for chk_file in os.listdir(src_dir):
        if chk_file.split('.')[-1] != 'xls' and chk_file.split('.')[-1] != 'xlsx':
            _move_err_sub_file(chk_file, err_dir, src_dir, tag)
            # print (u'发现错误!\n'
            #        u'%s 不是Excel文件, 将把它移动到 [错误的原始报表] 中...' % chk_file)
            # sub_file = os.path.join(src_dir, chk_file)
            # cp_to_file = os.path.join(err_dir, chk_file)
            # shutil.move(sub_file, cp_to_file)
            # print (u'移动完成!\n')
    print (u'检查完毕, OK!')
    return


def _chk_profit_src(src_dir, err_dir):
    print (u'正在检查 %s 是否符合 利润表 的原始文件规则...' % src_dir)
    tag = 'chk_excel_profit'
    for sub_dir in os.listdir(src_dir):
        sub_dir = os.path.join(src_dir, sub_dir)
        if os.path.isdir(sub_dir):
            if len(os.listdir(sub_dir)) != 0:
                for chk_file in os.listdir(sub_dir):
                    if chk_file.split('.')[-1] != 'xls' and chk_file.split('.')[-1] != 'xlsx':
                        _move_err_sub_file(sub_dir, err_dir, src_dir, tag)
                        continue
            else:
                _move_err_sub_file(sub_dir, err_dir, src_dir, tag)
        else:
            tag = 'chk_profit_src'
            _move_err_sub_file(sub_dir, err_dir, src_dir, tag)
            # print (u'发现错误!\n'
            #        u'%s 目录下只能存放文件夹, %s 是一个文件\n'
            #        u'将把它移动到 [错误的原始报表] 中...' % (src_dir, sub_dir))
            # sub_file = os.path.join(src_dir, sub_dir)
            # cp_to_file = os.path.join(err_dir, sub_dir)
            # shutil.move(sub_file, cp_to_file)
            # print (u'移动完成!\n')
    print (u'检查完毕, OK!')
    return


def chk_folder(src_dir, dst_dir, err_dir):
    # err_dir 为原始报表中, 发现错误的文件/文件夹存放的地方
    # src_dir 为原始报表存放的文件夹
    # dst_dir 为处理后报表存放的文件夹
    # 1: 检查 err_dir 是否为空, 不为空则彻底删除并重新创建
    print (u'正在检查 %s ...\n' % err_dir)
    if os.path.exists(err_dir) and os.path.isdir(err_dir):
        print (u'%s 是一个已存在的文件夹, 正在进行清理...' % err_dir)
        _del_path(err_dir, 1)
    elif os.path.isfile(err_dir):
        print (u'%s 是一个已存在的文件, 准备删除...' % err_dir)
        os.remove(err_dir)
        os.makedirs(err_dir)
    else:
        print (u'没有找到 %s , 准备创建该文件夹...' % err_dir)
        os.makedirs(err_dir)
    print (u'检查完毕!\n\n'
           u'开始检查 %s 和 %s ...' % (src_dir, dst_dir))
    if not os.path.isdir(src_dir) or not os.path.isdir(dst_dir):
        # 没有 src_dir 和 dst_dir 时的处理
        print(u'发现错误...\n'
              u'%s 和 %s\n'
              u'必须存在, 且必须是 [文件夹] , 请确认!\n' % (src_dir, dst_dir))
        print (u'请按任意键退出...')
        raw_input()
        exit(1)
    elif len(os.listdir(src_dir)) == 0:
        # src_dir 为空文件夹时的处理
        print(u'发现错误...\n'
              u'原始文件夹 %s 不能为空文件夹\n' % src_dir)
        print (u'请按任意键退出...')
        raw_input()
        exit(1)
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
                print (u'选择了1(自动备份), 系统将在清空 [处理后报表-历史备份] 文件夹后, 再进行备份...')
                mypath = os.path.dirname(dst_dir)
                cp_to_dir = os.path.join(mypath, u'处理后报表-历史备份')
                _del_path(cp_to_dir)
                print (u'清空完毕, 开始备份...')
                shutil.move(dst_dir, cp_to_dir)
                os.makedirs(dst_dir)
                print (u'备份完毕!')
                break
            elif get_var == '2':
                print (u'选择了2(手动备份), 系统将在3秒后退出\n')
                time.sleep(3)
                exit(0)
            elif get_var == '3':
                print (u'选择了3(直接删除), 系统将直接删除 %s 文件内容' % dst_dir)
                _del_path(dst_dir, 1)
                print (u'删除完毕!')
                break

    print (u'检查完毕!\n\n'
           u'开始检查子文件/子文件夹...')
    pjt_name = os.path.dirname(src_dir).split(os.path.sep)[-1]
    if pjt_name == u'金蝶报表切割':
        _chk_kingdee_src(src_dir, err_dir)
    elif pjt_name == u'利润表':
        _chk_profit_src(src_dir, err_dir)
    return


def copy_files(src_dir, dst_dir, tag):
    print (u'开始复制文件!\n'
           u'正在将原始文件/文件夹 从 %s'
           u'复制到 %s' % (src_dir, dst_dir))
    for filename in os.listdir(src_dir):
        rd_cp_file = os.path.join(src_dir, filename)
        cp_to_file = os.path.join(dst_dir, filename)
        if tag == 'tree':
            shutil.copytree(rd_cp_file, cp_to_file)
        elif tag == 'file':
            shutil.copyfile(rd_cp_file, cp_to_file)
    print (u'文件复制完毕\n')
    return


def copy_sheet(from_file, cp_to_file, from_sht_num, cp_to_sht_num=2, cp_to_sht_name=None):
    # from_file 和 cp_to_file 是待处理文件的完整路径
    # print (u'common - copy_sheet 接收到数据:'
    #        u'%s, %s, %s, %s, %s' % (from_file, cp_to_file, from_sht_num, cp_to_sht_num, cp_to_sht_name))
    if from_sht_num < 1 or cp_to_sht_num < 2:
        print (u'无效的sheet表编号\n')
        return
    if from_file != cp_to_file:
        print (u'开始复制sheet表\n'
               u'将 %s 的 sheet表%s\n'
               u'复制到 %s 的 sheet表%s' % (from_file, from_sht_num, cp_to_file, cp_to_sht_num))
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
        print (u'开始复制sheet表\n'
               u'将 %s 的 sheet表%s\n'
               u'复制到 sheet表%s' % (cp_to_file, from_sht_num, cp_to_sht_num))
        xlapp = Dispatch('Excel.Application')
        xlbook = xlapp.Workbooks.Open(Filename=cp_to_file)
        xlshts = xlbook.Worksheets
        xlshts(from_sht_num).Copy(None, xlshts(cp_to_sht_num - 1))
        xlbook.Save()
        xlbook.Close(SaveChanges=True)
    print (u'sheet表复制完成, 开始修改sheet名称...')

    if cp_to_file.split('.')[-1] == 'xls':
        rb = xlrd.open_workbook(cp_to_file, formatting_info=True)
        wb = copy(rb)
        if cp_to_sht_name is not None:
            ws = wb.get_sheet(1)
            ws.set_name(cp_to_sht_name)
        if cp_to_sht_name != 'Page1_Copy':
            wb.set_active_sheet(0)
        wb.save(cp_to_file)
    elif cp_to_file.split('.')[-1] == 'xlsx':
        wb = load_workbook(cp_to_file)
        if cp_to_sht_name is not None:
            wb.worksheets[1].title = cp_to_sht_name
        if cp_to_sht_name != 'Page1_Copy':
            wb.active = 0
        wb.save(cp_to_file)
        wb.close()
    if cp_to_sht_name is not None:
        print (u'已将复制后的sheet名称修改为 %s' % cp_to_sht_name)
    else:
        print (u'检测到无需修改sheet名称. 将跳过本步骤...')

    print (u'sheet工作表 复制完成!\n')
    return


def change_sheet_name(filename, sht_name, sht_num=0):
    file_type = filename.split('.')[-1]
    if file_type == 'xls':
        rb = xlrd.open_workbook(filename, formatting_info=True)
        wb = copy(rb)
        ws = wb.get_sheet(sht_num)
        ws.set_name(sht_name)
        wb.save(filename)
    elif file_type == 'xlsx':
        wb = load_workbook(filename)
        wb.worksheets[sht_num].title = sht_name
        wb.save(filename)
        wb.close()
    return
