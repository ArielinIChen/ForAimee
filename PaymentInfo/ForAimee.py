# encoding:utf-8
import os
import sys
import time
import datetime

from common import create_logfile, check_work_folder, output_info_redirect, create_xl_file

BASE_DIR = u'E:\诺互银行'
payment_total_dir = os.path.join(BASE_DIR, u'付款总计')
accountant_src_file_dir = os.path.join(BASE_DIR, u'会计给到原件')
bank_check_dir = os.path.join(BASE_DIR, u'银行核对')
bank_info_dir = os.path.join(BASE_DIR, u'银行信息资料')
voucher_dir = os.path.join(BASE_DIR, u'银行制单模板')

while True:
    print (u'脚本启动! \n'
           u'请选择需要操作的项目: \n'
           u'1: 银行导出数据核对 \n'
           u'2: 付款信息合并 \n'
           u'3: 付款总表拆分为分表 \n'
           u'4: 直接退出 \n')
    choice = raw_input()

    logfile_path = create_logfile(BASE_DIR)
    check_path_list = []    # .append([文件/路径名称, 'dir'/'file'])

    if choice == '1':
        print u'即将进行 - 银行导出数据核对\n'

        from BankExcelCheck import payment_check
        filename = u'银行核对公式.xlsx'
        filePath = os.path.join(bank_check_dir, filename)
        check_path_list.append([filePath, 'file'])

        check_work_folder(check_path_list, logfile_path)
        # check_folder('BankExcelCheck', logfile_path)
        output_info_redirect(u'开始执行银行数据核对...', logfile_path)

        payment_check(filePath)

        output_info_redirect(u'核对完毕! \n系统将在2秒后退出\n', logfile_path)
        time.sleep(2)
        break

    elif choice == '2' or choice == '3':
        while True:
            print (u'请输入付款日期(以YYYY-MM-DD格式, 如: 2019-01-01):')
            pay_date = raw_input()
            print (u'付款日期是 %s, 确认请按1, 重新输入请按2' % pay_date)
            confirm = raw_input()
            if confirm == '1':
                break
            else:
                continue

        pay_year = pay_date.split('-')[0]
        pay_month = pay_date.split('-')[1]
        pay_day = pay_date.split('-')[2]
        filename_date = '.'.join([str(pay_month), str(pay_day)])
        payment_total_xl_name = ''.join([filename_date, u'付款总计.xlsx'])
        payment_total_xl_path = os.sep.join([payment_total_dir, filename_date, payment_total_xl_name])
        split_from_xl_name = ''.join([filename_date, u'日款.xlsx'])
        split_from_xl_path = os.sep.join([payment_total_dir, filename_date, split_from_xl_name])
        bank_check_xl_path = os.path.join(bank_check_dir, u'银行核对公式.xlsx')
        bank_info_xl_path = os.path.join(bank_info_dir, u'制单银行-门店对应表.xlsx')
        payee_info_xl_path = os.path.join(voucher_dir, u'打款资料库.xlsx')

        check_path_list.append([accountant_src_file_dir, 'dir'])
        check_path_list.append([bank_check_xl_path, 'file'])
        check_path_list.append([payee_info_xl_path, 'file'])
        check_path_list.append([bank_info_xl_path, 'file'])

        if choice == '2':
            check_path_list.append([payment_total_xl_path, 'file'])
        if choice == '3':
            check_path_list.append([split_from_xl_path, 'file'])

        check_work_folder(check_path_list, logfile_path)

        if choice == '2':
            print u'即将进行 - 付款信息合并\n'

            from PaymentTotal import get_accountant_src_xl_data, merge_payment_xl
            from DependedInfo import shop_depart_relationship

            text = u'开始从 %s 获取 门店 - 记账部门 关系信息... \n' % bank_info_xl_path
            output_info_redirect(text, logfile_path)
            shop_depart_relation_data = shop_depart_relationship(bank_info_xl_path, logfile_path)

            text = u'获取完毕 \n\n' \
                   u'开始从 %s 获取 付款信息... \n' % accountant_src_file_dir
            output_info_redirect(text, logfile_path)

            accountant_src_data = get_accountant_src_xl_data(accountant_src_file_dir, shop_depart_relation_data,
                                                             pay_date, logfile_path)

            text = u'获取完毕 \n\n'
            output_info_redirect(text, logfile_path)

            merge_data = accountant_src_data['data_list']
            if 'no_depart_data_list' in accountant_src_data:
                no_depart_data_list = accountant_src_data['no_depart_data_list']
                no_depart_data_xl_path = os.path.join(voucher_dir, u'未匹配记账部门的付款信息.xlsx')

                output_info_redirect(u'即将把没有记账部门的付款信息, 写入 %s 中... \n' % no_depart_data_xl_path, logfile_path)
                create_xl_file(no_depart_data_xl_path, no_depart_data_list)
                output_info_redirect(u'写入完毕! \n', logfile_path)

            output_info_redirect(u'开始把付款信息 合并到 %s \n' % payment_total_xl_path, logfile_path)
            merge_payment_xl(payment_total_xl_path, merge_data, logfile_path)
            output_info_redirect(u'合并完毕! \n系统将在2秒后退出\n', logfile_path)
            time.sleep(2)

        elif choice == '3':
            print u'即将进行 - 付款总表拆分为分表\n'

            from MakeVoucher import get_voucher_info, create_voucher_file_use_openpyxl
            from DependedInfo import get_payee_info, get_paying_bank_info

            print u'开始[创建/检查]最终归档文件夹(E:\诺互银行\银行制单模板[华东/华中/华南/华北/有错误的付款明细])\n'
            now_time = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            for _dir_name in [u'华东', u'华中', u'华南', u'华北', u'有错误的付款明细']:
                _region_path = os.path.join(voucher_dir, _dir_name)
                if os.path.exists(_region_path):
                    if os.listdir(_region_path):
                        if _dir_name == u'有错误的付款明细':
                            print u'即将删除 %s 目录下的所有文件\n'
                            for _del_file in os.listdir(_region_path):
                                os.remove(os.path.join(_region_path, _del_file))
                            print u'删除完成. \n'
                        else:
                            backup_dir_name = _dir_name + now_time
                            output_info_redirect(u'%s 文件夹已存在并且有历史数据, 即将把它重命名为 %s; 并创建新文件夹\n'
                                                 % (_dir_name, backup_dir_name), logfile_path)
                            backup_region_path = os.path.join(voucher_dir, backup_dir_name)
                            os.rename(_region_path, backup_region_path)
                            os.makedirs(_region_path)
                else:
                    output_info_redirect(u'%s 文件夹不存在, 即将创建该文件夹\n' % _region_path, logfile_path)
                    os.makedirs(_region_path)

            output_info_redirect(u'归档文件夹检查完毕!\n\n开始从 %s 获取收款人信息.. \n'
                                 % payee_info_xl_path,  logfile_path)
            payee_data = get_payee_info(payee_info_xl_path)

            text = u'收款人信息 获取完毕 \n\n' \
                   u'开始从 %s 获取[付款账号-记账部门-所属地区]对应关系... \n' % bank_check_xl_path
            output_info_redirect(text, logfile_path)
            paying_bank_data = get_paying_bank_info(bank_check_xl_path)

            text = u'[付款账号-记账部门-所属地区]对应关系 获取完毕 \n\n' \
                   u'开始从 %s 获取付款信息... \n' % split_from_xl_name
            output_info_redirect(text, logfile_path)
            tmp_voucher_data = get_voucher_info(split_from_xl_path, pay_date, payee_data, paying_bank_data, logfile_path)
            output_info_redirect(u'付款信息 获取完毕 \n\n', logfile_path)

            if 'error_data' in tmp_voucher_data:
                error_data = tmp_voucher_data['error_data']
                error_data_xl_path = os.sep.join([voucher_dir, u'有错误的付款明细', u'error_data.xlsx'])
                output_info_redirect(u'存在错误的付款明细, 即将把相关条目写入 %s 文件 \n\n' % error_data_xl_path, logfile_path)
                create_xl_file(error_data_xl_path, error_data)
                output_info_redirect(u'写入完毕!\n\n', logfile_path)

            voucher_data = tmp_voucher_data['voucher_data']
            text = u'付款信息 获取完毕 \n\n' \
                   u'开始拆分付款信息到 %s ...\n' % voucher_dir
            output_info_redirect(text, logfile_path)
            create_voucher_file_use_openpyxl(voucher_dir, voucher_data, logfile_path)

            output_info_redirect(u'总表拆分已完成! \n系统将在2秒后退出\n', logfile_path)
            time.sleep(2)
        break

    elif choice == '4':
        print u'选择了4[直接退出], 将在2秒后退出...'
        time.sleep(2)
        break

    else:
        print u'选择错误, 请重新选择!'

sys.exit(0)
