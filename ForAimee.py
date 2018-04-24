# encoding:utf-8
from common import chk_folder, copy_files
import KingdeeFileCut, ProfitCalculate

while True:
    print (u'脚本启动! \n'
           u'请选择需要操作的项目: \n'
           u'1: 处理金蝶导出文件 \n'
           u'2: 利润表 \n'
           u'3: 直接退出 \n')
    choice = raw_input()

    if choice == '1':
        print(u'选择了1')
        print (u'##' * 14 + '\n' +
               u'#  开始运行 金蝶报表切割脚本  #\n' +
               u'##' * 14)
        src_dir = u'E:\报表处理\金蝶报表切割\原始报表'
        dst_dir = u'E:\报表处理\金蝶报表切割\处理后报表'
        err_dir = u'E:\报表处理\金蝶报表切割\错误的原始报表'

        chk_folder(src_dir, dst_dir, err_dir)
        copy_files(src_dir, dst_dir, tag='file')
        KingdeeFileCut.core_method(dst_dir)
        break
    elif choice == '2':
        print(u'选择了2')
        print (u'##' * 14 + '\n' +
               u'#  开始运行  利润表统计脚本  #\n' +
               u'##' * 14)
        src_dir = u'E:\报表处理\利润表\原始报表'
        dst_dir = u'E:\报表处理\利润表\处理后报表'
        err_dir = u'E:\报表处理\利润表\错误的原始报表'

        chk_folder(src_dir, dst_dir, err_dir)
        copy_files(src_dir, dst_dir, tag='tree')
        ProfitCalculate.core_method(dst_dir, err_dir)
        break
    elif choice == '3':
        print(u'选择了3, 将直接退出')
        break
    else:
        print(u'选择错误, 请重新选择')
