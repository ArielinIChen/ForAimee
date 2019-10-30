1. 银行导出数据核对
2. 付款信息合并
3. 付款总表拆分为分表

把py文件和setup.py放在同一级，然后执行一下步骤
1 pip install cython or conda install cython
2 pyinstaller --onefile --hidden-import pandas._libs.tslibs.timedeltas program.py
  (D:\Python27\Scripts\pyinstaller.exe --onefile --hidden-import pandas._libs.tslibs.timedeltas ForAimee.py)

参考于: https://stackoverflow.com/questions/51078506/attributeerror-type-object-pandas-libs-tslib-tsobject-has-no-attribute-re
