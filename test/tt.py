'''
这是测试用的文件，非必要文件
'''
from pandas import read_excel

a = 'F:/filePython/PycharmProject/V3Policy/PolicyFile/20220118133133/ServerPolicy/PolicyUpload20220118133133.xlsx'
b = r'F:\jm\POS\Policy数据导入模板.xlsx'

TPFOProUsecols = ['ProcessFlowName', 'StepID', 'PPID', 'EQID', 'BackupEQID', ]
df = read_excel(b, sheet_name='ProcessFlow', usecols=TPFOProUsecols)
print(df)

# NewFile_dfTPFO = 'F:\\' + "TPFO.xls"
# df.to_excel(NewFile_dfTPFO, sheet_name='Machine', index=False)