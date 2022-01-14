from pandas import read_excel

a = r'F:\filePython\PycharmProject\V3Policy\PolicyFile\20220113202749\ServerPolicy\PolicyUpload20220113202749.xlsx'

TPFOProUsecols = ['ProcessFlowName', 'StepID', 'PPID', 'EQID', 'BackupEQID', ]
df = read_excel(a, sheet_name='ProcessFlow', usecols=TPFOProUsecols)
print(df)

NewFile_dfTPFO = 'F:\\' + "TPFO.xls"
df.to_excel(NewFile_dfTPFO, sheet_name='Machine', index=False)