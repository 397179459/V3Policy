'''
@wanglei03  @V0019729
这个文件是用来处理TPFO的，主要用到ProcessFlow，ReworkFlow，StripperFlow 都要用到，处理逻辑都大概类似
'''
import logging
import os
import zipfile

from pandas import read_excel

import policy
import TFOMParse

logger = logging.getLogger()
# logger.setLevel(logging.INFO)
logger.setLevel(policy.logLevel)


# 处理 ProcessFlow sheet
def processFlow(oldPath):
    # 处理ProcessFlow
    logging.debug('===============  开始处理 ProcessFlow Sheet ===============')
    TPFOProUsecols = ['ProcessFlowName', 'StepID', 'PPID', 'EQID', 'BackupEQID', ]
    # TPFOProUsecols=['FactoryName', 'ProductSpecName', 'ProcessFlowName',
    #             'StepID', 'PPID', 'EQID', 'BackupEQID',]
    # op = oldPath.replace('\\', '//')

    dfProcess = read_excel(oldPath, sheet_name='ProcessFlow', usecols=TPFOProUsecols)
    dfP = dfProcess

    # 去除单元格字符前后空格
    dfP = dfP.applymap(lambda x: str(x).strip() if type(x) == str else x)

    logging.debug('## 开始 删除 Shipp operation ...')
    global noNeedOper
    noNeedOper = ['1X999', '2X998', '2X999', '4X999', '3X999', '5X999', ]
    dfP = dfP[~dfP['StepID'].isin(noNeedOper)]

    logging.debug('## 开始 处理拆分 Main 设备 ...')
    # 删除backup eq之后，对于main eq按照换行符拆分成多行，并且设置rolltype为main
    dfPAllM = dfP.drop(['BackupEQID'], axis=1)
    dfPAllM = dfPAllM.drop(['EQID'], axis=1).join(
        dfP['EQID'].str.split('\n', expand=True).stack().reset_index(level=1, drop=True).rename('EQID'))
    dfPAllM['rollType'] = 'Main'
    logging.debug('## 找到 {0} 行Main 设备'.format(dfPAllM.shape[0]))

    logging.debug('## 开始 处理拆分 BackUp 设备 ...')
    # 删除主EQ列之后，判断一下 backup 是否全部为空，如果全部为空则给赋值一个空的dataframe；如果不为空，处理逻辑同上
    dfPAllB = dfP.drop(['EQID'], axis=1)
    dfPAllBNotN = dfPAllB[dfPAllB['BackupEQID'].notnull()]
    if (dfPAllBNotN.size > 0):
        dfPAllB = dfPAllBNotN.drop(['BackupEQID'], axis=1).join(
            dfPAllBNotN['BackupEQID'].str.split('\n', expand=True).stack().reset_index(level=1, drop=True).rename(
                'EQID'))
        dfPAllB['rollType'] = 'Backup'
        logging.debug('## 找到 {0} 行backup 设备'.format(dfPAllB.shape[0]))
    else:
        dfPAllB = dfPAllBNotN
        logging.debug('## 没有 backup 设备')

    logging.debug('## 开始 合并 ProcessFlow中的main & backup EQ ...')
    dfPAllNew = dfPAllM.append(dfPAllB)
    dfPAllNew = dfPAllNew.reset_index(drop=True)

    newPRSColumns = {'ProcessFlowName': 'processFlowName',
                     'StepID': 'processOperationName', 'PPID': 'machineRecipeName', 'EQID': 'machineName', }
    addCols(dfPAllNew, 'ProcessFlow', newPRSColumns)

    logging.debug('===============  处理完成 ProcessFlow Sheet ===============\n')


# 处理 rework 和 stripper sheet
def ExcelTPFO(oldPath, sheetName):
    logging.debug('===========  开始处理 {0} Sheet  ==========='.format(sheetName))

    #     TPFORSUsecols=['FactoryName', 'ProductSpecName', 'ProcessFlow', 'POName', 'EQID', 'BackupEQID', 'PPID',]
    TPFORSUsecols = ['ProcessFlow', 'POName', 'EQID', 'BackupEQID', 'PPID', ]
    dfRS = read_excel(oldPath, sheet_name=sheetName, usecols=TPFORSUsecols)
    #     sheet是否有数据
    if dfRS.size > 0:
        # 去除单元格字符前后空格
        dfRS = dfRS.applymap(lambda x: str(x).strip() if type(x) == str else x)

        # 删除 ship operation
        logging.debug('## 开始 删除 Ship operation ...')
        dfRS = dfRS[~dfRS['POName'].isin(noNeedOper)]

        # Rework 和Stripper Flow 存在一个 Flow 跨多行的情况，针对FlowName 单独处理一下
        dfRS['ProcessFlow'] = dfRS['ProcessFlow'].fillna(method='ffill')

        # 删除backup eq之后，对于main eq按照换行符拆分成多行，并且设置rolltype为main
        logging.debug('## 开始 处理拆分 Main 设备 ...')
        dfRSAllM = dfRS.drop(['BackupEQID'], axis=1)
        dfRSAllM = dfRSAllM.drop(['EQID'], axis=1).join(
            dfRSAllM['EQID'].str.split('\n', expand=True).stack().reset_index(level=1, drop=True).rename('EQID'))
        dfRSAllM['rollType'] = 'Main'
        logging.debug('## 找到 {0} 行 Main 设备'.format(dfRSAllM.shape[0]))

        # 删除主EQ列之后，判断一下 backup 是否全部为空，如果全部为空则给赋值一个空的dataframe；如果不为空，处理逻辑同上
        logging.debug('## 开始 处理拆分 BackUp 设备 ...')
        dfRSAllB = dfRS.drop(['EQID'], axis=1)
        dfRSAllBNotN = dfRSAllB[dfRSAllB['BackupEQID'].notnull()]
        if (dfRSAllBNotN.size > 0):
            dfRSAllB = dfRSAllBNotN.drop(['BackupEQID'], axis=1).join(
                dfRSAllBNotN['BackupEQID'].str.split('\n', expand=True).stack().reset_index(level=1, drop=True).rename(
                    'EQID'))
            dfRSAllB['rollType'] = 'Backup'
            logging.debug('## 找到 {0} 行backup 设备'.format(dfRSAllB.shape[0]))
            logging.debug('## 开始 合并 ProcessFlow中的main & backup EQ ...')
            dfRSAll = dfRSAllM.append(dfRSAllB)
            dfRSAll = dfRSAll.reset_index(drop=True)
        else:
            logging.debug('## 没有 backup 设备')
            #         dfRSAllB = dfRSAllBNotN.rename({'BackupEQID':'EQID'})
            dfRSAll = dfRSAllM.reset_index(drop=True)
    # 如果是空sheet，就是空数据
    else:
        dfRSAll = dfRS

    logging.debug('===========  处理完成 {0} Sheet  ===========\n'.format(sheetName))
    return dfRSAll


# 进入处理 rework 和stripper的主程序
def RSFlow(oldPath):
    newRSColumns = {'ProcessFlow': 'processFlowName',
                    'POName': 'processOperationName', 'PPID': 'machineRecipeName', 'EQID': 'machineName', }

    dfRAllNew = ExcelTPFO(oldPath, 'ReworkFlow')
    addCols(dfRAllNew, 'ReworkFlow', newRSColumns)

    dfSAllNew = ExcelTPFO(oldPath, 'StripperFlow')
    addCols(dfSAllNew, 'StripperFlow', newRSColumns)


# 添加common的 column，都是一些默认的 version之类的
def addCols(df, sheetName, newPRSCols):
    logging.debug('## 开始 添加默认值 column  ...')

    # policy 获取的前端传的参数
    FactoryName = policy.FACTORY
    ProductSpecName = policy.PRODUCTSPEC

    # 添加Factory 和 product spec
    df[['factoryName', 'productSpecName']] = df.apply(lambda x: (FactoryName, ProductSpecName), axis=1,
                                                      result_type='expand')
    # 添加各种version，目前都是00001
    df[['productSpecVersion', 'processFlowVersion', 'processOperationVersion']] = df.apply(
        lambda x: ('00001', '00001', '00001',), axis=1, result_type='expand')
    # 各种flag
    df[['checkLevel', 'rmsFlag', 'dispatchState']] = df.apply(lambda x: ('N', 'N', 'N',), axis=1, result_type='expand')
    df[['dispatchPriority', 'ecRecipeFlag', 'ecRecipeName', 'maskCycleTarget']] = df.apply(lambda x: ('', '', '', ''),
                                                                                           axis=1, result_type='expand')
    logging.debug('## 开始 规范化 column  ...')
    # newPRSColumns = {'FactoryName':'factoryName', 'ProductSpecName':'productSpecName','ProcessFlowName':'processFlowName',
    #               'StepID':'processOperationName','PPID':'machineRecipeName','EQID':'machineName',}
    df = df.rename(columns=newPRSCols)

    # 标准的 column 位置
    stdTPFOColumns = ['factoryName', 'productSpecName', 'productSpecVersion', 'processFlowName', 'processFlowVersion',
                      'processOperationName', 'processOperationVersion', 'machineName', 'rollType', 'machineRecipeName',
                      'checkLevel', 'rmsFlag', 'dispatchState',
                      'dispatchPriority', 'ecRecipeFlag', 'ecRecipeName', 'maskCycleTarget']

    df = df[stdTPFOColumns]

    #     now = time.strftime("%Y%m%d%H%M%S", time.localtime())
    #     firstDir = pre + now
    #     os.mkdir(firstDir)
    firstDir = policy.TARGET_DIR
    newTPFOFileDir = firstDir + '\\' + sheetName + "TPFO"
    # 因为 Modeler 导入的时候对于文件名和sheet名很规范，所有这里采用不同·文件夹的方式输出
    os.mkdir(newTPFOFileDir)
    NewFile_dfTPFO = newTPFOFileDir + '\\' + "TPFO.xls"
    df.to_excel(NewFile_dfTPFO, sheet_name='Machine', index=False)
    logging.debug("Success ！ %s 导出成功" % NewFile_dfTPFO)


# 将转换之后的Excel转换成压缩包，以供下载
def zipdir(path, ziph):
    # 必须要这样写，要不然可能就丢失文件。缺点就是可能压缩后的文件深度比较深
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))
    logging.debug('Zip OK')


def zip():
    sourceFile = policy.FIRST_DIR
    zipFileName = '{0}.zip'.format(policy.now)
    targetZip = policy.TARGET_ZIP + '\\' + zipFileName
    zipfdownload = zipfile.ZipFile(targetZip, 'w', zipfile.ZIP_DEFLATED)
    zipdir(sourceFile, zipfdownload)
    zipfdownload.close()

    return zipFileName


# 所有处理程序的入口
def main(oldPath):
    logging.debug('## 开始 读取{0} .......'.format(oldPath))
    # os.mkdir(firstDir)
    # 先单独处理ProcessFlow
    processFlow(oldPath)
    # 再处理 rework 和stripper sheet
    RSFlow(oldPath)
    # 转换TFOM，这一过程最费时
    TFOMParse.main(oldPath)
    # 压缩转换后的文件
    zipFile = zip()

    # 返回压缩后的zip文件名，下载页面需要用到
    return zipFile
