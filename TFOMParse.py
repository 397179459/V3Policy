'''
@wanglei03  @V0019729
这个文件是用来处理TFOM的，主要用到ProcessFlow sheet
'''
import logging
import os
import pandas as pd
from pandas import read_excel
import policy

logger = logging.getLogger()
# logger.setLevel(logging.INFO)
logger.setLevel(policy.logLevel)


# 处理TFOM的主程序
def main(fileAbsPath):
    logging.debug('## 开始 读取{0} .......'.format(fileAbsPath))
    # TFOM 需要用到的字段，要求Excel里面的字段必须是这个名字，大小写，空格之类的严格遵守
    TFOMProUsecols = ['ProcessFlowName', 'Category1',
                      'StepID', 'EQID', 'BackupEQID', 'LotFrequency', 'SheetFrequency', 'SamplingSheet']
    # TFOMProUsecols=['FactoryName', 'ProductSpecName', 'ProcessFlowName', 'Category1',
    #                 'StepID', 'EQID', 'BackupEQID', 'LotFrequency', 'SheetFrequency', 'SamplingSheet']
    dfMProcess = read_excel(fileAbsPath, sheet_name='ProcessFlow', usecols=TFOMProUsecols)
    logging.debug('## {0} 成功读取，准备处理.......'.format(fileAbsPath))

    # 这里暂存index，为了最后的 FlowPriority 排序使用
    dfMP = dfMProcess
    dfMP['rn'] = dfMP.index

    # 去除所有单元格的前后空格
    dfMP = dfMP.applymap(lambda x: str(x).strip() if type(x) == str else x)

    logging.debug('## 开始 删除虚拟站点以及 repair ...')
    # TFOM里面需要剔除虚拟检测，修复等设备
    noNeedEQ = ['3ATV', '3CTV', '3TTV', '3ATB', '3ATL', '3ATR', '3TTR', ]
    # 取出前几个字符匹配
    dfMPNoR = dfMP[~dfMP.EQID.str[:4].isin(noNeedEQ)]
    dfMPNoR = dfMPNoR.reset_index(drop=True)

    logging.debug('## 开始 合并主站点的Main EQ和Backup EQ ...')
    # 将main operation的主设备和备用设备相连，sampling 设备就不需要了；EQID也要判空，比如ship站点不带设备的main operation
    # dfMPNoR.loc[(dfMPNoR.Category1 == 'Main') & (dfMPNoR.EQID.notnull()), 'NEWEQID'] = dfMPNoR['EQID'].map(str) + '\n' + dfMPNoR['BackupEQID'].map(str)
    dfMPNoR.loc[(dfMPNoR.Category1 == 'Main') & (dfMPNoR.EQID.notnull()), 'NEWEQID'] = dfMPNoR['EQID'].map(str) + (
        dfMPNoR['BackupEQID'].apply(lambda x: '\n' + x if str(x).startswith('3') else ''))
    # 删除旧的两列设备名
    dfMPNoR = dfMPNoR.drop(['EQID', 'BackupEQID'], axis=1)

    logging.debug('## 开始 处理抽检频率，层数，位置 ...')
    # 抽检频率，数量和层数，如果没有就为空
    cols = ['LotFrequency', 'SheetFrequency', 'SamplingSheet']
    dfMPNoR[cols] = dfMPNoR[cols].fillna(0)
    # 规范化，如果全抽就规范为 All
    dfMPNoR['SamplingSheet'] = dfMPNoR['SamplingSheet'].apply(lambda x: 'All' if str(x).startswith('A') else x)

    logging.debug('## 开始 拆分主站点多行设备 ...')
    dfMPMain = dfMPNoR.drop(['NEWEQID'], axis=1).join(
        dfMPNoR['NEWEQID'].str.split('\n', expand=True).stack().reset_index(level=1, drop=True).rename('EQ'))

    logging.debug('## 开始 找所有主站点的行号，并去重 ...')
    # 找到所有的 main operation 的行号，为了下一步sampling的进出operation准备
    mainList = dfMPMain[dfMPMain["Category1"] == 'Main'].index.tolist()
    # 多行main的index，只需要一个index就够了
    ml = list(set(mainList))
    ml = sorted(ml)

    # 找到所有的 Sampling 的行号
    sl = dfMPMain[dfMPMain["Category1"] == 'Sampling'].index.tolist()
    sl = sorted(sl)
    # logging.debug(sl)
    logging.debug('## 共 {0} 行 sampling ...，'.format(len(sl)))

    # logging.debug('## 开始 找所有sampling的上一个main operation ...')
    '''
    sampling 站点的上一个main，也就是TFOM的operation，这里是指从哪个main进抽检的；
    思路就是对每行sampling，在列表[sl]，用每个元素在main operation 列表[ml] 中找大于当前[sl][i]元素的，再往前减1个就是进该sampling的主站点；
    后面的sampling return主站点也是类似的思路，不过是找到比当前sampling大的第一个元素就是return operation
    '''
    # 进入sampling的 main operation
    sml = []
    # 对每行sampling去找
    for i in range(len(sl)):
        #     这里考虑到最后一个元素如果找不到，就挂在最后一个main operation，不过目前最后一个都是ship，基本上都能找到
        flag = False
        for j in range(len(ml)):
            if (ml[j] > sl[i]):
                flag = True
                break  # 找到即停止
        #     sml.append(ml[j-1])
        #  往前减 1
        sml.append(ml[j - 1] if (flag) else ml[-1])
    logging.debug('## 共 {0} 行 sampling 上一个main operation...'.format(len(sml)))

    # 思路在上面已经解释，stol[]是用来存
    stol = []
    for i in range(len(sl)):
        flag = False
        for j in range(len(ml)):
            if (ml[j] > sl[i]):
                flag = True
                break
        #     logging.debug(ml[j-1] if (flag) else ml[-1])
        stol.append(ml[j] if (flag) else ml[-1])
    logging.debug('## 共 {0} 行 sampling 的return main operation...'.format(len(stol)))
    # logging.debug(stol)

    # logging.debug(len(sl), len(sml), len(stol))

    logging.debug('## 开始 匹配sampling的进入和return main operation ...')
    df = dfMPMain.copy(deep=True)

    for i in range(len(sl)):
        '''
        用1个指针在上面的3个列表中同步遍历，最后都是需要在sampling的那一行上做修改，因为main 同一个index可能对应多行，
        所以这里需要判断，如果是series，需要转化成列表，取第一个值，这里取哪个值都可以，都一样
        '''
        if (type(df['ProcessFlowName'][sml[i]]) == str or type(df['ProcessFlowName'][sml[i]]) == int):
            MainFlow = df['ProcessFlowName'][sml[i]]
        else:
            MainFlow = df['ProcessFlowName'][sml[i]].tolist()[0]

        if (type(df['StepID'][sml[i]]) == str or type(df['StepID'][sml[i]]) == int):
            fromMainOper = df['StepID'][sml[i]]
        else:
            fromMainOper = df['StepID'][sml[i]].tolist()[0]

        if (type(df['StepID'][stol[i]]) == str or type(df['StepID'][stol[i]]) == int):
            toMainOper = df['StepID'][stol[i]]
        else:
            toMainOper = df['StepID'][stol[i]].tolist()[0]

        df.at[sl[i], 'processFlowName'] = MainFlow
        df.at[sl[i], 'processOpertionName'] = fromMainOper
        df.at[sl[i], 'returnOperationName'] = toMainOper

    logging.debug('## 开始 对多行main和sampling做笛卡尔组合...')
    dfm = df[df["Category1"] == 'Main']

    dfs = df[df["Category1"] == 'Sampling']

    dfms = pd.merge(dfs, dfm, left_on='processOpertionName', right_on='StepID')

    logging.debug('## 开始 规范化列名...')
    newCols = {'rn_x': 'rn', 'ProcessFlowName_x': 'toProcessFlowName', 'StepID_x': 'toProcessOperationName',
               'LotFrequency_x': 'lotSamplingCount', 'SheetFrequency_x': 'productSamplingCount',
               'SamplingSheet_x': 'productSamplingPosition',
               'processFlowName_x': 'processFlowName', 'processOpertionName_x': 'processOperationName',
               'returnOperationName_x': 'returnOperationName', 'EQ_y': 'machineName'}

    stdl = list(newCols.values())

    dfms = dfms.rename(columns=newCols)
    dfms = dfms[stdl]

    logging.debug('## 开始 添加 Version column ...')
    dfms[['factoryName', 'productSpecName']] = dfms.apply(lambda x: (policy.FACTORY, policy.PRODUCTSPEC), axis=1,
                                                          result_type='expand')
    dfms[['processFlowVersion', 'processOperationVersion', 'toProcessFlowVersion', 'toProcessOperationVersion',
          'returnOperationVer']] = dfms.apply(lambda x: ('00001', '00001', '00001', '00001', '00001',), axis=1,
                                              result_type='expand')
    # dfms[['checkLevel','rmsFlag', 'dispatchState']]=dfPAllNew.apply(lambda x:('N','N', 'N',),axis=1,result_type='expand')
    # dfms[['dispatchPriority','ecRecipeFlag', 'ecRecipeName', 'maskCycleTarget']]=dfPAllNew.apply(lambda x:('','', '', ''),axis=1,result_type='expand')

    logging.debug('## 开始 处理 FlowPriority ...')
    dfms['flowPriority'] = dfms.groupby(['processOperationName', 'machineName'], axis=0)['rn'].rank(ascending=True)
    # dfms.head(20)

    logging.debug('## 开始 规范化最终列名...')
    stdColumns = ['factoryName', 'processFlowName', 'processFlowVersion', 'processOperationName',
                  'processOperationVersion',
                  'machineName', 'toProcessFlowName', 'toProcessFlowVersion', 'toProcessOperationName',
                  'toProcessOperationVersion',
                  'flowPriority', 'lotSamplingCount', 'productSamplingCount', 'productSamplingPosition',
                  'returnOperationName', 'returnOperationVer', ]
    dfms = dfms[stdColumns]
    # dfms

    newTFOMFileDir = policy.TARGET_DIR + "\\TFOM"
    # 因为 Modeler 导入的时候对于文件名和sheet名很规范，所有这里采用不同文件夹的方式输出
    os.mkdir(newTFOMFileDir)
    NewFile_dfTFOM = newTFOMFileDir + '\\' + "TFOM.xls"
    dfms.to_excel(NewFile_dfTFOM, sheet_name='Sampling', index=False)
    logging.debug("Success ！ %s 导出成功" % NewFile_dfTFOM)
