import controls
import os
# def merge(action,curr,sheetnum,resfileName):
#     if action=='merge':
#         path=os.getcwd()
#         resfile=controls.fileInfo(os.path.join(path,resfileName))
#         content=[]
#         if curr=='curr':
#             pathfiles=[]
#             zs=os.listdir(path)
#             for z in zs:
#                 pathfiles.append(os.path.join(path,z))
#         else:
#             pathfiles=controls.pathCommon(path)['files']
#         for pathfile in pathfiles:
#             if pathfile.endswith('xls'):
#                 xlsx=controls.fileInfo(pathfile)
#                 xlsx.xlsToXlsx()
#         if sheetnum=='1':
#             for pathfile in pathfiles:
#                 if pathfile.endswith('.xlsx'):
#                     xlsx=controls.fileInfo(pathfile)
#                     content=content+xlsx.getFileContent(sheetName=None,containTitle=True)
#             resfile.expXlsx(content=content)
#         elif sheetnum=='2':
#             for pathfile in pathfiles:
#                 #print('pathfile={}'.format(pathfile))
#                 if pathfile.endswith('.xlsx'):
#                     xlsx=controls.fileInfo(pathfile)
#                     content=xlsx.getFileContent(sheetName=None,containTitle=True)
#                     sheetName=os.path.split(pathfile)[1]
#                     #print('ws={}'.format(sheetName))
#                     resfile.expXlsx(content=content,mode='',sheetName=sheetName)
#         else:
#             pass
#     elif action=='file':
#         pass
#     else:
#         pass
#     return 1
def getDict(words):
    baseDict={
        '合并xlsx':'merge',
        '不包含子文件夹':'0',
        '包含子文件夹':'1',
        '一个工作表':'1',
        '多个工作表':'2',
    }
    return baseDict[words]
#merge(action='merge',curr='',sheetnum='2',resfileName='res')
def uni(pathfiles,type='create'):
    res=[]
    if type=='create':
        for pathfile in pathfiles:
            if pathfile.endswith('.xls'):
                xls=controls.fileInfo(pathfile)
                xls.xlsToXlsx()
                res.append(xls.fileName+'x')
            else:
                res.append(pathfile)
    else:
        xlslist=[]
        xlsxlist=[]
        for z in pathfiles:
            if z.endswith('.xls'):
                xlslist.append(z)
            elif z.endswith('.xlsx'):
                xlsxlist.append(z)
            else:
                pass
        for xlsx in xlsxlist:
            for xls in xlsxlist:
                if xlsx==xls+'x':
                    os.remove(xlsx)
                    xls.remove(xls)
    return res
def merge2(path,pathaction,sheetnums,resname):
    switch = getDict(pathaction)
    # if switch == '0':
    #     pathfiles = []
    #     zs = os.listdir(path)
    #     for z in zs:
    #         pathfiles.append(os.path.join(path, z))
    # else:
    #     pathfiles = controls.pathCommon(path)
    pathfiles = controls.pathCommon(path=path,type=switch)['files']
    pathfiles=uni(pathfiles)
    switch=getDict(sheetnums)
    resfile=controls.fileInfo(os.path.join(path,resname))
    content=[]
    if switch=='1':
        for pathfile in pathfiles:
            xlsx=controls.fileInfo(pathfile)
            content=content+xlsx.getFileContent(sheetName=None,containTitle=True)
        resfile.expXlsx(content=content)
    else:
        for pathfile in pathfiles:
            xlsx=controls.fileInfo(pathfile)
            content=xlsx.getFileContent(sheetName=None,containTitle=True)
            sheetName=os.path.split(xlsx.fileName)[1]
            resfile.expXlsx(content=content,mode='',sheetName=sheetName)
    uni(pathfiles,type='clear')
    return resfile.fileName

def interface(action,path,pathaction,sheetnums,resname):
    '''
    :param action:操作方式
    :param path: 文件路径
    :param pathaction: 是否包含子文件夹
    :param sheetnums: 结果的sheet数量
    :param resname: 保存的文件名关键字
    :return:
    '''

    if getDict(action)=='merge':
        return merge2(path,pathaction,sheetnums,resname)



# interface(action='合并xlsx',path='C:\\Users\\xjk-lenovo\\Desktop\\20191219交换生',pathaction='包含子文件夹'
# ,sheetnums='多个工作表',resname='res')