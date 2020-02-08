import controls
import os
import time
def checkVersion(path,pathaction):
    '''高低版本文件'''
    pathfiles = controls.pathCommon(path=path, type=getDict(pathaction))['files']
    xlslist=[]
    xlsxlist=[]
    for file in pathfiles:
        if file.endswith('.xls'):
            xlslist.append(file)
        elif file.endswith('.xlsx'):
            xlsxlist.append(file)
        else:
            pass
    res=''
    for xlsx in xlsxlist:
        for xls in xlslist:
            if xlsx==xls+'x':
                res=res+xlsx
    return res
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
            for xls in xlslist:
                if xlsx==xls+'x':
                    os.remove(xlsx)
                    xlslist.remove(xls)
    return res
def merge2(path,pathaction,sheetnums,resname):
    start=time.perf_counter()
    switch = getDict(pathaction)
    #pathfiles=[]
    pathfiles = controls.pathCommon(path=path,type=switch)['files']
    #print(pathfiles)
    pathfiles=uni(pathfiles)
    #print(pathfiles)
    print('总共有{}个文件'.format(str(len(pathfiles))))
    switch=getDict(sheetnums)
    resfile=controls.fileInfo(os.path.join(path,resname))
    content=[]
    count=1
    if switch=='1':
        for pathfile in pathfiles:
            print('正在处理第{}个文件,文件名为{}'.format(str(count),pathfile))
            xlsx=controls.fileInfo(pathfile)
            content=content+xlsx.getFileContent(sheetName=None,containTitle=True)
            count=count+1
        resfile.expXlsx(content=content)
    else:
        for pathfile in pathfiles:
            #print(pathfile)
            print('正在处理第{}个文件,文件名为{}'.format(str(count),pathfile))
            xlsx=controls.fileInfo(pathfile)
            content=xlsx.getFileContent(sheetName=None,containTitle=True)
            #print(content)
            sheetName=os.path.split(xlsx.fileName)[1]
            resfile.expXlsx(content=content,mode='',sheetName=sheetName)
            count=count+1
    uni(pathfiles=controls.pathCommon(path=path,type=getDict(pathaction))['files'],type='clear')
    end=time.perf_counter()
    print('合并结束,合并的文件名为{}（含路径）,共耗时{}秒'.format(resfile.fileName,str(end-start)))
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




# interface(action='合并xlsx',path='C:\\Users\\xjk-lenovo\\Desktop\\20191219交换生',pathaction='不包含子文件夹'
# ,sheetnums='一个工作表',resname='res')