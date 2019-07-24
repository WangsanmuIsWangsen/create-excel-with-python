import xlsxwriter
import os
import sys

# create a new Excel file and add a worksheet
workbook = xlsxwriter.Workbook('apk内存监控情况统计.xlsx')
worksheet = workbook.add_worksheet()
#set format
bold=workbook.add_format({
        'bold':  True,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'left',  # 水平对齐方式
        'valign': 'vcenter',  # 垂直对齐方式
        'fg_color': '#F4B084',  # 单元格背景颜色
        'text_wrap': True,  # 是否自动换行
})
worksheet.set_column('A:A', 35)   #设置列的长度
worksheet.set_column('C:C', 11)
worksheet.set_column('D:D', 40)
worksheet.set_column('E:E', 12)
worksheet.set_column('F:F', 30)
worksheet.set_column('G:G', 20)
worksheet.write(0,0,"手机型号-版本",bold)
worksheet.write(0,1,"最高内存",bold)
worksheet.write(0,2,"超限值次数",bold)
worksheet.write(0,3,"日志路径",bold)
worksheet.write(0,4,"模块名称",bold)
worksheet.write(0,5,"对应问题单",bold)
worksheet.write(0,6,"定位责任人",bold)
def devices_info(file_path):
    count = 0
    try:
        with open(file_path, 'r') as f:
            line = f.readline()
            while line:
                if count == 2:
                    devices_name=line
                count+=1
                line=f.readline()
    except FileNotFoundError:
        print(file_path + " has lost")
    return devices_name
def analyze_typeandsize(typesize_info):   #将typesize_info二维列表的内容按照
    x = []  # 最高内存
    y = []  # 超限值次数
    z = []  # 模块名称
    type = typesize_info [0][0]
    max = 0   #记录最大值
    count = 0   #记录出现次数
    for i in typesize_info:
        if i[0] == type:
            count += 1
            if int(i[1].split('M')[0]) > max:
                max=int(i[1].split('M')[0])
        else:
            x.append(max)
            y.append(count)
            z.append(type)
            type=i[0]
            max=int(i[1].split('M')[0])
            count = 1
    #最后一组没有录入
    x.append(max)
    y.append(count)
    z.append(type)
    return (x,y,z)

def deal_hprof(info):   #返回type和size
    list = []
    line=info.split('/')
    type_size=line[5]
    type=type_size.split('_')[0]
    typename=type.split('.')[-1]
    sizeinfo=type_size.split('_')[2]
    size=sizeinfo.split('.')[0]
    list.append(typename)
    list.append(size)
    return list
def analyze_hprofinfo(hprof_path):
    type_size = []
    try:
        with open(hprof_path, 'r') as f:
            line = f.readline()
            while line:
                eachline = line.split()    #六个元素
                type_size.append(deal_hprof(eachline[4]))   #存放在二维列表
                line=f.readline()
            res=analyze_typeandsize(type_size)      #分析二维列表返回x,y,z元组
    except FileNotFoundError:
        print(hprof_path + " has lost")
    return res
#这里需要添加一个参数记录行数。最后要将这个参数返回
def insert_excal(folder, filename, hang):   #filename就是日志路径
    devices_name=devices_info(folder+"\\devicesInfo.txt")  #An的信息
    list = os.listdir(folder)
    #count = 0
    count_hang = hang
    if os.path.exists(folder+"\\hprof_record.txt"):
        info_list=analyze_hprofinfo(folder+"\\hprof_record.txt")    #返回([],[],[])
        for j in range(0,len(info_list[0])):
            worksheet.write(count_hang,0,devices_name)
            worksheet.write(count_hang,1,info_list[0][j])
            worksheet.write(count_hang,2,info_list[1][j])
            worksheet.write(count_hang,3,filename)
            worksheet.write(count_hang,4,info_list[2][j])
            count_hang += 1

    return count_hang   #main里面需要行数加上返回值

def main(argv):
    list = os.listdir(argv[1]) #传入的参数为当前目录
    count = 0
    count_hang = 1
    for i in range(0, len(list)):
        path = os.path.join(argv[1], list[i])  # 访问文件夹下所有文件
        if os.path.isdir(path):  # 文件是否是一个目录
            res_hang = insert_excal(path, list[i], count_hang)  # path是外部传入的参数
            count_hang=res_hang
            count += 1
    workbook.close()
    print("create excal successfully!")


if __name__ == '__main__':
    main(sys.argv)