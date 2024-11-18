# coding : utf-8
import tkinter
import tkinter.filedialog
import os
import xml.etree.ElementTree as ET
import xlwt
 
 
# 设置从第1行开始插入excel，因为第0行是标题
row_num = 1
# 设置用例标题所在列
casename_col_num = 3
# 设置所属产品所在列
project_col_num = 1
project_name = '项目名片'
# 设置模块所在列
module_col_num = 2
# 设置预期所在列
result_col_num = 6
# 设置用例类型所在列
casetype_col_num = 7
testcase_type = '功能测试'
# 设置用例状态所在列
status_col_num = 9
testcase_status = '正常'
 
 
# 选freemind文件
def select_file():
    global file_path
    file_path = tkinter.filedialog.askopenfilename(title='选择一个freemind文件', filetypes=[('freemind文件', '.mm')])
    return file_path
 
 
# 选Excel生成目录
def select_folder():
    global folder_path
    folder_path = tkinter.filedialog.askdirectory(title='选择一个文件夹')
    return folder_path
 
 
# 生成窗口
def create_frame():
    win = tkinter.Tk()
    win.geometry('360x150')
    win.title("选择freemind文件目录")
    file = tkinter.Button(win, text="选择freemind文件", height=2, width=20, fg="blue", bg="gray", command=select_file)
    file.pack()
    folder = tkinter.Button(win, text="选择生成的Excel目录", height=2, width=20, fg="blue", bg="gray", command=select_folder)
    folder.pack()
    excel = tkinter.Button(win, text="生成Excel文件", height=1, width=12, bg="gray", command=win.quit)
    excel.pack()
    win.mainloop()
 
 
# 设置excel标题
def set_excel_header():
    col = 0
    header = ['用例编号', '所属产品', '所属模块', '用例标题', '前置条件', '用例步骤', '预期', '用例类型', '适用阶段', '用例状态']
    for name in header:
        write_excel(name, 0, col)
        col += 1
 
 
# 生成用例标题的模块名,格式为模块1__模块2__模块3__
def set_testcase_module(element, testcase_title, text_index, second_node):
    global g_row
    
    testcase_title = testcase_title + '__' + element.get('TEXT')
    text_index += 1
    for child in element:
        if child is None or child.tag != 'node': continue
        set_case(child, testcase_title, text_index, second_node)
        name = child.get('TEXT')
        if child.find('node') is None and name is not None:
            # 计算根节点长度，用例标题的模块不必包含根节点
            title_length = len(first_node.get('TEXT')) + len(second_node.get('TEXT')) + 6
            title = testcase_title[title_length:] + '__' + name
            # print(f"title={title},text_index={text_index}")
            module_name = title.split('__')[0]

            # 写入用例编号
            write_excel(g_row, g_row, 0)
            # 写入用例标题
            write_excel(title[(len(module_name)+2):], g_row, casename_col_num)
            # 写入预期结果
            write_excel(name, g_row, result_col_num)
            # 写入用例类型
            write_excel(testcase_type, g_row, casetype_col_num)  
            # 写入用例状态
            write_excel(testcase_status, g_row, status_col_num)  
            write_excel(project_name, g_row, project_col_num)
            write_excel(module_name, g_row, module_col_num)
            g_row += 1
 
 
# 组成用例
def set_case(element, name, text_index, second_node):
    global g_row
    if element is not None:
        set_testcase_module(element, name, text_index, second_node)
 
 
# 写入Excel文件
def write_excel(name, row, col):
    ws.write(row, col, name)
 
 
if __name__ == '__main__':
    g_row = 1
    file_path = ''
    folder_path = ''
    file_name = 'freemind2excel_new.xls'
    create_frame()
    print('选择的freemind文件为：' + file_path)
    tree = ET.ElementTree(file=file_path)
    root = tree.getroot()
    first_node = root.find('node')
    second_node = first_node.find('node')
    wb = xlwt.Workbook()  # 创建工作簿
    ws = wb.add_sheet('freemind2excel')  # 指定工作簿名称
    set_excel_header()  # 写表头
    text_content = ''
    text_index = 1
    set_testcase_module(first_node, text_content, text_index, second_node)

    file_path = folder_path + '/' + file_name
    if os.path.exists(file_path): os.remove(path=file_path)
    wb.save(file_path)
    if not os.path.exists(file_path):
        print('Excel生成失败')
    else:
        print('Excel生成成功，路径为：' + file_path)
        print('测试用例条数为：' + str(g_row- 1))