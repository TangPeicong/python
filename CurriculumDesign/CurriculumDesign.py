import csv
import os
from tkinter.font import BOLD
import pymongo
from tkinter import *
from tkinter import ttk
from datetime import *
from pymongo import mongo_client


################# 系统初始化设置 #################
# 窗口标题
window_title = '学生信息管理系统'
# 窗口尺寸
window_width = 800
window_height = 550
# 窗口背景颜色
window_bg = "#f8f8f9"
# 窗口内边距
window_padx = 10
window_pady = 10
# 禁止窗口改变大小
window_resizable = False

# 设置控制台尺寸
Frame_Control_height = 360

# 数据库信息
DB_url = 'mongodb://localhost:27017/'
DB_DataBase = 'admin'
DB_Collection = 'students'

# Treeview表头
TreeviewHeadings = ['姓名', '学号', '年龄', '性别', '宿舍', '高等数学', '数据结构', '线性代数']

# 按钮颜色
button_bg = '#808695'
button_fg = "#ffffff"
################# ############# #################

class Application(Frame):
    ''' GUI主体框架类 '''
    # 创建类私有变量
    __DB_Collection = {}    # Mongo数据集
    __DataSet = {}  # 学生信息
    __treeviewSort_reverses = {
            '学号':False,
            '姓名':False,
            '年龄': False,
            '性别': False,
            '宿舍': False,
            '高等数学': False,
            '数据结构': False,
            '线性代数': False
        }
    def __init__(self, master = None):
        ''' 构造函数 '''
        # 初始化Frame尺寸
        super().__init__(
            master,
            width = window_width,
            height = window_height,
            bg = window_bg
        )
        # 保持Frame尺寸不被子控件影响
        self.pack_propagate(0)
        self.grid_propagate(0)
        self.pack()
        # 创建控件
        self.__CreateWiget()
        self.App_Initlization()
        self.__ConnectDB()
        self.__readDB()
        # 连接数据库
    
    def App_Initlization(self):
        ''' GUI 初始化设置 '''
        # ###############################
        # 修复TreeView设置颜色失效的bug
        def fix_map(option):
            return [elm for elm in style.map("Treeview", query_opt = option)
            if elm[:2] != ("!disabled", "!selected")]
        style = ttk.Style(self)
        style.map(
            "Treeview",
            foreground = fix_map("foreground"),
            background = fix_map("background")
        )
        # ###############################
        
        # ###############################
        # 设置初始Log
        # 解除Text_Log的禁用
        self.insertLog('------------------ WELCOME TO STUDENTS MANAGE SYSTEM -----------------\n')
        self.insertLog('****************************  系统使用说明 ****************************\n')
        self.insertLog('(1). 添加学生信息时,需输入该学生所有信息,若信息缺失,则无法成功添加;\n')
        self.insertLog('(2). 删除学生信息时,仅可通过学号、姓名删除,其他信息输入无效;\n')
        self.insertLog('(3). 修改学生信息时,仅可按学号修改,即学号不可修改,其他信息均可修改;\n')
        self.insertLog('(4). 查询学生信息时;可同时匹配多项信息查询;空输入表示显示所有学生信息.\n')
        self.insertLog('=======================================================================\n')
        
        # ###############################
    
    #################### UI 功能函数 ##################
    def clearTree(self, tree):
        x=tree.get_children()
        for item in x:
            tree.delete(item)
            
    def setTreeStyle(self, tree):
        # 处理表格项的样式
            items = tree.get_children()
            t_count = 0
            for item in items:
                if(t_count % 2 == 0):
                    tree.item(item, tag = 'oddrow')
                else:
                    tree.item(item, tag = 'evenrow')
                t_count = t_count + 1
            tree.tag_configure('oddrow',font=('楷体', 10), foreground='#657180')
            tree.tag_configure('evenrow', font=('楷体', 10), background = '#f8f8f9', foreground='#657180')

    def insertLog(self, text, time = False, splitH = False, splitE = False):
        ''' 插入日志 '''
        self.Text_Log.config(state = NORMAL)
        if(splitH == True):
            self.Text_Log.insert(END, '==============================================================\n')
        if(time == True):
            self.Text_Log.insert(END, datetime.strftime(datetime.now(),'[%Y-%m-%d %H:%M:%S] '))
        self.Text_Log.insert(END, text)
        if(splitE == True):
            self.Text_Log.insert(END, '==============================================================\n')
        self.Text_Log.config(state = DISABLED)
    # Treeview 点击表头按列升序/降序排序
    def __treeviewSort_column(self, tv, col):
        reversed = self.__treeviewSort_reverses[col]
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key = lambda x:str(x[0]), reverse=reversed)
        
        for index,(val, k) in enumerate(l):
            tv.move(k, '', index)
        tv.heading(col, command = lambda: self.__treeviewSort_column(tv, col))
        if(self.__treeviewSort_reverses[col]):
            self.__treeviewSort_reverses[col] = False
        else:
            self.__treeviewSort_reverses[col] = True
        self.setTreeStyle(tv)
    ###################################################
    
    def __ConnectDB(self):
        ''' 连接Mongo数据库 获得集合 '''
        DB_Client = pymongo.MongoClient(DB_url)
        DataBase = DB_Client[DB_DataBase]
        self.__DB_Collection = DataBase[DB_Collection]
    
    def showOnUI(self):
        # 同步UI显示
        self.clearTree(self.Treeview_main)  # 清空Treeview
        for item in self.__DataSet:
            self.Treeview_main.insert(
                "",
                END,
                values = (
                    item['name'],
                    item['num'],
                    item['age'],
                    item['sex'],
                    item['dormitory'],
                    item['score1'],
                    item['score2'],
                    item['score3']
                )
            )
            self.setTreeStyle(self.Treeview_main)   # 设置Treeview样式
    
    def __readDB(self):
        ''' 读取Mongo数据库的数据 '''
        self.__DataSet = self.__DB_Collection.find()
        self.showOnUI()
    
    
    def __btnCmd_Add(self):
        ''' [添加]按钮事件 '''
        _num = self.Entry_Num.get()
        _name = self.Entry_Name.get()
        _age = self.Entry_Age.get()
        _sex = self.Entry_Sex.get()
        _dor = self.Entry_Dor.get()
        _s1 = self.Entry_Score1.get()
        _s2 = self.Entry_Score2.get()
        _s3 = self.Entry_Score3.get()
        # 判断是否有空输入
        if(len(_num) != 10 or _num.isdigit() == False or _name == '' or
           _age.isdigit() == False or _sex not in ['男', '女'] or _dor == '' or
           _s1 == '' or _s2 == '' or _s3 == ''):
            self.insertLog('*** Error *** : 添加失败;输入有误,请检查.\n', time = True)
            self.insertLog('*** warring *** : 年龄必须为整数;性别必须是[男]或[女];不得存在空输入.\n', time = False)
            self.insertLog('*** warring *** : 学号需为10位整数.\n', time = False, splitE = True)
            self.Text_Log.see(END)
        else:
            findResult = self.__DB_Collection.count_documents({
                    'num': _num
                })
            if(findResult == 0):
                addResult = self.__DB_Collection.insert_one({
                'num': _num,
                'name': _name,
                'age': _age,
                'sex': _sex,
                'dormitory': _dor,
                'score1': _s1,
                'score2': _s2,
                'score3': _s3
                })
                self.insertLog('*** Succefully *** : 添加成功, 已同步UI显示.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                self.__readDB()
            else:
                self.insertLog('*** error *** : 添加失败, 学号重复.\n', time = True, splitE = True)
                self.Text_Log.see(END)
        
    def __btnCmd_Delete(self):
        ''' [删除]按钮事件 '''
        _num = self.Entry_Num.get()
        _name = self.Entry_Name.get()
        if(_num == '' and _name == ''):
            self.insertLog('*** error *** : 删除失败, 空输入.\n', time = True, splitE = True)
            self.Text_Log.see(END)
        elif(_num != '' and _name != ''):
            findResult = self.__DB_Collection.count_documents({
                'num': _num,
                'name': _name
            })
            if(findResult != 0):
                self.__DB_Collection.delete_one({
                    'num': _num,
                    'name': _name
                })
                self.insertLog('*** Succefully *** : 删除成功, 已同步UI显示.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                self.__readDB()
            else:
                self.insertLog('*** error *** : 删除失败, 该学生不存在.\n', time = True, splitE = True)
                self.Text_Log.see(END)
        elif(_num != ''):
            findResult = self.__DB_Collection.count_documents({
                'num': _num,
            })
            if(findResult != 0):
                self.__DB_Collection.delete_one({
                    'num': _num,
                })
                self.insertLog('*** Succefully *** : 删除成功, 已同步UI显示.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                self.__readDB()
            else:
                self.insertLog('*** error *** : 删除失败, 该学生不存在.\n', time = True, splitE = True)
                self.Text_Log.see(END)
        elif(_name != ''):
            findResult = self.__DB_Collection.count_documents({
                'name': _name
            })
            if(findResult != 0):
                self.__DB_Collection.delete_one({
                    'name': _name
                })
                self.insertLog('*** Succefully *** : 删除成功, 已同步UI显示.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                self.__readDB()
            else:
                self.insertLog('*** error *** : 删除失败, 该学生不存在.\n', time = True, splitE = True)
                self.Text_Log.see(END)
            
        
    def __btnCmd_Modify(self):
        ''' [修改]按钮事件 '''
        _num = self.Entry_Num.get()
        _name = self.Entry_Name.get()
        _age = self.Entry_Age.get()
        _sex = self.Entry_Sex.get()
        _dor = self.Entry_Dor.get()
        _s1 = self.Entry_Score1.get()
        _s2 = self.Entry_Score2.get()
        _s3 = self.Entry_Score3.get()
        if(_num == ''):
            self.insertLog('*** error *** : 修改失败, 学号空输入!\n', time = True, splitE = True)
            self.Text_Log.see(END)
        else:
            ''''''
            modifyVal = {}
            ####################
            if(_name != ''):
                modifyVal['name'] = _name
            if(_age != ''):
                modifyVal['age'] = _age
            if(_sex != ''):
                modifyVal['sex'] = _sex
            if(_dor != ''):
                modifyVal['dormitory'] = _dor
            if(_s1 != ''):
                modifyVal['score1'] = _s1
            if(_s2 != ''):
                modifyVal['score2'] = _s2
            if(_s3 != ''):
                modifyVal['score3'] = _s3
            ####################
            if(modifyVal != {}):
                self.__DB_Collection.update_one(
                    {
                        'num': _num
                    },
                    {
                        "$set": modifyVal
                    }
                )
                self.insertLog('*** Succefully *** : 修改成功, 已同步UI显示.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                self.__readDB()
            else:
                self.insertLog('*** Warring *** : 空输入, 不作修改.\n', time = True, splitE = True)
                self.Text_Log.see(END)
                
            
    def __btnCmd_Search(self):
        ''' [查询]按钮事件 '''
        _num = self.Entry_Num.get()
        _name = self.Entry_Name.get()
        _age = self.Entry_Age.get()
        _sex = self.Entry_Sex.get()
        _dor = self.Entry_Dor.get()
        _s1 = self.Entry_Score1.get()
        _s2 = self.Entry_Score2.get()
        _s3 = self.Entry_Score3.get()
        option = {}
        ret = {
            'num': 1,
            'name': 1,
            'age': 1,
            'sex': 1,
            'dormitory': 1,
            'score1': 1,
            'score2': 1,
            'score3': 1
        }
        if(_num != ''):
            option['num'] = _num
        if(_name != ''):
             option['name'] = _name
        if(_age != ''):
            option['age'] = _age
        if(_sex != ''):
            option['sex'] = _sex
        if(_dor != ''):
            option['dormitory'] = _dor
        if(_s1 != ''):
            option['score1'] = _s1
        if(_s2 != ''):
            option['score2'] = _s2
        if(_s3 != ''):
            option['score3'] = _s3
        ######### 查询 #########
        res = {}
        if(option != {}):
            res = self.__DB_Collection.find(
                option,
                ret
                )
            self.insertLog('*** Succefully *** : 查询成功, 已同步UI显示.\n', time = True, splitE = True)
            self.Text_Log.see(END)
            self.__readDB()
        else:
            res = self.__DB_Collection.find()
            self.insertLog('*** Warring *** : 空输入, 显示所有学生信息.\n', time = True, splitE = True)
            self.Text_Log.see(END)
            self.__readDB()
        self.__DataSet = res
        self.showOnUI()
    
    def __CreateWiget(self):
        ''' 创建控件 '''
        self.Treeview_main = ttk.Treeview(
            self,
            show = 'headings',
        )
        # 设置Treeview列宽
        viewWidth = int(window_width / len(TreeviewHeadings)) - window_padx * 2
        self.Treeview_main['columns'] = (TreeviewHeadings)
        for col in TreeviewHeadings:  # 定义表头和列
            self.Treeview_main.column(col, width = viewWidth, anchor = N)
            self.Treeview_main.heading(col, text = col,
                                       command=lambda _col=col: self.__treeviewSort_column(self.Treeview_main, _col))
        # 设置表头字体
        ttk.Style().configure('Treeview.Heading', font = ('楷体', 10))
        self.Treeview_main.pack(fill= X, side= TOP, padx= window_padx)
        
        # [Frame - Control]
        self.Frame_Control = Frame(
            self,
            width = window_width - window_padx * 2,
            height = Frame_Control_height,
            bg = window_bg
        )
        self.Frame_Control.pack_propagate(0)
        self.Frame_Control.grid_propagate(0)
        self.Frame_Control.pack(pady = 5)
        
        # [Frame - Entrys]
        self.Frame_Entrys = Frame(
            self.Frame_Control,
            width = window_width - window_padx * 2,
            height = 180,
            bg = window_bg,
            pady = 15
        )
        self.Frame_Entrys.pack_propagate(0)
        self.Frame_Entrys.grid_propagate(0)
        self.Frame_Entrys.pack()
        # [Frame - Log]
        self.Frame_Log = Frame(
            self.Frame_Control,
            width = window_width - window_padx * 2,
            height = 120,
            bg = window_bg
        )
        self.Frame_Log.pack_propagate(0)
        self.Frame_Log.grid_propagate(0)
        self.Frame_Log.pack()
        ###################### Entrys ########################
        EntrysPady = 5
        # [学号]
        self.Label_Num = Label(
            self.Frame_Entrys,
            text = '学号:',
            bg = window_bg
        )
        self.Label_Num.grid(row = 0, column = 0, padx = 10, pady = EntrysPady)
        self.Entry_Num = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Num.grid(row = 0, column = 1, pady = EntrysPady)
        
        # [姓名]
        self.Label_Name = Label(
            self.Frame_Entrys,
            text = '姓名:',
            bg = window_bg
        )
        self.Label_Name.grid(row = 1, column = 0, padx = 10, pady = EntrysPady)
        self.Entry_Name = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Name.grid(row = 1, column = 1, pady = EntrysPady)
        
        # [年龄]
        self.Label_Age = Label(
            self.Frame_Entrys,
            text = '年龄:',
            bg = window_bg
        )
        self.Label_Age.grid(row = 2, column = 0, padx = 10, pady = EntrysPady)
        self.Entry_Age = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Age.grid(row = 2, column = 1, pady = EntrysPady)
        
        # [性别]
        self.Label_Sex = Label(
            self.Frame_Entrys,
            text = '性别:',
            bg = window_bg
        )
        self.Label_Sex.grid(row = 3, column = 0, padx = 10, pady = EntrysPady)
        self.Entry_Sex = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Sex.grid(row = 3, column = 1, pady = EntrysPady)
        self.Label_split = Label(
            self.Frame_Entrys,
            text = '',
            bg = window_bg,
        )
        self.Label_split.grid(row = 0, column = 2, padx = 10)
        # [宿舍]
        self.Label_Dor = Label(
            self.Frame_Entrys,
            text = '宿舍:',
            bg = window_bg,
        )
        self.Label_Dor.grid(row = 0, column = 3, padx = 10, pady = EntrysPady)
        self.Entry_Dor = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Dor.grid(row = 0, column = 4, pady = EntrysPady)
        
        # [高等数学]
        self.Label_Score1 = Label(
            self.Frame_Entrys,
            text = '高等数学:',
            bg = window_bg,
        )
        self.Label_Score1.grid(row = 1, column = 3, padx = 10, pady = EntrysPady)
        self.Entry_Score1 = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Score1.grid(row = 1, column = 4, pady = EntrysPady)
        
        # [数据结构]
        self.Label_Score2 = Label(
            self.Frame_Entrys,
            text = '数据结构:',
            bg = window_bg,
        )
        self.Label_Score2.grid(row = 2, column = 3, padx = 10, pady = EntrysPady)
        self.Entry_Score2 = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Score2.grid(row = 2, column = 4, pady = EntrysPady)
        
        # [线性代数]
        self.Label_Score3 = Label(
            self.Frame_Entrys,
            text = '线性代数:',
            bg = window_bg,
        )
        self.Label_Score3.grid(row = 3, column = 3, padx = 10, pady = EntrysPady)
        self.Entry_Score3 = Entry(
            self.Frame_Entrys,
            width = 25,
        )
        self.Entry_Score3.grid(row = 3, column = 4, pady = EntrysPady)
        ###################### Buttons #######################
        buttonFont = ('微软雅黑', 8)
        buttonPadx = 20
        buttonWidth = 14
        # [Button - Add]
        self.Button_Add = Button(
            self.Frame_Entrys,
            width = buttonWidth,
            font = buttonFont,
            text = '添加',
            command = self.__btnCmd_Add
        )
        self.Button_Add.grid(row = 0, column = 5, padx = buttonPadx)
        
        # [Button - Delete]
        self.Button_Delete = Button(
            self.Frame_Entrys,
            width = buttonWidth,
            font = buttonFont,
            text = '删除',
            command = self.__btnCmd_Delete
        )
        self.Button_Delete.grid(row = 1, column = 5, padx = buttonPadx)
        
        # [Button - Modify]
        self.Button_Modify = Button(
            self.Frame_Entrys,
            width = buttonWidth,
            font = buttonFont,
            text = '修改',
            command = self.__btnCmd_Modify
        )
        self.Button_Modify.grid(row = 2, column = 5, padx = buttonPadx)
        
        # [Button - Search]
        self.Button_Search = Button(
            self.Frame_Entrys,
            width = buttonWidth,
            font = buttonFont,
            text = '查询',
            command = self.__btnCmd_Search
        )
        self.Button_Search.grid(row = 3, column = 5, padx = buttonPadx)
        ######################################################
        
        ####################### Logs #########################
        self.Text_Log = Text(
            self.Frame_Log,
            width = window_width - window_padx * 3,
            height = 100,
            bg = window_bg,
            padx = 10,
            font = ('宋体', 10)
        )
        self.Text_Log.config(state = DISABLED)
        self.Text_Log.pack()
        ######################################################

def Initlization(window):
    ''' 窗口初始化 '''
    # 设置窗口标题
    window.title(window_title)
    # 设置窗口尺寸并居中
    system_screen_width = window.winfo_screenwidth()
    system_screen_height = window.winfo_screenheight()
    window_offset_x = (system_screen_width - window_width) / 2
    window_offset_y = (system_screen_height - window_height) / 2
    window.geometry('%dx%d+%d+%d'%(window_width, window_height, window_offset_x, window_offset_y))
    # 设置窗口背景色
    window.configure(bg = window_bg)
    # 禁止窗口修改大小
    if(window_resizable == False):
        window.resizable(width = False, height = False)
        
if __name__ == '__main__':
    root = Tk()
    Initlization(root)
    app = Application(root)
    root.mainloop()
