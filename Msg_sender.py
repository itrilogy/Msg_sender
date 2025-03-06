# 库引用
from tkinter import * # GUI界面基础组件
from tkinter import ttk # GUI界面高级组件
from tkinter.messagebox import showinfo # 消息提示框
from tkinter.messagebox import showerror # 错误提示框
from tkinter.messagebox import askyesno # 是/否确认框
from tkinter.filedialog import askopenfilename # 文件选择对话框
from tkinter.simpledialog import askstring # 字符串输入对话框

from typing import Dict # 类型注解支持
from markdown import markdown # markdown文本转HTML支持
from tkhtmlview import HTMLScrolledText # HTML渲染支持
import html # HTML编码解码支持
import xlrd # Excel文件读取支持
import requests, json # HTTP请求和JSON数据处理支持
import datetime # 日期时间处理支持
import time # 时间操作支持
import os # 操作系统接口支持
import tempfile # 临时文件操作支持
import configparser # 配置文件读写支持
import random # 随机数生成支持
import logging
import sys

# 配置日志处理
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# 重定向标准错误输出到/dev/null以过滤系统级别的提示信息
if sys.platform == 'darwin':
    sys.stderr = open(os.devnull, 'w')

# 全局配置常量
CONFIG_FILE = 'config.ini' # 配置文件路径
Process_info='' # 进度信息
README_FILE='about.md' # 关于信息文件路径
TOKEN_FILE='token_access.conf' # 访问令牌文件路径
COPYRIGHT_INFO='\nSending by API消息发送助手\nCopyright © 2023-2025 Kwangwah Hung\nThis software is open source and free to use under the MIT License.\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software.' # 版权信息

# 全局配置对象
CONF_OBJ = configparser.ConfigParser()
CONF_OBJ_default_section = 'default'

# 配置管理类
class ConfigManager:
    def __init__(self, config_file):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.default_section = "default"
        self.required_fields = ["agentid", "cropid", "screctid"]
        self.load_config()

    def load_config(self):
        """加载配置文件，如果不存在则创建默认配置"""
        try:
            if not os.path.exists(self.config_file):
                self._create_default_config()
            self.config.read(self.config_file, encoding='utf-8')
        except Exception as e:
            showerror("错误", f"加载配置文件失败：{str(e)}")

    def _create_default_config(self):
        """创建默认配置文件"""
        default_section = f'自动配置项{random.randint(1000, 9999)}'
        self.config[default_section] = {
            'agentid': '',
            'cropid': '',
            'screctid': '',
            self.default_section: 'true'
        }
        self.save_config()

    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception as e:
            showerror("错误", f"保存配置文件失败：{str(e)}")

    def get_default_section(self):
        """获取默认配置节"""
        for section in self.config.sections():
            if self.config.getboolean(section, self.default_section, fallback=False):
                return section
        return None

    def validate_config(self, section):
        """验证配置项是否完整"""
        if not self.config.has_section(section):
            return False
        return all(self.config.has_option(section, field) for field in self.required_fields)

# 初始化配置管理器
config_manager = ConfigManager(CONFIG_FILE)

# 定义企业微信消息发送类
# 配置应用相关初始信息
class WeChat:
    """企业微信消息发送类"""
    def __init__(self, config_section=None):
        """初始化企业微信配置
        Args:
            config_section: 配置节名称，如果为None则使用默认配置
        """
        if config_section is None:
            config_section = config_manager.get_default_section()
            if not config_section:
                raise ValueError("未找到默认配置节")
        
        if not config_manager.validate_config(config_section):
            raise ValueError(f"配置节 {config_section} 不完整或无效")
        
        config = config_manager.config[config_section]
        self.CORPID = config.get('cropid')  # 企业ID
        self.CORPSECRET = config.get('screctid')  # 应用Secret
        self.AGENTID = config.get('agentid')  # 应用Agentid
        self.ACCESS_TOKEN_PATH = TOKEN_FILE  # 存放access_token的路径
        # self.AGENTID = "1000025"  # 应用Agentid
        # self.ACCESS_TOKEN_PATH = "access_token.conf" # 存放access_token的路径
# 根据初始配置信息获得登录信息access_token
    def _get_access_token(self):       
        url = f'https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={self.CORPID}&corpsecret={self.CORPSECRET}'
        res = requests.get(url=url)
        # print(url)
        # print()
        return json.loads(res.text)['access_token']
#将获取到的access_token保存到本地
    def _save_access_token(self, cur_time):
        with open(TOKEN_FILE, "w")as f:
            access_token = self._get_access_token()
            # 保存获取时间以及access_token
            f.write("\t".join([str(cur_time), access_token]))
        return access_token

        """
        读取"access_token.conf"中的access_token
        """
    def get_access_token(self):
        cur_time = time.time()
        try:
            with open(TOKEN_FILE, "r")as f:
                t, access_token = f.read().split()
                # 判断access_token是否有效
                if 0 < cur_time-float(t) < 7200:
                    return access_token
                else:
                    return self._save_access_token(cur_time)
        except:
            return self._save_access_token(cur_time)

        """
        发送消息(可以定义消息类型)，
        匹配message的生成内容，可以定义一个创建消息类型的类build_message。
        可以增加传递参数，touser,通过表单的内容获取
        """
    def send_message(self, message, msg_type,to_users):
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={self.get_access_token()}"
        send_values = {
            "touser": to_users,
            "msgtype": msg_type,
            "agentid": conf_AGENTID,
            msg_type: {
                "content": message
            },
        }
        send_message = (bytes(json.dumps(send_values,ensure_ascii=False), 'utf-8'))
        res = requests.post(url, send_message)
        return res.json()['errmsg']


# 发送文件功能，暂不开放
        
        # 先将文件上传到临时媒体库
        
    def _upload_file(self, file):
        url = f"https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={self.get_access_token()}&type=file"
        data = {"file": open(file, "rb")}
        res = requests.post(url, files=data)
        return res.json()['media_id']

        
        # 发送文件
        
    def send_file(self, file):
        media_id = self._upload_file(file) # 先将文件上传至临时媒体库
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={self.get_access_token()}"
        send_values = {
            "touser": self.TOUSER,
            "msgtype": "file",
            "agentid": self.AGENTID,
            "file": {
                "media_id": media_id
            },
        }
        send_message = (bytes(json.dumps(send_values,ensure_ascii=False), 'utf-8'))
        res = requests.post(url, send_message)
        return res.json()['errmsg']
# 帮助窗口代码类
class AboutWindow(Toplevel):
    def __init__(self, parent, text):
        super().__init__(parent)
        self.title("关于本软件")
        self.geometry("450x400")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        # 计算弹出窗口相较主窗口左上角的坐标
        self.geometry("+{}+{}".format(win.winfo_x()+50, win.winfo_y()))
        # 计算弹出窗口相较主窗口中心的坐标
        # x = win.winfo_x() + win.winfo_width()//2 - self.winfo_width()//2
        # y = win.winfo_y() + win.winfo_height()//2 - self.winfo_height()//2
        # self.geometry("+{}+{}".format(x, y))
        self.text = text
        self.transient(parent)
        self.grab_set() 
        self.create_widgets()

    def create_widgets(self):
        frame = Frame(self)
        frame.pack(side=TOP, padx=20, pady=10, fill=BOTH, expand=1)

        # 使用HTMLScrolledText渲染HTML
        html_widget = HTMLScrolledText(frame)
        html_widget.pack(side=LEFT, fill=BOTH, expand=1)
        html_widget.set_html(self.text)

        # closeButton = Button(self, text="关闭", width=10, command=self.cancel)
        # closeButton.pack(side=BOTTOM, padx=20, pady=10)

    def cancel(self):
        self.destroy()
# 配置窗口代码类
class cfgmanager(Toplevel):
    widget_dic: Dict[str, Widget] = {}



    def __init__(self,parent):
        super().__init__(parent)
        self.title("参数配置")
        # 设置父窗口关系
        self.transient(parent)
        # 设置窗口大小、居中
        width = 540
        height = 200
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)
        # 设置窗口模态
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.grab_set()
        self.widget_dic["tk_list_box_list_CfgItem"] = self.__tk_list_box_list_CfgItem(self)
        self.widget_dic["tk_label_lab_CfgItem"] = self.__tk_label_lab_CfgItem(self)
        self.widget_dic["tk_label_lab_AGENTID"] = self.__tk_label_lab_AGENTID(self)
        self.widget_dic["tk_label_lab_CORPID"] = self.__tk_label_lab_CORPID(self)
        self.widget_dic["tk_label_lab_CORPSECRET"] = self.__tk_label_lab_CORPSECRET(self)
        self.widget_dic["tk_input_text_AGENTID"] = self.__tk_input_text_AGENTID(self)
        self.widget_dic["tk_input_text_CROPID"] = self.__tk_input_text_CROPID(self)
        self.widget_dic["tk_input_text_CORPSECRET"] = self.__tk_input_text_CORPSECRET(self)
        self.widget_dic["tk_button_btn_CfgSave"] = self.__tk_button_btn_CfgSave(self)
        self.widget_dic["tk_button_btn_ItemAdd"] = self.__tk_button_btn_ItemAdd(self)
        self.widget_dic["tk_button_btn_ItemDel"] = self.__tk_button_btn_ItemDel(self)
        self.widget_dic["tk_button_btn_CfgDef"] = self.__tk_button_btn_CfgDef(self)

        
        self.widget_dic["tk_list_box_list_CfgItem"].bind("<ButtonRelease-1>", self.show_config)
        self.widget_dic["tk_list_box_list_CfgItem"].bind("<Double-Button-1>", self.edit_section)
        self.widget_dic["tk_button_btn_CfgDef"].bind("<ButtonRelease-1>", self.default_section)
        self.widget_dic["tk_button_btn_CfgSave"].bind("<ButtonRelease-1>", self.write_to_config)
        self.widget_dic["tk_button_btn_ItemAdd"].bind("<ButtonRelease-1>", self.add_new_section)
        self.widget_dic["tk_button_btn_ItemDel"].bind("<ButtonRelease-1>", self.remove_section)
        self.after(100, self.read_config)
        self.after(200, self.select_default_config)

    
    def cancel(self):
        # win.init_chk()
        self.destroy()

        # 自动隐藏滚动条
    def scrollbar_autohide(self,bar,widget):
        self.__scrollbar_hide(bar,widget)
        widget.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        bar.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        widget.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
        bar.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
    
    def __scrollbar_show(self,bar,widget):
        bar.lift(widget)

    def __scrollbar_hide(self,bar,widget):
        bar.lower(widget)
        
    def __tk_list_box_list_CfgItem(self,parent):
        lb = Listbox(parent)
        lb.place(x=20, y=40, width=150, height=120)
        return lb

    def __tk_label_lab_CfgItem(self,parent):
        label = Label(parent,text="配置项目",anchor="center")
        label.place(x=20, y=10, width=150, height=30)
        return label

    def __tk_label_lab_AGENTID(self,parent):
        label = Label(parent,text="AGENTID",anchor="center")
        label.place(x=190, y=40, width=80, height=25)
        return label

    def __tk_label_lab_CORPID(self,parent):
        label = Label(parent,text="CORPID",anchor="center")
        label.place(x=190, y=75, width=80, height=25)
        return label

    def __tk_label_lab_CORPSECRET(self,parent):
        label = Label(parent,text="CORPSECRET",anchor="center")
        label.place(x=190, y=110, width=80, height=25)
        return label

    def __tk_input_text_AGENTID(self,parent):
        ipt = Entry(parent)
        ipt.place(x=280, y=40, width=240, height=25)
        return ipt

    def __tk_input_text_CROPID(self,parent):
        ipt = Entry(parent)
        ipt.place(x=280, y=75, width=240, height=25)
        return ipt

    def __tk_input_text_CORPSECRET(self,parent):
        ipt = Entry(parent)
        ipt.place(x=280, y=110, width=240, height=25)
        return ipt

    def __tk_button_btn_CfgSave(self,parent):
        btn = Button(parent, text="保存")
        btn.place(x=280, y=145, width=75, height=30)
        return btn

    def __tk_button_btn_ItemAdd(self,parent):
        btn = Button(parent, text="+")
        btn.place(x=190, y=145, width=40, height=30)
        return btn

    def __tk_button_btn_ItemDel(self,parent):
        btn = Button(parent, text="-")
        btn.place(x=235, y=145, width=40, height=30)
        return btn

    def __tk_button_btn_CfgDef(self,parent):
        btn = Button(parent, text="设置为默认值")
        btn.place(x=360, y=145, width=160, height=30)
        return btn
    
    def select_default_config(self):
        """选中默认配置项并显示其详细信息"""
        try:
            # 获取所有列表项
            items = self.widget_dic["tk_list_box_list_CfgItem"].get(0, END)
            # 查找默认配置项
            for i, item in enumerate(items):
                if '(默认)' in item:
                    # 选中该项
                    self.widget_dic["tk_list_box_list_CfgItem"].selection_clear(0, END)
                    self.widget_dic["tk_list_box_list_CfgItem"].selection_set(i)
                    # 创建一个虚拟的事件对象
                    class Event:
                        def __init__(self, widget):
                            self.widget = widget
                    # 触发显示配置信息
                    self.show_config(Event(self.widget_dic["tk_list_box_list_CfgItem"]))
                    break
        except Exception as e:
            showerror("错误", f"选择默认配置项失败：{str(e)}")

    def read_config(self):
        # # 检查配置文件是否存在并进行处理
        # if not os.path.exists(CONFIG_FILE):
        #     # 文件不存在，提示用户是否创建新文件
        #     res = askyesno('提示', '配置文件不存在，是否创建新的配置文件？')
        #     if res:
        #         # 用户确认创建新文件，使用tempfile模块生成一个随机的文件名
        #         with tempfile.NamedTemporaryFile(mode='w', delete=False) as f:
        #             new_file = f.name
        #         # 将读取的config对象保存到新文件中
        #         CONF_OBJ.read_dict({
        #             # 'DEFAULT': {},
        #             '配置项1': {
        #                 'agentid': '',
        #                 'cropid': '',
        #                 'screctid': '',
        #                 CONF_OBJ_default_section: 'False'
        #             }
        #         })
        #         with open(new_file, 'w') as f:
        #             CONF_OBJ.write(f)
        #          # 重命名新文件，相当于剪切+粘贴，覆盖原来的配置文件
        #         os.rename(new_file, CONFIG_FILE)
        #         CONF_OBJ.read(CONFIG_FILE)
        #     else:
        #             return
        # 文件存在，读取配置文件
        # win.chk_config()
        # win.init_chk()
        CONF_OBJ.read(CONFIG_FILE)
        # 清除列表框文件
        self.widget_dic["tk_list_box_list_CfgItem"].delete(0, END)
        
        for section in CONF_OBJ.sections():
            # 判断该section是否为默认选项
            if CONF_OBJ.getboolean(section, CONF_OBJ_default_section):
                self.widget_dic["tk_list_box_list_CfgItem"].insert(END, f'{section} (默认)') 
                # 添加“(默认)”标志
            else:
                self.widget_dic["tk_list_box_list_CfgItem"].insert(END, section)
            
            # 列表排序
            items = list(self.widget_dic["tk_list_box_list_CfgItem"].get(0, END))  # 获取所有选项
            items.sort()  # 对选项进行排序
            self.widget_dic["tk_list_box_list_CfgItem"].delete(0, END)  # 清空列表框
            for item in items:
                self.widget_dic["tk_list_box_list_CfgItem"].insert(END, item)  # 按照顺序添加排序后的选项
        
        win.init_chk()

  


    def show_config(self,event):
        """显示选中配置项的详细信息"""
        try:
            # 获取当前选中的列表项
            selection = event.widget.curselection()
            if not selection:
                showerror("错误", "请选择配置项")
                return
            
            # 获取当前选定的section的值
            section = event.widget.get(selection[0]).replace(' (默认)', '')
            
            # 验证配置节是否有效
            if not config_manager.validate_config(section):
                showerror("错误", f"配置节 {section} 不完整或无效")
                return
            
            # 获取配置信息并显示
            settings = config_manager.config[section]
            self.widget_dic["tk_input_text_AGENTID"].delete(0, END)
            self.widget_dic["tk_input_text_AGENTID"].insert(0, settings.get("agentid", ""))
            self.widget_dic["tk_input_text_CROPID"].delete(0, END)
            self.widget_dic["tk_input_text_CROPID"].insert(0, settings.get("cropid", ""))
            self.widget_dic["tk_input_text_CORPSECRET"].delete(0, END)
            self.widget_dic["tk_input_text_CORPSECRET"].insert(0, settings.get("screctid", ""))
        except Exception as e:
            showerror("错误", f"显示配置信息失败：{str(e)}")

    def edit_section(self,event):
        # 获取点击的列表项的内容
        old_section = self.widget_dic["tk_list_box_list_CfgItem"].get(self.widget_dic["tk_list_box_list_CfgItem"].curselection()).replace(' (默认)', '')

        # 弹出输入框并获取用户输入
        new_section = askstring("编辑配置项", "请输入新的配置项名称：", initialvalue=old_section)

        # 如果用户点击了取消按钮或者新的section名称为空，则不更新config文件并退出函数
        if not new_section:
            return
        elif new_section==old_section:
            return
        else:
            try:
                # 使用ConfigManager更新配置
                config_manager.config.add_section(new_section)
                for key in config_manager.config[old_section]:
                    config_manager.config[new_section][key] = config_manager.config[old_section][key]
                
                config_manager.config.remove_section(old_section)
                config_manager.save_config()
                
                # 刷新配置显示
                self.read_config()
                showinfo("成功", "配置项重命名成功")
            except Exception as e:
                showerror("错误", f"重命名配置项失败：{str(e)}")

    def add_new_section(self,event):
        """添加新的配置项"""
        try:
            while True:
                # 弹出输入框获取用户输入的配置项名称
                section_name = askstring("新建配置项", "请输入配置项名称：")
                
                # 如果用户取消输入或输入为空，则退出
                if not section_name:
                    return
                
                # 检查是否与现有配置项重名
                if section_name in config_manager.config.sections():
                    showerror("错误", "配置项名称已存在，请使用其他名称")
                    continue
                
                # 如果没有重名，跳出循环
                break
            
            # 添加新配置节
            config_manager.config.add_section(section_name)
            config_manager.config[section_name] = {
                "agentid": "",
                "cropid": "",
                "screctid": "",
                config_manager.default_section: "false"
            }
            
            config_manager.save_config()
            self.read_config()
            showinfo("成功", "新配置项添加成功")
        except Exception as e:
            showerror("错误", f"添加配置项失败：{str(e)}")

    def remove_section(self,event):
        """删除选中的配置项"""
        try:
            # 获取选中项
            selection = self.widget_dic["tk_list_box_list_CfgItem"].curselection()
            if not selection:
                showerror("错误", "请选择要删除的配置项")
                return
            
            section = self.widget_dic["tk_list_box_list_CfgItem"].get(selection[0]).replace(' (默认)', '')
            
            # 确认是否删除
            if not askyesno("确认", f"确定要删除配置项 {section} 吗？"):
                return
            
            # 检查是否为默认配置
            if config_manager.config.getboolean(section, config_manager.default_section, fallback=False):
                showerror("错误", "不能删除默认配置项")
                return
            
            # 删除配置节
            config_manager.config.remove_section(section)
            config_manager.save_config()
            
            # 清空输入框
            self.widget_dic["tk_input_text_AGENTID"].delete(0, END)
            self.widget_dic["tk_input_text_CROPID"].delete(0, END)
            self.widget_dic["tk_input_text_CORPSECRET"].delete(0, END)
            
            # 重新加载配置文件
            config_manager.load_config()
            
            # 刷新列表显示
            self.widget_dic["tk_list_box_list_CfgItem"].delete(0, END)
            for section in config_manager.config.sections():
                if config_manager.config.getboolean(section, config_manager.default_section):
                    self.widget_dic["tk_list_box_list_CfgItem"].insert(END, f'{section} (默认)')
                else:
                    self.widget_dic["tk_list_box_list_CfgItem"].insert(END, section)
            
            # 对列表项进行排序
            items = list(self.widget_dic["tk_list_box_list_CfgItem"].get(0, END))
            items.sort()
            self.widget_dic["tk_list_box_list_CfgItem"].delete(0, END)
            for item in items:
                self.widget_dic["tk_list_box_list_CfgItem"].insert(END, item)
            
            showinfo("成功", "配置项删除成功")
        except Exception as e:
            showerror("错误", f"删除配置项失败：{str(e)}")

    def default_section(self,event):
        # 获取点击的列表项的内容
        section = self.widget_dic["tk_list_box_list_CfgItem"].get(self.widget_dic["tk_list_box_list_CfgItem"].curselection()).replace(' (默认)', '')

        # 更新config文件，将选中的section设置为默认，并保存修改
        for section_name in CONF_OBJ.sections():
            if section_name == section:
                CONF_OBJ[section_name][CONF_OBJ_default_section] = "true"
            else:
                CONF_OBJ[section_name][CONF_OBJ_default_section] = "false"
        with open(CONFIG_FILE, 'w') as configfile:
            CONF_OBJ.write(configfile)

         # 清空列表框并重新添加更新后的config文件中的section
        self.read_config()

    def write_to_config(self,event):
        """保存配置信息到文件"""
        try:
            # 获取选中项
            selection = self.widget_dic["tk_list_box_list_CfgItem"].curselection()
            if not selection:
                showerror("错误", "请选择要保存的配置项")
                return
            
            section = self.widget_dic["tk_list_box_list_CfgItem"].get(selection[0]).replace(' (默认)', '')
            
            # 获取输入框中的值
            input_agentid = self.widget_dic["tk_input_text_AGENTID"].get().strip()
            input_cropid = self.widget_dic["tk_input_text_CROPID"].get().strip()
            input_screctid = self.widget_dic["tk_input_text_CORPSECRET"].get().strip()
            
            # 验证输入值不为空
            if not all([input_agentid, input_cropid, input_screctid]):
                showerror("错误", "所有字段都必须填写")
                return
            
            # 保存配置
            is_default = config_manager.config.getboolean(section, config_manager.default_section, fallback=False)
            config_manager.config[section] = {
                "agentid": input_agentid,
                "cropid": input_cropid,
                "screctid": input_screctid,
                config_manager.default_section: str(is_default).lower()
            }
            
            config_manager.save_config()
            self.read_config()
            showinfo("成功", "配置保存成功")
        except Exception as e:
            showerror("错误", f"保存配置失败：{str(e)}")
# 主窗口绘制代码类
class WinGUI(Tk):
    widgets: Dict[str, Widget] = {}  # 存储控件对象的字典
    

    def __init__(self):
        super().__init__()
        self.__win()
        # 将不同类型的控件及其相关参数放置在同一函数中，提高代码可读性和可维护性
        self.__create_listbox()
        self.__create_buttons()
        self.__create_labels()
        self.__create_text()
        self.__other_adapter()

    def __win(self):
        self.title("API消息发送助手")
        # 设置窗口大小、居中
        width = 700
        height = 400
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)

        # 自动隐藏滚动条
    def scrollbar_autohide(self,bar,widget):
        self.__scrollbar_hide(bar,widget)
        widget.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        bar.bind("<Enter>", lambda e: self.__scrollbar_show(bar,widget))
        widget.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
        bar.bind("<Leave>", lambda e: self.__scrollbar_hide(bar,widget))
    
    def __scrollbar_show(self,bar,widget):
        bar.lift(widget)

    def __scrollbar_hide(self,bar,widget):
        bar.lower(widget)

    # 列表框创建函数    
    def __create_listbox(self):
        # 相同类型的控件及其相关参数放置在同一函数中，提高代码可读性和可维护性
        self.widgets["lb_All_Item"] = Listbox(self, font=("宋体", 12), selectmode=MULTIPLE)
        self.widgets["lb_All_Item"].place(x=10, y=60, width=150, height=310)

        vbar = Scrollbar(self)
        self.widgets["lb_All_Item"].configure(yscrollcommand=vbar.set)
        vbar.config(command=self.widgets["lb_All_Item"].yview)
        vbar.place(x=145, y=60, width=15, height=310)
        self.scrollbar_autohide(vbar, self.widgets["lb_All_Item"])

        self.widgets["lb_Select_Item"] = Listbox(self, font=("宋体", 12))
        self.widgets["lb_Select_Item"].place(x=230, y=60, width=150, height=260)

        vbar2 = Scrollbar(self)
        self.widgets["lb_Select_Item"].configure(yscrollcommand=vbar2.set)
        vbar2.config(command=self.widgets["lb_Select_Item"].yview)
        vbar2.place(x=365, y=60, width=15, height=260)
        self.scrollbar_autohide(vbar2, self.widgets["lb_Select_Item"])
    
        self.widgets["lb_User_Item"] = Listbox(self, font=("宋体", 12))
        self.widgets["lb_User_Item"].place(x=230, y=330, width=150, height=40)

    #按钮创建函数
    def __create_buttons(self):
        self.widgets["btn_ItemAdd"] = Button(self, text="添加", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_ItemAdd"].place(x=170, y=60, width=50, height=40)
        self.widgets["btn_ItemAdd"].configure(command=self.btn_ItemAdd_click)


        self.widgets["btn_ItemDel"] = Button(self, text="删除", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_ItemDel"].place(x=170, y=110, width=50, height=40)
        self.widgets["btn_ItemDel"].configure(command=self.btn_ItemDel_click)

        self.widgets["btn_ItemUp"] = Button(self, text="向上", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_ItemUp"].place(x=170, y=185, width=50, height=40)
        self.widgets["btn_ItemUp"].configure(command=lambda: self.btn_move_item(-1))

        self.widgets["btn_ItemDown"] = Button(self, text="向下", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_ItemDown"].place(x=170, y=235, width=50, height=40)
        self.widgets["btn_ItemDown"].configure(command=lambda: self.btn_move_item(1))

        self.widgets["btn_ItemToUsers"] = Button(self, text="用户", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_ItemToUsers"].place(x=170, y=330, width=50, height=40)
        self.widgets["btn_ItemToUsers"].configure(command=self.btn_ItemToUsers_click)


        self.widgets["btn_GeneratePreview"] = Button(self, text="生成预览", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_GeneratePreview"].place(x=390, y=330, width=70, height=40)
        self.widgets["btn_GeneratePreview"].configure(command=self.btn_GeneratePreview_click)

        self.widgets["btn_PreviousPage"] = Button(self, text="上页 -", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_PreviousPage"].place(x=465, y=330, width=50, height=40)
        self.widgets["btn_PreviousPage"].configure(command=self.btn_PreviousPage_click)

        self.widgets["btn_NextPage"] = Button(self, text="下页 +", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_NextPage"].place(x=520, y=330, width=50, height=40)
        self.widgets["btn_NextPage"].configure(command=self.btn_NextPage_click)

        self.widgets["btn_Send"] = Button(self, text="单发", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_Send"].place(x=585, y=330, width=50, height=40)
        self.widgets["btn_Send"].configure(command=self.msg_single_send)


        self.widgets["btn_Send2"] = Button(self, text="群发", font=("宋体", 12), bg="lightblue", activebackground="blue")
        self.widgets["btn_Send2"].place(x=640, y=330, width=50, height=40)
        self.widgets["btn_Send2"].configure(command=self.msg_list_send)
    
    # 标签创建函数
    def __create_labels(self):
        self.widgets["lab_All_Item"] = Label(self, text="素材数据可用信息列", font=("宋体", 12), anchor="center")
        self.widgets["lab_All_Item"].place(x=10, y=30, width=150, height=30)

        self.widgets["lab_Select_Item"] = Label(self, text="已选择使用信息列", font=("宋体", 12), anchor="center")
        self.widgets["lab_Select_Item"].place(x=230, y=30, width=150, height=30)

        self.widgets["lab_Preview"] = Label(self, text="预览区域", font=("宋体", 12), anchor="center")
        self.widgets["lab_Preview"].place(x=390, y=30, width=80, height=30)
    # 文本框创建函数
    def __create_text(self):
        # self.widgets["text_Preview"] = Text(self, font=("宋体", 12), state="disabled")
        self.widgets["text_Preview"] = HTMLScrolledText(self, bg='SystemButtonFace',font=("宋体", 12), state="disabled")
        self.widgets["text_Preview"].place(x=390, y=60, width=300, height=260)
        
    def __other_adapter(self):
        global status_var,title_var
        global title_info
        title_var=IntVar()
        status_var=IntVar()
        title_info=''
    # 创建状态栏（配置信息概览和处理进度）

        # self.widgets["status_bar"] = Label(self, text=CONFIG_FILE, bd=1, relief=SUNKEN, anchor=W)
        # self.widgets["status_bar"].pack(side=BOTTOM, fill=X)

        # self.widgets["process_bar"] = Label(self, text=CONFIG_FILE, bd=1, relief=SUNKEN, anchor=E)
        # self.widgets["process_bar"].pack(side=BOTTOM, fill=X)

        self.widgets["status_bar_frame"] = Frame(self)
        self.widgets["status_bar_frame"].pack(side=BOTTOM, fill=X)

        self.widgets["process_bar_line"]= ttk.Progressbar(self.widgets["status_bar_frame"], orient='horizontal',mode='determinate')
        self.widgets["process_bar_line"].pack(pady=10,side=RIGHT,fill=X, expand=True)

        self.widgets["status_bar"] = Label(self.widgets["status_bar_frame"], text=CONFIG_FILE, bd=1, relief=SUNKEN,bg="blue", width=50, font=("Arial", 10))
        self.widgets["status_bar"].pack(side=LEFT, fill=X, expand=True)

        self.widgets["process_bar"] = Label(self.widgets["status_bar_frame"], text=Process_info, bd=1, relief=SUNKEN,bg="blue", width=30, font=("Arial", 10))
        self.widgets["process_bar"].pack(side=RIGHT, fill=X, expand=True)


        
    # 创建复选框，固定消息格式
        self.widgets["msg_format_set"] = Checkbutton(self, text="使用Markdown格式", variable=status_var)
        self.widgets["msg_format_set"].place(x=540, y=30, width=140, height=30)
    # 创建标题
        self.widgets["msg_title_set"] = Checkbutton(self, text="使用标题", variable=title_var,command=self.set_title_info)
        self.widgets["msg_title_set"].place(x=460, y=30, width=80, height=30)
        # self.widgets["btn_Send2"] = Button(self, text="群发", font=("宋体", 12), bg="lightblue", activebackground="blue")
        # self.widgets["btn_Send2"].place(x=640, y=330, width=50, height=40)
        # self.widgets["btn_Send2"].configure(command=self.msg_list_send)
# 主窗口交互代码类
class Win(WinGUI):
    file_status=0

    def __init__(self):
        super().__init__()
        self.menu = self.create_menu()
        self.config(menu=self.menu)
        self.Preview_table = []
        self.after(100,self.chk_config)
        self.after(100,self.conf_reload)
        # self.widgets["msg_format_set"].variavle.get
  
    def conf_reload(self):
        chk_result=self.init_chk()
        if chk_result == 1:
        # 需要弹出配置窗口
            self.menu_Config()
            # return
        else:
            showinfo("提示", "配置数据读取完成")

    def create_menu(self):
        menu = Menu(self, tearoff=False)
        file_menu = Menu(menu, tearoff=False)
        file_menu.add_command(label="打开消息素材文件...", command=self.menu_Openfile)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.menu_Quit)
        menu.add_cascade(label="文件", menu=file_menu)

        Config_menu = Menu(menu, tearoff=False)
        Config_menu.add_command(label="应用配置", command=self.menu_Config)
        Config_menu.add_command(label="关于", command=self.menu_About)
        menu.add_cascade(label="帮助", menu=Config_menu)

        return menu

    def menu_Config(self):
        seeting_window = cfgmanager(win)
        seeting_window.grab_set()

    def menu_About(self):
        with open(README_FILE, 'r') as file:
            md_text = file.read()
            text = markdown(md_text)
        about_window = AboutWindow(win, text)
        about_window.grab_set()

    def menu_Quit(self):
        self.quit()
    
    def menu_Openfile(self):
        # global file_path   # 使用全局变量
    # 打开文件选择对话框
        self.file_path = askopenfilename(filetypes=(("Excel files", "*.xls"), ("所有文件", "*.*")))
        if not self.file_path:
            self.file_status=0
            return
        # 加载工资文件
        try:
            workbook = xlrd.open_workbook(self.file_path)
            sheet_names = workbook.sheet_names()
            if not sheet_names:
                showerror("错误", "无效的Excel文件")
                self.file_status=0
                return
            sheet = workbook.sheet_by_name(sheet_names[0])
            # 获取表头
            header = [cell.value for cell in sheet.row(0)]
            # 更新列表项
            self.widgets["lb_All_Item"].delete(0, END)
            for h in header:
                self.widgets["lb_All_Item"].insert(END, h)
            self.file_status=1
        except Exception as e:
            showerror("错误", str(e))
            self.file_status=0
            return
        
    def btn_ItemAdd_click(self):
        selected_items = self.widgets["lb_All_Item"].curselection()  # 获取选中的列表项
        for index in selected_items:
            item = self.widgets["lb_All_Item"].get(index)  # 获取列表项文本
            self.widgets["lb_Select_Item"].insert(END, item)  # 添加到已选择使用信息列
        self.widgets["lb_Select_Item"].selection_clear(0, END)
        self.widgets["lb_Select_Item"].activate(END)
        self.widgets["lb_Select_Item"].selection_set(END)


    def btn_ItemDel_click(self):
        selected_items = self.widgets["lb_Select_Item"].curselection()  # 获取选中的列表项
        for index in selected_items:
            self.widgets["lb_Select_Item"].delete(index)  # 从已选择使用信息列中删除选中的列表项

    # 上下移动列表框中的选中项
    def btn_move_item(self,step):
        selected_items = self.widgets["lb_Select_Item"].curselection()
        if len(selected_items) != 1:
            return
        index = selected_items[0]
        pos = index + step
        if pos < 0 or pos >= self.widgets["lb_Select_Item"].size():
            return
        item = self.widgets["lb_Select_Item"].get(index)
        self.widgets["lb_Select_Item"].delete(index)
        self.widgets["lb_Select_Item"].insert(pos, item)
        self.widgets["lb_Select_Item"].selection_clear(0, END)
        self.widgets["lb_Select_Item"].activate(pos)
        self.widgets["lb_Select_Item"].selection_set(pos)

    def btn_ItemToUsers_click(self):
        selected_items = self.widgets["lb_All_Item"].curselection()  # 获取选中的列表项
        self.widgets["lb_User_Item"].delete(0, END)  # 清空原有内容
        for index in selected_items:
            item = self.widgets["lb_All_Item"].get(index)  # 获取列表项文本
            self.widgets["lb_User_Item"].insert(END, item)  # 添加到用户列
        self.widgets["lb_User_Item"].selection_clear(0, END)
        self.widgets["lb_User_Item"].activate(END)
        self.widgets["lb_User_Item"].selection_set(END)

    def btn_GeneratePreview_click(self):
        # self.Preview_items = [self.widgets["lb_Select_Item"].get(i) for i in self.widgets["lb_Select_Item"].curselection()]  # 获取选择的列表项
        self.Preview_table=[]
        # self.current_row_index=1
        User_check_items=self.widgets["lb_User_Item"].get(0, END)
        MsgBody_check_items=self.widgets["lb_Select_Item"].get(0, END)
        Preview_items = User_check_items+MsgBody_check_items
        if self.file_status==0:
            result = askyesno("提示", "未打开任何xls数据表格，是否打开?")
            if result:
                self.menu_Openfile()
                return
            else:
                return
        elif len(User_check_items) == 0 and len(MsgBody_check_items) == 0:  # 没有选择任何列表项
            showinfo("提示", "缺少接收者信息和信息主体！")
            return
        elif len(User_check_items) == 0:
            showinfo("提示", "缺少接收者信息！请选定好用户字段。")
            return
        elif len(MsgBody_check_items) == 0:
            showinfo("提示", "缺少信息主体！请选定好使用的信息列。")
            return
        try:
            with open(self.file_path, 'rb') as f:
                Preview_workbook = xlrd.open_workbook(file_contents=f.read())
                Preview_sheet = Preview_workbook.sheet_by_index(0)
                   # 获取数据子集
            for Preview_sheet_row_index in range(Preview_sheet.nrows):
                Preview_table_convert = {}
                for Preview_sheet_column_name in Preview_items:
                    if Preview_sheet_column_name not in Preview_sheet.row_values(0):
                        showerror("提示", '打开的文件中未能找到对应的列.')
                    Preview_sheet_column_index = Preview_sheet.row_values(0).index(Preview_sheet_column_name)
                    cell_value = Preview_sheet.cell_value(Preview_sheet_row_index, Preview_sheet_column_index)
                    Preview_table_convert[Preview_sheet_column_name] = cell_value
                self.Preview_table.append(Preview_table_convert)
            self.current_row_index=1 #给出预览索引初始位置
            self.show_row()  
        except Exception as e:
            showerror("提示", str(e))

    # 显示table的行数据
    def show_row(self):
        format_status = status_var.get()
        # 逻辑判断标题栏复选框并作出
        if title_var.get()==1:
            self.row_str=self.format_title_info+'\n' #携带非空标题
            # 界定插入行的标记
            row_replace=1
        else:
            self.row_str='\n' #携带空标题
            # 界定插入行的标记
            row_replace=1
        # self.row_str = ''
        for key, value in self.Preview_table[self.current_row_index].items():
            # self.row_str += f"{key}: {value}\n" # 原始数据形态
            # markdown自定义后形态判定
            if format_status==1:
                # 携带markdown格式标记
                # self.row_str += f"> <font color=\"info\">**{key}**</font>: {value}  \n"
                self.row_str += f"> **<font color=\"info\">{key}</font>**: {value}  \n"
            else:
                # 不携带markdown标记
                self.row_str += f"{key}: {value}  \n"
            # 另外一种markdown格式形式样式
            # self.row_str += f"### <font color=\"info\">**{key}**</font>: *{ value }*  \n"
        self.row_str += f'{COPYRIGHT_INFO}  \n'
        # 获取第一个键值的值，作为touser的用户名
        self.to_users = self.Preview_table[self.current_row_index][list(self.Preview_table[self.current_row_index].keys())[0]]
        # 文本预览插入非发送字符的提示和标记，定位在row_replace定义的行数，识别以两个空格+\n结尾的字符。
        self.row_str = self.row_str.replace('  \n', '  \n ---(接收用户，此行不作为发送信息)---  \n', row_replace)
        # 消息主体保留第一行，不保留第二、三行，包括接收者以及接受着注释
        self.current_message = '\n'.join(self.row_str.split('\n')[:1] + self.row_str.split('\n')[3:])
        # self.current_message = '  \n'.join(self.row_str.split('  \n')[2:])        
        self.widgets["text_Preview"].set_html(markdown(self.row_str)) #显示在预览栏目里的信息主体


    def btn_PreviousPage_click(self):
        try:
            if self.current_row_index > 1:
                self.current_row_index -= 1
                self.show_row()
            else:
                showinfo("提示", "无向上数据")
        except AttributeError:
            self.btn_GeneratePreview_click()

        # 其他实现途径，已被优化。
        # current_row = int(self.widgets["text_Preview"].index("end-1c").split('.')[0])  # 获取当前显示的最后一行的行号
        # if current_row <= 2:  # 已经到达第一行，无法向上翻页
        #     return
        # self.widgets["text_Preview"].yview_scroll(-self.preview_page_size, "units")  # 上移一页
        # self.show_preview_page(-1)  # 切换预览页，并重新显示
        
    def btn_NextPage_click(self):
        # if not self.current_row_index:
        #     self.btn_GeneratePreview_click()
        try:
            if self.current_row_index < len(self.Preview_table)-1:
                self.current_row_index += 1
                self.show_row()
            else:
                showinfo("提示", "无向下数据")
        except AttributeError:
            self.btn_GeneratePreview_click()


    def init_chk(self):
        global conf_AGENTID,conf_CROPID,conf_SCRECTID
        CONF_OBJ = configparser.ConfigParser()
        CONF_OBJ.read(CONFIG_FILE)
        default_values = set()   #集合中存储所有默认值
        default_count = 0   #存储value为true的DEFAULT_SECTION_VALUE的数量
        default_sections = []   #存储value为true的DEFAULT_SECTION_VALUE的DEFAULT_SECTION
        for section in CONF_OBJ.sections():
            if CONF_OBJ.has_option(section, CONF_OBJ_default_section):
                chk_value = CONF_OBJ.get(section, CONF_OBJ_default_section)
                if chk_value == "true":
                    default_values.add(chk_value)
                    default_count += 1
                    default_sections.append(section)
        if default_count == 0:
            showinfo("提示", "无默认配置项，进行默认设置.")
            return 1
        elif default_count == 1:
            if CONF_OBJ.get(default_sections[0], CONF_OBJ_default_section) == "true":
                cfg_default_section = default_sections[0]
                conf_AGENTID = CONF_OBJ.get(cfg_default_section, "agentid")
                conf_CROPID = CONF_OBJ.get(cfg_default_section, "cropid")
                conf_SCRECTID = CONF_OBJ.get(cfg_default_section, "screctid")
                # showinfo("提示", "配置数据读取完成")
                self.widgets["status_bar"].config(text="当前配置文件为：" + CONFIG_FILE+"； 当前配置项为：" + cfg_default_section)
                # 创建发送类对象，并引用全局变量conf_AGENTID、conf_CROPID、conf_SCRECTID
                # self.mod_sending =WeChat()
                print(conf_AGENTID,conf_CROPID,conf_SCRECTID)
                return 0
            else:
                showinfo("错误", "无法读取默认配置项")
                return 1
        elif default_count > 1:
            default_text = " ".join(default_sections)
            showinfo("提示", "存在多个默认选项，请从以下分区中选择一个配置为默认：\n\n{}".format(default_text))
            return 1


    def chk_config(self):
        # 检查配置文件是否存在并进行处理
        if not os.path.exists(CONFIG_FILE):
            # 文件不存在，提示用户是否创建新文件
            showinfo('提示', '配置文件不存在，即将创建新的配置文件。')
            # 用户确认创建新文件，使用tempfile模块生成一个随机的文件名
            with tempfile.NamedTemporaryFile(mode='w', delete=False) as f:
                new_file = f.name
            # 将读取的config对象保存到新文件中
            CONF_OBJ.read_dict({
            # 'DEFAULT': {},
                write_item_section: {
                    write_item_agentid: '',
                    write_item_cropid: '',
                    write_item_screctid: '',
                    CONF_OBJ_default_section: 'true'
                }
            })
            with open(new_file, 'w') as f:
                CONF_OBJ.write(f)
                # 重命名新文件，相当于剪切+粘贴，覆盖原来的配置文件
            os.rename(new_file, CONFIG_FILE)

        else:
            return
        # 文件存在，读取配置文件

    def msg_single_send(self):
        # 创建发送类对象，并引用全局变量conf_AGENTID、conf_CROPID、conf_SCRECTID
        self.mod_sending =WeChat() 
        # self.mod_sending._get_access_token()       
        try:
            self.mod_sending.send_message(self.current_message,'markdown',self.to_users)
            self.widgets["process_bar"].config(text="消息已发送至self.to_users，发送完成")
            self.widgets["process_bar_line"]['value']=100
        except AttributeError:
            # self.btn_GeneratePreview_click()
            showinfo("提示", "无预览数据，请确认是否具有预览数据")

        return

    def msg_list_send(self):
        # 创建发送类对象，并引用全局变量conf_AGENTID、conf_CROPID、conf_SCRECTID
        self.mod_sending =WeChat() 
        # decide_push=0
        # self.mod_sending._get_access_token()
        self.btn_GeneratePreview_click()
        self.send_range=len(self.Preview_table)
        result = askyesno("提示", "是否批量发送?")        
        if result:
            for i in range(1, self.send_range):
                # self.to_users = self.Preview_table[self.current_row_index][list(self.Preview_table[self.current_row_index].keys())[0]]
                self.current_row_index =i
                self.show_row()
                self.mod_sending.send_message(self.current_message,'markdown',self.to_users) 
                
                # 延时函数中插入插入进度条函数，带入当前i的值以及范围值
                self.after(100,self.process_bar_moving(i,self.send_range))
                # self.widgets["process_bar"].config(text=f"消息已发送至{self.to_users}，发送完成。当前第{i}位/总共{len(self.Preview_table)-1}位")
                # self.widgets["process_bar_line"]['value']=i/self.send_range*100
                # time.sleep(0.2)
        else:
            return
    
    def set_title_info(self):
        global title_info
        if title_var.get()==1:
            self.format_title_info='###### '+title_info
            new_title_info = askstring("编辑标题", "请输入新的标题:", initialvalue=title_info)
            if new_title_info:
                title_info=new_title_info
                self.format_title_info='###### '+new_title_info
        else:
            self.format_title_info=''
    

    def process_bar_moving(self,process_bar_value,process_bar_send_range):
        self.widgets["process_bar"].config(text=f"{self.to_users}已发送，第{process_bar_value}位/共{process_bar_send_range-1}位")
        self.widgets["process_bar_line"]['value']=process_bar_value/(process_bar_send_range-1)*100
        

if __name__ == "__main__":
    win = Win()
    win.mainloop()