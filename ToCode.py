from GetData import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

class ToCode():
    def __init__(self):

        self.win=Tk()
        win=self.win

        try:
            win.iconbitmap('config/icon.ico')
        except:
            messagebox.showerror(title='一个很确定的错误框',message="虽然报错了，但是代码不可能错的，必然是配置文件有问题")
            return
        

        frame=LabelFrame(win,text='说明')
        frame.pack()

        message='''
        本软件由河北大学JoJen设计
        配合B站UP主武汉大学Yophone同学所编写的Word论文插件使用
        主要功能为将Excel编写的配置文件转换为VBA代码
        让不了解编程的同学们也可以方便的制作自己适用的插件
        软件本身存在不少bug，但是我懒得改
        软件取决于心情可能会更新也可能不会
        本软件只负责依据自定义的配置信息生成代码
        插件的安装和使用可以参照Yophnoe同学的教程
        https://www.bilibili.com/video/BV1xT4y15777
        插件问题可以加QQ群1072641745交流
        软件问题可以加我私人QQ：2041463817
        有更新的话可能发布在这里：https://github.com/JoJenH/-VBA-
        '''

        Message(frame,text=message).pack()
        Button(text='点它选择写好的配置文件就没了',command=self.click).pack()
        win.mainloop()

    def click(self):
        win=self.win
        
        config_name=filedialog.askopenfilename(title='选取配置文件',filetypes=[('Excel文档','xlsx'),('Excel文档','xls')])
        if config_name=='':
            return
        
        code_name='config/code.txt'
        code=''
        try:
            with open(code_name,'r',encoding='utf-8') as f:
                code=f.read()
        except:
            messagebox.showerror(title='一个很确定的错误框',message="虽然报错了，但是代码不可能错的，必然是配置文件有问题")
            return

        data=GetData(config_name)
        data_dic=data.get_data()
        
        for data in data_dic:
            code=code.replace(data,str(data_dic[data]))
        
        win.clipboard_clear()
        win.clipboard_append(code)
        
        if messagebox.askokcancel(title='一个很不确定的消息框', message='不出意外的话,估摸着应该是成功了\n且相关代码已复制至剪贴板，是否需要另存为文件？'):
            fname = filedialog.asksaveasfilename(title=u'保存文件', filetypes=[("文本文档", ".txt")],initialfile='new_code')
            try:
                with open(fname+'.txt','w',encoding='utf-8') as f:
                    f.write(code)
            except:
                messagebox.showerror(title='一个很确定的错误框',message="虽然不知道为啥，但就是报错了")
                return
            messagebox.showinfo(title="又一个很不确定的消息框",message="不出意外的话，估摸着应该是写入到当前目录了")
       
if __name__=='__main__':
    ToCode()