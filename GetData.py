import xlrd

class GetData:
    def __init__(self,file_name):
        xlsx=xlrd.open_workbook(file_name)
        self.table=xlsx.sheet_by_name(xlsx.sheet_names()[0])

        self.row_dic={}# 列标字典
        self.col_dic={}# 行标字典
        self.data_dic={}#参数字典

        for i in range(2,self.table.ncols):
            self.col_dic[str(self.table.cell_value(0,i))]=i

        format_name=''
        for i in range(1,self.table.nrows):
            if self.table.cell_value(i,0)!='':
                format_name=self.table.cell_value(i,0)
            self.row_dic[format_name+str(self.table.cell_value(i,1))]=i

    def get_value(self,row,col):
        return self.table.cell_value(self.row_dic[row],self.col_dic[col])

    def get_data(self):
        for row in self.row_dic:
            for col in self.col_dic:
                self.data_dic[row+col]=str(self.get_value(row,col))
        return self.data_dic
