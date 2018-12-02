#!/usr/bin/env python
#-*- coding:utf-8 -*-
# author:an
# datetime:2018/11/30 13:41
# software: PyCharm

import wx
import wx.grid
import numpy
import xlrd
class GridFrame(wx.Frame):
    def __init__(self, parent,data_path):
        wx.Frame.__init__(self, parent,title = u"邵天竺保研管理系统",size = wx.Size((1275,500)))
        RowSize = 20
        ColSize = 50
       # Create a wxGrid object
        grid = wx.grid.Grid(self, -1)
        # Then we call CreateGrid to set the dimensions of the grid
        # (100 rows and 10 columns in this example)

        data = self.load_file(data_path)#读取数据
        data = self.update_list(data)#将数据排序
        nrows = len(data)
        ncols = len(data[0])
        grid.CreateGrid(nrows, ncols)

        for row in range(0,nrows):
            data_row = data[row]
            if row == 0:
                for i in range(0,ncols):
                    grid.SetCellBackgroundColour(row,i,wx.GREEN)
                    if i<6:
                        grid.SetColLabelValue(i, str(data_row[i]))
                    else:
                        grid.SetColLabelValue(i,'课程'+str(int(data_row[i])))
            else:
                grid.SetRowSize(row, RowSize)
                #grid.SetColFormatFloat(row, 6, 2)
                grid.SetRowLabelSize(40)
                for col in range(0,ncols):
                    d = str(data_row[col])
                    #print(row,col)
                    if row<20:
                        grid.SetCellBackgroundColour(row, col, wx.GREEN)
                    grid.SetColSize(col, ColSize)
                    grid.SetCellValue(row-1, col,d)
        button = wx.Button(grid,label = '更新/提交',pos = (1190,5),size = (80,25))
        self.Show()
    def load_file(self,file_path):
        # data_path=data_path.decode('utf-8')
        data = xlrd.open_workbook(file_path)  # 获取Excel的数据
        # 获取sheet
        table_data = data.sheet_by_index(0)
        data_list = list()
        for row in range(0, table_data.nrows):
            data_list.append(table_data.row_values(row))
        return data_list
    def tackSecond(self,element):
        return element[1]
    def update_list(self,data_list):
        for row in range(1,len(data_list)):
            #for col in range(6,len(data_list[row])):
            data_list[row][5] = round(sum(data_list[row][i] for i in range(6,len(data_list[row])))/18,2)
            data_list[row][1] = data_list[row][2] + data_list[row][3] + data_list[row][4] +data_list[row][5]
        first_row = data_list[0]
        del data_list[0]
        data_list.sort(key= self.tackSecond,reverse=True)
        print(type(numpy.array(data_list)[:,1].argsort()))
        data_list.insert(0, first_row)
        print(data_list)
        return data_list





        # 获取一行的数值，例如第5行
        #rowvalue = table.row_values(5)
        # And set grid cell contents as strings
        # We can specify that some cells are read.only
        # grid.SetCellValue(0, 3, 'This is read.only')
        # grid.SetReadOnly(0, 3)
        #
        # # Colours can be specified for grid cell contents
        # grid.SetCellValue(3, 3, 'green on grey')
        # grid.SetCellTextColour(3, 3, wx.GREEN)
        # grid.SetCellBackgroundColour(3, 3, wx.LIGHT_GREY)
        #
        # # We can specify the some cells will store numeric
        # # values rather than strings. Here we set grid column 5
        # # to hold floating point values displayed with width of 6
        # # and precision of 2
        # grid.SetColFormatFloat(5, 6, 2)
        # grid.SetCellValue(0, 6, '3.1415')







if __name__ == '__main__':
    app = wx.App(0)
    frame = GridFrame(None,'data.xlsx')
    app.MainLoop()