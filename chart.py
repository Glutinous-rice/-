import openpyxl
from openpyxl.chart import BarChart,Series,Reference
wb=openpyxl.Workbook()
ws=wb.active
x=[('A','B','C'),           #新建一个list
   (1,20,30),
   (21,50,60),
   (3,80,90),
   (9,50,60),
   (15,40,30),
   (12,20,30),
   (7,10,50)
   ]
for i in x:
    ws.append(i)    #把这些都添加到单元格里

chart1=BarChart()    #创建图表
chart1.title='ziMu'
chart1.type= 'col'  #设置水平
chart1.style=10
chart1.x_axis.title='big'     #X轴名字
chart1.y_axis.title='small'    #Y轴名字

cats=Reference(ws,min_row=2,max_row=7,min_col=1)    #采集X轴数据
datas=Reference(ws,min_row=1,max_row=7,min_col=2,max_col=3)   #采集图表内数据
chart1.add_data(datas,titles_from_data=True)        #把数据添加到图表内
chart1.set_categories(cats)                         #设置X轴
ws.add_chart(chart1,'A10')                         #将图标添加到A10起始的单元格

from copy import deepcopy
chart2=deepcopy(chart1)
chart2.title='Horizontal Bar Chart'
chart2.type='bar'
ws.add_chart(chart2,'K10')


chart3=deepcopy(chart1)     #复制chart1的属性
chart3.type='col'        #通过将类型分别设置为 col或bar，在垂直和水平条形图之间切换。
chart3.title='Stacked Chart'
chart3.overlap=100     #使用堆叠图表时，重叠需要设置为100
chart3.grouping = "stacked"
ws.add_chart(chart3,'A27')


chart4=deepcopy(chart1)
chart4.type='bar'      #通过将类型分别设置为 col或bar，在垂直和水平条形图之间切换。
chart4.title='Percent Stacked Chart'
chart4.overlap=100   #使用堆叠图表时，重叠需要设置为100
chart4.grouping = "percentStacked"     #设置类型
ws.add_chart(chart4,'K27')


wb.save('charts-copy.xlsx')







