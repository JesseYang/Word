ExportExcel
===========

运行方式
-----------
首先执行ps -ef | grep 'rackup'找到正在运行的rackup，确定是本项目rackup进程情况下，sudo kill -9 #{process id}杀掉进程

项目文件夹命令行下切换到jruby环境，执行rackup config.ru &> log &，http服务会自动并在后台运行，日志文件写入到项目文件夹下的log中，log中可以查看服务端口号

参数定义：
-----------
>
title: 标题
> 
category\_label: 类别标签，逗号分割
>
series\_label: 系列标签，逗号分割
>
chart\_type: 图表类型，可以是ring（全环图），half\_ring（半环图），stack（条形图），bar1（二维柱状图），bar2（一维柱状图）
>
value\_axis: 坐标轴文字
>
data: 数据，一个系列内逗号分割，系列之间'-'分割

导出一张图片示例：
-----------
发送get请求到118\.194\.61\.82:9292/export.json
>
118\.194\.61\.82:9292/export.json?title=图标题&data=10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10&category\_label=一,二,三,四,五,六,七,八,九,十,十一,十二,十三,十四,十五,十六,十七&series\_label=哈哈&chart\_type=ring&value\_axis=population
>
118\.194\.61\.82:9292/export.json?title=图标题&data=10,20,30,40,50&category\_label=一,二,三,四,五&series\_label=哈哈&chart\_type=half\_ring&value\_axis=population
>
118\.194\.61\.82:9292/export.json?title=图标题&data=10,10,10,10-20,20,20,20-30,30,30,30-40,40,40,40-50,50,50,50&category\_label=一,二,三,四&series\_label=one,two,three,four,five&chart\_type=bar1&value\_axis=population

导出多张图片示例：
-----------
发送post请求到118\.194\.61\.82:9292/export.json，参数export_data为数组，每个元素代表一张图的数据，结构为
>
data：二维数组，元素为数值
>
category\_label：数组，元素为类别标签
>
series\_label：数组，元素为系列标签
>
title：字符串，图标题
>
chart\_type：字符串，图类型
>
value\_axis：字符串，轴显示文字
