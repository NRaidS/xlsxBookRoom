# Introduction

本项目用于模拟维护一个教室预约系统。

模拟预约并显示下一周的教室预约情况。

#  安装流程

### 下载python运行环境

安装python运行环境（解释器）：
可用教程： [超详细的Python安装和环境搭建](https://blog.csdn.net/qq_53280175/article/details/121107748)

推荐安装Python3.9或以上版本

### 下载项目文件

使用git下载，[git安装教程:](https://www.cnblogs.com/xiaoliu66/p/9404963.html)

> git clone https://github.com/NRaidS/xlsxBookRoom.git

或直接在github网页端下载后直接解压。

### 依赖库安装

打开cmd，切换命令行到本工程所在目录:
> cd clstmeSystem

执行以下命令安装依赖:
>python -m pip install -U pip

>pip install -r requirements.txt


### 运行程序

直接点击main.py文件执行，或者进入在cmd中输入
>python main.py

### 基础表格文件介绍

stuNo.xlsx存放学生学号和姓名

classInfo.xlsx存放班级信息

clstme.xlsx存放预约信息


# 使用手册

项目使用样例视频：[基于py实现的教室预约系统](https://www.bilibili.com/video/BV1K24y1R7Wx/?share_source=copy_web&vd_source=fa6ceda0e61840db504a95bcd25f6d74)
本项目功能主要有登录，菜单，预约教室，显示目前全部的预约信息，查询所有教室，查询本人预约情况，修改本人预约信息，退出系统
其功能相互关系如下图所示。

![image](image/menu.png )


### 0、登录

在登录系统根据提示分别输入学号和姓名。验证成功后即可进入菜单操作。
### 1、菜单

在菜单中可以输入数字1-6分别对应6个功能：预约教室，显示目前全部的预约信息，查询所有教室，查询本人预约情况，修改本人预约信息，退出系统

### 2、预约教室

![image](image/book.png )

### 3、显示目前全部的预约信息

读取clstme.xlsx文件并输出成一张图表供查看。
### 4、查询所有教室

读取classInfo.xlsx文件，并将信息输出到控制台。
### 5、查询本人预约情况

读取clstme.xlsx文件，只输出本人预约的信息。
### 6、修改本人预约信息

![image](image/change.png )

### 7、退出系统

退出程序

# 不足之处

1、显示目前全部的预约信息时，制作出的图表不够美观。

2、显示本人预约信息时，并没有严格按照时间顺序进行排序。

3、需手动添加教室，学生。

4、在修改信息时，修改前后并没有实现排他锁和共享锁，在多线程操作时可能会导致读脏数据等一系列问题。
