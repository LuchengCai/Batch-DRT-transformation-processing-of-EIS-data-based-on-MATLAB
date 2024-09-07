# 基于MATLAB实现EIS数据的批量DRT变换处理

[实操演示视频](https://member.bilibili.com/platform/upload-manager/article)

*******************************本项目基于MATLAB实现EIS数据的批量DRT变换处理*********************************

使用说明

作者：Lucheng Cai，浙江大学材料科学与工程学院Wei-Qiang Han老师课题组直博生

1.使用DRT拟合可修改以下参数（line55-line61）

2.本代码默认使用贝叶斯法拟合（Bayesian Run），如不能使用此法，将使用简单拟合（Simple Run）。具体可见代码运行完毕后的提示。

3.用户需提供包含EIS数据的excel表，每个EIS数据包含‘频率’‘阻抗实部’‘阻抗虚部’分列保存，不同EIS数据放在excel表不同sheet中，sheet命名应符合规则（尽量以英文字母和数字组合）。具体见EIS数据模板。

4.用户需根据EIS数据中‘频率’‘阻抗实部’‘阻抗虚部’的位置修改line97-line99的代码，具体见该行相应注释。


使用方法

1.新建项目文件夹，下载本项目所有文件（后缀.m）到项目文件夹

2.将多个EIS数据放入到项目文件夹中的excel表（新建）中，具体要求见上述 使用说明3.

3.打开matlab，将matlab左侧文件夹设置为项目文件夹，打开本项目中的import_data.m文件

4.修改拟合参数（line55-line61），一般默认即可。

5.关键！关键！关键！根据EIS数据中‘频率’‘阻抗实部’‘阻抗虚部’的位置修改line97-line99的代码，具体见该行相应注释。

6.点击matlab上方运行按钮，选定目标excel表，等待代码运行完毕。

7.关键！关键！关键！作图。在生成的totalDRT的excel表中，每两列为一组EIS数据，各组的前一列为tau，后一列为MAP。整体复制到origin中，前者为横轴，后者为纵轴即可作图。


本项目在香港科技大学Francesco Ciucci老师工作的基础上进行改进，具体请见：https://github.com/ciuccislab/DRTtools

在此对Francesco Ciucci老师抱以敬意！


感谢好朋友Chenchen Xu的指导解惑（github：https://github.com/InvincibleGuy777）


[How to use this demo](https://member.bilibili.com/platform/upload-manager/article)

*******************************Batch DRT transformation processing of EIS data based on MATLAB*********************************

Instructions:

Author: Lucheng Cai, a doctoral candidate in the research group of Professor Wei-Qiang Han at the School of Materials Science and Engineering, Zhejiang University.

1.Parameters for DRT fitting can be modified (line 55-line 61).

2.This code defaults to using Bayesian Run for fitting. If this method cannot be used, a Simple Run fitting will be applied. Details can be seen in the prompt after running the code.

3.Users need to provide an Excel containing EIS data. Each EIS data should include columns for 'Frequency,' 'real part of impedance (Re),' and 'imaginary part of impedance (Im).' 

  Different EIS data should be placed in different sheets of the Excel, and the sheet names should follow the naming convention (preferably a combination of letters and numbers). 

  Refer to the EIS data template for specifics.

4.Users should modify the code at lines 97-99 according to the positions of 'Frequency,' 'Re,' and 'Im' in the EIS data.

Directions:

1.Create a new project folder and download all files of this project (with the .m extension) into the project folder.

2.Place multiple EIS data into an Excel (create a new one) in the project folder, following the requirements mentioned in the Instructions above.

3.Open Matlab, set the folder on the left side to the project folder, and open the import_data.m file in this project.

4.Modify the fitting parameters (line 55-line 61), usually the default settings are sufficient.

5.Crucial! Crucial! Crucial! Modify the code at lines 97-99 according to the positions of 'Frequency,' 'Re,' and 'Im' in the EIS data. 

  Refer to the corresponding comments in those lines for specifics.

6.Click the run button at the top of Matlab, select the target Excel table, and wait for the code to finish running.

7.Crucial! Crucial! Crucial! Plotting. In the generated total DRT Excel table, each two columns represent one set of EIS data, where the first column is tau and the second column is MAP. 

 Copy the data to another sheet, using the former as the x-axis and the latter as the y-axis for plotting.


This project is an improvement based on the work of Professor Francesco Ciucci at the Hong Kong University of Science and Technology. 

For more details, please visit: https://github.com/ciuccislab/DRTtools

A respectful acknowledgement to Professor Francesco Ciucci!


Thanks to my friend Chenchen Xu for guidance and assistance (github: https://github.com/InvincibleGuy777)
