本程序主要用于将office文件转为pdf文件  

支持word转pdf，excel转pdf，ppt转pdf，txt转pdf，html转pdf，等等
支持MS Office和WPS Office产生的上述三种文件转换

本程序保护两部分，一部分是调用安装的Office COM组件；另外一部分是exe可执行工具

Office2PDF 编译后产生Office2PDF.dll
PDFUtil 编译后生成一个exe命令行可执行文件；调用Office2PDF.dll将输入的参数所指文件转为pdf。用法如下：  
..PDFUtil.exe sourceOfficeFile targetPDFFile