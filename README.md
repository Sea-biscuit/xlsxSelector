# xlsxSelector
一个简单的Excel脚本，支持分割、合并、查重，同时进行空白格删除
### 关于一些小问题
输出的格式声明了是UTF-8 BOM编码方式，这个方案只适合在Windows系统内使用，使用Linux和MacOS时使用该脚本可能会产生意想不到的bug  
在dist文件中存放了两个xlsxSelector文件，存储在***xlsxSelector文件夹内的exe文件***需要计算机部署pip环境并且安装**pandas+numpy+openpyxl+fuzzywuzzy**这四个库，而***最外层的exe文件***则可直接运行——区别是可直接运行的exe会卡一点
  
### 更新说明
1.更新了读取文件地址的逻辑，使得交互更加便捷  
2.更新了GUI版本，基于Web实现  
3.初步实现了功能完善  
4.更新了补佳乐使用方法 :>  
  
##### Todo
√ 模糊查重  
√ GUI  
□ 简化代码  