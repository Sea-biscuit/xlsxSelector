# xlsxSelector
一个简单的Excel截取表格导出脚本
### 关于一些小问题1
首先是读取Excel表格行数的问题，我为了去除第一行的文字，将起始读取位置设在了第二行——在这个意义下，第二行的内容会被脚本识别为第一行，因此在
<img width="929" height="90" alt="image" src="https://github.com/user-attachments/assets/dd0d657f-e2ac-4ce1-b33c-4d3971e64321" />
此处，需要将所选的行数减去1
### 关于一些小问题2
导出的时候命名会以xxxx_part1.xlsx开始输出，part1开始到partxxx，暂时无法修改输出格式

#### 一些碎碎念
大量的数据就不要用Excel文件了...几十万行的表用一下MySQl方便又安全XD
感谢AI
科技改变生活
