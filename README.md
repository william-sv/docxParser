# docxParser
解析docx文档，输出文章内容及图片（图片默认输出图片位置~）

#### 方案

1. 使用 zipfile 获取 docx文档中的 document.xml 数据
2. 使用 xml.dom.minidom.parseString() 解析文档，取出docx文档中的段落内容
3. 读取document.xml.rels 文件，匹配 图片资源的位置


#### 用法

##### 获取word文档内容，包含文字及图片所在位置（图片以 {image1} 表示）
 ```python
 # params: file_path  word文档绝对路径
 # return: <p>文字</p>{image1}<p>文字</p>
 DocxParser().get_content(file_path)
 ```
 
 --- 
 
 ##### 导出文档图片
 ```python
 # params: file_path word文档绝对路径
 # params: save_path  图片保存地址
 # return: 最终图片地址：图片保存地址/word/media/
 DocxParser().get_media(file_path，save_path)
 ```
 
 
   
 
 
 
 
