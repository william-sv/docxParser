# docxParser
解析docx文档，输出文章内容及图片（图片默认输出图片位置~）

用法
 DocxParser().get_content(file_path) 获取word文档内容，包含文字及图片所在位置（图片以 {image1} 表示）
 file_path  word文档绝对路径
 输出： <p>文字</p>{image1}<p>文字</p>
 DocxParser().get_media(file_path，save_path)  导出文档图片
 file_path word文档绝对路径
 save_path  图片保存地址 （最终图片地址：图片保存地址/word/media/）
 
 
