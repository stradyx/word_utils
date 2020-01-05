# word_utils
两种方式实现Word文件的生成和读取

### 使用poi生成和读取

在读取word解析为html文件时，需要判断.doc和.docx不同的文件格式使用不同的解析方法

在生成word的时候生成的是.docx文件，需要替换的位置用 {{name}} 代替

#### 使用aspose-words读取

两种类型直接读取，不过没有授权文件的话会有水印生成

#### 使用freemarker生成

目前就.doc文件可以生成，而且模板文件生成比较麻烦

1. 替换位置的格式和poi不同，是 ${name}
2. 将.doc文件另存为.xml文件（会有下划线消失之类的问题出现，需手动修改）
3. 将.xml文件内容复制到模板文件.ftl中


PS:目前为止，两种方法生成的word文件，poi都不能读取，而aspose都能正常读取
