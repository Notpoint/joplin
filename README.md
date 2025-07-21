audiotoword.py 使用说明
脚本作用

批量遍历指定目录内的所有 .m4a 文件

若输出目录已存在同名 .txt，直接跳过；否则调用 Whisper-large-v3 语音模型转写并写入同名 .txt

md2word.py 使用说明
转换规则
普通段落 → Word 段落

图片 ``

从 RESOURCE_FOLDER 里按文件名查找并插入，宽度默认 4 英寸

音频链接 [xxx.m4a](../_resources/yyy.m4a)

原 Markdown 行保留

在其下一段追加同名 yyy.txt 的文本，文件从 TRANSCRIPT_FOLDER 查找

如未找到则写入 [未找到转录文件: yyy.txt]

其它行按原样写入