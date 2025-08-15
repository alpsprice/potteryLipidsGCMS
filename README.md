# potteryLipidsGCMS

Update logs:
Aug 15, 2025: First upload.


description: 
生成（陶残脂肪酸）化合物列表的脚本; //a python script to generate compounds list, originally designed for GCMS data of pottery lipids
如果用的是 DB-5Ht 的柱子，岛津分析软件 + nist 数据库，分析陶残化合物，本脚本的数据库可直接用。其他情况可能需要调整部分代码~
本脚本基于岛津软件自动识别生成峰表和化合物名称，功能仅限于统计和制表，有些峰自动识别都认不出来/认错，本脚本救不了，哦。
不会编程，代码全靠与 AI 的反复协商。欢迎留言反馈，查看留言不及时请见谅。

File for analysis：
一整个文件夹的，.txt 格式的，如下所示的处理结果表，可由岛津气相软件生成，允许一个 txt 包含多个处理结果；//a folder full of .txt file of the following format; Can be exported from Shimadzu GCMS software.

[Header]
Data File Name	D:\Documents\250710\bigPot.qgd
Output Date	2025/7/16
Output Time	12:45:46

[MC Peak Table]
# of Peaks	4
Mass	TIC
Peak#	Ret.Time	Proc.From	Proc.To	Mass	Area	Height	A/H	Conc.	Mark	Name	Ret. Index	Area%	Height%	SI	CAS #
1	16.366	16.314	16.594	TIC	2142910	299824	7.15	56.18	   	Hexadecanoic acid, methyl ester	1930	56.18	55.68	97	112-39-0
2	16.628	16.594	16.721	TIC	129504	32851	3.94	3.40	 V 	Hexadecanoic acid, methyl ester	1958	3.40	6.10	90	112-39-0
3	18.037	18.008	18.067	TIC	9958	6009	1.66	0.26	   	8-Octadecenoic acid, methyl ester	2110	0.26	1.12	91	2345-29-1
4	18.263	18.202	18.600	TIC	1531948	199749	7.67	40.16	   	Methyl stearate	2138	40.16	37.10	97	112-61-8
