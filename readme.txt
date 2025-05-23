/*
--------------------------------------------------------------------------------
注意：excel扩展名必须是.xlsx格式，如果不是请用excel另存为.xlsx格式，手动修改后缀无效。
	  excel文件更新可能有延时，请等待log框提示完成后，再操作下一个功能按钮。
	  使用软件前，请关闭需要操作的excel文件。
--------------------------------------------------------------------------------
基础功能：

1.测试log整理：
1.1 新建.xlsx文件，名称自定义（eg:op01.xlsx),新建5个sheet,sheet名分别为：AllLog, sort, SortSelectTrans, toSheetAll, limit；
1.2 将测试log中的数据拷贝到ALLlog中，将ALLlog数据拷贝到“sort”sheet中，删除sort中的异常数据，按SN号排序；
1.3 复制sort中的行：UUTSerialNumber行和所有SN号的行数据。粘贴到SortSelectTrans中，粘贴选择转置，删除空测试项；
1.4 复制sort中的行：UUTSerialNumber行 + limit high + limit Low + 单位 + 比较 这5行。粘贴到limit中，将limit high 和 limit low调换行，即limit high在limit low上方；

2.MSA模板整理：
2.1 根据op01.xlsx中SortSelectTrans里的test item，在MSA模板中给每个test item创建单独的sheet，sheet名一般取test item 前8位（sheet名中最好保留数字+最少一个字母）。
	必须确保SortSelectTrans所有的test item都有单独的sheet（空测试项除外）；测试sheet name 顺序：从右->左 对应 test item 小->大。
2.2 MSA模板中的Summary中的数据需要手动整理。

3.软件介绍
3.1 路径(.xlsx路径)
	SourcePath：整理log的excel路径 + 文件名；功能按钮：SelectPathS选择文件
	TargetPath：MSA模板的excel路径 + 文件名；功能按钮：SelectPathT选择文件

3.2 参数：根据名称理解
	Extract data: SourcePath：将op01.xlsx中工作表SortSelectTrans中的数据提取到 toSheetAll中；功能按钮：ExtractData
	Paste to GRR: SourcePath to TargetPath：将op01.xlsx中工作表toSheetAll中的数据拷贝粘贴到MSA模板中；功能按钮：PasteToGRR
	Paste to GRR: Limit：将op01.xlsx中工作表limit中的数据拷贝粘贴到MSA模板中；功能按钮：PasteLimit
	Summary:将GRR模板中 单表里的测试数据用公式关联到Summary指定位置；格式：开始行\n开始列1,开始列2,开始列3，开始列4，开始列5\nSummary sheet名
			开始行，列，Summary名都用'\n'隔开。列之间用','隔开.开始列数据依次是："LowLimit", "HighLimit", "CP", "CPK", "GRR Value" 对应的列标位置

=======================================================================================================
扩展功能，可跳过：

3.3 其他功能
	按钮：ExtracSheetToTxt：将Target路径下模板中的所有sheet名（Summary除外），提取出来并生成txt文件
	按钮：CopyPaste：拷贝粘贴PasteParam中的值,注意：值和公式不能一起复制粘贴，否则手动打开excel异常；请将值和公式分两次分别复制粘贴,每组数据用'\n'分割。
	按钮：DeleteRange：删除DeleteParam中的值
	按钮：DeleteSheet：删除工作表，保留ReserveCnt个或删除SheetName个。
	按钮：CreatSheet：新建工作表。在PosSheet的左侧开始创建
	按钮：RemaneSheet：重命名工作表
	按钮：ReadIni：读取INI配置文件到UI控件。
	按钮：WriteIni：将UI控件数据写入INI配置文件
	按钮：<-----：在General: SheetName中。单击可切换General功能是否使能。
	SheetName：指定作用域:CopyPaste，DeleteRange，DeleteSheet，CreatSheet，RemaneSheet的工作表。
	ReserveCnt：如果值是”:cnt“格式(cnt取正整数)，则表示只保留cnt个工作表格，其他的删除。如果不是”:cnt“格式，则删除SheetName中的工作表。
	General: CopyPaste And Delete
		ExcelPath:excel路径，作用域：General
		PasteParam参数格式：StartRow,StartCol,EndRow,EndCol,StartRowDest,StartColDest\nStartRow2......  // 开始行，开始列，结束行，结束列，目标行，目标列（行列都是正整数,从1开始）,每组数据用'\n'分割。
		DeleteParam参数格式：StartCol,StartCol,EndRow,EndColnStartRow1\nStartRow2......					// 开始行，开始列，结束行，结束列（行列都是正整数）,每组数据用'\n'分割。
		ReserveCnt参数格式：:cnt																		// 保留cnt个sheet。注意最后一个sheet不在统计范围（保留Summary工作表）。
		PosSheet参数：																					//新建sheet时，在PosSheet的左侧开始创建
=======================================================================================================

4. 软件操作说明：
4.1 选择路径；
4.2 设置参数；//基础功能基本上只要按实际测试项个数修改TotalItem，路径改一下，其他的按默认值即可。
4.3 提取数据；--功能按钮：ExtractData
4.4 拷贝提取的测试数据到MSA模板；--功能按钮：PasteToGRR
4.5 拷贝limit 到 MSA模板；--功能按钮：PasteLimit
4.6 拷贝关联单表的公式到 MSA 模板的Summary工作表；-- 功能按钮：SummaryFL，保持默认即可（除非Summary名不一样）
5.6 手动整理MSA模板中的summary里的序号，测试项名，单位等固定数据

5.范例：
详见可执行文件路径下eg test data文件夹中的文件。


*/