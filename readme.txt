/*
--------------------------------------------------------------------------------
ע�⣺excel��չ��������.xlsx��ʽ�������������excel���Ϊ.xlsx��ʽ���ֶ��޸ĺ�׺��Ч��
	  excel�ļ����¿�������ʱ����ȴ�log����ʾ��ɺ��ٲ�����һ�����ܰ�ť��
	  ʹ�����ǰ����ر���Ҫ������excel�ļ���
--------------------------------------------------------------------------------
1.����log����
1.1 �½�.xlsx�ļ��������Զ��壨eg:op01.xlsx),�½�5��sheet,sheet���ֱ�Ϊ��AllLog, sort, SortSelectTrans, toSheetAll, limit��
1.2 ������log�е����ݿ�����ALLlog�У���ALLlog���ݿ�������sort��sheet�У�ɾ���쳣���ݣ���SN������
1.3 ����sort�е��У�UUTSerialNumber�к�����SN�ŵ������ݡ�ճ����SortSelectTrans�У�ճ��ѡ��ת�ã�ɾ���ղ����
1.4 ����sort�е��У�UUTSerialNumber�� + limit high + limit Low + ��λ + �Ƚ� ��5�С�ճ����limit�У���limit high �� limit low�����У���limit high��limit low�Ϸ���

2.MSAģ������
2.1 ����op01.xlsx��SortSelectTrans���test item����MSAģ���и�ÿ��test item����������sheet��sheet��һ��ȡtest item ǰ8λ��sheet������ñ�������+����һ����ĸ����
	����ȷ�����е�test item���е�����sheet���ղ�������⣩������sheet name ˳�򣺴���->�� ��Ӧ test item С->��
2.2 MSAģ���е�Summary�е�������Ҫ�ֶ�����

3.�������
3.1 ·��(.xlsx·��)
	SourcePath������log��excel·�� + �ļ��������ܰ�ť��SelectPathSѡ���ļ�
	TargetPath��MSAģ���excel·�� + �ļ��������ܰ�ť��SelectPathTѡ���ļ�

3.2 �����������������
	Extract data: SourcePath����op01.xlsx�й�����SortSelectTrans�е�������ȡ�� toSheetAll�У����ܰ�ť��ExtractData
	Paste to GRR: SourcePath to TargetPath����op01.xlsx�й��ʱ�toSheetAll�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteToGRR
	Paste to GRR: Limit����op01.xlsx�й��ʱ�limit�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteLimit

=======================================================================================================
	��������չ���ܣ���������

3.3 ��������
	��ť��ExtracSheetToTxt����Target·����ģ���е�����sheet����Summary���⣩����ȡ����������txt�ļ�
	��ť��CopyPaste������ճ��PastePara�е�ֵ
	��ť��DeleteRange��ɾ��DeletePara�е�ֵ
	��ť��DeleteSheet��ɾ������������ReserveCnt����ɾ��SheetName����
	��ť��CreatSheet���½�������
	��ť��RemaneSheet��������������
	��ť��ReadIni����ȡINI�����ļ���UI�ؼ���
	��ť��WriteIni����UI�ؼ�����д��INI�����ļ�
	SheetName��ָ��������:CopyPaste��DeleteRange��DeleteSheet��CreatSheet��RemaneSheet�Ĺ�����
	ReserveCnt�����ֵ�ǡ�:cnt����ʽ(cntȡ������)�����ʾֻ����cnt���������������ɾ����������ǡ�:cnt����ʽ����ɾ��SheetName�еĹ�����
	General: CopyPaste And Delete
		ExcelPath:excel·����������General
		PastePara������ʽ��StartRow,StartCol,EndRow,EndCol,StartRowDest,StartColDest// ��ʼ�У���ʼ�У������У������У�Ŀ���У�Ŀ���У����ж�����������
		DeletePara������ʽ��StartCol,StartCol,EndRow,EndCol							// ��ʼ�У���ʼ�У������У������У����ж�����������
		ReserveCnt������ʽ��:cnt													// ����cnt��sheet��ע�����һ��sheet����ͳ�Ʒ�Χ������Summary��������
=======================================================================================================

4. �������˵����
4.1 ѡ��·����
4.2 ���ò�����
4.3 ��ȡ���ݣ�--���ܰ�ť��ExtractData
4.4 ������ȡ���ݵ�MSAģ�壻--���ܰ�ť��PasteToGRR
4.5 ����limit �� MSAģ�壻--���ܰ�ť��PasteLimit
5.6 �ֶ�����MSAģ���е�summary���ݣ�

5.������
�����ִ���ļ�·����eg test data�ļ����е��ļ���


*/