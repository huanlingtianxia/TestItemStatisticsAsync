/*
ע�⣺excel��չ������ʱ.xlsx��ʽ��

1.����log����
1.1 �½�.xlsx�ļ��������Զ��壨eg:op01.xlsx),�½�5��sheet,sheet���ֱ�Ϊ��AllLog, sort, SortSelectTrans, toSheetAll, limit��
1.2 ������log�е����ݿ�����ALLlog�У���ALLlog���ݿ�������sort��sheet�У�ɾ���쳣���ݣ���SN������
1.3 ����sort�е��У�UUTSerialNumber�к�����SN�ŵ������ݡ�ճ����SortSelectTrans�У�ճ��ѡ��ת�ã�ɾ���ղ����
1.4 ����sort�е��У�UUTSerialNumber�� + limit high + limit Low + ��λ + �Ƚ� ��5�С�ճ����limit�У���limit high �� limit low�����У���limit high���Ϸ���

2.MSAģ������
2.1 ����op01.xlsx��SortSelectTrans���test item����MSAģ���и�ÿ��test item����������sheet��sheet��ȡtest item ǰ8λ������ȷ�����е�item���е�����sheet���ղ�������⣩��
2.2 MSAģ���е�Summary�е�������Ҫ�ֶ�����

3.�������
3.1 ·��(.xlsx·��)
	SourcePath������log��excel·�� + �ļ��������ܰ�ť��SelectPathSѡ���ļ�
	TargetPath��MSAģ���excel·�� + �ļ��������ܰ�ť��SelectPathTѡ���ļ�

3.2 �����������������
	Extract data: SourcePath����op01.xlsx SortSelectTrans�е�������ȡ�� toSheetAll�У����ܰ�ť��ExtractData
	Paste to GRR: SourcePath to TargetPath����op01.xlsx�й��ʱ�toSheetAll�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteToGRR
	Paste to GRR: Limit����op01.xlsx�й��ʱ�limit�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteLimit

3.3 ��������
	��ť��ExtracSheetToTxt����Target·����ģ���е�����sheet����Summary���⣩����ȡ����������txt�ļ�
	��ť��CopyPaste������ճ��PastePara�е�ֵ
	��ť��DeleteRange��ɾ��DeletePara�е�ֵ
	��ť��DeleteSheet��ɾ������������ReserveCnt����ɾ��SheetName����
	��ť��CreatSheet���½�������
	SheetName����չCopyPaste��DeleteRange��DeleteSheet�Ĺ�����������
	ReserveCnt�����ֵ�ǡ�:cnt����ʽ(cntȡ������)�����ʾֻ����cnt���������������ɾ����������ǡ�:cnt����ʽ����ɾ��SheetName�еĹ�����
	General: CopyPaste And Delete
		PastePara������ʽ��StartCol,StartCol,EndRow,EndCol,StartRowDest,StartColDest// ��ʼ�У���ʼ�У������У������У�Ŀ���У�Ŀ���У����ж�����������
		DeletePara������ʽ��StartCol,StartCol,EndRow,EndCol							// ��ʼ�У���ʼ�У������У������У����ж�����������
		ReserveCnt������ʽ��:cnt													// ����cnt��sheet��ע�����һ��sheet����ͳ�Ʒ�Χ������Summary��������

4. �������˵����
4.1 ѡ��·����
4.2 ���ò�����
4.3 ��ȡ���ݣ�
4.4 ������ȡ���ݵ�MSAģ�壻
4.5 ����limit �� MSAģ�壻
5.6 �ֶ�����MSAģ���е�summary���ݣ�












*/