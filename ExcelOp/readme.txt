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
	Paste to GRR module: SourcePath to TargetPath����op01.xlsx toSheetAll�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteToGRR
	Paste to GRR: Limit����op01.xlsx limit�е����ݿ���ճ����MSAģ���У����ܰ�ť��PasteLimit

3.3 ��������
	��ť��ExtracSheetToTxt����Target·����ģ���е�����sheet����Summary���⣩����ȡ����������txt�ļ�
	��ť��CopyPaste������ճ��PastePara�е�ֵ
	��ť��DeleteRange��ɾ��DeletePara�е�ֵ
	��ť��DeleteSheet��ɾ������������ReserveCnt����
	SheetName����չCopyPaste��DeleteRange�Ĺ�����������
	//��ť��CopyGRRModuleAndDelete���ɺ��ԣ����ڿ���MSAģ����16�еĹ�ʽ��17�� �� ɾ��11.xx������ĵ�17�е����ݺ͹�ʽ��

4. �������˵����
4.1 ѡ��·����
4.2 ���ò�����
4.3 ��ȡ���ݣ�
4.4 ������ȡ���ݵ�MSAģ�壻
4.5 ����limit �� MSAģ�壻
5.6 �ֶ�����MSAģ���е�summary���ݣ�












*/