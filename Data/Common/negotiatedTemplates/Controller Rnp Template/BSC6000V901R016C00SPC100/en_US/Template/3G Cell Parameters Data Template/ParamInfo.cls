VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FieldName As String              '������Ϣ
Public FieldType As Integer
Public DisplayName, CHSDisplayName As String   '��ʾ����, ������ʾ����
Public RangeDef As New Collection       'ȡֵ��Χ�б�
Public EnumRangeDef As New Collection   'ö��ȡֵ��Χ�б�
Public FieldTypeName As String          '�ֶ�����������ѡINT��STRING��LIST��ö�٣�, BITMAP
Public FieldLocation  As Integer        '�ֶε���λ��
Public FieldMinValue As Double      '�����ֶε������Сֵ�����ַ����ֶ������СֵΪ���ȵķ�Χ��ö��ֵ����Ϊ��
Public FieldMaxValue As Double
Public BitmapField As String             'bitmap���ͱ�ʶ
Public FieldISMustGive As String        ' �ֶ��Ƿ���YES /  NO