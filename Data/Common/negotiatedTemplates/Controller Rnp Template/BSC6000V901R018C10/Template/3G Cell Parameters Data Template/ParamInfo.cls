VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FieldName As String              '参数信息
Public FieldType As Integer
Public DisplayName, CHSDisplayName As String   '显示名称, 中文显示名称
Public RangeDef As New Collection       '取值范围列表
Public EnumRangeDef As New Collection   '枚举取值范围列表
Public FieldTypeName As String          '字段类型名，可选INT，STRING，LIST（枚举）, BITMAP
Public FieldLocation  As Integer        '字段的列位置
Public FieldMinValue As Double      '整型字段的最大最小值，对字符型字段最大最小值为长度的范围，枚举值类型为空
Public FieldMaxValue As Double
Public BitmapField As String             'bitmap类型标识
Public FieldISMustGive As String        ' 字段是否必填，YES /  NO
