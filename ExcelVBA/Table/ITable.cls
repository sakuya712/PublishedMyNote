VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'''<classSummary>�e�[�u���̒l��ݒ肷�邱�Ƃ�񑩂���C���^�[�t�F�[�X
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

'''<summary> �l��ݒ肷�郁�\�b�h��񑩂���
'''<param> �e�[�u��, ��ƂȂ�Z��
'''<returns> �l���������R���N�V����
Public Function Create(ByVal Table As Variant, ByVal Cell As Range) As Collection
End Function

'������ PersonTable�N���X
'Private Function ITable_Create(ByVal Table As Variant, ByVal Cell As Range) As Collection
'    Dim Coll As Collection
'    Set Coll = New Collection
'    Dim NewItems As PersonTable
'    Dim Row As Long
'    For Row = LBound(Table, 1) To UBound(Table, 1)
'        Set NewItems = New PersonTable
'        NewItems.Init Table, Row
'        Coll.Add NewItems
'    Next
'    Set ITable_Create = Coll
'End Function