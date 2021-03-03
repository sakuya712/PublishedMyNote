VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'''<classSummary> �e�[�u���̃f�[�^���i�[���邽�߂̃N���X
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

Implements ITable

'�C���X�^���X�ϐ�
Private Tokyo_ As Double
Private Osaka_ As Double
Private Nagoya_ As Double

'''<summary> �����{�X(�擾)
'''<param> �Ȃ�
'''<returns> �����{�X
Public Property Get Tokyo() As Double
    Tokyo = Tokyo_
End Property

'''<summary> ���x�X(�擾)
'''<param> �Ȃ�
'''<returns> ���x�X
Public Property Get Osaka() As Double
    Osaka = Osaka_
End Property

'''<summary> ���É��x�X(�擾)
'''<param> �Ȃ�
'''<returns> ���É��x�X
Public Property Get Nagoya() As Double
    Nagoya = Nagoya_
End Property

'''<summary> ITable�Ŗ񑩂����l��Collection���쐬����
'''<param> �e�[�u��, ��ƂȂ�Z��
'''<returns> �l��������Collection
Private Function ITable_Create(ByVal Table As Variant, ByVal Cell As Range) As Collection
    Dim Coll As Collection
    Set Coll = New Collection
    Dim NewItems As SalesTable
    Dim Row As Long
    For Row = LBound(Table, 1) To UBound(Table, 1)
        Set NewItems = New SalesTable
        NewItems.Init Table, Row
        Coll.Add NewItems
    Next
    Set ITable_Create = Coll
End Function

'''<summary> �l��ݒ肷��
'''<param> �e�[�u��
'''<returns> �Ȃ�
Public Function Init(ByVal Table As Variant, ByVal Row As Long)
    Tokyo_ = Table(Row, 1)
    Osaka_ = Table(Row, 2)
    Nagoya_ = Table(Row, 3)
End Function
