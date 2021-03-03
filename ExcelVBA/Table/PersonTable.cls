VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonTable"
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
Private Name_ As String
Private Address_ As String
Private PhoneNumber_ As String

'''<summary> ���O(�擾)
'''<param> �Ȃ�
'''<returns> ���O
Public Property Get Name() As String
    Name = Name_
End Property

'''<summary> �Z��(�擾)
'''<param> �Ȃ�
'''<returns> �Z��
Public Property Get Address() As String
    Address = Address_
End Property

'''<summary> �d�b�ԍ�(�擾)
'''<param> �Ȃ�
'''<returns> �d�b�ԍ�
Public Property Get PhoneNumber() As String
    PhoneNumber = PhoneNumber_
End Property

'''<summary> ITable�Ŗ񑩂����l��Collection���쐬����
'''<param> �e�[�u��, ��ƂȂ�Z��
'''<returns> �l��������Collection
Private Function ITable_Create(ByVal Table As Variant, ByVal Cell As Range) As Collection
    Dim Coll As Collection
    Set Coll = New Collection
    Dim NewItems As PersonTable
    Dim Row As Long
    For Row = LBound(Table, 1) To UBound(Table, 1)
        Set NewItems = New PersonTable
        NewItems.Init Table, Row
        Coll.Add NewItems
    Next
    Set ITable_Create = Coll
End Function

'''<summary> �l��ݒ肷��
'''<param> �e�[�u��
'''<returns> �Ȃ�
Public Function Init(ByVal Table As Variant, ByVal Row As Long)
        Name_ = Table(Row, 1)
        Address_ = Table(Row, 2)
        PhoneNumber_ = Table(Row, 3)
End Function