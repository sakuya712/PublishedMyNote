VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'''<classSummary>�e�[�u�����Q�Ƃ��邽�߂̃N���X
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

Private Table_() As Variant
Private SelectCell_ As Range
Private TableWs_ As Worksheet
Private Items_ As Collection

'''<summary> �e�[�u�����擾����(�R���X�g���N�^)
'''<param> ITable��Implements�����N���X�A��ƂȂ�Z��
'''<returns>�Ȃ�
Public Function Init(ByVal TableClass As ITable, ByVal SelectCell As Range)
    Set SelectCell_ = SelectCell
    Set TableWs_ = SelectCell.Parent
    '�e�[�u�����擾����
    If TableWs_.FilterMode Then TableWs_.ShowAllData
    Dim RowBottom As Long
    Dim ColumnRight As Long
    With TableWs_
        RowBottom = .Cells(.Rows.Count, SelectCell.Column).End(xlUp).Row
        ColumnRight = SelectCell.End(xlToRight).Column
        Table_() = .Range(SelectCell, .Cells(RowBottom, ColumnRight))
    End With
    Set Items_ = TableClass.Create(Table_(), SelectCell)
End Function

'''<summary> �e�[�u���̍s����
'''<param> �Ȃ�
'''<returns> �e�[�u���̍s����
Public Property Get Rows() As Long
    Rows = UBound(Table_(), 1)
End Property

'''<summary> �e�[�u���̗񑍐�
'''<param> �Ȃ�
'''<returns> �e�[�u���̗񑍐�
Public Property Get Columns() As Long
    Columns = UBound(Table_(), 2)
End Property

'''<summary> �e�[�u���ԍ����w�肵��Range�I�u�W�F�N�g���擾����
'''<param> �e�[�u���̍s�ԍ�, �e�[�u���̗�ԍ�
'''<returns> 2���z��
Public Function GetCell(Row, Column) As Range
    Dim TableRow As Long
    Dim TableColumn As Long
    TableRow = SelectCell_.Row + Row - 1
    TableColumn = SelectCell_.Column + Column - 1
    Set GetCell = TableWs_.Cells(TableRow, TableColumn)
End Function

'''<summary> �e�[�u����2���z��ŕԂ�
'''<param> �Ȃ�
'''<returns> 2���z��
Public Function GetTable() As Variant
    GetTable = Table_
End Function

'''<summary> �e���ڂ��Z�b�g���ꂽ�N���X��Ԃ�
'''<param> �e�[�u���̍s�ԍ�
'''<returns>Item�^(�e���ڂ̒l)
'''<remarks>�f�t�H���g�����o
Public Function Items(ByVal Row As Long) As ITable
Attribute Items.VB_UserMemId = 0
    On Error GoTo Err01
    Set Items = Items_(Row)
    Exit Function
Err01:
    err.Raise 9001, , "���݂��Ȃ��s�ԍ��ł�"
End Function

'''<summary> �f�X�g���N�^
'''<param> �Ȃ�
'''<returns>�Ȃ�
Private Sub Class_Terminate()
    Set Items_ = Nothing
    Set SelectCell_ = Nothing
    Set TableWs_ = Nothing
End Sub

'''<summary> �����������T�|�[�g���� ���P�Ƃł͎g�p�ł��܂���
'''<param> �Ȃ�
'''<returns> �Ȃ�
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
        Set NewEnum = Items_.[_NewEnum]
End Function