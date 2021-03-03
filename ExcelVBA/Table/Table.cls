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
'''<classSummary>テーブルを参照するためのクラス
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

Private Table_() As Variant
Private SelectCell_ As Range
Private TableWs_ As Worksheet
Private Items_ As Collection

'''<summary> テーブルを取得する(コンストラクタ)
'''<param> ITableをImplementsしたクラス、基準となるセル
'''<returns>なし
Public Function Init(ByVal TableClass As ITable, ByVal SelectCell As Range)
    Set SelectCell_ = SelectCell
    Set TableWs_ = SelectCell.Parent
    'テーブルを取得する
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

'''<summary> テーブルの行総数
'''<param> なし
'''<returns> テーブルの行総数
Public Property Get Rows() As Long
    Rows = UBound(Table_(), 1)
End Property

'''<summary> テーブルの列総数
'''<param> なし
'''<returns> テーブルの列総数
Public Property Get Columns() As Long
    Columns = UBound(Table_(), 2)
End Property

'''<summary> テーブル番号を指定してRangeオブジェクトを取得する
'''<param> テーブルの行番号, テーブルの列番号
'''<returns> Range型
Public Function GetCell(Row, Column) As Range
    Dim TableRow As Long
    Dim TableColumn As Long
    TableRow = SelectCell_.Row + Row - 1
    TableColumn = SelectCell_.Column + Column - 1
    Set GetCell = TableWs_.Cells(TableRow, TableColumn)
End Function

'''<summary> テーブルを2次配列で返す
'''<param> なし
'''<returns> 2次配列
Public Function GetTable() As Variant
    GetTable = Table_()
End Function

'''<summary> 各項目がセットされたクラスを返す
'''<param> テーブルの行番号
'''<returns>Item型(各項目の値)
'''<remarks>デフォルトメンバ
Public Function Items(ByVal Row As Long) As ITable
Attribute Items.VB_UserMemId = 0
    On Error GoTo Err01
    Set Items = Items_(Row)
    Exit Function
Err01:
    err.Raise 9001, , "存在しない行番号です"
End Function

'''<summary> デストラクタ
'''<param> なし
'''<returns>なし
Private Sub Class_Terminate()
    Set Items_ = Nothing
    Set SelectCell_ = Nothing
    Set TableWs_ = Nothing
End Sub

'''<summary> 反復処理をサポートする ※単独では使用できません
'''<param> なし
'''<returns> なし
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = Items_.[_NewEnum]
End Function
