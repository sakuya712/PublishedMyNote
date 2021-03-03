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
'''<classSummary>テーブルの値を設定することを約束するインターフェース
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

'''<summary> 値を設定するメソッドを約束する
'''<param> テーブル, 基準となるセル
'''<returns> 値が入ったコレクション
Public Function Create(ByVal Table As Variant, ByVal Cell As Range) As Collection
End Function

'実装例 PersonTableクラス
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
