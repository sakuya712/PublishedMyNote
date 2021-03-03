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
'''<classSummary> テーブルのデータを格納するためのクラス
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

Implements ITable

'インスタンス変数
Private Tokyo_ As Double
Private Osaka_ As Double
Private Nagoya_ As Double

'''<summary> 東京本店(取得)
'''<param> なし
'''<returns> 東京本店
Public Property Get Tokyo() As Double
    Tokyo = Tokyo_
End Property

'''<summary> 大阪支店(取得)
'''<param> なし
'''<returns> 大阪支店
Public Property Get Osaka() As Double
    Osaka = Osaka_
End Property

'''<summary> 名古屋支店(取得)
'''<param> なし
'''<returns> 名古屋支店
Public Property Get Nagoya() As Double
    Nagoya = Nagoya_
End Property

'''<summary> ITableで約束した値のCollectionを作成する
'''<param> テーブル, 基準となるセル
'''<returns> 値が入ったCollection
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

'''<summary> 値を設定する
'''<param> テーブル
'''<returns> なし
Public Function Init(ByVal Table As Variant, ByVal Row As Long)
    Tokyo_ = Table(Row, 1)
    Osaka_ = Table(Row, 2)
    Nagoya_ = Table(Row, 3)
End Function

