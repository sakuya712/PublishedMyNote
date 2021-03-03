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
'''<classSummary> テーブルのデータを格納するためのクラス
'''<updateDate> 2021/03/01
'''<version> Ver.100
'===============================================================================

Implements ITable

'インスタンス変数
Private Name_ As String
Private Address_ As String
Private PhoneNumber_ As String

'''<summary> 名前(取得)
'''<param> なし
'''<returns> 名前
Public Property Get Name() As String
    Name = Name_
End Property

'''<summary> 住所(取得)
'''<param> なし
'''<returns> 住所
Public Property Get Address() As String
    Address = Address_
End Property

'''<summary> 電話番号(取得)
'''<param> なし
'''<returns> 電話番号
Public Property Get PhoneNumber() As String
    PhoneNumber = PhoneNumber_
End Property

'''<summary> ITableで約束した値のCollectionを作成する
'''<param> テーブル, 基準となるセル
'''<returns> 値が入ったCollection
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

'''<summary> 値を設定する
'''<param> テーブル
'''<returns> なし
Public Function Init(ByVal Table As Variant, ByVal Row As Long)
        Name_ = Table(Row, 1)
        Address_ = Table(Row, 2)
        PhoneNumber_ = Table(Row, 3)
End Function
