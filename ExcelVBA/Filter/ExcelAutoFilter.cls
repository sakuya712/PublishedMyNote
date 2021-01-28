VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelAutoFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'''<classSummary> フィルタ情報を値渡しでコピーする。一時的にフィルタ保持するためのクラス
'''<updateDate> 2021/01/21
'''<version> Ver.100
'===============================================================================

Private FilterRange_ As Range
Private Filters_ As Collection

'''<summary> 実質的なコンストラクタ
'''<param> 対象のワークシート
'''<returns> なし
Public Function Init(ByVal Ws As Worksheet) As ExcelAutoFilter
    If Ws.AutoFilterMode Then
        Set FilterRange_ = Ws.AutoFilter.Range
        Set Filters_ = New Collection
        Dim Filter As ExcelFilter
        Dim f As Filter
        For Each f In Ws.AutoFilter.Filters
            Set Filter = New ExcelFilter
            Filters_.Add Filter.Init(f)
        Next
    End If
    Set Init = Me
End Function

'''<summary> 保持しているフィルタに再設定する
'''<param> 対象のワークシート,[フィルタを掛ける範囲](指定しない場合はコンストラクタ時と同じ範囲)
'''<returns> なし
Public Function Restore(ByVal Ws As Worksheet, Optional ByVal NowFilterRange As Range = Nothing)
    If Filters Is Nothing Then Exit Function
    If NowFilterRange Is Nothing Then Set NowFilterRange = Me.FilterRange
    Dim Col As Long
    For Col = 1 To Me.Filters.Count
        Dim f As ExcelFilter
        Set f = Me.Filters(Col)
        If f.IsFilter = False Then GoTo Continue
        If VBA.IsEmpty(f.Criteria2) Then
            If f.Operator = 0 Then
                NowFilterRange.AutoFilter Col, f.Criteria1
            Else
                NowFilterRange.AutoFilter Col, f.Criteria1, f.Operator
            End If
        Else
            NowFilterRange.AutoFilter Col, f.Criteria1, f.Operator, f.Criteria2
        End If
Continue:
    Next
End Function

'''<summary> フィルタの項目範囲を返す
'''<param> なし
'''<returns> Range型
Public Property Get FilterRange() As Range
    Dim Row As Long
    Dim LeftCol As Long
    Dim RightCol As Long
    Dim Ws As Worksheet
    Set Ws = FilterRange_.Parent
    Row = FilterRange_.Row
    LeftCol = FilterRange_.Column
    RightCol = LeftCol + FilterRange_.Columns.Count - 1
    Set FilterRange = Ws.Range(Cells(Row, LeftCol), Cells(Row, RightCol))
End Property

'''<summary> フィルタのコレクションを返す
'''<param> なし
'''<returns> フィルタのコレクション(Collection<ExcelFilter>型)
Public Property Get Filters() As Collection
    Set Filters = Filters_
End Property

