VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'''<classSummary> Filterクラスを値渡しで保持するクラス
'''<updateDate> 2021/01/21
'''<version> Ver.100
'===============================================================================

Private Count_ As Long
Private Criteria1_ As Variant
Private Criteria2_ As Variant
Private On_ As Boolean
Private Operator_ As XlAutoFilterOperator

'''<summary> 実質的なコンストラクタ
'''<param> Filterクラス
'''<returns> なし
Public Function Init(ByVal f As Filter) As ExcelFilter
    If f.On Then
        On Error Resume Next
        Count_ = f.Count
        If VBA.TypeName(f.Criteria1) = "Interior" Then
            Criteria1_ = f.Criteria1.Color
        Else
            Criteria1_ = f.Criteria1
        End If
        Criteria2_ = f.Criteria2
        Operator_ = f.Operator
    Else
        Count_ = 0
    End If
    On_ = f.On
    Set Init = Me
End Function

'''<summary> 条件の数を返す
'''<param> なし
'''<returns> 条件の数
'''<remarks> 1ならCriteria1に条件が2なら、Criteria1とCriteria2に条件が、3以上ならCriteria1に一致している値の配列が入る
Public Property Get Count() As Long
    Count = Count_
End Property

'''<summary> フィルター処理される範囲内で指定された列に対して最初にフィルター処理される値を返します
'''<param> なし
'''<returns> フィルタ条件
Public Property Get Criteria1() As Variant
    Criteria1 = Criteria1_
End Property

'''<summary> フィルター処理される範囲内で指定の列に対して 2 番目にフィルター処理される値を返します
'''<param> なし
'''<returns> フィルタ条件
Public Property Get Criteria2() As Variant
    Criteria2 = Criteria2_
End Property

'''<summary> フィルター情報があるならTrueを返す
'''<param> なし
'''<returns> 真偽値
Public Property Get IsFilter() As Boolean
    IsFilter = On_
End Property

'''<summary> 指定したフィルターによって適用される2つの条件を関連付ける演算子を表す**XlAutoFilterOperator** 値を返します。
'''<param> なし
'''<returns> XlAutoFilterOperator
Public Property Get Operator() As XlAutoFilterOperator
    Operator = Operator_
End Property
