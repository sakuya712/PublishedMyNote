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
'''<classSummary> Filter�N���X��l�n���ŕێ�����N���X
'''<updateDate> 2021/01/21
'''<version> Ver.100
'===============================================================================

Private Count_ As Long
Private Criteria1_ As Variant
Private Criteria2_ As Variant
Private On_ As Boolean
Private Operator_ As XlAutoFilterOperator

'''<summary> �����I�ȃR���X�g���N�^
'''<param> Filter�N���X
'''<returns> �Ȃ�
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

'''<summary> �����̐���Ԃ�
'''<param> �Ȃ�
'''<returns> �����̐�
'''<remarks> 1�Ȃ�Criteria1�ɏ�����2�Ȃ�ACriteria1��Criteria2�ɏ������A3�ȏ�Ȃ�Criteria1�Ɉ�v���Ă���l�̔z�񂪓���
Public Property Get Count() As Long
    Count = Count_
End Property

'''<summary> �t�B���^�[���������͈͓��Ŏw�肳�ꂽ��ɑ΂��čŏ��Ƀt�B���^�[���������l��Ԃ��܂�
'''<param> �Ȃ�
'''<returns> �t�B���^����
Public Property Get Criteria1() As Variant
    Criteria1 = Criteria1_
End Property

'''<summary> �t�B���^�[���������͈͓��Ŏw��̗�ɑ΂��� 2 �ԖڂɃt�B���^�[���������l��Ԃ��܂�
'''<param> �Ȃ�
'''<returns> �t�B���^����
Public Property Get Criteria2() As Variant
    Criteria2 = Criteria2_
End Property

'''<summary> �t�B���^�[��񂪂���Ȃ�True��Ԃ�
'''<param> �Ȃ�
'''<returns> �^�U�l
Public Property Get IsFilter() As Boolean
    IsFilter = On_
End Property

'''<summary> �w�肵���t�B���^�[�ɂ���ēK�p�����2�̏������֘A�t���鉉�Z�q��\��**XlAutoFilterOperator** �l��Ԃ��܂��B
'''<param> �Ȃ�
'''<returns> XlAutoFilterOperator
Public Property Get Operator() As XlAutoFilterOperator
    Operator = Operator_
End Property