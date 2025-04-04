VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderAccuracySummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private lng合計件数 As Long
Private cur合計金額 As Currency
Private cur合計金額確率 As Currency
Private lng当期件数 As Long
Private cur当期金額 As Currency
Private cur当期金額確率 As Currency
Private lng次年度件数 As Long
Private cur次年度金額 As Currency
Private cur次年度金額確率 As Currency

'************************************************
'扱う営業データサマリクラス(受注見込み別)
'************************************************

'合計件数プロパティ(取得)
Public Property Get 合計件数() As Long
   合計件数 = lng合計件数
End Property

'合計件数プロパティ(変更)
Public Property Let 合計件数(ByVal vNewValue As Long)
   lng合計件数 = vNewValue
End Property

'合計金額プロパティ(取得)
Public Property Get 合計金額() As Currency
   合計金額 = cur合計金額
End Property

'合計金額プロパティ(変更)
Public Property Let 合計金額(ByVal vNewValue As Currency)
   cur合計金額 = vNewValue
End Property

'合計金額確率プロパティ(取得)
Public Property Get 合計金額確率() As Currency
   合計金額確率 = cur合計金額確率
End Property

'合計金額確率プロパティ(変更)
Public Property Let 合計金額確率(ByVal vNewValue As Currency)
   cur合計金額確率 = vNewValue
End Property

'当期件数プロパティ(取得)
Public Property Get 当期件数() As Long
   当期件数 = lng当期件数
End Property

'当期件数プロパティ(変更)
Public Property Let 当期件数(ByVal vNewValue As Long)
   lng当期件数 = vNewValue
End Property

'当期金額プロパティ(取得)
Public Property Get 当期金額() As Currency
   当期金額 = cur当期金額
End Property

'当期金額プロパティ(変更)
Public Property Let 当期金額(ByVal vNewValue As Currency)
   cur当期金額 = vNewValue
End Property

'当期金額確率プロパティ(取得)
Public Property Get 当期金額確率() As Currency
   当期金額確率 = cur当期金額確率
End Property

'当期金額確率プロパティ(変更)
Public Property Let 当期金額確率(ByVal vNewValue As Currency)
   cur当期金額確率 = vNewValue
End Property

'次年度件数プロパティ(取得)
Public Property Get 次年度件数() As Long
   次年度件数 = lng次年度件数
End Property

'次年度件数プロパティ(変更)
Public Property Let 次年度件数(ByVal vNewValue As Long)
   lng次年度件数 = vNewValue
End Property

'次年度金額プロパティ(取得)
Public Property Get 次年度金額() As Currency
   次年度金額 = cur次年度金額
End Property

'次年度金額プロパティ(変更)
Public Property Let 次年度金額(ByVal vNewValue As Currency)
   cur次年度金額 = vNewValue
End Property

'次年度金額確率プロパティ(取得)
Public Property Get 次年度金額確率() As Currency
   次年度金額確率 = cur次年度金額確率
End Property

'次年度金額確率プロパティ(変更)
Public Property Let 次年度金額確率(ByVal vNewValue As Currency)
   cur次年度金額確率 = vNewValue
End Property
