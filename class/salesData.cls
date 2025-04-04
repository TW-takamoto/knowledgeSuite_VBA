Option Explicit

'************************************************
'扱う営業データクラス
'************************************************

Private str営業担当 As String
Private str事業区分 As String
Private str顧客名称 As String
Private strWorkNo As String
Private str案件名 As String
Private strGRP As String
Private str区分1 As String
Private str区分2 As String
Private str受注見込 As String
Private strフェーズ As String
Private str開始年月 As String
Private str検収年月 As String
Private str売上年月 As String
Private str最終更新日 As String
Private lng次年度計上 As Long
Private str売上金額 As String
Private str売上1月 As String
Private str売上2月 As String
Private str売上3月 As String
Private str売上4月 As String
Private str売上5月 As String
Private str売上6月 As String
Private str売上7月 As String
Private str売上8月 As String
Private str売上9月 As String
Private str売上10月 As String
Private str売上11月 As String
Private str売上12月 As String

'営業担当プロパティ(取得)
Public Property Get 営業担当() As String
   営業担当 = str営業担当
End Property

'営業担当プロパティ(変更)
Public Property Let 営業担当(ByVal vNewValue As String)
   str営業担当 = vNewValue
End Property

'営業区分プロパティ(取得)
Public Property Get 事業区分() As String
   事業区分 = str事業区分
End Property

'営業区分プロパティ(変更)
Public Property Let 事業区分(ByVal vNewValue As String)
   str事業区分 = vNewValue
End Property

'顧客名称プロパティ(取得)
Public Property Get 顧客名称() As String
   顧客名称 = str顧客名称
End Property

'顧客名称プロパティ(変更)
Public Property Let 顧客名称(ByVal vNewValue As String)
   str顧客名称 = vNewValue
End Property

'WorkNoプロパティ(取得)
Public Property Get WorkNo() As String
   WorkNo = strWorkNo
End Property

'WorkNoプロパティ(変更)
Public Property Let WorkNo(ByVal vNewValue As String)
   strWorkNo = vNewValue
End Property

'案件名プロパティ(取得)
Public Property Get 案件名() As String
   案件名 = str案件名
End Property

'案件名プロパティ(変更)
Public Property Let 案件名(ByVal vNewValue As String)
   str案件名 = vNewValue
End Property

'GRPプロパティ(取得)
Public Property Get GRP() As String
   GRP = strGRP
End Property

'GRPプロパティ(変更)
Public Property Let GRP(ByVal vNewValue As String)
   strGRP = vNewValue
End Property

'区分1プロパティ(取得)
Public Property Get 区分1() As String
   区分1 = str区分1
End Property

'区分1プロパティ(変更)
Public Property Let 区分1(ByVal vNewValue As String)
   str区分1 = vNewValue
End Property

'区分2プロパティ(取得)
Public Property Get 区分2() As String
   区分2 = str区分2
End Property

'区分2プロパティ(変更)
Public Property Let 区分2(ByVal vNewValue As String)
   str区分2 = vNewValue
End Property

'受注見込プロパティ(取得)
Public Property Get 受注見込() As String
   受注見込 = str受注見込
End Property

'受注見込プロパティ(変更)
Public Property Let 受注見込(ByVal vNewValue As String)
   str受注見込 = vNewValue
End Property

'フェーズプロパティ(取得)
Public Property Get フェーズ() As String
   フェーズ = strフェーズ
End Property

'フェーズプロパティ(変更)
Public Property Let フェーズ(ByVal vNewValue As String)
   strフェーズ = vNewValue
End Property

'開始年月プロパティ(取得)
Public Property Get 開始年月() As String
   開始年月 = str開始年月
End Property

'開始年月プロパティ(変更)
Public Property Let 開始年月(ByVal vNewValue As String)
   str開始年月 = vNewValue
End Property

'検収年月プロパティ(取得)
Public Property Get 検収年月() As String
   検収年月 = str検収年月
End Property

'検収年月プロパティ(変更)
Public Property Let 検収年月(ByVal vNewValue As String)
   str検収年月 = vNewValue
End Property

'売上年月プロパティ(取得)
Public Property Get 売上年月() As String
   売上年月 = str売上年月
End Property

'売上年月プロパティ(変更)
Public Property Let 売上年月(ByVal vNewValue As String)
   str売上年月 = vNewValue
End Property

'最終更新日プロパティ(取得)
Public Property Get 最終更新日() As String
   最終更新日 = str最終更新日
End Property

'最終更新日プロパティ(変更)
Public Property Let 最終更新日(ByVal vNewValue As String)
   str最終更新日 = vNewValue
End Property

'次年度計上プロパティ(取得)
Public Property Get 次年度計上() As Long
   次年度計上 = lng次年度計上
End Property

'次年度計上プロパティ(変更)
Public Property Let 次年度計上(ByVal vNewValue As Long)
   lng次年度計上 = vNewValue
End Property

'売上金額プロパティ(取得)
Public Property Get 売上金額() As String
   売上金額 = str売上金額
End Property

'売上金額プロパティ(変更)
Public Property Let 売上金額(ByVal vNewValue As String)
   str売上金額 = vNewValue
End Property

'売上1月プロパティ(取得)
Public Property Get 売上1月() As String
   売上1月 = str売上1月
End Property

'売上1月プロパティ(変更)
Public Property Let 売上1月(ByVal vNewValue As String)
   str売上1月 = vNewValue
End Property

'売上2月プロパティ(取得)
Public Property Get 売上2月() As String
   売上2月 = str売上2月
End Property

'売上2月プロパティ(変更)
Public Property Let 売上2月(ByVal vNewValue As String)
   str売上2月 = vNewValue
End Property

'売上3月プロパティ(取得)
Public Property Get 売上3月() As String
   売上3月 = str売上3月
End Property

'売上3月プロパティ(変更)
Public Property Let 売上3月(ByVal vNewValue As String)
   str売上3月 = vNewValue
End Property

'売上4月プロパティ(取得)
Public Property Get 売上4月() As String
   売上4月 = str売上4月
End Property

'売上4月プロパティ(変更)
Public Property Let 売上4月(ByVal vNewValue As String)
   str売上4月 = vNewValue
End Property

'売上5月プロパティ(取得)
Public Property Get 売上5月() As String
   売上5月 = str売上5月
End Property

'売上5月プロパティ(変更)
Public Property Let 売上5月(ByVal vNewValue As String)
   str売上5月 = vNewValue
End Property

'売上6月プロパティ(取得)
Public Property Get 売上6月() As String
   売上6月 = str売上6月
End Property

'売上6月プロパティ(変更)
Public Property Let 売上6月(ByVal vNewValue As String)
   str売上6月 = vNewValue
End Property

'売上7月プロパティ(取得)
Public Property Get 売上7月() As String
   売上7月 = str売上7月
End Property

'売上7月プロパティ(変更)
Public Property Let 売上7月(ByVal vNewValue As String)
   str売上7月 = vNewValue
End Property

'売上8月プロパティ(取得)
Public Property Get 売上8月() As String
   売上8月 = str売上8月
End Property

'売上8月プロパティ(変更)
Public Property Let 売上8月(ByVal vNewValue As String)
   str売上8月 = vNewValue
End Property

'売上9月プロパティ(取得)
Public Property Get 売上9月() As String
   売上9月 = str売上9月
End Property

'売上9月プロパティ(変更)
Public Property Let 売上9月(ByVal vNewValue As String)
   str売上9月 = vNewValue
End Property

'売上10月プロパティ(取得)
Public Property Get 売上10月() As String
   売上10月 = str売上10月
End Property

'売上10月プロパティ(変更)
Public Property Let 売上10月(ByVal vNewValue As String)
   str売上10月 = vNewValue
End Property

'売上11月プロパティ(取得)
Public Property Get 売上11月() As String
   売上11月 = str売上11月
End Property

'売上11月プロパティ(変更)
Public Property Let 売上11月(ByVal vNewValue As String)
   str売上11月 = vNewValue
End Property

'売上12月プロパティ(取得)
Public Property Get 売上12月() As String
   売上12月 = str売上12月
End Property

'売上12月プロパティ(変更)
Public Property Let 売上12月(ByVal vNewValue As String)
   str売上12月 = vNewValue
End Property
