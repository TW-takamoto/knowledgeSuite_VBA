Option Explicit

'************************************************
'エクスポートされるデータ
'************************************************
Public Enum NegotiationData
    操作
    レコードID
    商談ID
    営業担当部署識別方法
    営業担当部署ID
    営業担当部署コード
    営業担当部署名
    顧客番号
    顧客名称
    WorkNo
    案件名
    事業区分
    grp
    区分1
    区分2
    営業担当者メールアドレス
    営業担当者名称
    受注見込
    フェーズ
    開始年月
    検収年月
    売上年月
    作成日
    作成者
    最終更新日
    最終更新者
    売上金額
    次年度計上
    売上1月
    売上2月
    売上3月
    売上4月
    売上5月
    売上6月
    売上7月
    売上8月
    売上9月
    売上10月
    売上11月
    売上12月
End Enum

'************************************************
'テーブルのカラム番号
'************************************************
Public Enum KnowledgeSuiteTableColumn
    区分1
    顧客名称
    案件名
    grp
    売上1月
    売上2月
    売上3月
    売上1Q
    売上4月
    売上5月
    売上6月
    売上2Q
    売上上期
    売上7月
    売上8月
    売上9月
    売上3Q
    売上10月
    売上11月
    売上12月
    売上4Q
    売下下期
    売上金額
End Enum

Public initialRowCountStockBlue As Long
Public initialRowCountSpotBlue As Long
Public lastRow As Long
Public KnowledgeSuiteTableStock_blue As ListObject
Public KnowledgeSuiteTableStock_green As ListObject
Public KnowledgeSuiteTableSpot_blue As ListObject
Public KnowledgeSuiteTableSpot_green As ListObject
Public ws As Worksheet

    
Sub SalesDataProcessing()
    On Error GoTo ErrorHandler
    
    'CSVエクスポートデータを保持するList
    Dim salesDatas() As salesData
    Dim initialSalesDataRow As Integer
    Dim filePath As Variant
    Dim val As Variant
    Dim outRow As Integer: outRow = 0
    Dim outArray() As salesData
    initialSalesDataRow = 21
    Dim tbl As ListObject
    
    Set ws = ActiveSheet

    ' テーブルの存在確認
    On Error Resume Next
    For Each tbl In ws.ListObjects
        If InStr(tbl.Name, "KnowledgeSuiteTableStock_blue") > 0 Then
            Set KnowledgeSuiteTableStock_blue = tbl
        ElseIf InStr(tbl.Name, "KnowledgeSuiteTableSpot_blue") > 0 Then
            Set KnowledgeSuiteTableSpot_blue = tbl
        ElseIf InStr(tbl.Name, "KnowledgeSuiteTableStock_green") > 0 Then
            Set KnowledgeSuiteTableStock_green = tbl
        ElseIf InStr(tbl.Name, "KnowledgeSuiteTableSpot_green") > 0 Then
            Set KnowledgeSuiteTableSpot_green = tbl
        End If
    Next tbl
    KnowledgeSuiteTableStock_blue.TableStyle = ""
    KnowledgeSuiteTableSpot_blue.TableStyle = ""
    KnowledgeSuiteTableStock_green.TableStyle = ""
    KnowledgeSuiteTableSpot_green.TableStyle = ""
    If Err.Number <> 0 Then
        MsgBox "必要なテーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    initialRowCountStockBlue = KnowledgeSuiteTableStock_blue.ListRows.Count
    initialRowCountSpotBlue = KnowledgeSuiteTableSpot_blue.ListRows.Count
    lastRow = KnowledgeSuiteTableSpot_blue.ListRows(KnowledgeSuiteTableSpot_blue.ListRows.Count).Range.row

    '指定したファイルを取得
    filePath = Application.GetOpenFilename()
    
    If filePath = False Then
        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbInformation
        Exit Sub
    End If
    
    ' CSVファイルを読み込む
    salesDatas = getSalesData(CStr(filePath))
    
    For Each val In salesDatas
        outRow = outRow + 1
        ReDim Preserve outArray(outRow)
        Set outArray(outRow - 1) = New salesData
        outArray(outRow - 1).区分1 = val.区分1
        outArray(outRow - 1).顧客名称 = val.顧客名称
        outArray(outRow - 1).案件名 = val.案件名
        outArray(outRow - 1).grp = val.grp
        outArray(outRow - 1).次年度計上 = val.次年度計上
        outArray(outRow - 1).売上1月 = val.売上1月
        outArray(outRow - 1).売上2月 = val.売上2月
        outArray(outRow - 1).売上3月 = val.売上3月
        outArray(outRow - 1).売上4月 = val.売上4月
        outArray(outRow - 1).売上5月 = val.売上5月
        outArray(outRow - 1).売上6月 = val.売上6月
        outArray(outRow - 1).売上7月 = val.売上7月
        outArray(outRow - 1).売上8月 = val.売上8月
        outArray(outRow - 1).売上9月 = val.売上9月
        outArray(outRow - 1).売上10月 = val.売上10月
        outArray(outRow - 1).売上11月 = val.売上11月
        outArray(outRow - 1).売上12月 = val.売上12月
        outArray(outRow - 1).売上金額 = val.売上金額
        outArray(outRow - 1).フェーズ = val.フェーズ
        outArray(outRow - 1).受注見込 = val.受注見込
    Next
    
    If Not (KnowledgeSuiteTableStock_blue.DataBodyRange Is Nothing) Then
        KnowledgeSuiteTableStock_blue.DataBodyRange.Delete
    End If
    If Not (KnowledgeSuiteTableSpot_blue.DataBodyRange Is Nothing) Then
        KnowledgeSuiteTableSpot_blue.DataBodyRange.Delete
    End If
    If Not (KnowledgeSuiteTableStock_green.DataBodyRange Is Nothing) Then
        KnowledgeSuiteTableStock_green.DataBodyRange.Delete
    End If
    If Not (KnowledgeSuiteTableSpot_green.DataBodyRange Is Nothing) Then
        KnowledgeSuiteTableSpot_green.DataBodyRange.Delete
    End If
    
    '以下でテーブルにデータを追加していく
    Dim i As Integer
    Dim N As Long
    Dim TF As String

    ' 確定フェーズ
    Dim validPhasesConfirmed As Variant
    validPhasesConfirmed = Array("契約締結", "検収書受領", "受注", "請求完了", "納品完了")
    Dim isValidPhase As Boolean

    ' 除外フェーズ
    Dim excludedPhases As Variant
    excludedPhases = Array("敗北（案件消滅）", "失注")

    ' 受注見込
    Dim validProspects As Variant
    validProspects = Array("A", "B")
    Dim isValidProspect As Boolean

    Dim phase As Variant
    Dim prospect As Variant
    For i = 0 To outRow - 1
        Dim lng売上1月 As Long: If outArray(i).売上1月 = "" Then lng売上1月 = 0 Else lng売上1月 = CLng(outArray(i).売上1月)
        Dim lng売上2月 As Long: If outArray(i).売上2月 = "" Then lng売上2月 = 0 Else lng売上2月 = CLng(outArray(i).売上2月)
        Dim lng売上3月 As Long: If outArray(i).売上3月 = "" Then lng売上3月 = 0 Else lng売上3月 = CLng(outArray(i).売上3月)
        Dim lng売上4月 As Long: If outArray(i).売上4月 = "" Then lng売上4月 = 0 Else lng売上4月 = CLng(outArray(i).売上4月)
        Dim lng売上5月 As Long: If outArray(i).売上5月 = "" Then lng売上5月 = 0 Else lng売上5月 = CLng(outArray(i).売上5月)
        Dim lng売上6月 As Long: If outArray(i).売上6月 = "" Then lng売上6月 = 0 Else lng売上6月 = CLng(outArray(i).売上6月)
        Dim lng売上7月 As Long: If outArray(i).売上7月 = "" Then lng売上7月 = 0 Else lng売上7月 = CLng(outArray(i).売上7月)
        Dim lng売上8月 As Long: If outArray(i).売上8月 = "" Then lng売上8月 = 0 Else lng売上8月 = CLng(outArray(i).売上8月)
        Dim lng売上9月 As Long: If outArray(i).売上9月 = "" Then lng売上9月 = 0 Else lng売上9月 = CLng(outArray(i).売上9月)
        Dim lng売上10月 As Long: If outArray(i).売上10月 = "" Then lng売上10月 = 0 Else lng売上10月 = CLng(outArray(i).売上10月)
        Dim lng売上11月 As Long: If outArray(i).売上11月 = "" Then lng売上11月 = 0 Else lng売上11月 = CLng(outArray(i).売上11月)
        Dim lng売上12月 As Long: If outArray(i).売上12月 = "" Then lng売上12月 = 0 Else lng売上12月 = CLng(outArray(i).売上12月)
        Dim lng売上金額 As Long: If outArray(i).売上金額 = "" Then lng売上金額 = 0 Else lng売上金額 = CLng(outArray(i).売上金額)
        Dim Current1QSummary As Long: Current1QSummary = lng売上1月 + lng売上2月 + lng売上3月
        Dim Current2QSummary As Long: Current2QSummary = lng売上4月 + lng売上5月 + lng売上6月
        Dim Current3QSummary As Long: Current3QSummary = lng売上7月 + lng売上8月 + lng売上9月
        Dim Current4QSummary As Long: Current4QSummary = lng売上10月 + lng売上11月 + lng売上12月
        Dim CurrentFirstHalfSummary As Long: CurrentFirstHalfSummary = Current1QSummary + Current2QSummary
        Dim CurrentSecondHalfSummary As Long: CurrentSecondHalfSummary = Current3QSummary + Current4QSummary

        ' 除外フェーズの確認
        For Each phase In excludedPhases
            If phase = outArray(i).フェーズ Then
                GoTo NextRow
            End If
        Next phase

        ' 次年度計上の確認
        If outArray(i).次年度計上 = "1" Then
            GoTo NextRow
        End If

        ' フェーズの確認
        isValidPhase = False
        For Each phase In validPhasesConfirmed
            If phase = outArray(i).フェーズ Then
                isValidPhase = True
                Exit For
            End If
        Next phase

        ' 受注見込の確認
        isValidProspect = False
        For Each prospect In validProspects
            If prospect = outArray(i).受注見込 Then
                isValidProspect = True
                Exit For
            End If
        Next prospect

        ' 行データを配列として準備
        Dim rowData(1 To 1, 1 To 23) As Variant
        rowData(1, KnowledgeSuiteTableColumn.区分1 + 1) = outArray(i).区分1
        rowData(1, KnowledgeSuiteTableColumn.顧客名称 + 1) = outArray(i).顧客名称
        rowData(1, KnowledgeSuiteTableColumn.案件名 + 1) = outArray(i).案件名
        rowData(1, KnowledgeSuiteTableColumn.grp + 1) = outArray(i).grp
        rowData(1, KnowledgeSuiteTableColumn.売上1月 + 1) = IIf(lng売上1月 = 0, " ", Int(lng売上1月))
        rowData(1, KnowledgeSuiteTableColumn.売上2月 + 1) = IIf(lng売上2月 = 0, " ", Int(lng売上2月))
        rowData(1, KnowledgeSuiteTableColumn.売上3月 + 1) = IIf(lng売上3月 = 0, " ", Int(lng売上3月))
        rowData(1, KnowledgeSuiteTableColumn.売上1Q + 1) = Int(Current1QSummary)
        rowData(1, KnowledgeSuiteTableColumn.売上4月 + 1) = IIf(lng売上4月 = 0, " ", Int(lng売上4月))
        rowData(1, KnowledgeSuiteTableColumn.売上5月 + 1) = IIf(lng売上5月 = 0, " ", Int(lng売上5月))
        rowData(1, KnowledgeSuiteTableColumn.売上6月 + 1) = IIf(lng売上6月 = 0, " ", Int(lng売上6月))
        rowData(1, KnowledgeSuiteTableColumn.売上2Q + 1) = Int(Current2QSummary)
        rowData(1, KnowledgeSuiteTableColumn.売上上期 + 1) = Int(CurrentFirstHalfSummary)
        rowData(1, KnowledgeSuiteTableColumn.売上7月 + 1) = IIf(lng売上7月 = 0, " ", Int(lng売上7月))
        rowData(1, KnowledgeSuiteTableColumn.売上8月 + 1) = IIf(lng売上8月 = 0, " ", Int(lng売上8月))
        rowData(1, KnowledgeSuiteTableColumn.売上9月 + 1) = IIf(lng売上9月 = 0, " ", Int(lng売上9月))
        rowData(1, KnowledgeSuiteTableColumn.売上3Q + 1) = Int(Current3QSummary)
        rowData(1, KnowledgeSuiteTableColumn.売上10月 + 1) = IIf(lng売上10月 = 0, " ", Int(lng売上10月))
        rowData(1, KnowledgeSuiteTableColumn.売上11月 + 1) = IIf(lng売上11月 = 0, " ", Int(lng売上11月))
        rowData(1, KnowledgeSuiteTableColumn.売上12月 + 1) = IIf(lng売上12月 = 0, " ", Int(lng売上12月))
        rowData(1, KnowledgeSuiteTableColumn.売上4Q + 1) = Int(Current4QSummary)
        rowData(1, KnowledgeSuiteTableColumn.売下下期 + 1) = Int(CurrentSecondHalfSummary)
        rowData(1, KnowledgeSuiteTableColumn.売上金額 + 1) = Int(lng売上金額)

        If outArray(i).区分1 = "フロー" And isValidPhase = True Then
            ' スポット青テーブルに行を追加
            Dim newRowSpotBlue As ListRow
            rowData(1, KnowledgeSuiteTableColumn.区分1 + 1) = "スポット"
            Set newRowSpotBlue = KnowledgeSuiteTableSpot_blue.ListRows.Add
            newRowSpotBlue.Range.Value = rowData

        ElseIf outArray(i).区分1 <> "フロー" And isValidPhase = True Then
            ' ストック青テーブルに行を追加
            Dim newRowStockBlue As ListRow
            Set newRowStockBlue = KnowledgeSuiteTableStock_blue.ListRows.Add
            newRowStockBlue.Range.Value = rowData
            
        ElseIf outArray(i).区分1 = "フロー" And isValidPhase = False And isValidProspect = True Then
            ' スポット緑テーブルに行を追加
            Dim newRowSpotGreen As ListRow
            rowData(1, KnowledgeSuiteTableColumn.区分1 + 1) = "スポット"
            Set newRowSpotGreen = KnowledgeSuiteTableSpot_green.ListRows.Add
            newRowSpotGreen.Range.Value = rowData
            
        ElseIf outArray(i).区分1 <> "スポット" And isValidPhase = False And isValidProspect = True Then
            ' ストック緑テーブルに行を追加
            Dim newRowStockGreen As ListRow
            Set newRowStockGreen = KnowledgeSuiteTableStock_green.ListRows.Add
            newRowStockGreen.Range.Value = rowData
        End If

NextRow:
    Next i

    ' データをGRPごとにグループ化する処理を追加
    GroupDataByGRP
    
    ' テーブルの書式設定
    SetupTableFormat KnowledgeSuiteTableStock_blue
    SetupTableFormat KnowledgeSuiteTableSpot_blue
    SetupTableFormat KnowledgeSuiteTableStock_green
    SetupTableFormat KnowledgeSuiteTableSpot_green
    
    MsgBox "完了しました。"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & " (エラー番号: " & Err.Number & ")", vbCritical
End Sub

' GRPごとにデータをグループ化する関数
Sub GroupDataByGRP()

    ' 各テーブルに対してGRPでソート
    If Not (KnowledgeSuiteTableStock_blue.DataBodyRange Is Nothing) Then
        With KnowledgeSuiteTableStock_blue.Sort
            .SortFields.Clear
            ' カスタムソートの順序を指定
            .SortFields.Add Key:=KnowledgeSuiteTableStock_blue.ListColumns(KnowledgeSuiteTableColumn.grp + 1).Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="次世代金融,国内マーケット,フロントソリューション,バックオフィスソリューション,デジタルコマース,システム運用,セキュリティサービス,グローバルマーケット,WT" ' カスタム順序を指定
            .Header = xlYes
            .Apply
        End With
    End If
    
    If Not (KnowledgeSuiteTableSpot_blue.DataBodyRange Is Nothing) Then
        With KnowledgeSuiteTableSpot_blue.Sort
            .SortFields.Clear
            .SortFields.Add Key:=KnowledgeSuiteTableSpot_blue.ListColumns(KnowledgeSuiteTableColumn.grp + 1).Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="次世代金融,国内マーケット,フロントソリューション,バックオフィスソリューション,デジタルコマース,システム運用,セキュリティサービス,グローバルマーケット,WT" ' カスタム順序を指定
            .Header = xlYes
            .Apply
        End With
    End If
    
    If Not (KnowledgeSuiteTableStock_green.DataBodyRange Is Nothing) Then
        With KnowledgeSuiteTableStock_green.Sort
            .SortFields.Clear
            .SortFields.Add Key:=KnowledgeSuiteTableStock_green.ListColumns(KnowledgeSuiteTableColumn.grp + 1).Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="次世代金融,国内マーケット,フロントソリューション,バックオフィスソリューション,デジタルコマース,システム運用,セキュリティサービス,グローバルマーケット,WT" ' カスタム順序を指定
            .Header = xlYes
            .Apply
        End With
    End If
    
    If Not (KnowledgeSuiteTableSpot_green.DataBodyRange Is Nothing) Then
        With KnowledgeSuiteTableSpot_green.Sort
            .SortFields.Clear
            .SortFields.Add Key:=KnowledgeSuiteTableSpot_green.ListColumns(KnowledgeSuiteTableColumn.grp + 1).Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="次世代金融,国内マーケット,フロントソリューション,バックオフィスソリューション,デジタルコマース,システム運用,セキュリティサービス,グローバルマーケット,WT" ' カスタム順序を指定
            .Header = xlYes
            .Apply
        End With
    End If
    
    Dim rowDifference As Long: rowDifference = 0
    ' GRPごとに小計を追加
    AddSubtotalsForGRP KnowledgeSuiteTableStock_blue, rowDifference
    AddSubtotalsForGRP KnowledgeSuiteTableSpot_blue, rowDifference
    AddSubtotalsForGRP KnowledgeSuiteTableStock_green, rowDifference
    AddSubtotalsForGRP KnowledgeSuiteTableSpot_green, rowDifference

    Dim nowRow As Long
    nowRow = KnowledgeSuiteTableSpot_blue.ListRows(KnowledgeSuiteTableSpot_blue.ListRows.Count).Range.row

    ' セルの追加
    If rowDifference > 0 Then
        Range("Y" & lastRow & ":" & "AN" & nowRow - 1).Insert Shift:=xlDown
    ElseIf rowDifference < 0 Then
        Range("Y" & lastRow + 1 & ":" & "AN" & nowRow).Delete Shift:=xlUp
    End If
End Sub

' テーブルにGRPごとの小計を追加する関数
Sub AddSubtotalsForGRP(tbl As ListObject, ByRef rowDifference As Long)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = tbl.Parent
    
    ' テーブルのデータを配列に取得
    Dim dataArray As Variant
    dataArray = tbl.DataBodyRange.Value
    
    ' GRPの一覧を取得
    Dim grpCollection As New Collection
    Dim grpExists As Boolean
    
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArray, 1)
        Dim grpValue As Variant
        grpValue = dataArray(i, KnowledgeSuiteTableColumn.grp + 1)
        
        If Not IsEmpty(grpValue) Then
            ' 既存のGRPかチェック
            grpExists = False
            On Error Resume Next
            grpExists = (TypeName(grpCollection(CStr(grpValue))) <> "Error")
            On Error GoTo 0
            
            If Not grpExists Then
                ' 新しいGRPを追加
                On Error Resume Next
                grpCollection.Add grpValue, CStr(grpValue)
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' テーブルをクリア
    ' tbl.DataBodyRange.Delete
    
    ' GRPごとにデータを追加し、小計を挿入
    Dim grp As Variant
    Dim grpArray() As Variant
    ReDim grpArray(1 To grpCollection.Count)
    
    ' Collectionの内容を配列に移す
    For i = 1 To grpCollection.Count
        grpArray(i) = grpCollection(i)
    Next i
    
    ' グループ処理の前に追加した小計行数を追跡する変数を追加
    Dim addedSubtotalRows As Long
    addedSubtotalRows = 0

    ' 各GRPについて処理
    For i = 1 To UBound(grpArray)
        grp = grpArray(i)
        
        ' 現在のGRPに一致するデータを追加
        Dim rowsAdded As Long
        Dim lastRowOfGRP As Long
        rowsAdded = 0
        lastRowOfGRP = 0
        For j = 1 To UBound(dataArray, 1)
            If dataArray(j, KnowledgeSuiteTableColumn.grp + 1) = grp Then
                
                rowsAdded = rowsAdded + 1
                lastRowOfGRP = j

            End If
        Next j
        
        ' GRPのデータが存在する場合のみ小計行を追加
        If rowsAdded > 0 Then
            ' 小計行を該当グループの最後に挿入
            Dim newRow As ListRow
            Dim k As Long
            ' 小計行を追加
            Set newRow = tbl.ListRows.Add(lastRowOfGRP + 1 + addedSubtotalRows)
            
            ' 小計行の書式設定
            newRow.Range.Interior.Color = RGB(235, 235, 235) ' より薄いグレー色に変更
            
            ' GRP名を設定
            newRow.Range(KnowledgeSuiteTableColumn.grp + 1).Value = grp & " 計"
            
            ' 数値列の合計を計算
            Dim startRow As Long
            Dim endRow As Long
            startRow = lastRowOfGRP - rowsAdded + addedSubtotalRows + 1
            endRow = lastRowOfGRP + addedSubtotalRows

            For k = KnowledgeSuiteTableColumn.売上1月 + 1 To KnowledgeSuiteTableColumn.売上金額 + 1
                ' 単純なSUMを使用してグループ内の合計を計算
                newRow.Range(k).Formula = "=SUM(" & _
                    tbl.ListRows(startRow).Range(k).Address(False, False) & ":" & _
                    tbl.ListRows(endRow).Range(k).Address(False, False) & ")"
            Next k

            ' 追加した小計行数をカウントアップ
            addedSubtotalRows = addedSubtotalRows + 1
        End If
    Next i
    
    ' 総合計行を追加（データがある場合のみ）
    If tbl.ListRows.Count > 0 Then
        ' 総合計行を追加
        Set newRow = tbl.ListRows.Add
        
        ' 総合計行のスタイル設定
        newRow.Range.Font.Bold = True
        newRow.Range.Interior.Color = RGB(220, 220, 220) ' より薄いグレー色に変更
        
        ' 総合計のラベルを設定
        newRow.Range(1).Value = "合計"
        
        ' 数値列の合計を計算 - 小計行を除外して計算
        For k = KnowledgeSuiteTableColumn.売上1月 + 1 To KnowledgeSuiteTableColumn.売上金額 + 1
            ' 各グループの小計行の合計を計算
            Dim sumFormula As String
            sumFormula = "=SUM("
            
            Dim r As Long
            Dim firstAddress As Boolean
            firstAddress = True
            
            For r = 1 To tbl.ListRows.Count - 1
                If InStr(tbl.ListRows(r).Range(KnowledgeSuiteTableColumn.grp + 1).Value, "計") > 0 Then
                    If Not firstAddress Then
                        sumFormula = sumFormula & ","
                    End If
                    sumFormula = sumFormula & tbl.ListRows(r).Range(k).Address
                    firstAddress = False
                End If
            Next r
            
            sumFormula = sumFormula & ")"
            newRow.Range(k).Formula = sumFormula
        Next k
    End If
    
    ' テーブルの更新後の行数計算
    If tbl.Name = "KnowledgeSuiteTableStock_blue" Then
        rowDifference = rowDifference + (tbl.ListRows.Count - initialRowCountStockBlue)
    ElseIf tbl.Name = "KnowledgeSuiteTableSpot_blue" Then
        rowDifference = rowDifference + (tbl.ListRows.Count - initialRowCountSpotBlue)
    End If
    
End Sub

' テーブルの書式設定を行う関数
Sub SetupTableFormat(tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' テーブルスタイルをなしに設定
    tbl.TableStyle = ""
    
    ' 数値列を右寄せに設定
    Dim k As Long
    For k = KnowledgeSuiteTableColumn.売上1月 + 1 To KnowledgeSuiteTableColumn.売上金額 + 1
        tbl.ListColumns(k).DataBodyRange.HorizontalAlignment = xlRight
    Next k

    ' 文字列の書式を設定
    For k = KnowledgeSuiteTableColumn.区分1 + 1 To KnowledgeSuiteTableColumn.grp + 1
        tbl.ListColumns(k).DataBodyRange.HorizontalAlignment = xlLeft
    Next k

    With tbl.Range
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 10
    End With
    

End Sub

'************************************************
'指定ディレクトリのCSVファイルから営業データを取得する
'************************************************
Function getSalesData(ByVal filePath As String) As salesData()
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim blnHeader As Boolean
    Dim var As Variant
    Dim csvResult() As salesData
    Dim lineCount As Integer
    Dim buf As String
    Dim headerLineCount As Integer
    Const headerLine As Integer = 2
    
    blnHeader = True
    lineCount = 0
    headerLineCount = 1


    ' ファイル番号を取得
    fileNum = FreeFile
    
    ' ファイルを開く
    Open filePath For Input As #fileNum
    
    ' ファイルを1行ずつ読み込む
    Do Until EOF(fileNum)
        Line Input #fileNum, buf
        
        If blnHeader = True Then
            If headerLineCount < headerLine Then
                blnHeader = True
                headerLineCount = headerLineCount + 1
            Else
                blnHeader = False
            End If
        Else
            var = splitWithDoubleQuotation(buf)
            
            ' 配列の要素数を確認
            If UBound(var) < 38 Then  ' NegotiationDataの最大インデックスより大きい値
                ' 要素数が足りない場合はスキップまたはログ出力
                Debug.Print "行 " & lineCount & " のデータ形式が不正です: " & buf
                GoTo ContinueLoop
            End If
            
            ReDim Preserve csvResult(lineCount)
            Set csvResult(lineCount) = New salesData
            
            ' エラーハンドリングを追加
            On Error Resume Next
            
            csvResult(lineCount).営業担当 = var(NegotiationData.営業担当者名称)
            csvResult(lineCount).事業区分 = var(NegotiationData.事業区分)
            csvResult(lineCount).顧客名称 = var(NegotiationData.顧客名称)
            csvResult(lineCount).WorkNo = var(NegotiationData.WorkNo)
            csvResult(lineCount).案件名 = var(NegotiationData.案件名)
            csvResult(lineCount).grp = var(NegotiationData.grp)
            csvResult(lineCount).区分1 = var(NegotiationData.区分1)
            csvResult(lineCount).区分2 = var(NegotiationData.区分2)
            csvResult(lineCount).受注見込 = var(NegotiationData.受注見込)
            csvResult(lineCount).フェーズ = var(NegotiationData.フェーズ)
            csvResult(lineCount).開始年月 = var(NegotiationData.開始年月)
            csvResult(lineCount).検収年月 = var(NegotiationData.検収年月)
            csvResult(lineCount).売上年月 = var(NegotiationData.売上年月)
            csvResult(lineCount).最終更新日 = var(NegotiationData.最終更新日)
            csvResult(lineCount).次年度計上 = val(var(NegotiationData.次年度計上))
            csvResult(lineCount).売上金額 = var(NegotiationData.売上金額)
            csvResult(lineCount).売上1月 = var(NegotiationData.売上1月)
            csvResult(lineCount).売上2月 = var(NegotiationData.売上2月)
            csvResult(lineCount).売上3月 = var(NegotiationData.売上3月)
            csvResult(lineCount).売上4月 = var(NegotiationData.売上4月)
            csvResult(lineCount).売上5月 = var(NegotiationData.売上5月)
            csvResult(lineCount).売上6月 = var(NegotiationData.売上6月)
            csvResult(lineCount).売上7月 = var(NegotiationData.売上7月)
            csvResult(lineCount).売上8月 = var(NegotiationData.売上8月)
            csvResult(lineCount).売上9月 = var(NegotiationData.売上9月)
            csvResult(lineCount).売上10月 = var(NegotiationData.売上10月)
            csvResult(lineCount).売上11月 = var(NegotiationData.売上11月)
            csvResult(lineCount).売上12月 = Trim(var(NegotiationData.売上12月))
            
            ' エラーチェック
            If Err.Number <> 0 Then
                Debug.Print "エラー発生: " & Err.Description & " (行 " & lineCount & ")"
                Err.Clear
                GoTo ContinueLoop
            End If
            
            On Error GoTo ErrorHandler
            
            lineCount = lineCount + 1
        End If
ContinueLoop:
    Loop
    
    ' ファイルを閉じる
    Close #fileNum
    
    getSalesData = csvResult
    Exit Function
    
ErrorHandler:
    Debug.Print "getSalesData関数でエラーが発生しました: " & Err.Description & " (エラー番号: " & Err.Number & ")"
    ' エラー時は空の配列を返す
    ReDim csvResult(0)
    getSalesData = csvResult
End Function

Function splitWithDoubleQuotation(ByVal str As String) As Variant
    Dim strTemp As String
    Dim quotCount As Long
    Dim i As Long
    Dim data As String: data = ""
    Dim strCount As Long
    Dim doubleQuoteFlag As Boolean                                              'True=ダブルクォートの中,False=ダブルクォートの外
    Dim retArray() As String
    
    For i = 1 To Len(str)
        strTemp = Mid(str, i, 1)                                                'strから現在の1文字を切り出す
        If strTemp = """" Then                                                  'strTempがダブルクォーテーションなら
            doubleQuoteFlag = Not doubleQuoteFlag                               'ダブルクォーテーションの時は、フラグを逆転させる
        ElseIf strTemp = "," And doubleQuoteFlag Then                           'strTempがカンマかつダブルクォートの中
            data = data & strTemp
        ElseIf strTemp = "," And Not doubleQuoteFlag Then                       'ダブルクォートの外のカンマ発見で区切り
            strCount = strCount + 1
            ReDim Preserve retArray(strCount)
            retArray(strCount - 1) = data
            data = ""
        Else
            data = data & strTemp
        End If
    Next i
    
    If Len(data) > 1 Then
        '最終列のデータを配列に格納
        strCount = strCount + 1
        ReDim Preserve retArray(strCount)
        retArray(strCount - 1) = data
    End If
    splitWithDoubleQuotation = retArray
End Function

Function LongFormatExcludeZero(ByVal inValue As Long) As String
    Dim resultStr As String
    If inValue = 0 Then
        resultStr = "0"
    Else
        resultStr = Format(inValue, "#,#")
    End If
    LongFormatExcludeZero = resultStr
End Function

Function CurrencyFormatExcludeZero(ByVal inValue As Currency) As String
    Dim resultStr As String
    If inValue = 0 Then
        resultStr = "0"
    Else
        resultStr = Format(inValue, "#,#")
    End If
    CurrencyFormatExcludeZero = resultStr
End Function

Function FormatExcludeZeroAndSpace(ByVal inValue As String) As String
    Dim resultStr As String
    If inValue = "0" Then
        resultStr = "0"
    ElseIf inValue = "" Then
        resultStr = ""
    Else
        resultStr = Format(inValue, "#,#")
    End If
    FormatExcludeZeroAndSpace = resultStr
End Function

Function IsEmptyArray(arrayTmp As Variant) As Boolean
    On Error GoTo ERROR_
    If arrayTmp Is Nothing Then
        IsEmptyArray = True
        Exit Function
    End If

    If (0 <= UBound(arrayTmp)) Then
        IsEmptyArray = False
    Else
        IsEmptyArray = True
    End If

    Exit Function
ERROR_:
    IsEmptyArray = True
End Function