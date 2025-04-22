Public lastRow As Long
Public KnowledgeSuiteTableStock_blue As ListObject
Public KnowledgeSuiteTableStock_green As ListObject
Public KnowledgeSuiteTableSpot_blue As ListObject
Public KnowledgeSuiteTableSpot_green As ListObject
Public ws As Worksheet

'KnowledgeSuiteTableのサマリを実行する
Sub KnowledgeSuiteSummaryCalc()

    '************************************************
    '営業データを計算する
    '************************************************
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
    If Err.Number <> 0 Then
        MsgBox "必要なテーブルが見つかりません。", vbExclamation
        Exit Sub
    End If

    lastRow = KnowledgeSuiteTableSpot_blue.ListRows(KnowledgeSuiteTableSpot_blue.ListRows.Count).Range.row

    AddSummaryForGRP KnowledgeSuiteTableStock_blue
    AddSummaryForGRP KnowledgeSuiteTableSpot_blue
    AddSummaryForGRP KnowledgeSuiteTableStock_green
    AddSummaryForGRP KnowledgeSuiteTableSpot_green

    ' サマリー部分の範囲全体に設定
    Dim summaryRange As Range
    Set summaryRange = Range(Cells(WholeSummaryInitialLine.ストック, SummaryColumns.売上1月), Cells(lastRow + 53, "X"))

    With summaryRange
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 10
    End With

    Set summaryRange = Range(Cells(lastRow + 4, SummaryColumns.売上1月), Cells(53, "X"))
    With summaryRange
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 10
    End With

    MsgBox "完了しました。"
End Sub

' テーブルにGRPごとのサマリを追加する関数
Sub AddSummaryForGRP(tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = tbl.Parent
        
    ' GRPごとにデータを追加し、小計を挿入
    Dim grp As Variant
    Dim grpArray() As Variant
    
    ' GRPの一覧を取得
    grpArray = Array(GroupConst.次世代金融, GroupConst.国内マーケット, GroupConst.フロント, GroupConst.バックオフィス, GroupConst.デジタルコマース, GroupConst.システム運用, GroupConst.セキュリティ, GroupConst.グローバルマーケット, GroupConst.ワークステクノロジー)

    ' 各GRPについて処理
    For i = 0 To UBound(grpArray)
        grp = grpArray(i)
        
        ' 数値列の合計を計算
        For k = KnowledgeSuiteTableColumn.売上1月 + 1 To KnowledgeSuiteTableColumn.売上金額 + 1
            ' 四半期や半期の列はスキップ
            If (k - 1) = KnowledgeSuiteTableColumn.売上1Q Or _
               (k - 1) = KnowledgeSuiteTableColumn.売上2Q Or _
               (k - 1) = KnowledgeSuiteTableColumn.売上3Q Or _
               (k - 1) = KnowledgeSuiteTableColumn.売上4Q Or _
               (k - 1) = KnowledgeSuiteTableColumn.売上上期 Or _
               (k - 1) = KnowledgeSuiteTableColumn.売下下期 Or _
               (k - 1) = KnowledgeSuiteTableColumn.売上金額 Then
                GoTo NextColumn
            End If
            
            ' 事業部全体のフィルター処理
            Dim totalFormula As String
            totalFormula = "=SUMIF(" & tbl.ListColumns(KnowledgeSuiteTableColumn.区分1 + 1).Range.Address & _
                            ",""" & "合計" & """," & tbl.ListColumns(k).Range.Address & ")"

            ' グループ毎のフィルター処理
            Dim totalGroupFormula As String
            ' 該当グループのデータだけをフィルタリングするための条件式
            totalGroupFormula = "=SUMIF(" & tbl.ListColumns(KnowledgeSuiteTableColumn.grp + 1).Range.Address & _
                            ",""" & grp + " 計" & """," & tbl.ListColumns(k).Range.Address & ")"
            
            ' KnowledgeSuiteTableColumnからSummaryColumnsへの対応付け
            Dim summaryCol As String
            Select Case k - 1  ' +1を相殺するために-1
                Case KnowledgeSuiteTableColumn.売上1月
                    summaryCol = SummaryColumns.売上1月
                Case KnowledgeSuiteTableColumn.売上2月
                    summaryCol = SummaryColumns.売上2月
                Case KnowledgeSuiteTableColumn.売上3月
                    summaryCol = SummaryColumns.売上3月
                Case KnowledgeSuiteTableColumn.売上4月
                    summaryCol = SummaryColumns.売上4月
                Case KnowledgeSuiteTableColumn.売上5月
                    summaryCol = SummaryColumns.売上5月
                Case KnowledgeSuiteTableColumn.売上6月
                    summaryCol = SummaryColumns.売上6月
                Case KnowledgeSuiteTableColumn.売上7月
                    summaryCol = SummaryColumns.売上7月
                Case KnowledgeSuiteTableColumn.売上8月
                    summaryCol = SummaryColumns.売上8月
                Case KnowledgeSuiteTableColumn.売上9月
                    summaryCol = SummaryColumns.売上9月
                Case KnowledgeSuiteTableColumn.売上10月
                    summaryCol = SummaryColumns.売上10月
                Case KnowledgeSuiteTableColumn.売上11月
                    summaryCol = SummaryColumns.売上11月
                Case KnowledgeSuiteTableColumn.売上12月
                    summaryCol = SummaryColumns.売上12月
            End Select
            
            ' 結果をセルに設定
            If InStr(tbl.Name, "KnowledgeSuiteTableStock_blue") > 0 Then

                ' 事業部全体のデータを計算
                Cells(WholeSummaryInitialLine.ストック, summaryCol) = totalFormula

                ' グループ毎のデータを計算
                Select Case grp
                    Case GroupConst.次世代金融
                        Cells(GroupSummaryInitialLine.次世代金融ストック, summaryCol) = totalGroupFormula
                    Case GroupConst.国内マーケット
                        Cells(GroupSummaryInitialLine.国内マーケットストック, summaryCol) = totalGroupFormula
                    Case GroupConst.フロント
                        Cells(GroupSummaryInitialLine.フロントソリューションストック, summaryCol) = totalGroupFormula
                    Case GroupConst.バックオフィス
                        Cells(GroupSummaryInitialLine.バックオフィスソリューションストック, summaryCol) = totalGroupFormula
                    Case GroupConst.デジタルコマース
                        Cells(GroupSummaryInitialLine.デジタルコマースストック, summaryCol) = totalGroupFormula
                    Case GroupConst.システム運用
                        Cells(GroupSummaryInitialLine.システム運用ストック, summaryCol) = totalGroupFormula
                    Case GroupConst.セキュリティ
                        Cells(GroupSummaryInitialLine.セキュリティサービスストック, summaryCol) = totalGroupFormula
                    Case GroupConst.グローバルマーケット
                        Cells(GroupSummaryInitialLine.グローバルマーケットストック, summaryCol) = totalGroupFormula
                    Case GroupConst.ワークステクノロジー
                        Cells(GroupSummaryInitialLine.ワークステクノロジーストック, summaryCol) = totalGroupFormula
                End Select
            ElseIf InStr(tbl.Name, "KnowledgeSuiteTableSpot_blue") > 0 Then
                
                ' 事業部全体のデータを計算
                Cells(WholeSummaryInitialLine.スポット, summaryCol) = totalFormula

                ' グループ毎のデータを計算
                Select Case grp
                    Case GroupConst.次世代金融
                        Cells(GroupSummaryInitialLine.次世代金融スポット, summaryCol) = totalGroupFormula
                    Case GroupConst.国内マーケット
                        Cells(GroupSummaryInitialLine.国内マーケットスポット, summaryCol) = totalGroupFormula
                    Case GroupConst.フロント
                        Cells(GroupSummaryInitialLine.フロントソリューションスポット, summaryCol) = totalGroupFormula
                    Case GroupConst.バックオフィス
                        Cells(GroupSummaryInitialLine.バックオフィスソリューションスポット, summaryCol) = totalGroupFormula
                    Case GroupConst.デジタルコマース
                        Cells(GroupSummaryInitialLine.デジタルコマーススポット, summaryCol) = totalGroupFormula
                    Case GroupConst.システム運用
                        Cells(GroupSummaryInitialLine.システム運用スポット, summaryCol) = totalGroupFormula
                    Case GroupConst.セキュリティ
                        Cells(GroupSummaryInitialLine.セキュリティサービススポット, summaryCol) = totalGroupFormula
                    Case GroupConst.グローバルマーケット
                        Cells(GroupSummaryInitialLine.グローバルマーケットスポット, summaryCol) = totalGroupFormula
                    Case GroupConst.ワークステクノロジー
                        Cells(GroupSummaryInitialLine.ワークステクノロジースポット, summaryCol) = totalGroupFormula
                End Select

            ElseIf InStr(tbl.Name, "KnowledgeSuiteTableStock_green") > 0 Then

                totalFormula = totalFormula & " + " & Cells(WholeSummaryInitialLine.ストック, summaryCol).Address
                ' 事業部全体のデータを計算
                Cells(lastRow + 4, summaryCol) = totalFormula

                ' グループ毎のデータを計算
                Select Case grp
                    Case GroupConst.次世代金融
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.次世代金融ストック, summaryCol).Address
                        Cells(lastRow + 10, summaryCol) = totalGroupFormula
                    Case GroupConst.国内マーケット
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.国内マーケットストック, summaryCol).Address
                        Cells(lastRow + 15, summaryCol) = totalGroupFormula
                    Case GroupConst.フロント
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.フロントソリューションストック, summaryCol).Address
                        Cells(lastRow + 20, summaryCol) = totalGroupFormula
                    Case GroupConst.バックオフィス
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.バックオフィスソリューションストック, summaryCol).Address
                        Cells(lastRow + 25, summaryCol) = totalGroupFormula
                    Case GroupConst.デジタルコマース
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.デジタルコマースストック, summaryCol).Address
                        Cells(lastRow + 30, summaryCol) = totalGroupFormula
                    Case GroupConst.システム運用
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.システム運用ストック, summaryCol).Address
                        Cells(lastRow + 35, summaryCol) = totalGroupFormula
                    Case GroupConst.セキュリティ
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.セキュリティサービスストック, summaryCol).Address
                        Cells(lastRow + 40, summaryCol) = totalGroupFormula
                    Case GroupConst.グローバルマーケット
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.グローバルマーケットストック, summaryCol).Address
                        Cells(lastRow + 45, summaryCol) = totalGroupFormula
                    Case GroupConst.ワークステクノロジー
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.ワークステクノロジーストック, summaryCol).Address
                        Cells(lastRow + 50, summaryCol) = totalGroupFormula
                End Select

            ElseIf InStr(tbl.Name, "KnowledgeSuiteTableSpot_green") > 0 Then

                totalFormula = totalFormula & " + " & Cells(WholeSummaryInitialLine.スポット, summaryCol).Address
                ' 事業部全体のデータを計算
                Cells(lastRow + 5, summaryCol) = totalFormula

                ' グループ毎のデータを計算
                Select Case grp
                    Case GroupConst.次世代金融
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.次世代金融スポット, summaryCol).Address
                        Cells(lastRow + 11, summaryCol) = totalGroupFormula
                    Case GroupConst.国内マーケット
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.国内マーケットスポット, summaryCol).Address
                        Cells(lastRow + 16, summaryCol) = totalGroupFormula
                    Case GroupConst.フロント
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.フロントソリューションスポット, summaryCol).Address
                        Cells(lastRow + 21, summaryCol) = totalGroupFormula
                    Case GroupConst.バックオフィス
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.バックオフィスソリューションスポット, summaryCol).Address
                        Cells(lastRow + 26, summaryCol) = totalGroupFormula
                    Case GroupConst.デジタルコマース
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.デジタルコマーススポット, summaryCol).Address
                        Cells(lastRow + 31, summaryCol) = totalGroupFormula
                    Case GroupConst.システム運用
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.システム運用スポット, summaryCol).Address
                        Cells(lastRow + 36, summaryCol) = totalGroupFormula
                    Case GroupConst.セキュリティ
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.セキュリティサービススポット, summaryCol).Address
                        Cells(lastRow + 41, summaryCol) = totalGroupFormula
                    Case GroupConst.グローバルマーケット
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.グローバルマーケットスポット, summaryCol).Address
                        Cells(lastRow + 46, summaryCol) = totalGroupFormula
                    Case GroupConst.ワークステクノロジー
                        totalGroupFormula = totalGroupFormula & " + " & Cells(GroupSummaryInitialLine.ワークステクノロジースポット, summaryCol).Address
                        Cells(lastRow + 51, summaryCol) = totalGroupFormula
                End Select
            End If
NextColumn:
        Next k
    Next i

End Sub