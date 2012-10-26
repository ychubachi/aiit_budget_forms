Attribute VB_Name = "Module2"
Sub 研究費ボタン_Click()
    Worksheets("研究費").Activate
End Sub

Sub 帳簿ボタン_Click()
    Worksheets("帳簿").Activate
End Sub

Sub 出張ボタン_Click()
    Worksheets("出張").Activate
End Sub

Sub 伝票印刷ボタン_Click()
    Dim myNum As Variant
    Dim oval As Excel.Shape
    Dim ovalX, ovalY, ovalW, ovalH As Integer
    Dim oval2 As Excel.Shape
    Dim oval2X, oval2Y, oval2W, oval2H As Integer

    myNum = Application.InputBox("印刷する伝票のNoを入力してください")
    
    If myNum <> False Then
        ' 伝票のNoを転記する
        Worksheets("内部利用").Range("伝票No").Value = myNum
        
        ' 帳簿の種類を判定する
        Dim 伝票種別 As Variant
        伝票種別 = Worksheets("内部利用").Range("伝票種別").Value
        
        
        Select Case 伝票種別
        Case "立替"
            ' 立替払承認届の場合
            Worksheets("立替払承認届").Activate
            With ActiveSheet
                ' 表示する
                .Visible = True
                ' 研究費種別をまるで囲む
                With Worksheets("内部利用")
                    ovalX = CInt(.Range("立替用研究費区分座標").Item(1).Value)
                    ovalY = CInt(.Range("立替用研究費区分座標").Item(2).Value)
                    ovalW = CInt(.Range("立替用研究費区分座標").Item(3).Value)
                    ovalH = CInt(.Range("立替用研究費区分座標").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                
                ' 立替払理由をまるで囲む
                With Worksheets("内部利用")
                    
                    oval2X = CInt(.Range("立替用理由区分座標").Item(1).Value)
                    oval2Y = CInt(.Range("立替用理由区分座標").Item(2).Value)
                    oval2W = CInt(.Range("立替用理由区分座標").Item(3).Value)
                    oval2H = CInt(.Range("立替用理由区分座標").Item(4).Value)
                End With
                            
                Set oval2 = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, oval2X, oval2Y, oval2W, oval2H)
                oval2.Fill.Transparency = 1#
                
                ' PDFファイルに出力する
                PDFファイル出力 "No" & myNum & "-立替払承認届.pdf"
                
                ' まるを削除する
                oval.Delete
                oval2.Delete
                
                ' 非表示にもどす
                .Visible = False
            End With
            ' 帳簿シートに戻る
            Worksheets("帳簿").Activate
        Case "発注"
            ' 発注情報通知書の場合
            Worksheets("発注情報等通知書").Activate
            With ActiveSheet
                ' 表示する
                .Visible = True
                ' 研究費種別をまるで囲む
                With Worksheets("内部利用")
                    ovalX = CInt(.Range("発注用研究費区分座標").Item(1).Value)
                    ovalY = CInt(.Range("発注用研究費区分座標").Item(2).Value)
                    ovalW = CInt(.Range("発注用研究費区分座標").Item(3).Value)
                    ovalH = CInt(.Range("発注用研究費区分座標").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                
                ' PDFファイルに出力する
                PDFファイル出力 "No" & myNum & "-発注情報通知書.pdf"
                
                
                ' まるを削除する
                oval.Delete
                
                ' 非表示にもどす
                .Visible = False
            End With
            ' 帳簿シートに戻る
            Worksheets("帳簿").Activate
        Case "旅費"
            ' 出張の場合、出張シートに遷移する
            MsgBox "出張シートから印刷してください"
            Worksheets("出張").Activate
        Case Else
            MsgBox "指定された伝票の出力には対応しておりません"
        End Select
           
    End If
End Sub


Sub 命令簿内訳書ボタン_Click()
    Dim myNum As Variant
    Dim area As Variant
    Dim oval As Excel.Shape
    Dim ovalX, ovalY, ovalW, ovalH As Integer

    myNum = Application.InputBox("印刷する伝票のNoを入力してください")
    ' myNum = 1
    
    If myNum <> False Then
        With Worksheets("内部利用（旅費）")
            ' 伝票のNoをシートに転記します
            .Range("旅行No").Value = myNum
        
            ' 国内か海外かを取得します
            area = .Range("内外").Value
        End With
        
        Select Case area
        Case "国内"
            ' 旅行命令簿
            Worksheets("旅行命令簿").Activate
            With ActiveSheet
                ' 表示する
                .Visible = True
    
                ' PDFファイルに出力する
                PDFファイル出力 "No" & myNum & "-旅行命令簿.pdf"
    
                ' 非表示にもどす
                .Visible = False
            End With
            
            ' 旅費計算内訳書
            Worksheets("旅費計算内訳書").Activate
            With ActiveSheet
                ' 表示する
                .Visible = True
                
                ' 研究費種別をまるで囲む
                With Worksheets("内部利用（旅費）")
                    ovalX = CInt(.Range("旅行区分座標").Item(1).Value)
                    ovalY = CInt(.Range("旅行区分座標").Item(2).Value)
                    ovalW = CInt(.Range("旅行区分座標").Item(3).Value)
                    ovalH = CInt(.Range("旅行区分座標").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                    
                ' PDFファイルに出力する
                PDFファイル出力 "No" & myNum & "-旅費計算内訳書.pdf"
                    
                ' まるを削除する
                oval.Delete
                    
                ' 非表示にもどす
                .Visible = False
            End With
        Case "海外"
            With Worksheets("様式１（旅行申請書）")
                .Activate
                .Visible = True
                PDFファイル出力 "No" & myNum & "-様式１（旅行申請書）.pdf"
                .Visible = False
            End With
            With Worksheets("様式２甲（旅行命令簿）")
                .Activate
                .Visible = True
                PDFファイル出力 "No" & myNum & "-様式２甲（旅行命令簿）.pdf"
                .Visible = False
            End With
            With Worksheets("様式２乙（旅行日程表）")
                .Activate
                .Visible = True
                PDFファイル出力 "No" & myNum & "-様式２乙（旅行日程表）.pdf"
                .Visible = False
            End With
            
            
        Case Else
            MsgBox "国内または海外のどちらかを選択してください。"
        End Select
        
        
        ' 出張シートに戻る
        Worksheets("出張").Activate
    End If
End Sub

Sub 出張復命書ボタン_Click()
    Dim myNum As Variant

    myNum = Application.InputBox("印刷する伝票のNoを入力してください")
    ' myNum = 1
    
    If myNum <> False Then
        ' 伝票のNoを転記する
        Worksheets("内部利用（旅費）").Range("B7").Value = myNum
        
        ' 旅行命令簿
        Worksheets("出張復命書").Activate
        With ActiveSheet
            ' 表示する
            .Visible = True

            ' 印刷プレビューを表示する
            PDFファイル出力 "No" & myNum & "出張復命書.pdf"
            ' 非表示にもどす
            .Visible = False
        End With
    End If
End Sub


Sub PDFファイル出力(ByVal myFileName As String)
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=myFileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    MsgBox "マイドキュメントにPDFファイル「" & myFileName & "」を作成しました"
End Sub

