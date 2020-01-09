Imports Microsoft.Office.Interop '声明1
Imports Microsoft.Office.Interop.Word '声明2
Public Class Form1

    Dim BridgeName, Place, RiverName, Angle, RiverWidth, BridgeSpan, BridgeWidthAll, WidthType As String

    Dim Word1 As Word.Application

    Dim CalcBook As Word.Document

    Dim table(100) As Word.Table

    Dim para(100) As Word.Paragraph

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        BridgeName = TextBox1.Text
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Place = TextBox2.Text
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        RiverName = TextBox3.Text
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Angle = TextBox4.Text
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        RiverWidth = TextBox5.Text
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        BridgeSpan = TextBox6.Text
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        BridgeWidthAll = TextBox7.Text
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "单幅" Then
            WidthType = 1
        ElseIf ComboBox1.Text = "双幅" Then
            WidthType = 2
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ActiveRange As Word.Range

        '启动word
        Word1 = CreateObject("word.application")

        '创建word文档，以指定的模板，可见
        CalcBook = Word1.Documents.Add("DesignNotes.dotm", False, , True)

        'para2 = "道路等级：桥面按非机动车道设计。"

        '文档激活，可见
        Word1.Visible = True
        CalcBook.Activate()

        '页面设置采用A3，横向，分两栏
        'With CalcBook.PageSetup
        '.PaperSize = Word.WdPaperSize.wdPaperA3
        '.Orientation = Word.WdOrientation.wdOrientLandscape
        '.TextColumns.SetCount(2)
        'End With

        'With CalcBook.Styles("正文").Font
        '.NameFarEast = "仿宋"
        '.NameAscii="仿宋"
        '.Name = "仿宋"
        '.Size = 12
        'End With

        'With CalcBook.Styles("正文").ParagraphFormat
        '.CharacterUnitFirstLineIndent = 2
        'End With

        ActiveRange = CalcBook.Range(Start:=0, [End]:=0)

        'CalcBook.Content.InsertAfter("本次设计" & BridgeName & "位于" & Place & "，跨越" & RiverName & "，桥梁中心线与河道中心线成" & Angle)

        '第一章，工程概况
        ActiveRange.Style = CalcBook.Styles("标题 1")
        ActiveRange.InsertAfter("工程概况" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        ActiveRange.Style = CalcBook.Styles("正文")
        ActiveRange.InsertAfter("本次设计" & BridgeName & "位于" & Place & "，跨越" & RiverName & "，桥梁中心线与河道中心线成" & Angle & Chr(13))
        ActiveRange.Start = ActiveRange.End

        '第二章，设计规范及依据
        ActiveRange.Style = CalcBook.Styles("标题 1")
        ActiveRange.InsertAfter("设计规范及依据" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        ActiveRange.Style = CalcBook.Styles("正文")
        ActiveRange.InsertAfter("规范内容" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        '第三章，批复执行情况
        ActiveRange.Style = CalcBook.Styles("标题 1")
        ActiveRange.InsertAfter("初步设计批复意见执行情况" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        ActiveRange.Style = CalcBook.Styles("正文")
        ActiveRange.InsertAfter("批复内容及回复" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        '第四章，工程地质
        ActiveRange.Style = CalcBook.Styles("标题 1")
        ActiveRange.InsertAfter("工程地质" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        ActiveRange.Style = CalcBook.Styles("正文")
        ActiveRange.InsertAfter("土层描述内容" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        '第五章，主要技术标准
        ActiveRange.Style = CalcBook.Styles("标题 1")
        ActiveRange.InsertAfter("主要技术标准" & Chr(13))
        ActiveRange.Start = ActiveRange.End

        ActiveRange.Style = CalcBook.Styles("正文")
        ActiveRange.InsertAfter("技术标准内容" & Chr(13))
        ActiveRange.Start = ActiveRange.End


        'CalcBook.ApplyListTemplate
    End Sub


End Class
