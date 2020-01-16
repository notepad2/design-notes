Imports Microsoft.Office.Interop '声明1
Imports Microsoft.Office.Interop.Word '声明2
'Imports System.Resources
Public Class Form1

    Dim BridgeName, Place, RiverName, Angle, RiverWidth, BridgeSpan, BridgeWidthAll, BridgeAllLong, WidthType As String

    Dim Word1 As Word.Application

    Dim CalcBook As Word.Document

    'Dim table(100) As Word.Table

    Dim para(100) As String

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
        ComboBox1.SelectedIndex = 1
        WidthType = ComboBox1.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim InsertTitle, InsertContent As String

        '启动word

        Word1 = CreateObject("word.application")

        '创建word文档，以指定的模板，可见
        CalcBook = Word1.Documents.Add("DesignNotes.dotm", False, , )
        'CalcBook = Word1.Documents.Open("E:\source.code\design-notes\design-notes\design-notes\Resources\DesignNotes.dotm")

        '文档激活，可见
        Word1.Visible = True
        CalcBook.Activate()

        Dim ActiveRange As Word.Range
        Dim InsertContent As String

        InsertContent = Nothing

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

        GenContent(1, "工程概况")
        para(1) = "本次设计" & BridgeName & "位于" & Place & "，跨越" & RiverName & "，桥梁中心线与河道中心线成" & Angle & "°角斜交，河道宽约" & RiverWidth & "m，桥梁采用" & BridgeSpan & "布跨，桥梁总长" & BridgeAllLong & "m，桥面全宽" & BridgeWidthAll & "m，按" & WidthType & "设计。"
        GenContent(0, para(1))

        GenContent(1, "设计规范及依据")
        GenContent(0, "规范及依据内容")

        GenContent(1, "初步设计批复意见执行情况")
        GenContent(0, Chr(13))

        GenContent(1, "工程地质")



        'InsertContent = "你说你还是喜欢孤单，其实你怕被我看穿"
        'GenContent(0, InsertContent)

        'CalcBook.Paragraphs.Add()


        'ActiveRange.ListFormat.ApplyNumberDefault()


    End Sub


    '函数，在文档末尾按格式插入文字内容，0为正文，123为标题123，4、5为自动编号正文，6为无间隔居中。
    Private Function GenContent(ByVal ContentType As Integer, ByVal ContentWords As String)

        Dim ActiveRange As Word.Range

        ActiveRange = CalcBook.Range(0, CalcBook.Sections.Last.Range.End)
        ActiveRange.Start = ActiveRange.End

        If ActiveRange.Start <> 0 Then
            ActiveRange.InsertAfter(Chr(13))
            ActiveRange.Start = ActiveRange.End
        End If

        If ContentType = 0 Then
            ActiveRange.Style = CalcBook.Styles("正文")
        ElseIf ContentType = 1 Then
            ActiveRange.Style = CalcBook.Styles("标题 1")
        ElseIf ContentType = 2 Then
            ActiveRange.Style = CalcBook.Styles("标题 2")
        ElseIf ContentType = 3 Then
            ActiveRange.Style = CalcBook.Styles("标题 3")
        ElseIf ContentType = 4 Then
            ActiveRange.Style = CalcBook.Styles("正文编号1、")
        ElseIf ContentType = 5 Then
            ActiveRange.Style = CalcBook.Styles("正文编号a）")
        ElseIf ContentType = 6 Then
            ActiveRange.Style = CalcBook.Styles("无间隔")
        End If

        ActiveRange.InsertAfter(ContentWords)

        GenContent = 0

    End Function

End Class
