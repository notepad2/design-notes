Imports Microsoft.Office.Interop '声明1
'Imports Microsoft.Office.Interop.Word '声明2
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


        '启动word
        Word1 = CreateObject("word.application")

        '创建word文档，以指定的模板，可见
        CalcBook = Word1.Documents.Add("DesignNotes.dotm", False, , True)

        'para2 = "道路等级：桥面按非机动车道设计。"

        '文档激活
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


        CalcBook.Content.InsertAfter(BridgeName)

    End Sub


End Class
