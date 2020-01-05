Imports Microsoft.Office.Interop '声明1
'Imports Microsoft.Office.Interop.Word '声明2
Public Class Form1

    Dim BridgeName, Place, RiverName, Angle, RiverWidth, BridgeSpan, BridgeWidthAll, WidthType As String



    Dim Word1 As Word.Application



    Dim CalcBook As Word.Document



    Dim table(100) As Word.Table
    Dim para(100) As Word.Paragraph

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Bridgename = TextBox1.Text
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


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim MyRange As Word.Range


        Word1 = CreateObject("word.application")

        CalcBook = CalcBook.Documents.Add
        'MyDoc.PageSetup.
        MyRange = CalcBook.Range(Start:=0, End:=0)

        'para2 = "道路等级：桥面按非机动车道设计。"

        Word1.Visible = True
        CalcBook.Activate()

        CalcBook.Content.InsertAfter(BridgeName)

    End Sub


End Class
