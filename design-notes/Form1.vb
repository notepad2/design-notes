Imports Microsoft.Office.Interop '声明1
Imports Microsoft.Office.Interop.Word '声明2
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MyWord As Word.Application
        Dim MyDoc As Word.Document
        Dim table1 As Word.Table
        Dim para1 As Word.Paragraph
        Dim para2 As Word.Paragraph
        Dim para3 As Word.Paragraph
        Dim MyRange As Word.Range


        MyWord = CreateObject("word.application")

        MyDoc = MyWord.Documents.Add
        'MyDoc.PageSetup.
        MyRange = MyDoc.Range(Start:=0, End:=0)

        MyWord.Visible = True
        MyDoc.Activate()



    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
