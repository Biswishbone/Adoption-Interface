Imports System.Data.SqlClient
Imports System.Web.UI.WebControls.Expressions

Public Class Adoption_Interface
    Inherits System.Web.UI.Page
    Public Shared cons As String = "server=COBBISSQL01.AD.ILSTU.EDU/BIS362;Database=BIS362;trusted_connection=yes;"
    Public Shared con As SqlConnection = New SqlConnection(cons)
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Image1.ImageUrl = "OIP.jpg"
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim adoptionid, dogid, adopterid, adoptiondate, returndate, returnreason, notes As String
        If TextBox2.Text = Nothing Or TextBox3.Text = Nothing Or TextBox4.Text = Nothing Or TextBox5.Text = Nothing Or TextBox6.Text = Nothing Or TextBox7.Text = Nothing Then
            MsgBox("All Fields Must Be Completed", vbExclamation, "Error")
            Exit Sub
        End If

        adoptionid = TextBox1.Text
        dogid = TextBox2.Text
        adopterid = TextBox3.Text
        adoptiondate = TextBox4.Text
        returndate = TextBox5.Text
        returnreason = TextBox6.Text
        notes = TextBox7.Text

        Dim cmdinsert As New SqlCommand("Insert into Adoption (Adoptionid, dogid, adopterid, adoptiondate, returndate, returnreason, notes) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7)", con)
        With cmdinsert.Parameters
            .Clear()
            .AddWithValue("@p1", adoptionid)
            .AddWithValue("@p2", dogid)
            .AddWithValue("@p3", adopterid)
            .AddWithValue("@p4", adoptiondate)
            .AddWithValue("@p5", returndate)
            .AddWithValue("@p6", returnreason)
            .AddWithValue("@p7", notes)
        End With

        Try
            If con.State = ConnectionState.Closed Then con.Open()
            cmdinsert.ExecuteNonQuery()
            Response.Write("Adoption Added Succesffully")
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
End Class
