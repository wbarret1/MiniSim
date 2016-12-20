''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' MiniSim - A VB.NET Mini CAPE-OPEN COSE
' (C) Jasper van Baten, AmsterCHEM 2009
'
' This code is intended as accompanying material for a CAPE-OPEN workshop
' at the 2009 Euro CAPE-OPEN conference (Munich, April 4)
'
' If you want to use this code, please contact jasper@amsterchem.com
'  for conditions and 
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Imports System.Windows.Forms

Public Class SelectPP

    Public Function GetSelection() As String
        GetSelection = selected
    End Function

    Private selected As String

    Sub New(ByVal ppm As CapeOpen.ICapeThermoPropertyPackageManager)
        InitializeComponent()
        'list packages:
        Dim packageList() As String
        Try
            packageList = ppm.GetPropertyPackageList
        Catch ex As Exception
            Throw New Exception("Failed to get list of property packages from PPM: " + ex.Message)
        End Try
        If packageList.Length = 0 Then Throw New Exception("Property Package Manager does not expose property packages")
        'fill list
        Dim i As Integer
        For i = 0 To packageList.Length - 1
            List.Items.Add(packageList(i))
        Next i
        List.SelectedItem = List.Items(0)
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        selected = List.SelectedItem
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
