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
Imports Microsoft.Win32

Public Class CapeOpenSelector

    Public Function GetSelectedCLSD() As String
        GetSelectedCLSD = CLSID
    End Function

    Private CLSID As String

    Class ListItem
        Public name As String
        Public clsID As String
        Public Overrides Function ToString() As String
            ToString = name
        End Function
    End Class

    Sub New(ByVal catID As String, ByVal title As String)
        InitializeComponent()
        'set window title:
        Me.Text = title
        'fill the list with all registered CAPE-OPEN components of the given catID
        Dim rk As RegistryKey = Registry.ClassesRoot.OpenSubKey("CLSID", False)
        Dim rk1 As RegistryKey
        Dim classes() As String = rk.GetSubKeyNames
        Dim i As Integer
        Dim name As String
        For i = 0 To classes.Length - 1
            rk1 = rk.OpenSubKey(classes(i) + "\Implemented Categories\" + catID)
            If (rk1 IsNot Nothing) Then
                name = Nothing
                'try to get name from cape description
                rk1 = rk.OpenSubKey(classes(i) + "\CapeDescription")
                If (rk1 IsNot Nothing) Then
                    Try
                        name = rk1.GetValue("Name")
                    Catch
                    End Try
                End If
                If (name Is Nothing) Then name = "CLSID: " + classes(i) 'no name from registry
                'add to list
                Dim it As New ListItem
                it.clsID = classes(i)
                it.name = name
                List.Items.Add(it)
            End If
        Next i
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If (List.SelectedIndex < 0) Then
            MsgBox("Please select an item", MsgBoxStyle.Critical, "Error:")
            Exit Sub
        End If
        Dim item As ListItem = List.SelectedItem
        CLSID = item.clsID
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub List_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles List.MouseDoubleClick
        OK_Button_Click(sender, e)
    End Sub

End Class
