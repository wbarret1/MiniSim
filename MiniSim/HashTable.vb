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

Public Class CaseInsentiveComparer
    'helper class to create case insentive hash table, as CAPE-OPEN strings are case insentive
    Implements IEqualityComparer
    Public Function Equals1(ByVal x As Object, ByVal y As Object) _
        As Boolean Implements IEqualityComparer.Equals
        Return String.Compare(x, y, StringComparison.InvariantCultureIgnoreCase) = 0
    End Function

    Public Function GetHashCode1(ByVal obj As Object) _
        As Integer Implements IEqualityComparer.GetHashCode
        Return obj.ToString().ToLower().GetHashCode()
    End Function
End Class

