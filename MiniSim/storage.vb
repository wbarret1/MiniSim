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

<ComClass(Storage.ClassId, Storage.InterfaceId, Storage.EventsId)> _
Public Class Storage
    'memory stream for storing and loading CAPE-OPEN PMCs
    Implements System.Runtime.InteropServices.ComTypes.IStream

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "76B1A614-AAC5-4e23-9FEE-00F86D7D1214"
    Public Const InterfaceId As String = "D039B8EC-0BF5-4edc-A3B7-6D9CBECBDCCE"
    Public Const EventsId As String = "02D9530B-1626-4a27-A75E-CB4B7E9DBD0D"
#End Region

    Sub New()
        'creates a writable stream
        readStream = False
        dataAlloc = 2048 'initial storage length
        dataLen = 0 'amount written
        ReDim data(dataAlloc)
        curpos = 0
    End Sub

    Sub New(ByVal data() As Byte, ByVal datasize As Integer)
        'creates a readible stream
        readStream = True
        Me.data = data
        Me.dataAlloc = datasize
        Me.dataLen = datasize
        curpos = 0
    End Sub

    Public readStream As Boolean
    Public data() As Byte
    Public curpos As Integer
    Public dataLen As Integer
    Public dataAlloc As Integer

    Public Sub Clone(ByRef ppstm As System.Runtime.InteropServices.ComTypes.IStream) Implements System.Runtime.InteropServices.ComTypes.IStream.Clone
        Throw New Exception("Not implemented")
    End Sub

    Public Sub Commit(ByVal grfCommitFlags As Integer) Implements System.Runtime.InteropServices.ComTypes.IStream.Commit
        Throw New Exception("Not implemented")
    End Sub

    Public Sub CopyTo(ByVal pstm As System.Runtime.InteropServices.ComTypes.IStream, ByVal cb As Long, ByVal pcbRead As System.IntPtr, ByVal pcbWritten As System.IntPtr) Implements System.Runtime.InteropServices.ComTypes.IStream.CopyTo
        Throw New Exception("Not implemented")
    End Sub

    Public Sub LockRegion(ByVal libOffset As Long, ByVal cb As Long, ByVal dwLockType As Integer) Implements System.Runtime.InteropServices.ComTypes.IStream.LockRegion
        'void implementation
    End Sub

    Public Sub Read(ByVal pv() As Byte, ByVal cb As Integer, ByVal pcbRead As System.IntPtr) Implements System.Runtime.InteropServices.ComTypes.IStream.Read
        If (cb + curpos > dataLen) Then Throw New Exception("trying to read past end of data")
        Dim i As Integer
        For i = 0 To cb - 1
            pv(i) = data(curpos + i)
        Next
        curpos += cb
        Marshal.WriteInt32(pcbRead, cb)
    End Sub

    Public Sub Revert() Implements System.Runtime.InteropServices.ComTypes.IStream.Revert
        Throw New Exception("Not implemented")
    End Sub

    Public Sub Seek(ByVal dlibMove As Long, ByVal dwOrigin As Integer, ByVal plibNewPosition As System.IntPtr) Implements System.Runtime.InteropServices.ComTypes.IStream.Seek
        Dim newpos As Integer
        Select Case dlibMove
            Case 0 'STREAM_SEEK_SET
                newpos = dlibMove
            Case 1 'STREAM_SEEK_CUR
                newpos = curpos + dlibMove
            Case 2 'STREAM_SEEK_END 
                newpos = dataLen + dlibMove
            Case Else
                Throw New Exception("Invalid seek origin")
        End Select
        If (newpos < 0) Then Throw New Exception("Attempt to seek before start of file")
        If (newpos > dataLen) Then Throw New Exception("Attempt to seek past end of file")
        curpos = newpos
        Marshal.WriteInt32(plibNewPosition, curpos)
    End Sub

    Public Sub SetSize(ByVal libNewSize As Long) Implements System.Runtime.InteropServices.ComTypes.IStream.SetSize
        If (libNewSize > dataAlloc) Then
            'realloc data
            dataAlloc = libNewSize
            ReDim Preserve data(dataAlloc)
        End If
    End Sub

    Public Sub Stat(ByRef pstatstg As System.Runtime.InteropServices.ComTypes.STATSTG, ByVal grfStatFlag As Integer) Implements System.Runtime.InteropServices.ComTypes.IStream.Stat
        Throw New Exception("Not implemented")
    End Sub

    Public Sub UnlockRegion(ByVal libOffset As Long, ByVal cb As Long, ByVal dwLockType As Integer) Implements System.Runtime.InteropServices.ComTypes.IStream.UnlockRegion
        'void implementation
    End Sub

    Public Sub Write(ByVal pv() As Byte, ByVal cb As Integer, ByVal pcbWritten As System.IntPtr) Implements System.Runtime.InteropServices.ComTypes.IStream.Write
        If (readStream) Then Throw New Exception("Stream is read-only")
        Dim i, newLen As Integer
        newLen = curpos + cb
        If (newLen > dataAlloc) Then
            dataAlloc = newLen + dataAlloc 'grow fast
            ReDim Preserve data(dataAlloc)
        End If
        If (newLen > dataLen) Then dataLen = newLen
        For i = 0 To cb - 1
            data(curpos + i) = pv(i)
        Next
        curpos += cb
        Marshal.WriteInt32(pcbWritten, cb)
    End Sub

End Class
