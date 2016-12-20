Imports Microsoft.Win32

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

Public Class MainWnd

    Const catIDUnitOperation As String = "{678C09A5-7D66-11D2-A67D-00105A42887F}" 'cat ID for listing unit operations from registry
    Const catIDPPM As String = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}" 'cat ID for listing property package managers from registry

    'data members

    Dim fileName As String = Nothing
    Dim unit As CapeOpen.ICapeUnit 'unit operation, this is the main interface
    Dim pp As Object 'property package, we require several interfaces
    'we will remember several flavours of the pp
    Friend ppCompounds As CapeOpen.ICapeThermoCompounds
    Friend ppCalcRoutine As CapeOpen.ICapeThermoPropertyRoutine
    Friend ppEqRoutine As CapeOpen.ICapeThermoEquilibriumRoutine
    Friend ppPhases As CapeOpen.ICapeThermoPhases
    Friend ppMOContext As CapeOpen.ICapeThermoMaterialContext
    Friend ppUniConst As CapeOpen.ICapeThermoUniversalConstant
    'compound data
    Friend compIDs As String() = Nothing
    Friend molWeights As Double() = Nothing
    Friend formulae() As String = Nothing
    Friend names() As String = Nothing
    Friend boilTemps() As Double = Nothing
    Friend casNos() As String = Nothing
    'phase data
    Friend phaseList As String() = Nothing
    Friend aggStates As String() = Nothing
    Friend keyComps As String() = Nothing
    'single phase property lookup
    Friend propertyTable As propTable
    'keep track of the current MO as known to the PP
    Friend currentMO As MaterialObject
    'for save & load
    Dim unitClsId As String
    Dim ppClsId As String
    Dim ppName As String

    'methods

    Sub New()
        'the application and main window are being created. We need to ask the user for a unit operation, 
        ' and we need to obtain a PP from a PPM; without either we cannot continue and we exit
        Dim index, index1 As Integer
        InitializeComponent()
        'register msim file
        Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Classes", True)
        rk.CreateSubKey(".msim", RegistryKeyPermissionCheck.ReadWriteSubTree).SetValue(Nothing, "MiniSim File")
        rk = rk.CreateSubKey("MiniSim File", RegistryKeyPermissionCheck.ReadWriteSubTree)
        rk.SetValue(Nothing, "MiniSim File")
        rk.CreateSubKey("Default icon").SetValue(Nothing, """" + Application.ExecutablePath + """,0")
        rk = rk.CreateSubKey("Shell", RegistryKeyPermissionCheck.ReadWriteSubTree)
        rk.SetValue(Nothing, "open")
        rk = rk.CreateSubKey("open", RegistryKeyPermissionCheck.ReadWriteSubTree)
        rk.SetValue(Nothing, "&Open")
        rk.CreateSubKey("Command", RegistryKeyPermissionCheck.ReadWriteSubTree).SetValue(Nothing, """" + Application.ExecutablePath + """ ""%1""")
        'create property table
        propertyTable = New propTable(Me)
        'check command line arguments
        Dim s As String = Environment.CommandLine
        If (s.Length > 0) Then
            'strip app name
            If (s.Substring(0, 1) = """") Then
                'quote delimited
                index = InStr(2, s, """")
                If (index > 0) Then s = s.Substring(index)
            Else
                'white space
                index = InStr(s, " ")
                index1 = InStr(s, vbTab)
                If (index > 0) Then If (index1 > 0) Then If index1 < index Then index = index1
                If (index > 0) Then s = s.Substring(index)
            End If
            s = s.Trim(" " + vbTab)
        End If
        If (s.Length > 2) Then
            'strip quotes
            If (s.Substring(0, 1) = """") Then
                s = s.Substring(1)
                index = s.IndexOf("""")
                If (index >= 0) Then s = s.Substring(0, index)
            End If
        End If
        If (s.Length > 0) Then
            'should be a file name
            fileName = s
            LoadFile()
        Else
            'init without file
            InitError.Text = "Select New or Open from the File menu"
        End If
    End Sub

    Sub NewFile()
        Dim selectDlg As CapeOpenSelector
        'prepare for a new file
        fileName = Nothing
        'first ask for a unit operation to simulate
        selectDlg = New CapeOpenSelector(catIDUnitOperation, "Select unit operation:")
        If selectDlg.ShowDialog <> Windows.Forms.DialogResult.OK Then
            InitError.Text = "Unit operation selection was cancelled"
            Exit Sub
        End If
        'create the unit
        unitClsId = selectDlg.GetSelectedCLSD
        unit = Activator.CreateInstance(Type.GetTypeFromCLSID(New System.Guid(unitClsId)))
        If (unit Is Nothing) Then
            InitError.Text = "Failed to create unit operation of the given type"
            Exit Sub
        End If
        'initialize the unit
        InitCAPEOPENObject(unit)
        'ask for a PPM (version 1.1 PPM is all we support in this application)
        selectDlg = New CapeOpenSelector(catIDPPM, "Select property package manager:")
        If selectDlg.ShowDialog <> Windows.Forms.DialogResult.OK Then
            InitError.Text = "PPM selection was cancelled"
            Exit Sub
        End If
        'create the PPM
        Dim ppm As CapeOpen.ICapeThermoPropertyPackageManager
        ppClsId = selectDlg.GetSelectedCLSD
        ppm = Activator.CreateInstance(Type.GetTypeFromCLSID(New System.Guid(ppClsId)))
        If (ppm Is Nothing) Then
            InitError.Text = "Failed to create property package manager of the given type"
            Exit Sub
        End If
        'initialize the ppm
        InitCAPEOPENObject(ppm)
        'ask for a list of property packages:
        Dim ppDialog As SelectPP = Nothing
        Try
            ppDialog = New SelectPP(ppm)
        Catch ex As Exception
            'failed to get list of PPs, see SelectPP.New
            InitError.Text = ErrorFromCOObject(ppm, ex)
            Exit Sub
        End Try
        If ppDialog.ShowDialog <> Windows.Forms.DialogResult.OK Then
            InitError.Text = "PP selection was cancelled"
            Exit Sub
        End If
        'create PP
        ppName = ppDialog.GetSelection()
        Try
            pp = ppm.GetPropertyPackage(ppName)
        Catch ex As Exception
            InitError.Text = "Failed to create property package: " + ErrorFromCOObject(ppm, ex)
            CleanUpCAPEOPENObject(ppm)
            Exit Sub
        End Try
        'done with PPM
        CleanUpCAPEOPENObject(ppm)
        'init PP
        InitCAPEOPENObject(pp)
        If Not dataFromPP() Then Exit Sub
        'initialization all ok, hide init error and show main controls
        SetControls(True)
        'revise the initial grid layout and create material objects
        UpdateMaterialsAndGrid()
    End Sub

    Function dataFromPP() As Boolean
        Dim i As Integer
        dataFromPP = False
        'we will work with this PP throughout the simulation. We will ask for a couple of details. First obtain interfaces:
        Try
            ppCompounds = pp
            ppCalcRoutine = pp
            ppEqRoutine = pp
            ppPhases = pp
            ppMOContext = pp
        Catch ex As Exception
            InitError.Text = "Failed to obtain interface from PP: " + ex.Message
            Exit Function
        End Try
        'do not fail on this one; it is of so little use that we will be ok if not having it
        Try
            ppUniConst = pp
        Catch
            ppUniConst = Nothing
        End Try
        'get list of compounds from PP
        Try
            ppCompounds.GetCompoundList(compIDs, formulae, names, boilTemps, molWeights, casNos)
        Catch ex As Exception
            InitError.Text = "Failed to obtain compound list from PP: " + ErrorFromCOObject(ppCompounds, ex)
            Exit Function
        End Try
        'sanity check
        If (compIDs.Length() = 0) Then
            InitError.Text = "PP does not expose compounds"
            Exit Function
        End If
        If (compIDs.Length() <> molWeights.Length) Then
            InitError.Text = "PP returns invalid list of molecular weights (invalid size)"
            Exit Function
        End If
        propertyTable.SetCompCount(compIDs.Length)
        'we need all mole weights
        For i = 0 To molWeights.Length - 1
            If (Double.IsNaN(molWeights(i))) Then
                InitError.Text = "Molecular weight unknown for compound " + compIDs(i) + " (this application requires all mole weights)"
                Exit Function
            End If
        Next i
        'get phase list from PP
        Try
            ppPhases.GetPhaseList(phaseList, aggStates, keyComps)
        Catch ex As Exception
            InitError.Text = "Failed to obtain phase list from PP: " + ErrorFromCOObject(ppPhases, ex)
            Exit Function
        End Try
        'sanity check
        If (phaseList.Length() = 0) Then
            InitError.Text = "PP does not expose phases"
            Exit Function
        End If
        dataFromPP = True
    End Function

    Sub SetControls(ByVal runMode As Boolean)
        InitError.Visible = Not runMode
        DataGrid.Visible = runMode
        Edit.Visible = runMode
        Status.Visible = runMode
        SolveBtn.Visible = runMode
        ReportList.Visible = runMode
        ShowReport.Visible = runMode
    End Sub

    Sub Reset()
        'drop all current objects
        unitClsId = Nothing
        ppClsId = Nothing
        ppName = Nothing
        DataGrid.Columns.Clear() 'will remove all materials
        If (unit IsNot Nothing) Then CleanUpCAPEOPENObject(unit)
        If (pp IsNot Nothing) Then CleanUpCAPEOPENObject(pp)
        unit = Nothing
        pp = Nothing
        ppCompounds = Nothing
        ppCalcRoutine = Nothing
        ppEqRoutine = Nothing
        ppPhases = Nothing
        ppMOContext = Nothing
        ppUniConst = Nothing
        compIDs = Nothing
        molWeights = Nothing
        formulae = Nothing
        names = Nothing
        boilTemps = Nothing
        casNos = Nothing
        phaseList = Nothing
        aggStates = Nothing
        keyComps = Nothing
        currentMO = Nothing
        'set window to non-running mode
        SetControls(False)
        Status.Text = String.Empty
        InitError.Text = String.Empty
    End Sub


    Protected Overrides Sub Finalize()
        'do some cleanup 
        If (unit IsNot Nothing) Then CleanUpCAPEOPENObject(unit)
        If (pp IsNot Nothing) Then CleanUpCAPEOPENObject(pp)
        'done
        MyBase.Finalize()
    End Sub

    Private Sub UpdateMaterialsAndGrid()
        'first revise the material; we add a material to all unconnected ports
        Dim port As CapeOpen.ICapeUnitPort
        Dim portCollection As CapeOpen.ICapeCollection
        Dim i As Integer, j As Integer
        'first wipe grid
        ReportList.Items.Clear()
        DataGrid.Columns.Clear()
        DataGrid.Rows.Clear()
        'we cannot solve at this point
        SolveBtn.Enabled = False
        Try
            portCollection = unit.ports
            'loop over all ports
            For i = 1 To portCollection.Count
                Try
                    port = portCollection.Item(i)
                    'get port name (make up a name if fail)
                    Dim name As String
                    Dim iden As CapeOpen.ICapeIdentification
                    Try
                        iden = port
                        name = iden.ComponentName
                    Catch ex As Exception
                        name = "Port " + i.ToString
                    End Try
                    'check port type
                    Try
                        If (port.portType <> CapeOpen.CapePortType.CAPE_MATERIAL) Then
                            MsgBox("Ignoring port " + name + ": not a material port", MsgBoxStyle.Information, "MiniSim:")
                            ReleaseIfCOM(port)
                            Continue For
                        End If
                    Catch ex As Exception
                        MsgBox("Failed to check port type from unit operation: " + ErrorFromCOObject(port, ex), MsgBoxStyle.Exclamation, "Error:")
                    End Try
                    'check port direction
                    Dim isInlet As Boolean = False
                    Try
                        Dim direc As CapeOpen.CapePortDirection
                        direc = port.direction
                        If (direc = CapeOpen.CapePortDirection.CAPE_INLET) Or (direc = CapeOpen.CapePortDirection.CAPE_INLET_OUTLET) Then isInlet = True
                    Catch ex As Exception
                        MsgBox("Failed to check port direction from unit operation: " + ErrorFromCOObject(port, ex), MsgBoxStyle.Exclamation, "Error:")
                    End Try
                    'get the connected object
                    Dim mat As MaterialObject = Nothing
                    Try
                        mat = port.connectedObject
                    Catch ex As Exception
                        'no connected object
                        mat = Nothing
                    End Try
                    If (mat Is Nothing) Then
                        'create a material and connect to port
                        mat = New MaterialObject
                        mat.Init(Me, name)
                        Try
                            port.Connect(mat)
                        Catch ex As Exception
                            MsgBox("Ignoring port " + name + ", failed to connect: " + ErrorFromCOObject(port, ex), MsgBoxStyle.Information, "MiniSim:")
                            ReleaseIfCOM(port)
                            Continue For
                        End Try
                        If isInlet Then
                            'set some default values
                            mat.T = 300.0
                            mat.P = 101325
                            mat.totFlowMassBasis = False
                            For j = 0 To compIDs.Length - 1
                                mat.overallComposition(j) = 1.0 / CType(compIDs.Length, Double)
                            Next j
                            mat.overallCompositionMassBasis = False
                        End If
                    End If
                    'set material in or outlet
                    mat.Inlet = isInlet
                    'add column to grid, remember material reference
                    Dim colIndex As Integer = DataGrid.Columns.Add(name, name)
                    DataGrid.Columns(colIndex).SortMode = DataGridViewColumnSortMode.NotSortable
                    DataGrid.Columns(colIndex).Tag = mat
                    'done with port
                    ReleaseIfCOM(port)
                Catch ex As Exception
                    'failed to get port from unit operation
                    MsgBox("Failed to get port " + i.ToString() + " from unit operation: " + ErrorFromCOObject(portCollection, ex), MsgBoxStyle.Exclamation, "Error:")
                End Try
            Next i
            'done with port collection
            ReleaseIfCOM(portCollection)
            If (DataGrid.Columns.Count > 0) Then
                'create the rows for the grid
                DataGrid.Rows.Add()
                DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = "Direction"
                DataGrid.Rows(DataGrid.Rows.Count - 1).ReadOnly = True
                DataGrid.Rows.Add()
                DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = "Temperature/[K]"
                DataGrid.Rows.Add()
                DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = "Pressure/[Pa]"
                DataGrid.Rows.Add()
                DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = "Total flow/[mol/s]"
                For i = 0 To compIDs.Length - 1
                    DataGrid.Rows.Add()
                    DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = "mole frac " + compIDs(i)
                Next i
                'add all phases
                For j = 0 To phaseList.Length - 1
                    DataGrid.Rows.Add() 'blank row
                    DataGrid.Rows(DataGrid.Rows.Count - 1).ReadOnly = True
                    DataGrid.Rows.Add()
                    DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = phaseList(j) + " fraction"
                    DataGrid.Rows(DataGrid.Rows.Count - 1).ReadOnly = True
                    For i = 0 To compIDs.Length - 1
                        DataGrid.Rows.Add()
                        DataGrid.Rows(DataGrid.Rows.Count - 1).HeaderCell.Value = phaseList(j) + " " + compIDs(i)
                        DataGrid.Rows(DataGrid.Rows.Count - 1).ReadOnly = True
                    Next i
                Next j
                Try
                    'set the grid content, as far as available
                    SetGridContent()
                    'validate 
                    DoValidate()
                Catch
                End Try
            Else
                Status.Text = "No valid ports present"
            End If
        Catch ex As Exception
            'failed to get port names
            Status.Text = "Failed to get port collection from unit operation: " + ErrorFromCOObject(unit, ex)
        End Try
        'fill report list
        Dim rep As CapeOpen.ICapeUnitReport
        Try
            rep = unit
            Dim repNames() As String
            Try
                repNames = rep.reports
                For i = 0 To repNames.Length - 1
                    ReportList.Items.Add(repNames(i))
                Next i
                If (ReportList.Items.Count) Then ReportList.SelectedIndex = 0
            Catch ex As Exception
                MsgBox("Failed to get report list: " + ex.Message, MsgBoxStyle.Critical, "Error:")
            End Try
        Catch
            'no reports
        End Try
    End Sub

    Private Sub SetGridContent()
        'set the content for the grid, as far as we have it available
        ' this method demonstrates the principles; it is not optimized for performance
        Dim i, j, k As Integer
        Dim d() As Double = Nothing
        Dim mat As MaterialObject
        Dim rdOnly As Boolean
        Dim rowIndex As Integer
        'loop over all columns
        For i = 0 To DataGrid.Columns.Count - 1
            'material for this column
            mat = DataGrid.Columns(i).Tag 'this is where we stored the material refernece in UpdateMaterialsAndGrid
            rowIndex = 0
            If (mat.isInlet) Then
                rdOnly = False
                DataGrid.Rows(rowIndex).Cells(i).Value = "inlet"
            Else
                rdOnly = True
                DataGrid.Rows(rowIndex).Cells(i).Value = "outlet"
            End If
            rowIndex += 1
            'T
            If Not Double.IsNaN(mat.T) Then DataGrid.Rows(rowIndex).Cells(i).Value = mat.T.ToString
            If (rdOnly) Then DataGrid.Rows(rowIndex).Cells(i).ReadOnly = True
            rowIndex += 1
            'P
            If Not Double.IsNaN(mat.P) Then DataGrid.Rows(rowIndex).Cells(i).Value = mat.P.ToString
            If (rdOnly) Then DataGrid.Rows(rowIndex).Cells(i).ReadOnly = True
            rowIndex += 1
            If (mat.isInlet) Then
                'we presume inlet flow and composition values to be in mole basis, as we set them only in mole basis
                If Not Double.IsNaN(mat.totFlow) Then DataGrid.Rows(rowIndex).Cells(i).Value = mat.totFlow.ToString
                rowIndex += 1
                For j = 0 To compIDs.Length - 1
                    'we presume overall composition for 
                    If Not Double.IsNaN(mat.overallComposition(j)) Then DataGrid.Rows(rowIndex).Cells(i).Value = mat.overallComposition(j).ToString
                    rowIndex += 1
                Next j
            Else
                'get totFlow using GetProp, will make sure that we have mass /mole conversion
                Try
                    mat.GetOverallProp("totalFlow", "mole", d)
                Catch ex As Exception
                    ReDim d(0)
                    d(0) = Double.NaN
                End Try
                If Not Double.IsNaN(d(0)) Then DataGrid.Rows(rowIndex).Cells(i).Value = d(0).ToString
                DataGrid.Rows(rowIndex).Cells(i).ReadOnly = True
                rowIndex += 1
                'get composition using GetProp, will make sure that we have mass /mole conversion
                Try
                    mat.GetOverallProp("fraction", "mole", d)
                Catch ex As Exception
                    ReDim d(compIDs.Length - 1)
                    For j = 0 To compIDs.Length - 1
                        d(j) = Double.NaN
                    Next j
                End Try
                For j = 0 To compIDs.Length - 1
                    'we presume overall composition for 
                    If Not Double.IsNaN(d(j)) Then DataGrid.Rows(rowIndex).Cells(i).Value = d(j).ToString
                    DataGrid.Rows(rowIndex).Cells(i).ReadOnly = True
                    rowIndex += 1
                Next j
            End If
            'no phase data if not in equilibrium
            If Not mat.inEquilibrium Then
                'just wipe values
                k = (2 + compIDs.Length) * phaseList.Length
                For j = 0 To k - 1
                    DataGrid.Rows(rowIndex).Cells(i).Value = String.Empty
                    rowIndex += 1
                Next j
                Continue For
            End If
            'all phases
            For k = 0 To phaseList.Length - 1
                If (mat.phaseExists(k)) Then
                    'skip line
                    rowIndex += 1
                    'phase fraction
                    Try
                        mat.GetSinglePhaseProp("phaseFraction", phaseList(k), "mole", d)
                        DataGrid.Rows(rowIndex).Cells(i).Value = d(0).ToString
                    Catch ex As Exception
                        DataGrid.Rows(rowIndex).Cells(i).Value = String.Empty
                    End Try
                    rowIndex += 1
                    'composition
                    Try
                        mat.GetSinglePhaseProp("fraction", phaseList(k), "mole", d)
                    Catch ex As Exception
                        ReDim d(compIDs.Length - 1)
                        For j = 0 To compIDs.Length - 1
                            d(i) = Double.NaN
                        Next j
                    End Try
                    For j = 0 To compIDs.Length - 1
                        'we presume overall composition for 
                        If Not Double.IsNaN(d(j)) Then DataGrid.Rows(rowIndex).Cells(i).Value = d(j).ToString
                        DataGrid.Rows(rowIndex).Cells(i).ReadOnly = True
                        rowIndex += 1
                    Next j
                Else
                    'skip phase
                    For j = 0 To compIDs.Length + 1
                        DataGrid.Rows(rowIndex).Cells(i).Value = String.Empty
                        rowIndex += 1
                    Next j
                End If
            Next k
        Next i
    End Sub

    Private Sub DoValidate()
        Dim i, j As Integer
        Dim mat As MaterialObject
        SolveBtn.Enabled = False
        'loop over all input materials to see whether complete and ok
        For i = 0 To DataGrid.Columns.Count - 1
            mat = DataGrid.Columns(i).Tag
            If (mat.isInlet) Then
                Dim complete As Boolean = True
                If (Double.IsNaN(mat.T)) Then
                    complete = False
                ElseIf (Double.IsNaN(mat.P)) Then
                    complete = False
                ElseIf (Double.IsNaN(mat.totFlow)) Then
                    complete = False
                Else
                    For j = 0 To compIDs.Length - 1
                        If (Double.IsNaN(mat.overallComposition(j))) Then
                            complete = False
                            Exit For
                        End If
                    Next j
                End If
                If Not complete Then
                    Status.Text = "Feed """ + mat.ComponentName + """ is not fully specified"
                    Exit Sub
                End If
            End If
        Next i
        'all feed streams are complete, check if also in equilibrium
        For i = 0 To DataGrid.Columns.Count - 1
            mat = DataGrid.Columns(i).Tag
            If (mat.isInlet) Then
                If Not mat.inEquilibrium Then
                    Status.Text = "Feed """ + mat.ComponentName + """ is not in thermodynamic equilibrium"
                    Exit Sub
                End If
            End If
        Next i
        'validate the unit operation 
        Try
            Dim s As String = Nothing
            If Not unit.Validate(s) Then
                Status.Text = "Unit: " + s
                Exit Sub
            End If
        Catch ex As Exception
            Status.Text = "Failed to validate unit: " + ErrorFromCOObject(unit, ex)
        End Try
        'all ok 
        Status.Text = "Ready to solve"
        SolveBtn.Enabled = True
    End Sub

    'utility functions

    Public Sub InitCAPEOPENObject(ByVal o As Object)
        'if IPersistStreamInit, we must call InitNew
        Dim iPI As CapeOpen.IPersistStreamInit
        Try
            iPI = o
            Try
                iPI.InitNew()
            Catch ex As Exception
                MsgBox("Failed to call InitNew: " + ex.Message, MsgBoxStyle.Exclamation, "InitNew failed:")
            End Try
        Catch
            'no IPersistStreamInit
        End Try
        'if ICapeUtilities is implemented, we must call Initialize
        Dim utils As CapeOpen.ICapeUtilities
        Try
            utils = o
            'implemented
            Try
                utils.Initialize()
            Catch ex As Exception
                MsgBox("Failed to call Initialize: " + ex.Message, MsgBoxStyle.Exclamation, "Initialize failed:")
            End Try
        Catch
            'not implemented
        End Try
    End Sub

    Public Sub CleanUpCAPEOPENObject(ByVal o As Object)
        'if ICapeUtilities is implemented, we must call Terminate
        Dim utils As CapeOpen.ICapeUtilities
        Try
            utils = o
            'implemented
            Try
                utils.Terminate()
            Catch ex As Exception
                MsgBox("Failed to call Terminate: " + ex.Message, MsgBoxStyle.Exclamation, "Terminate failed:")
            End Try
        Catch
            'not implemented
        End Try
        'release the object
        ReleaseIfCOM(o)
    End Sub

    Private Sub Edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Edit.Click
        'edit the unit operation
        Dim utils As CapeOpen.ICapeUtilities
        Try
            utils = unit
            Try
                Dim oldStatus As String = Status.Text
                Status.Text = "Editing unit operation"
                utils.Edit()
                Status.Text = oldStatus
                'update materials and grid
                UpdateMaterialsAndGrid()
            Catch ex As Exception
                MsgBox("Failed to edit unit operation: " + ex.Message, MsgBoxStyle.Exclamation, "Edit:")
            End Try
        Catch ex As Exception
            'no utilities
            MsgBox("Cannot edit unit operation: no ICapeUtilities interface", MsgBoxStyle.Exclamation, "Edit:")
        End Try
    End Sub

    Private Sub CloseBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseBtn.Click
        Me.Close()
    End Sub

    Private Sub SolveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SolveBtn.Click
        'attempt to solve
        Try
            unit.Calculate()
            Status.Text = "Unit is solved"
        Catch ex As Exception
            Status.Text = "Failed to solve: " + ErrorFromCOObject(unit, ex)
        End Try
        'update grid
        SetGridContent()
    End Sub

    Public Shared Function ErrorFromCOObject(ByVal o As Object, ByVal e As Exception) As String
        'get an error description from a CAPE-OPEN object
        Dim err As CapeOpen.ECapeUser
        Dim res As String
        Try
            err = o
            Try
                res = err.description
                'if no description, use e
                If (res Is Nothing) Then
                    If (e IsNot Nothing) Then res = e.Message
                ElseIf (res = String.Empty) Then
                    If (e IsNot Nothing) Then res = e.Message
                End If
            Catch ex As Exception
                res = "Failed to get error from CAPE-OPEN object: " + ex.Message
            End Try
        Catch ex As Exception
            If e Is Nothing Then
                res = "Unknown error (ECapeUser not exposed by object)"
            Else
                res = e.Message
            End If
        End Try
        Return res
    End Function

    Private Sub DataGrid_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGrid.CellValidating
        'data was typed in the grid. See if we can accept it and act on it
        e.Cancel = False
        'check valid row
        If ((e.RowIndex < 1) Or (e.RowIndex > compIDs.Length + 3)) Then Exit Sub
        'get material
        Dim mat As MaterialObject
        mat = DataGrid.Columns(e.ColumnIndex).Tag
        If Not mat.isInlet Then Exit Sub 'read only data
        'check empty
        If (e.FormattedValue = String.Empty) Then
            Status.Text = "Invalid data entry: empty string"
            e.Cancel = True
            Exit Sub
        End If
        'get double data:
        Dim d As Double = Double.NaN
        Try
            d = CType(e.FormattedValue, Double)
        Catch ex As Exception
            e.Cancel = True
            Status.Text = "Invalid data entry: cannot convert to numeric value"
            Exit Sub
        End Try
        'which data?
        If (e.RowIndex = 1) Then
            'temperature
            mat.T = d
        ElseIf (e.RowIndex = 2) Then
            'pressure
            mat.P = d
        ElseIf (e.RowIndex = 3) Then
            'flow
            mat.totFlowMassBasis = False
            mat.totFlow = d
        Else
            'composition
            mat.overallComposition(e.RowIndex - 4) = d
            mat.overallCompositionMassBasis = False
        End If
        If Not CheckCompleteAndFlash(mat) Then
            'skip validate
            SetGridContent()
            Exit Sub
        End If
        'refill list
        SetGridContent()
        'validate
        DoValidate()
    End Sub

    Private Function CheckCompleteAndFlash(ByVal mat As MaterialObject) As Boolean
        Dim i As Integer
        CheckCompleteAndFlash = True
        'check all data complete
        Dim complete As Boolean
        complete = True
        If (Double.IsNaN(mat.T)) Then
            complete = False
        ElseIf (Double.IsNaN(mat.P)) Then
            complete = False
        Else
            For i = 0 To compIDs.Length - 1
                If (Double.IsNaN(mat.overallComposition(i))) Then
                    complete = False
                    Exit For
                End If
            Next i
        End If
        If (complete) Then
            Try
                Dim phaseStatus(phaseList.Length - 1) As Integer
                For i = 0 To phaseList.Length - 1
                    phaseStatus(i) = CapeOpen.CapePhaseStatus.CAPE_UNKNOWNPHASESTATUS
                Next i
                mat.SetPresentPhases(phaseList, phaseStatus)
                'flash the material
                Dim spec1(2) As String, spec2(2) As String
                spec1(0) = "temperature"
                spec1(1) = String.Empty
                spec1(2) = "Overall"
                spec2(0) = "pressure"
                spec2(1) = String.Empty
                spec2(2) = "Overall"
                Try
                    mat.CalcEquilibrium(spec1, spec2, "unspecified")
                Catch ex As Exception
                    Status.Text = "Flash failed: " + ErrorFromCOObject(mat, ex)
                    CheckCompleteAndFlash = False
                    SolveBtn.Enabled = False
                    Exit Function
                End Try
            Catch ex As Exception
                Status.Text = "Failed to set list of phase before flash: " + ErrorFromCOObject(mat, ex)
                CheckCompleteAndFlash = False
                SolveBtn.Enabled = False
                Exit Function
            End Try
        End If
    End Function

    Private Sub ReleaseIfCOM(ByVal o As Object)
        If o.GetType().IsCOMObject Then Marshal.ReleaseComObject(o)
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        Dim dlg As New About
        dlg.ShowDialog()
    End Sub

    Private Function GetNumericFormat() As IFormatProvider
        Dim n As New System.Globalization.NumberFormatInfo
        n.NumberDecimalDigits = 16
        n.NumberDecimalSeparator = "."
        n.NumberGroupSeparator = ""
        Dim noGroups() As Integer = {}
        n.NumberGroupSizes = noGroups
        n.NegativeSign = "-"
        Return n
    End Function

    Sub SaveFile()
        Dim writer As System.Xml.XmlWriter
        Dim errorTxt As String
        Dim i, j As Integer
        Dim format As IFormatProvider = GetNumericFormat()
        'create file
        Try
            writer = System.Xml.XmlWriter.Create(fileName)
        Catch ex As Exception
            errorTxt = "Failed to save: " + ex.Message
            Status.Text = errorTxt
            MsgBox(errorTxt, MsgBoxStyle.Critical, "Save:")
            Exit Sub
        End Try
        Try
            writer.WriteStartDocument(True)
            writer.WriteStartElement("Minisim")
            'unit
            writer.WriteStartElement("Unit")
            writer.WriteAttributeString("CLSID", unitClsId)
            'save unit
            Dim istr As CapeOpen.IPersistStream = Nothing
            Dim istrp As CapeOpen.IPersistStreamInit = Nothing
            Dim o As Object = Nothing
            Try
                istr = unit
                o = istr
            Catch
                'no IPersistStream, try IPersistStreamInit
                Try
                    istrp = unit
                    o = istrp
                Catch
                End Try
            End Try
            If (o IsNot Nothing) Then
                Try
                    Dim s As New Storage
                    If (istr IsNot Nothing) Then istr.Save(s, True) Else istrp.Save(s, True)
                    writer.WriteAttributeString("datalen", s.dataLen.ToString(format))
                    writer.WriteBase64(s.data, 0, s.dataLen)
                Catch ex As Exception
                    Throw New Exception(ErrorFromCOObject(o, ex))
                End Try
            End If
            writer.WriteEndElement()
            'save PP
            writer.WriteStartElement("PP")
            writer.WriteAttributeString("CLSID", ppClsId)
            writer.WriteAttributeString("Name", ppName)
            writer.WriteEndElement()
            'save all feed streams
            For i = 0 To DataGrid.Columns.Count - 1
                Dim mat As MaterialObject
                mat = DataGrid.Columns(i).Tag
                If (mat.isInlet) Then
                    writer.WriteStartElement("Feed")
                    writer.WriteAttributeString("Name", mat.ComponentName)
                    If Not Double.IsNaN(mat.T) Then writer.WriteAttributeString("T", mat.T.ToString(format))
                    If Not Double.IsNaN(mat.P) Then writer.WriteAttributeString("P", mat.P.ToString(format))
                    If Not Double.IsNaN(mat.totFlow) Then writer.WriteAttributeString("flow", mat.totFlow.ToString(format))
                    For j = 0 To compIDs.Length - 1
                        If Not Double.IsNaN(mat.overallComposition(j)) Then writer.WriteAttributeString("X" + (j + 1).ToString(format), mat.overallComposition(j).ToString(format))
                    Next j
                    writer.WriteEndElement()
                End If
            Next i
            writer.WriteEndElement()
            'done 
            writer.Close()
        Catch ex As Exception
            writer.Close()
            errorTxt = "Failed to save: " + ex.Message
            Status.Text = errorTxt
            MsgBox(errorTxt, MsgBoxStyle.Critical, "Save:")
            Exit Sub
        End Try
        'all ok
        Status.Text = "Saved " + fileName
    End Sub

    Sub LoadFile()
        Dim errorTxt As String
        Dim i As Integer
        Dim fname As String = fileName
        Dim val As String
        Reset() 'wipes file name as well
        fileName = fname
        Dim format As IFormatProvider = GetNumericFormat()
        Dim reader As System.Xml.XmlReader
        'create file
        Try
            reader = System.Xml.XmlReader.Create(fileName)
        Catch ex As Exception
            errorTxt = "Failed to load: " + ex.Message
            Status.Text = errorTxt
            MsgBox(errorTxt, MsgBoxStyle.Critical, "Load:")
            Exit Sub
        End Try
        Reset()
        Try
            reader.MoveToContent()
            ' Parse the file and display each of the nodes.
            While reader.Read()
                Select Case reader.NodeType
                    Case System.Xml.XmlNodeType.Element
                        If (reader.Name = "Unit") Then
                            val = reader.GetAttribute("CLSID")
                            If (val Is Nothing) Then Throw New Exception("unit misses CLSID in file")
                            'create unit
                            unitClsId = val
                            unit = Activator.CreateInstance(Type.GetTypeFromCLSID(New System.Guid(unitClsId)))
                            If (unit Is Nothing) Then Throw New Exception("Failed to create unit operation of the given type")
                            'get storage interface
                            Dim istr As CapeOpen.IPersistStream = Nothing
                            Dim istrp As CapeOpen.IPersistStreamInit = Nothing
                            Dim o As Object = Nothing
                            Dim datalen As Integer = 0
                            Dim unitLoaded As Boolean = False
                            Try
                                istr = unit
                                o = istr
                            Catch
                                'no IPersistStream, try IPersistStreamInit
                                Try
                                    istrp = unit
                                    o = istrp
                                Catch
                                End Try
                            End Try
                            val = reader.GetAttribute("datalen")
                            If (val IsNot Nothing) Then
                                datalen = Integer.Parse(val, format)
                                If (o IsNot Nothing) Then
                                    Dim data(datalen) As Byte
                                    If Not reader.Read() Then Throw New Exception("Unexpected end of file")
                                    If Not reader.NodeType = Xml.XmlNodeType.Text Then Throw New Exception("Unexpected file content")
                                    reader.ReadContentAsBase64(data, 0, datalen)
                                    Dim s As New Storage(data, datalen)
                                    Try
                                        If (istr IsNot Nothing) Then istr.Load(s) Else istrp.Load(o)
                                    Catch ex As Exception
                                        Throw New Exception("Unit Operation Load failed: " + ErrorFromCOObject(o, ex))
                                    End Try
                                    unitLoaded = True
                                End If
                            End If
                            'initialize the unit
                            If Not unitLoaded Then
                                If istrp IsNot Nothing Then
                                    'call initnew
                                    Try
                                        istrp.InitNew()
                                    Catch ex As Exception
                                        MsgBox("Failed to call InitNew: " + ex.Message, MsgBoxStyle.Exclamation, "InitNew failed:")
                                    End Try
                                End If
                            End If
                            'if ICapeUtilities is implemented, we must call Initialize
                            Dim utils As CapeOpen.ICapeUtilities
                            Try
                                utils = unit
                                'implemented
                                Try
                                    utils.Initialize()
                                Catch ex As Exception
                                    MsgBox("Failed to call Initialize: " + ex.Message, MsgBoxStyle.Exclamation, "Initialize failed:")
                                End Try
                            Catch
                                'not implemented
                            End Try
                        ElseIf reader.Name = "PP" Then
                            If unit Is Nothing Then Throw New Exception("no unit operation data in file")
                            val = reader.GetAttribute("CLSID")
                            If (val Is Nothing) Then Throw New Exception("PP misses CLSID in file")
                            'create the PPM
                            Dim ppm As CapeOpen.ICapeThermoPropertyPackageManager
                            ppClsId = val
                            ppm = Activator.CreateInstance(Type.GetTypeFromCLSID(New System.Guid(ppClsId)))
                            If (ppm Is Nothing) Then Throw New Exception("Failed to create property package manager of the type that was saved")
                            'initialize the ppm
                            InitCAPEOPENObject(ppm)
                            'create PP
                            val = reader.GetAttribute("Name")
                            If (val Is Nothing) Then
                                CleanUpCAPEOPENObject(ppm)
                                Throw New Exception("PP misses Name in file")
                            End If
                            ppName = val
                            Try
                                pp = ppm.GetPropertyPackage(ppName)
                            Catch ex As Exception
                                Dim s As String = "Failed to create property package: " + ErrorFromCOObject(ppm, ex)
                                CleanUpCAPEOPENObject(ppm)
                                Throw New Exception(s)
                            End Try
                            'done with PPM
                            CleanUpCAPEOPENObject(ppm)
                            'init PP
                            InitCAPEOPENObject(pp)
                            'get data from PP
                            If Not dataFromPP() Then Throw New Exception(InitError.Text)
                            'ok to update now
                            SetControls(True)
                            UpdateMaterialsAndGrid()
                        ElseIf reader.Name = "Feed" Then
                            If unit Is Nothing Then Throw New Exception("no unit operation data in file")
                            If pp Is Nothing Then Throw New Exception("no PP data in file")
                            val = reader.GetAttribute("Name")
                            If (val IsNot Nothing) Then
                                'find the feed
                                Dim mat As MaterialObject = Nothing
                                For i = 0 To DataGrid.Columns.Count - 1
                                    Dim mat1 As MaterialObject = DataGrid.Columns(i).Tag
                                    If (mat1.ComponentName = val) Then
                                        mat = mat1
                                        Exit For
                                    End If
                                Next i
                                If (mat IsNot Nothing) Then
                                    val = reader.GetAttribute("T")
                                    If (val Is Nothing) Then mat.T = Double.NaN Else mat.T = Double.Parse(val, format)
                                    val = reader.GetAttribute("P")
                                    If (val Is Nothing) Then mat.P = Double.NaN Else mat.P = Double.Parse(val, format)
                                    val = reader.GetAttribute("flow")
                                    If (val Is Nothing) Then mat.totFlow = Double.NaN Else mat.totFlow = Double.Parse(val, format)
                                    mat.totFlowMassBasis = False
                                    For i = 0 To compIDs.Length - 1
                                        val = reader.GetAttribute("X" + (i + 1).ToString)
                                        If (val Is Nothing) Then mat.overallComposition(i) = Double.NaN Else mat.overallComposition(i) = Double.Parse(val, format)
                                    Next i
                                    'check if complete and flash
                                    CheckCompleteAndFlash(mat)
                                End If
                            End If

                        End If
                End Select
            End While
            'sanity check: 
            If unit Is Nothing Then Throw New Exception("no unit operation data in file")
            If pp Is Nothing Then Throw New Exception("no PP data in file")
            reader.Close()
            'refill list
            SetGridContent()
            'validate
            DoValidate()
        Catch ex As Exception
            'fail
            Reset()
            reader.Close()
            InitError.Text = "Load error: " + ex.Message
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        'browse for msim file
        Dim fdlg As New OpenFileDialog
        fdlg.Filter = "MiniSim files (*.msim)|*.msim|All files (*.*)|*.*"
        If (fdlg.ShowDialog = Windows.Forms.DialogResult.OK) Then
            'load
            fileName = fdlg.FileName
            LoadFile()
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If Not DataGrid.Visible Then
            'cannot save
            MsgBox("Cannot save: current file is not valid", MsgBoxStyle.Critical, "Save")
            Exit Sub
        End If
        If (fileName Is Nothing) Then
            'use save as instead
            SaveAsToolStripMenuItem_Click(sender, e)
            Exit Sub
        End If
        'save with current file name
        SaveFile()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If Not DataGrid.Visible Then
            'cannot save
            MsgBox("Cannot save: current file is not valid", MsgBoxStyle.Critical, "Save")
            Exit Sub
        End If
        'prompt for file
        Dim fdlg As New SaveFileDialog
        fdlg.Filter = "MiniSim files (*.msim)|*.msim|All files (*.*)|*.*"
        If (fdlg.ShowDialog = Windows.Forms.DialogResult.OK) Then
            'file name has been selected
            fileName = fdlg.FileName
            SaveFile()
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Reset()
        NewFile()
    End Sub

    Private Sub ShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowReport.Click
        If (ReportList.SelectedIndex >= 0) Then
            Dim repName As String = ReportList.Items(ReportList.SelectedIndex)
            If repName IsNot Nothing Then
                If repName <> String.Empty Then
                    'obtain report
                    Dim rep As CapeOpen.ICapeUnitReport
                    rep = unit
                    Try
                        rep.selectedReport = repName
                        Dim repText As String = Nothing
                        rep.ProduceReport(repText)
                        If (repText IsNot Nothing) Then
                            'show
                            Dim repDlg As ShowReport
                            repDlg = New ShowReport(repName, repText)
                            repDlg.ShowDialog()
                        End If
                    Catch ex As Exception
                        MsgBox("Failed to get selected report: " + ErrorFromCOObject(rep, ex), MsgBoxStyle.Critical, "Report:")
                    End Try
                End If
            End If
        End If
    End Sub

End Class

