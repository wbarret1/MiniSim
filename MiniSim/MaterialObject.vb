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

<ComClass(MaterialObject.ClassId, MaterialObject.InterfaceId, MaterialObject.EventsId)> _
Public Class MaterialObject
    Inherits CapeOpen.CCapeObject 'implements ICapeIdentification as well as error interfaces
    Implements CapeOpen.ICapeThermoMaterial
    Implements CapeOpen.ICapeThermoCompounds
    Implements CapeOpen.ICapeThermoEquilibriumRoutine
    Implements CapeOpen.ICapeThermoPhases
    Implements CapeOpen.ICapeThermoPropertyRoutine
    Implements CapeOpen.ICapeThermoUniversalConstant

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "939d4e44-510e-4e13-a7bd-e54736974fca"
    Public Const InterfaceId As String = "401bb9f5-4f1f-4597-918b-afd42ae34e1c"
    Public Const EventsId As String = "b933c87e-96da-4bbf-9871-c5d6aeda77a0"
#End Region


    'these are not required for an MO; but deriving from the CCapeObject class makes it required
    Public Overrides Sub Initialize()
    End Sub

    Public Overrides Sub Terminate()
    End Sub

    'data members
    Friend isInlet As Boolean
    Private simulator As MainWnd

    Public Property Inlet() As Boolean
        Get
            Inlet = isInlet
        End Get
        Set(ByVal value As Boolean)
            isInlet = value
        End Set
    End Property

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    'the material will need access to the properties of the simulator, call init after creation
    Public Sub Init(ByVal wnd As MainWnd, ByVal name As String)
        Dim i As Integer
        simulator = wnd
        'init ICapeIdentification members
        ComponentName = name
        ComponentDescription = "MiniSim Material Object"
        'dimension data members
        ReDim overallComposition(simulator.compIDs.Length - 1)
        ReDim phaseComposition(simulator.phaseList.Length - 1, simulator.compIDs.Length - 1)
        ReDim phaseFractions(simulator.phaseList.Length - 1)
        ReDim phaseExists(simulator.phaseList.Length - 1)
        overallProperties = New Hashtable(New CaseInsentiveComparer)
        ReDim phaseProperties(simulator.phaseList.Length - 1)
        For i = 0 To simulator.phaseList.Length - 1
            phaseProperties(i) = New Hashtable(New CaseInsentiveComparer)
        Next i
        ReDim phaseCompositionMassBasis(simulator.phaseList.Length - 1)
        ReDim phaseFractionMassBasis(simulator.phaseList.Length - 1)
        'wipe all
        ClearAllProps()
    End Sub

    'property storage flags
    Enum PropFlags
        StoredInMassBasis = 1
    End Enum

    'class for property storage
    Private Class PropertyStorage
        'we remember the property values and how we stored it
        Friend values() As Double
        Friend flags As PropFlags
    End Class

    'data members related to property storage; we keep special values for the special properties
    ' the simulator has access to all our data (it is our special Friend)
    Friend inEquilibrium As Boolean
    Friend T, P As Double 'this MO supports a single T and P for all phases; even though CAPE-OPEN v1.1 allows to set/get T/P for all phases
    Friend totFlow As Double
    Friend totFlowMassBasis As Boolean
    Friend overallComposition() As Double
    Friend overallCompositionMassBasis As Boolean
    Friend phaseComposition(,) As Double
    Friend phaseCompositionMassBasis() As Boolean
    Friend phaseFractions() As Double
    Friend phaseFractionMassBasis() As Boolean
    Friend phaseExists() As Boolean
    'all other properties go in case insentive hash tables, stored by property name and phase
    Friend overallProperties As Hashtable
    Friend phaseProperties() As Hashtable

    'ICapeThermoMaterial methods

    Public Sub ClearAllProps() Implements CapeOpen.ICapeThermoMaterial.ClearAllProps
        Dim i, j As Integer
        'overall properties
        inEquilibrium = False
        P = Double.NaN
        T = Double.NaN
        totFlow = Double.NaN
        For i = 0 To simulator.compIDs.Length - 1
            overallComposition(i) = Double.NaN
        Next i
        overallProperties.Clear() 'rather than removing all entries, it would likely be more efficient to set the entry content to NaN, as this saves memory allocation
        'phase properties
        For j = 0 To simulator.phaseList.Length - 1
            For i = 0 To simulator.compIDs.Length - 1
                phaseComposition(j, i) = Double.NaN
            Next i
            phaseFractions(j) = Double.NaN
            phaseExists(j) = False
            phaseProperties(j).Clear() 'rather than removing all entries, it would likely be more efficient to set the entry content to NaN, as this saves memory allocation
        Next j
    End Sub

    Public Sub CopyFromMaterial(ByRef source As Object) Implements CapeOpen.ICapeThermoMaterial.CopyFromMaterial
        Dim i As Integer
        'we presume source is a material object implemented by this application
        Dim srcMat As MaterialObject = source
        'copy global data
        inEquilibrium = srcMat.inEquilibrium
        T = srcMat.T
        P = srcMat.P
        totFlow = srcMat.totFlow
        totFlowMassBasis = srcMat.totFlowMassBasis
        'rather than cloning the data, it would likely be more efficient to copy the data
        ' as the memory is already allocated
        overallComposition = srcMat.overallComposition.Clone
        overallCompositionMassBasis = srcMat.overallCompositionMassBasis
        phaseComposition = srcMat.phaseComposition.Clone
        phaseFractions = srcMat.phaseFractions.Clone
        phaseExists = srcMat.phaseExists.Clone
        overallProperties = srcMat.overallProperties.Clone
        For i = 0 To simulator.phaseList.Length - 1
            phaseProperties(i) = srcMat.phaseProperties(i).Clone
        Next i
        'we do not mark this material as inlet
        isInlet = False
    End Sub

    Public Function CreateMaterial() As Object Implements CapeOpen.ICapeThermoMaterial.CreateMaterial
        'create a new material
        Dim newMat As New MaterialObject
        newMat.Init(simulator, "Copy of " + ComponentName) 'we should garantuee a unique name, really
        CreateMaterial = newMat
    End Function

    Public Sub GetOverallProp(ByVal propName As String, ByVal basis As String, ByRef results As Object) Implements CapeOpen.ICapeThermoMaterial.GetOverallProp
        Dim d() As Double
        Dim flow As Double
        Dim i As Integer
        Dim massBasis As Boolean
        'check special properies
        If (SameString(propName, "Temperature")) Then
            'request for overall temperature
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for temperature", "ICapeThermoMaterial", "GetOverallProp")
            If Double.IsNaN(T) Then throwError("ValueError", "Temperature is not set", "ICapeThermoMaterial", "GetOverallProp")
            ReDim d(0)
            d(0) = T
            results = d
        ElseIf (SameString(propName, "Pressure")) Then
            'request for overall pressure
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for pressure", "ICapeThermoMaterial", "GetOverallProp")
            If Double.IsNaN(P) Then throwError("ValueError", "Pressure is not set", "ICapeThermoMaterial", "GetOverallProp")
            ReDim d(0)
            d(0) = P
            results = d
        ElseIf (SameString(propName, "Flow")) Then
            'request for overall component flows
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for flow", "ICapeThermoMaterial", "GetOverallProp")
            End If
            'flow is fraction * totalFlow
            If Double.IsNaN(totFlow) Then throwError("ValueError", "totalFlow is not set", "ICapeThermoMaterial", "GetOverallProp")
            ReDim d(simulator.compIDs.Length - 1)
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(overallComposition(i)) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetOverallProp")
                d(i) = overallComposition(i)
            Next i
            'set flow
            flow = totFlow
            'convert flow and fraction to proper basis
            If (massBasis) Then
                If Not (overallCompositionMassBasis) Then fractionToMass(d)
                If Not (totFlowMassBasis) Then scalarToMassOverall(flow)
            Else
                If (overallCompositionMassBasis) Then fractionToMole(d)
                If (totFlowMassBasis) Then scalarToMoleOverall(flow)
            End If
            For i = 0 To simulator.compIDs.Length - 1
                d(i) *= flow 'are now component flows
            Next
            results = d
        ElseIf (SameString(propName, "totalFlow")) Then
            'request for overall total flow
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for flow", "ICapeThermoMaterial", "GetOverallProp")
            End If
            If Double.IsNaN(totFlow) Then throwError("ValueError", "totalFlow is not set", "ICapeThermoMaterial", "GetOverallProp")
            ReDim d(0)
            d(0) = totFlow
            If (massBasis) Then
                If Not (totFlowMassBasis) Then scalarToMassOverall(d(0))
            Else
                If totFlowMassBasis Then scalarToMoleOverall(d(0))
            End If
            results = d
        ElseIf (SameString(propName, "fraction")) Then
            'request for overall composition
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(overallComposition(i)) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetOverallProp")
            Next i
            d = overallComposition.Clone
            If (massBasis) Then
                If Not (overallCompositionMassBasis) Then fractionToMass(d)
            Else
                If (overallCompositionMassBasis) Then fractionToMole(d)
            End If
            results = d
        ElseIf (SameString(propName, "phaseFraction")) Then
            'request for overall phase fraction
            throwError("InvalidProperty", "phaseFraction is not a valid overall property", "ICapeThermoMaterial", "GetOverallProp")
        Else
            'not a special property, look up
            Dim err As String = Nothing
            Dim info As propTable.PropInfo = simulator.propertyTable.GetPropInfo(propName, err)
            If (info Is Nothing) Then
                'property look up failed
                throwError("InvalidProperty", err, "ICapeThermoMaterial", "GetOverallProp")
            End If
            'check basis
            If (info.BasisOrder <> 0) Then
                If (SameString(basis, "mole")) Then
                    massBasis = False
                ElseIf (SameString(basis, "mass")) Then
                    massBasis = True
                Else
                    throwError("BasisError", "Expected mole or mass basis for " + propName, "ICapeThermoMaterial", "GetOverallProp")
                End If
            Else
                If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "GetOverallProp")
            End If
            'get value
            Dim values As PropertyStorage
            values = overallProperties.Item(propName)
            If (values Is Nothing) Then throwError("ValueError", "Values not set for " + propName, "ICapeThermoMaterial", "GetOverallProp")
            'as we do not re-use storage, but toss it instead, we do not have to check whether values are absent
            d = values.values.Clone
            Dim propConvOrder = info.BasisOrder
            If (propConvOrder <> 0) Then
                'check mass basis conversion
                If (massBasis) Then
                    If (values.flags And PropFlags.StoredInMassBasis) Then propConvOrder = 0 'no conversion required
                Else
                    If ((values.flags And PropFlags.StoredInMassBasis) = 0) Then
                        propConvOrder = 0 'no conversion required
                    Else
                        'conversion from mass to mole, opposite direction 
                        propConvOrder = -propConvOrder
                    End If
                End If
            End If
            If (propConvOrder <> 0) Then
                overallMoleToMassConversion(d, propConvOrder, info.Derivative)
            End If
            results = d
        End If
    End Sub

    Public Sub GetOverallTPFraction(ByRef temperature As Double, ByRef pressure As Double, ByRef composition As Object) Implements CapeOpen.ICapeThermoMaterial.GetOverallTPFraction
        Dim i As Integer
        'check values
        If Double.IsNaN(T) Then throwError("ValueError", "Temperature is not set", "ICapeThermoMaterial", "GetOverallTPFraction")
        If Double.IsNaN(P) Then throwError("ValueError", "Pressure is not set", "ICapeThermoMaterial", "GetOverallTPFraction")
        'fraction is mole basis
        Dim d() As Double
        d = overallComposition.Clone
        If (overallCompositionMassBasis) Then
            'convert
            fractionToMole(d)
        Else
            'check and use
            For i = 0 To simulator.compIDs.Length - 1
                If (Double.IsNaN(d(i))) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetOverallTPFraction")
            Next i
        End If
        'all ok
        temperature = T
        pressure = P
        composition = d
    End Sub

    Public Sub GetPresentPhases(ByRef phaseLabels As Object, ByRef phaseStatus As Object) Implements CapeOpen.ICapeThermoMaterial.GetPresentPhases
        'get present phases and their status
        Dim i, count As Integer
        count = 0
        For i = 0 To simulator.phaseList.Length - 1
            If (phaseExists(i)) Then count += 1
        Next i
        Dim phases(count - 1) As String
        count = 0
        For i = 0 To simulator.phaseList.Length - 1
            If (phaseExists(i)) Then
                phases(count) = simulator.phaseList(i)
                count += 1
            End If
        Next i
        phaseLabels = phases
        Dim status(count - 1) As Integer
        If (inEquilibrium) Then
            For i = 0 To count - 1
                status(i) = CapeOpen.CapePhaseStatus.CAPE_ATEQUILIBRIUM
            Next i
        Else
            For i = 0 To count - 1
                status(i) = CapeOpen.CapePhaseStatus.CAPE_UNKNOWNPHASESTATUS
            Next i
        End If
        phaseStatus = status
    End Sub

    Public Sub GetSinglePhaseProp(ByVal propName As String, ByVal phaseLabel As String, ByVal basis As String, ByRef results As Object) Implements CapeOpen.ICapeThermoMaterial.GetSinglePhaseProp
        Dim d() As Double
        Dim flow, factor As Double
        Dim i As Integer
        Dim massBasis As Boolean
        Dim phaseIndex As Integer
        'get the phase index
        phaseIndex = getPhaseIndex(phaseLabel)
        'check if phase is present
        If Not phaseExists(phaseIndex) Then throwError("PhaseNotPresent", "Phase """ + phaseLabel + """ is not present", "ICapeThermoMaterial", "GetSinglePhaseProp")
        'check special properies
        If (SameString(propName, "Temperature")) Then
            'request for temperature
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for temperature", "ICapeThermoMaterial", "GetSinglePhaseProp")
            If Double.IsNaN(T) Then throwError("ValueError", "Temperature is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            ReDim d(0)
            d(0) = T
            results = d
        ElseIf (SameString(propName, "Pressure")) Then
            'request for pressure
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for pressure", "ICapeThermoMaterial", "GetSinglePhaseProp")
            If Double.IsNaN(P) Then throwError("ValueError", "Pressure is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            ReDim d(0)
            d(0) = P
            results = d
        ElseIf (SameString(propName, "Flow")) Then
            'request for component flows
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for flow", "ICapeThermoMaterial", "GetSinglePhaseProp")
            End If
            'flow is fraction * totalFlow * phaseFraction
            If Double.IsNaN(phaseFractions(phaseIndex)) Then throwError("ValueError", "phaseFraction is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            If Double.IsNaN(totFlow) Then throwError("ValueError", "totalFlow is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            ReDim d(simulator.compIDs.Length - 1)
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(phaseComposition(phaseIndex, i)) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
                d(i) = phaseComposition(phaseIndex, i)
            Next i
            'set flow
            flow = totFlow
            'set phase fraction factor
            factor = phaseFractions(phaseIndex)
            'convert flow and fraction to proper basis
            If (massBasis) Then
                If Not (phaseCompositionMassBasis(phaseIndex)) Then fractionToMass(d)
                If Not (totFlowMassBasis) Then scalarToMassPhase(phaseIndex, flow)
                If Not (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMass(phaseIndex, factor)
            Else
                If (phaseCompositionMassBasis(phaseIndex)) Then fractionToMole(d)
                If (totFlowMassBasis) Then scalarToMolePhase(phaseIndex, flow)
                If (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMole(phaseIndex, factor)
            End If
            factor *= flow
            For i = 0 To simulator.compIDs.Length - 1
                d(i) *= factor 'are now component flows for the given phase
            Next
            results = d
        ElseIf (SameString(propName, "totalFlow")) Then
            'request for total flow of phase
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for flow", "ICapeThermoMaterial", "GetSinglePhaseProp")
            End If
            If Double.IsNaN(totFlow) Then throwError("ValueError", "totalFlow is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            If Double.IsNaN(phaseFractions(phaseIndex)) Then throwError("ValueError", "phaseFraction is not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
            ReDim d(0)
            d(0) = totFlow
            factor = phaseFractions(phaseIndex)
            If (massBasis) Then
                If Not (totFlowMassBasis) Then scalarToMassPhase(phaseIndex, d(0))
                If Not (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMass(phaseIndex, factor)
            Else
                If totFlowMassBasis Then scalarToMolePhase(phaseIndex, d(0))
                If (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMole(phaseIndex, factor)
            End If
            d(0) *= factor
            results = d
        ElseIf (SameString(propName, "fraction")) Then
            'request for composition
            ReDim d(simulator.compIDs.Length - 1)
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(phaseComposition(phaseIndex, i)) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetSinglePhaseProp")
                d(i) = phaseComposition(phaseIndex, i)
            Next i
            If (massBasis) Then
                If Not (phaseCompositionMassBasis(phaseIndex)) Then fractionToMass(d)
            Else
                If (phaseCompositionMassBasis(phaseIndex)) Then fractionToMole(d)
            End If
            results = d
        ElseIf (SameString(propName, "phaseFraction")) Then
            'request for phase fraction
            ReDim d(0)
            d(0) = phaseFractions(phaseIndex)
            If (massBasis) Then
                If Not (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMass(phaseIndex, d(0))
            Else
                If (phaseFractionMassBasis(phaseIndex)) Then phaseFractionToMole(phaseIndex, d(0))
            End If
            results = d
        Else
            'not a special property, look up
            Dim err As String = Nothing
            Dim info As propTable.PropInfo = simulator.propertyTable.GetPropInfo(propName, err)
            If (info Is Nothing) Then
                'property look up failed
                throwError("InvalidProperty", err, "ICapeThermoMaterial", "GetSinglePhaseProp")
            End If
            'check basis
            If (info.BasisOrder <> 0) Then
                If (SameString(basis, "mole")) Then
                    massBasis = False
                ElseIf (SameString(basis, "mass")) Then
                    massBasis = True
                Else
                    throwError("BasisError", "Expected mole or mass basis for " + propName, "ICapeThermoMaterial", "GetSinglePhaseProp")
                End If
            Else
                If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "GetSinglePhaseProp")
            End If
            'get value
            Dim values As PropertyStorage
            values = phaseProperties(phaseIndex).Item(propName)
            If (values Is Nothing) Then throwError("ValueError", "Values not set for " + propName, "ICapeThermoMaterial", "GetSinglePhaseProp")
            'as we do not re-use storage, but toss it instead, we do not have to check whether values are absent
            d = values.values.Clone
            Dim propConvOrder = info.BasisOrder
            If (propConvOrder <> 0) Then
                'check mass basis conversion
                If (massBasis) Then
                    If (values.flags And PropFlags.StoredInMassBasis) Then propConvOrder = 0 'no conversion required
                Else
                    If ((values.flags And PropFlags.StoredInMassBasis) = 0) Then
                        propConvOrder = 0 'no conversion required
                    Else
                        'conversion from mass to mole, opposite direction 
                        propConvOrder = -propConvOrder
                    End If
                End If
            End If
            If (propConvOrder <> 0) Then
                phaseMoleToMassConversion(phaseIndex, d, propConvOrder, info.Derivative)
            End If
            results = d
        End If
    End Sub

    Public Sub GetTPFraction(ByVal phaseLabel As String, ByRef temperature As Double, ByRef pressure As Double, ByRef composition As Object) Implements CapeOpen.ICapeThermoMaterial.GetTPFraction
        Dim i As Integer, phaseIndex As Integer
        'check values
        If Double.IsNaN(T) Then throwError("ValueError", "Temperature is not set", "ICapeThermoMaterial", "GetTPFraction")
        If Double.IsNaN(P) Then throwError("ValueError", "Pressure is not set", "ICapeThermoMaterial", "GetTPFraction")
        'get the phase index
        phaseIndex = getPhaseIndex(phaseLabel)
        'check if phase is present
        If Not phaseExists(phaseIndex) Then throwError("PhaseNotPresent", "Phase """ + phaseLabel + """ is not present", "ICapeThermoMaterial", "GetTPFraction")
        'fraction is mole basis
        Dim d(simulator.compIDs.Length - 1) As Double
        For i = 0 To simulator.compIDs.Length - 1
            d(i) = phaseComposition(phaseIndex, i)
        Next i
        If (phaseCompositionMassBasis(phaseIndex)) Then
            'convert
            fractionToMole(d)
        Else
            'check and use
            For i = 0 To simulator.compIDs.Length - 1
                If (Double.IsNaN(d(i))) Then throwError("ValueError", "one or more compositions not set", "ICapeThermoMaterial", "GetTPFraction")
            Next i
        End If
        'all ok
        temperature = T
        pressure = P
        composition = d
    End Sub

    Public Sub GetTwoPhaseProp(ByVal propName As String, ByVal phaseLabels As Object, ByVal basis As String, ByRef results As Object) Implements CapeOpen.ICapeThermoMaterial.GetTwoPhaseProp
        'get phases
        Dim phases() As String = Nothing
        Dim phaseIndex1, phaseIndex2 As Integer
        Try
            phases = phaseLabels
        Catch ex As Exception
            throwError("PhaseLabelError", "Invalid object for phase labels: string array expected", "ICapeThermoMaterial", "GetTwoPhaseProp", ex)
        End Try
        If (phases.Length <> 2) Then throwError("PhaseLabelError", "Invalid object for phase labels: expected two phases", "ICapeThermoMaterial", "GetTwoPhaseProp")
        'identify phases
        phaseIndex1 = getPhaseIndex(phases(0))
        phaseIndex2 = getPhaseIndex(phases(1))
        If (phaseIndex1 = phaseIndex2) Then throwError("PhaseLabelError", "Invalid object for phase labels: phases cannot be the same", "ICapeThermoMaterial", "GetTwoPhaseProp")
        'check basis
        If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "GetTwoPhaseProp")
        'find property; to make it unique in the hash table we store it under phase 1, with name phaseLabel2:propname
        Dim values As PropertyStorage
        values = phaseProperties(phaseIndex1).Item(phases(1) + ":" + propName)
        If (values Is Nothing) Then throwError("ValueError", "Values not set for " + propName, "ICapeThermoMaterial", "GetTwoPhaseProp")
        'as we do not re-use storage, but toss it instead, we do not have to check whether values are absent
        results = values.values.Clone
    End Sub

    Public Sub SetOverallProp(ByVal propName As String, ByVal basis As String, ByVal values As Object) Implements CapeOpen.ICapeThermoMaterial.SetOverallProp
        Dim d() As Double
        Dim massBasis As Boolean
        Dim i As Integer
        'check special properies
        If (SameString(propName, "Temperature")) Then
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for temperature", "ICapeThermoMaterial", "SetOverallProp")
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for temperature, scalar expected", "ICapeThermoMaterial", "SetOverallProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetOverallProp")
            T = d(0)
        ElseIf (SameString(propName, "Pressure")) Then
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for pressure", "ICapeThermoMaterial", "SetOverallProp")
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for pressure, scalar expected", "ICapeThermoMaterial", "SetOverallProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetOverallProp")
            P = d(0)
        ElseIf (SameString(propName, "Flow")) Then
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for flow", "ICapeThermoMaterial", "SetOverallProp")
            End If
            d = values
            If (d.Length <> simulator.compIDs.Length) Then throwError("InvalidNumberOfValues", "Invalid number of values for flow, expected one value for each compound", "ICapeThermoMaterial", "SetOverallProp")
            For i = 0 To d.Length - 1
                If (Double.IsNaN(d(i))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetOverallProp")
            Next i
            'we set totFlow and overallComposition
            totFlow = 0
            For i = 0 To d.Length - 1
                totFlow += d(i)
            Next i
            totFlowMassBasis = massBasis
            If (totFlow = 0) Then
                'unknown composition
                For i = 0 To d.Length - 1
                    overallComposition(i) = Double.NaN
                Next i
            Else
                'set composition
                For i = 0 To d.Length - 1
                    overallComposition(i) = d(i) / totFlow
                Next i
                overallCompositionMassBasis = massBasis
            End If
        ElseIf (SameString(propName, "totalFlow")) Then
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for totalFlow", "ICapeThermoMaterial", "SetOverallProp")
            End If
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for totalFlow, scalar expected", "ICapeThermoMaterial", "SetOverallProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetOverallProp")
            totFlow = d(0)
            totFlowMassBasis = massBasis
        ElseIf (SameString(propName, "fraction")) Then
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for composition", "ICapeThermoMaterial", "SetOverallProp")
            End If
            d = values
            If (d.Length <> simulator.compIDs.Length) Then throwError("InvalidNumberOfValues", "Invalid number of values for cpmposition, expected one value for each compound", "ICapeThermoMaterial", "SetOverallProp")
            For i = 0 To d.Length - 1
                If (Double.IsNaN(d(i))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetOverallProp")
            Next i
            overallComposition = d
            overallCompositionMassBasis = massBasis
        ElseIf (SameString(propName, "phaseFraction")) Then
            throwError("InvalidProperty", "phaseFraction is not a valid property for the overall phase", "ICapeThermoMaterial", "SetOverallProp")
        Else
            'not a special property
            Dim err As String = Nothing
            Dim info As propTable.PropInfo = simulator.propertyTable.GetPropInfo(propName, err)
            If (info Is Nothing) Then
                'property look up failed
                throwError("InvalidProperty", err, "ICapeThermoMaterial", "SetOverallProp")
            End If
            'check basis
            massBasis = False
            If (info.BasisOrder <> 0) Then
                If (SameString(basis, "mole")) Then
                    '... see above
                ElseIf (SameString(basis, "mass")) Then
                    massBasis = True
                Else
                    throwError("BasisError", "Expected mole or mass basis for " + propName, "ICapeThermoMaterial", "SetOverallProp")
                End If
            Else
                If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "SetOverallProp")
            End If
            'get value
            Dim storage As PropertyStorage
            storage = overallProperties.Item(propName)
            If (storage Is Nothing) Then
                'create storage
                storage = New PropertyStorage
                overallProperties.Item(propName) = storage
            End If
            'as we do not re-use storage, but toss it instead, we do not have to check whether values are absent
            d = values
            If (d.Length <> info.DataCount) Then throwError("InvalidNumberOfValues", "Invalid number of values for property " + propName, "ICapeThermoMaterial", "SetOverallProp")
            storage.values = d
            storage.flags = 0
            If (massBasis) Then storage.flags = storage.flags Or PropFlags.StoredInMassBasis
        End If
        'we can assume that at this point we are not in equilibrium 
        inEquilibrium = False
    End Sub

    Public Sub SetPresentPhases(ByVal phaseLabels As Object, ByVal phaseStatus As Object) Implements CapeOpen.ICapeThermoMaterial.SetPresentPhases
        'we ignore phase status in this MO; it would be nice to keep track of the phase status in case it serves as an initial
        ' guess for equilibrium calculations performed by the PP
        Dim i As Integer
        Dim labs() As String
        labs = phaseLabels
        'set all phases as not present
        For i = 0 To simulator.phaseList.Length - 1
            phaseExists(i) = False
        Next i
        'set the ones present that are specified
        For i = 0 To labs.Length - 1
            Dim index As Integer = getPhaseIndex(labs(i))
            phaseExists(index) = True
        Next i
        inEquilibrium = False
    End Sub

    Public Sub SetSinglePhaseProp(ByVal propName As String, ByVal phaseLabel As String, ByVal basis As String, ByVal values As Object) Implements CapeOpen.ICapeThermoMaterial.SetSinglePhaseProp
        Dim d() As Double
        Dim massBasis As Boolean
        Dim i, phaseIndex As Integer
        'get phase index
        phaseIndex = getPhaseIndex(phaseLabel)
        'check special properies
        If (SameString(propName, "Temperature")) Then
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for temperature", "ICapeThermoMaterial", "SetSinglePhaseProp")
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for temperature, scalar expected", "ICapeThermoMaterial", "SetSinglePhaseProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetSinglePhaseProp")
            T = d(0)
        ElseIf (SameString(propName, "Pressure")) Then
            If basis <> String.Empty Then throwError("BasisError", "Expected no basis for pressure", "ICapeThermoMaterial", "SetSinglePhaseProp")
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for pressure, scalar expected", "ICapeThermoMaterial", "SetSinglePhaseProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetSinglePhaseProp")
            P = d(0)
        ElseIf (SameString(propName, "Flow")) Then
            'setting component flows for individual phases is not supported by this MO
            ' the MO stored total flow on overall basis; the conversion can only be done if all data is available at the same time
            throwError("NotImplemented", "Setting compound flows for individual phases is not supported by this material object", "ICapeThermoMaterial", "SetSinglePhaseProp")
        ElseIf (SameString(propName, "totalFlow")) Then
            'setting flows for individual phases is not supported by this MO
            ' the MO stored total flow on overall basis; the conversion can only be done if all data is available at the same time
            throwError("NotImplemented", "Setting flows for individual phases is not supported by this material object", "ICapeThermoMaterial", "SetSinglePhaseProp")
        ElseIf (SameString(propName, "fraction")) Then
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for composition", "ICapeThermoMaterial", "SetSinglePhaseProp")
            End If
            d = values
            If (d.Length <> simulator.compIDs.Length) Then throwError("InvalidNumberOfValues", "Invalid number of values for cpmposition, expected one value for each compound", "ICapeThermoMaterial", "SetSinglePhaseProp")
            For i = 0 To d.Length - 1
                If (Double.IsNaN(d(i))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetSinglePhaseProp")
                phaseComposition(phaseIndex, i) = d(i)
            Next i
            phaseCompositionMassBasis(phaseIndex) = massBasis
        ElseIf (SameString(propName, "phaseFraction")) Then
            If (SameString(basis, "mole")) Then
                massBasis = False
            ElseIf (SameString(basis, "mass")) Then
                massBasis = True
            Else
                throwError("BasisError", "Expected mole or mass basis for phaseFraction", "ICapeThermoMaterial", "SetSinglePhaseProp")
            End If
            d = values
            If (d.Length <> 1) Then throwError("InvalidNumberOfValues", "Invalid number of values for temperature, scalar expected", "ICapeThermoMaterial", "SetSinglePhaseProp")
            If (Double.IsNaN(d(0))) Then throwError("InvalidValues", "NaN values present in data", "ICapeThermoMaterial", "SetSinglePhaseProp")
            phaseFractions(phaseIndex) = d(0)
            phaseFractionMassBasis(phaseIndex) = massBasis
        Else
            'not a special property
            Dim err As String = Nothing
            Dim info As propTable.PropInfo = simulator.propertyTable.GetPropInfo(propName, err)
            If (info Is Nothing) Then
                'property look up failed
                throwError("InvalidProperty", err, "ICapeThermoMaterial", "SetSinglePhaseProp")
            End If
            'check basis
            massBasis = False
            If (info.BasisOrder <> 0) Then
                If (SameString(basis, "mole")) Then
                    '... see above
                ElseIf (SameString(basis, "mass")) Then
                    massBasis = True
                Else
                    throwError("BasisError", "Expected mole or mass basis for " + propName, "ICapeThermoMaterial", "SetSinglePhaseProp")
                End If
            Else
                If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "SetSinglePhaseProp")
            End If
            'get value
            Dim storage As PropertyStorage
            storage = phaseProperties(phaseIndex).Item(propName)
            If (storage Is Nothing) Then
                'create storage
                storage = New PropertyStorage
                phaseProperties(phaseIndex).Item(propName) = storage
            End If
            'as we do not re-use storage, but toss it instead, we do not have to check whether values are absent
            d = values
            If (d.Length <> info.DataCount) Then throwError("InvalidNumberOfValues", "Invalid number of values for property " + propName, "ICapeThermoMaterial", "SetSinglePhaseProp")
            storage.values = d
            storage.flags = 0
            If (massBasis) Then storage.flags = storage.flags Or PropFlags.StoredInMassBasis
        End If
        'we can assume that at this point we are not in equilibrium 
        inEquilibrium = False
    End Sub

    Public Sub SetTwoPhaseProp(ByVal propName As String, ByVal phaseLabels As Object, ByVal basis As String, ByVal values As Object) Implements CapeOpen.ICapeThermoMaterial.SetTwoPhaseProp
        'get phases
        Dim phases() As String = Nothing
        Dim phaseIndex1, phaseIndex2 As Integer
        Try
            phases = phaseLabels
        Catch ex As Exception
            throwError("PhaseLabelError", "Invalid object for phase labels: string array expected", "ICapeThermoMaterial", "SetTwoPhaseProp", ex)
        End Try
        If (phases.Length <> 2) Then throwError("PhaseLabelError", "Invalid object for phase labels: expected two phases", "ICapeThermoMaterial", "SetTwoPhaseProp")
        'identify phases
        phaseIndex1 = getPhaseIndex(phases(0))
        phaseIndex2 = getPhaseIndex(phases(1))
        If (phaseIndex1 = phaseIndex2) Then throwError("PhaseLabelError", "Invalid object for phase labels: phases cannot be the same", "ICapeThermoMaterial", "SetTwoPhaseProp")
        'check basis
        If basis <> String.Empty Then throwError("BasisError", "Expected no basis for " + propName, "ICapeThermoMaterial", "SetTwoPhaseProp")
        'check value count; surfaceTension is scalar, we presume all others to be vector
        ' composition derivative are order higher, with 2 sets of values
        Dim expected As Integer
        Dim index As Integer
        Dim pName As String, derivName As String = Nothing
        Dim isCompositionDerivative As Boolean = False
        index = InStr(propName, ".")
        If (index = 0) Then
            'no derivative
            pName = propName
        Else
            pName = Left(propName, index - 1)
            derivName = Mid(propName, index)
            'check deriv name
            If ((SameString(derivName, "dmolfraction")) Or (SameString(derivName, "dmoles"))) Then
                isCompositionDerivative = True
            ElseIf Not ((SameString(derivName, "dtemperature")) Or (SameString(derivName, "dpressure"))) Then
                'invalid derivative
                throwError("PropertyError", "Invalid or unknown derivative: " + derivName, "ICapeThermoMaterial", "SetTwoPhaseProp")
            End If
        End If
        If (SameString(pName, "surfaceTension")) Then expected = 1 Else expected = simulator.compIDs.Length
        If (isCompositionDerivative) Then expected *= simulator.compIDs.Length * 2
        'check values
        Dim vals() As Double = Nothing
        Try
            vals = values
        Catch ex As Exception
            throwError("ValueError", "Expected double array for values: " + ex.Message, "ICapeThermoMaterial", "SetTwoPhaseProp", ex)
        End Try
        If (vals.Length <> expected) Then throwError("ValueError", "Unexpected number of values", "ICapeThermoMaterial", "SetTwoPhaseProp")
        'store
        Dim storage As PropertyStorage
        Dim key As String = phases(1) + ":" + propName
        storage = phaseProperties(phaseIndex1).Item(key)
        If (storage Is Nothing) Then
            storage = New PropertyStorage
            phaseProperties(phaseIndex1).Item(key) = storage
        End If
        storage.flags = 0
        storage.values = vals
    End Sub

    'ICapeThermoCompounds methods

    Public Function GetCompoundConstant(ByVal props As Object, ByVal compIds As Object) As Object Implements CapeOpen.ICapeThermoCompounds.GetCompoundConstant
        'we pass this call on to the PP
        ' mind that if the compIDs argument is empty, an MO that uses a sub-set of compounds should pass a list
        Try
            GetCompoundConstant = simulator.ppCompounds.GetCompoundConstant(props, compIds)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetCompoundConstant of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetCompoundConstant", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Sub GetCompoundList(ByRef compIds As Object, ByRef formulae As Object, ByRef names As Object, ByRef boilTemps As Object, ByRef molwts As Object, ByRef casnos As Object) Implements CapeOpen.ICapeThermoCompounds.GetCompoundList
        'we do not pass this call on to the PP, as the PP will be asking us about which compounds we are using
        ' (which in our case is the full list returned by the PP)
        compIds = simulator.compIDs
        formulae = simulator.formulae
        names = simulator.names
        boilTemps = simulator.boilTemps
        molwts = simulator.molWeights
        casnos = simulator.casNos
    End Sub

    Public Function GetConstPropList() As Object Implements CapeOpen.ICapeThermoCompounds.GetConstPropList
        'we pass this call on to the PP
        Try
            GetConstPropList = simulator.ppCompounds.GetConstPropList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetConstPropList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetConstPropList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Function GetNumCompounds() As Integer Implements CapeOpen.ICapeThermoCompounds.GetNumCompounds
        'we do not pass this call on; the PP may be asking us
        Return simulator.compIDs.Length
    End Function

    Public Sub GetPDependentProperty(ByVal props As Object, ByVal pressure As Double, ByVal compIds As Object, ByRef propVals As Object) Implements CapeOpen.ICapeThermoCompounds.GetPDependentProperty
        'we pass this call on to the PP
        ' mind that if the compIDs argument is empty, an MO that uses a sub-set of compounds should pass a list
        Try
            simulator.ppCompounds.GetPDependentProperty(props, pressure, compIds, propVals)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetPDependentProperty of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetPDependentProperty", ex)
        End Try
    End Sub

    Public Function GetPDependentPropList() As Object Implements CapeOpen.ICapeThermoCompounds.GetPDependentPropList
        'we pass this call on to the PP
        Try
            GetPDependentPropList = simulator.ppCompounds.GetPDependentPropList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetPDependentPropList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetPDependentPropList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Sub GetTDependentProperty(ByVal props As Object, ByVal temperature As Double, ByVal compIds As Object, ByRef propVals As Object) Implements CapeOpen.ICapeThermoCompounds.GetTDependentProperty
        'we pass this call on to the PP
        ' mind that if the compIDs argument is empty, an MO that uses a sub-set of compounds should pass a list
        Try
            simulator.ppCompounds.GetTDependentProperty(props, temperature, compIds, propVals)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetPDependentProperty of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetTDependentProperty", ex)
        End Try
    End Sub

    Public Function GetTDependentPropList() As Object Implements CapeOpen.ICapeThermoCompounds.GetTDependentPropList
        'we pass this call on to the PP
        Try
            GetTDependentPropList = simulator.ppCompounds.GetPDependentPropList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetTDependentPropList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCompounds, ex), "ICapeThermoCompounds", "GetTDependentPropList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    'ICapeThermoEquilibriumRoutine methods

    Public Sub CalcEquilibrium(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) Implements CapeOpen.ICapeThermoEquilibriumRoutine.CalcEquilibrium
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppEqRoutine.CalcEquilibrium(specification1, specification2, solutionType)
            'we are now in equilibrium
            inEquilibrium = True
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CalcEquilibrium of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppEqRoutine, ex), "ICapeThermoEquilibriumRoutine", "CalcEquilibrium", ex)
        End Try
    End Sub

    Public Function CheckEquilibriumSpec(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) As Boolean Implements CapeOpen.ICapeThermoEquilibriumRoutine.CheckEquilibriumSpec
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppEqRoutine.CheckEquilibriumSpec(specification1, specification2, solutionType)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CheckEquilibriumSpec of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppEqRoutine, ex), "ICapeThermoEquilibriumRoutine", "CheckEquilibriumSpec", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    'ICapeThermoPhases methods

    Public Function GetNumPhases() As Integer Implements CapeOpen.ICapeThermoPhases.GetNumPhases
        'we have this info, lets not forward to PP
        GetNumPhases = simulator.phaseList.Length
    End Function

    Public Function GetPhaseInfo(ByVal phaseLabel As String, ByVal phaseAttribute As String) As Object Implements CapeOpen.ICapeThermoPhases.GetPhaseInfo
        'forward the function
        Try
            GetPhaseInfo = simulator.ppPhases.GetPhaseInfo(phaseLabel, phaseAttribute)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetPhaseInfo of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppPhases, ex), "ICapeThermoPhases", "GetPhaseInfo", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Sub GetPhaseList(ByRef phaseLabels As Object, ByRef stateOfAggregation As Object, ByRef keyCompoundId As Object) Implements CapeOpen.ICapeThermoPhases.GetPhaseList
        'we have this info, lets not forward to PP
        phaseLabels = simulator.phaseList
        stateOfAggregation = simulator.aggStates
        keyCompoundId = simulator.keyComps
    End Sub

    'ICapeThermoPropertyRoutine methods

    Public Sub CalcAndGetLnPhi(ByVal phaseLabel As String, ByVal temperature As Double, ByVal pressure As Double, ByVal moleNumbers As Object, ByVal fFlags As Integer, ByRef lnPhi As Object, ByRef lnPhiDT As Object, ByRef lnPhiDP As Object, ByRef lnPhiDn As Object) Implements CapeOpen.ICapeThermoPropertyRoutine.CalcAndGetLnPhi
        'forward the function
        Try
            simulator.ppCalcRoutine.CalcAndGetLnPhi(phaseLabel, temperature, pressure, moleNumbers, fFlags, lnPhi, lnPhiDT, lnPhiDP, lnPhiDn)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CalcAndGetLnPhi of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "CalcAndGetLnPhi", ex)
        End Try
    End Sub

    Public Sub CalcSinglePhaseProp(ByVal props As Object, ByVal phaseLabel As String) Implements CapeOpen.ICapeThermoPropertyRoutine.CalcSinglePhaseProp
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppCalcRoutine.CalcSinglePhaseProp(props, phaseLabel)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CalcSinglePhaseProp of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "CalcSinglePhaseProp", ex)
        End Try
    End Sub

    Public Sub CalcTwoPhaseProp(ByVal props As Object, ByVal phaseLabels As Object) Implements CapeOpen.ICapeThermoPropertyRoutine.CalcTwoPhaseProp
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppCalcRoutine.CalcTwoPhaseProp(props, phaseLabels)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CalcTwoPhaseProp of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "CalcTwoPhaseProp", ex)
        End Try
    End Sub

    Public Function CheckSinglePhasePropSpec(ByVal propName As String, ByVal phaseLabel As String) As Boolean Implements CapeOpen.ICapeThermoPropertyRoutine.CheckSinglePhasePropSpec
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppCalcRoutine.CheckSinglePhasePropSpec(propName, phaseLabel)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CheckSinglePhasePropSpec of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "CheckSinglePhasePropSpec", ex)
        End Try
    End Function

    Public Function CheckTwoPhasePropSpec(ByVal propName As String, ByVal phaseLabels As Object) As Boolean Implements CapeOpen.ICapeThermoPropertyRoutine.CheckTwoPhasePropSpec
        'set the current MO
        SetCurrentMO()
        'forward the function
        Try
            simulator.ppCalcRoutine.CheckTwoPhasePropSpec(propName, phaseLabels)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "CheckTwoPhasePropSpec of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "CheckTwoPhasePropSpec", ex)
        End Try
    End Function

    Public Function GetSinglePhasePropList() As Object Implements CapeOpen.ICapeThermoPropertyRoutine.GetSinglePhasePropList
        'set the current MO
        SetCurrentMO() 'although we do not think anybody will use it
        'forward the function
        Try
            GetSinglePhasePropList = simulator.ppCalcRoutine.GetSinglePhasePropList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetSinglePhasePropList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "GetSinglePhasePropList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Function GetTwoPhasePropList() As Object Implements CapeOpen.ICapeThermoPropertyRoutine.GetTwoPhasePropList
        'set the current MO
        SetCurrentMO() 'although we do not think anybody will use it
        'forward the function
        Try
            GetTwoPhasePropList = simulator.ppCalcRoutine.GetTwoPhasePropList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetTwoPhasePropList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppCalcRoutine, ex), "ICapeThermoPropertyRoutine", "GetTwoPhasePropList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    'ICapeThermoUniversalConstant methods

    Public Function GetUniversalConstant(ByVal constantId As String) As Object Implements CapeOpen.ICapeThermoUniversalConstant.GetUniversalConstant
        If (simulator.ppUniConst Is Nothing) Then throwError("NotImplemented", "Property package does not implement universal constants", "ICapeThermoUniversalConstant", "GetUniversalConstantList")
        'forward the function
        Try
            GetUniversalConstant = simulator.ppUniConst.GetUniversalConstant(constantId)
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetUniversalConstant of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppUniConst, ex), "ICapeThermoPropertyRoutine", "GetUniversalConstant", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    Public Function GetUniversalConstantList() As Object Implements CapeOpen.ICapeThermoUniversalConstant.GetUniversalConstantList
        If (simulator.ppUniConst Is Nothing) Then throwError("NotImplemented", "Property package does not implement universal constants", "ICapeThermoUniversalConstant", "GetUniversalConstantList")
        'forward the function
        Try
            GetUniversalConstantList = simulator.ppUniConst.GetUniversalConstantList()
        Catch ex As Exception
            'get the error
            throwError("PPFailure", "GetUniversalConstantList of PP failed: " + MainWnd.ErrorFromCOObject(simulator.ppUniConst, ex), "ICapeThermoPropertyRoutine", "GetUniversalConstantList", ex)
            Throw New Exception("Code should not reach this point")
        End Try
    End Function

    'exception implementation 

    <ComClass(UserException.ClassId, UserException.InterfaceId, UserException.EventsId)> _
Class UserException
        Inherits System.ApplicationException
        Implements CapeOpen.ECapeUser
        Implements CapeOpen.ECapeRoot
        Implements CapeOpen.ECapeUnknown

#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "886320AC-D422-4f1c-B40C-740BBF52EB8A"
        Public Const InterfaceId As String = "400B3A7B-D5B3-4cb5-B148-ED59256E4709"
        Public Const EventsId As String = "DDF26305-63CB-4256-81F6-4AAD19F5A23E"
#End Region

        Private _name, _desc, _iname, _scope As String
        Private _exception As System.Exception

        Sub New(ByVal name As String, ByVal desc As String, ByVal iname As String, ByVal scope As String, ByVal HResult As Int32)
            MyBase.New(desc)
            MyBase.HResult = HResult
            _name = name
            _desc = desc
            _iname = iname
            _scope = scope
            _exception = Nothing
        End Sub
        Sub New(ByVal name As String, ByVal desc As String, ByVal iname As String, ByVal scope As String, ByVal exception As System.Exception)
            MyBase.New(desc, exception)
            MyBase.HResult = CapeOpen.CapeErrorInterfaceHR.ECapeUnknownHR
            Dim comEx As COMException = exception
            If Not IsNothing(comEx) Then
                MyBase.HResult = comEx.ErrorCode
            End If
            _name = name
            _desc = desc
            _iname = iname
            _scope = scope
            _exception = exception
        End Sub

        Public ReadOnly Property Name() As String Implements CapeOpen.ECapeRoot.Name
            Get
                Return _name
            End Get
        End Property

        Public ReadOnly Property code() As Integer Implements CapeOpen.ECapeUser.code
            Get
                Return 0 'not supported
            End Get
        End Property

        Public ReadOnly Property description() As String Implements CapeOpen.ECapeUser.description
            Get
                Return _desc
            End Get
        End Property

        Public ReadOnly Property interfaceName() As String Implements CapeOpen.ECapeUser.interfaceName
            Get
                Return _iname
            End Get
        End Property

        Public ReadOnly Property moreInfo() As String Implements CapeOpen.ECapeUser.moreInfo
            Get
                Return "This Material Object is part of a work-shop. Ask Jasper van Baten or Bill Barrett for more info"
            End Get
        End Property

        Public ReadOnly Property operation() As String Implements CapeOpen.ECapeUser.operation
            Get
                Return "N/A" 'not implemented
            End Get
        End Property

        Public ReadOnly Property scope() As String Implements CapeOpen.ECapeUser.scope
            Get
                Return _scope
            End Get
        End Property
    End Class

    Private Sub throwError(ByVal name As String, ByVal desc As String, ByVal iname As String, ByVal scope As String, ByVal exception As System.Exception)
        'uses the implementation of CCapeObject
        throwException(New UserException(name, desc, iname, scope, exception))
    End Sub

    Private Sub throwError(ByVal name As String, ByVal desc As String, ByVal iname As String, ByVal scope As String)
        'uses the implementation of CCapeObject
        throwException(New UserException(name, desc, iname, scope, CapeOpen.CapeErrorInterfaceHR.ECapeUnknownHR))
    End Sub

    'utility functions

    Public Shared Function SameString(ByVal a As String, ByVal b As String) As Boolean
        'CAPE-OPEN strings are case insensitive
        Return (String.Compare(a, b, StringComparison.InvariantCultureIgnoreCase) = 0)
    End Function

    Friend Function getPhaseIndex(ByVal phaseLabel As String) As Integer
        'utiltity function for getting phase index given a phase name
        Dim i As Integer
        For i = 0 To simulator.phaseList.Length - 1
            If (SameString(simulator.phaseList(i), phaseLabel)) Then Return i
        Next i
        throwError("UnknownPhase", "Unknown phase label """ + phaseLabel + """", "N/A", "getPhaseIndex")
    End Function

    Private Sub fractionToMass(ByRef d() As Double)
        'convert fractions from mole to mass basis; only pass arrays of the right size
        Dim tot As Double
        Dim i As Integer
        For i = 0 To simulator.compIDs.Length
            d(i) *= simulator.molWeights(i)
            tot += d(i)
        Next i
        If (tot <= 0) Then throwError("ZeroFraction", "Total composition is zero or negative", "N/A", "fractionToMass")
        tot = 1.0 / tot
        For i = 0 To simulator.compIDs.Length
            d(i) *= tot
        Next i
    End Sub

    Private Sub fractionToMole(ByRef d() As Double)
        'convert fractions from mass to mole basis; only pass arrays of the right size
        Dim tot As Double
        Dim i As Integer
        For i = 0 To simulator.compIDs.Length
            d(i) /= simulator.molWeights(i)
            tot += d(i)
        Next i
        If (tot <= 0) Then throwError("ZeroFraction", "Total composition is zero or negative", "N/A", "fractionToMass")
        tot = 1.0 / tot
        For i = 0 To simulator.compIDs.Length
            d(i) *= tot
        Next i
    End Sub

    Private Function OverallMixtureMW() As Double
        Dim MW As Double
        Dim i As Integer
        'special case for single compound
        If (simulator.compIDs.Length = 1) Then Return simulator.molWeights(0)
        'mixture MW
        MW = 0
        If (overallCompositionMassBasis) Then
            'overall composition in mass basis, first make into mole basis
            Dim d() As Double = overallComposition.Clone
            fractionToMole(d)
            For i = 0 To simulator.compIDs.Length - 1
                MW += d(i) * simulator.molWeights(i)
            Next i
        Else
            'overall composition in mole basis, simple addition
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(overallComposition(i)) Then throwError("MissingComposition", "One or more compositions missing in mass/mole basis conversion", "N/A", "OverallMixtureMW")
                MW += overallComposition(i) * simulator.molWeights(i)
            Next i
        End If
        Return MW
    End Function

    Private Function PhaseMixtureMW(ByVal phaseIndex As Integer) As Double
        Dim MW As Double
        Dim i As Integer
        'special case for single compound
        If (simulator.compIDs.Length = 1) Then Return simulator.molWeights(0)
        'mixture MW
        MW = 0
        If (phaseCompositionMassBasis(phaseIndex)) Then
            'composition in mass basis, first make into mole basis
            Dim d(simulator.phaseList.Length - 1) As Double
            For i = 0 To simulator.compIDs.Length - 1
                d(i) = phaseComposition(phaseIndex, i)
            Next i
            fractionToMole(d)
            For i = 0 To simulator.compIDs.Length - 1
                MW += d(i) * simulator.molWeights(i)
            Next i
        Else
            'overall composition in mole basis, simple addition
            For i = 0 To simulator.compIDs.Length - 1
                If Double.IsNaN(phaseComposition(phaseIndex, i)) Then throwError("MissingComposition", "One or more compositions missing in mass/mole basis conversion", "N/A", "PhaseMixtureMW")
                MW += phaseComposition(phaseIndex, i) * simulator.molWeights(i)
            Next i
        End If
        Return MW
    End Function

    Private Sub scalarToMassOverall(ByRef d As Double)
        d *= 0.001 * OverallMixtureMW()
    End Sub

    Private Sub scalarToMoleOverall(ByRef d As Double)
        d /= 0.001 * OverallMixtureMW()
    End Sub

    Private Sub scalarToMassPhase(ByVal phaseIndex As Integer, ByRef d As Double)
        d *= 0.001 * PhaseMixtureMW(phaseIndex)
    End Sub

    Private Sub scalarToMolePhase(ByVal phaseIndex As Integer, ByRef d As Double)
        d /= 0.001 * PhaseMixtureMW(phaseIndex)
    End Sub

    Private Sub overallMoleToMassConversion(ByRef d() As Double, ByVal order As Integer, ByVal deriv As propTable.DerivativeType)
        If (deriv = propTable.DerivativeType.moles) Then
            'we do not support this mole/mass conversion
            throwError("BasisConversion", "Unknown basis conversion for mole number derivative; please use only mole basis", "N/A", "overallMoleToMassConversion")
        End If
        If (d.Length > simulator.compIDs.Length) Then
            'order higher than vector; this can happen in version 1.0 (e.g. 'pure enthalpy') but not in version 1.1
            throwError("BasisConversion", "Unknown basis conversion for this property type", "N/A", "overallMoleToMassConversion")
        End If
        'conversion factor for mixture
        Dim factor As Double
        If (order = 1) Then
            'e.g. mol/m3
            factor = 0.001 * OverallMixtureMW()
        ElseIf (order = -1) Then
            'e.g. J/mol
            factor = 1.0 / (0.001 * OverallMixtureMW())
        Else
            'should not happen, see property table
            throwError("BasisConversion", "Unknown conversion order for mole/mass conversion", "N/A", "overallMoleToMassConversion")
        End If
        'vector; in version 1.0 we would have to think wether this requires mixture molecular weight (e.g. enthalpy.DmolFraction)
        ' or component values (e.g. 'pure enthalpy'). In version 1.1 we can safely assume we need mixture molecular weight
        'hence, we apply the above factor to all value types
        Dim i As Integer
        For i = 0 To d.Length - 1
            d(i) *= factor
        Next i
    End Sub

    Private Sub phaseMoleToMassConversion(ByVal phaseIndex As Integer, ByRef d() As Double, ByVal order As Integer, ByVal deriv As propTable.DerivativeType)
        If (deriv = propTable.DerivativeType.moles) Then
            'we do not support this mole/mass conversion
            throwError("BasisConversion", "Unknown basis conversion for mole number derivative; please use only mole basis", "N/A", "phaseMoleToMassConversion")
        End If
        If (d.Length > simulator.compIDs.Length) Then
            'order higher than vector; this can happen in version 1.0 (e.g. 'pure enthalpy') but not in version 1.1
            throwError("BasisConversion", "Unknown basis conversion for this property type", "N/A", "phaseMoleToMassConversion")
        End If
        'conversion factor for mixture
        Dim factor As Double
        If (order = 1) Then
            'e.g. mol/m3
            factor = 0.001 * PhaseMixtureMW(phaseIndex)
        ElseIf (order = -1) Then
            'e.g. J/mol
            factor = 1.0 / (0.001 * PhaseMixtureMW(phaseIndex))
        Else
            'should not happen, see property table
            throwError("BasisConversion", "Unknown conversion order for mole/mass conversion", "N/A", "phaseMoleToMassConversion")
        End If
        'vector; in version 1.0 we would have to think wether this requires mixture molecular weight (e.g. enthalpy.DmolFraction)
        ' or component values (e.g. 'pure enthalpy'). In version 1.1 we can safely assume we need mixture molecular weight
        'hence, we apply the above factor to all value types
        Dim i As Integer
        For i = 0 To d.Length - 1
            d(i) *= factor
        Next i
    End Sub

    Private Function phaseFractionToMole(ByVal phaseIndex As Integer, ByRef fraction As Double) As Double
        Dim tot As Double = 0
        Dim myMW As Double
        Dim i As Integer, j As Integer
        Dim composition(simulator.phaseList.Length - 1) As Double
        For i = 0 To simulator.phaseList.Length - 1
            If (phaseExists(i)) Then
                If (Double.IsNaN(phaseFractions(i))) Then throwError("MissingData", "Missing phase fractions for phase fraction basis conversion", "N/A", "phaseFractionToMole")
                If (phaseFractionMassBasis(i)) Then throwError("NotSupported", "Inconsistent phase fraction bases for phase fraction basis conversion; not supported", "N/A", "phaseFractionToMole")
                For j = 0 To simulator.compIDs.Length - 1
                    If (Double.IsNaN(phaseComposition(i, j))) Then throwError("MissingData", "Missing phase composition for phase fraction basis conversion", "N/A", "phaseFractionToMole")
                    composition(j) = phaseComposition(i, j)
                Next j
                If (phaseCompositionMassBasis(i)) Then
                    'convert to mole
                    Dim tot1 As Double = 0
                    For j = 0 To simulator.compIDs.Length - 1
                        composition(i) /= simulator.molWeights(i)
                        tot1 += composition(i)
                    Next j
                    tot1 = 1.0 / tot1
                    For j = 0 To simulator.compIDs.Length - 1
                        composition(i) *= tot1
                    Next j
                End If
                Dim phaseMW As Double = 0
                For j = 0 To simulator.compIDs.Length - 1
                    phaseMW += composition(j) * simulator.molWeights(i)
                Next j
                If (i = phaseIndex) Then myMW = phaseMW
                If (phaseMW <= 0) Then throwError("InvalidData", "total phase molecular weight is zero or negative", "N/A", "phaseFractionToMole")
                tot += phaseFractions(i) / phaseMW
            End If
        Next i
        If (tot <= 0) Then throwError("InvalidData", "total molar phase fraction is zero or negative", "N/A", "phaseFractionToMole")
        phaseFractionToMole = fraction / myMW / tot
    End Function

    Private Function phaseFractionToMass(ByVal phaseIndex As Integer, ByRef fraction As Double) As Double
        Dim tot As Double = 0
        Dim myMW As Double
        Dim i As Integer, j As Integer
        Dim composition(simulator.phaseList.Length - 1) As Double
        For i = 0 To simulator.phaseList.Length - 1
            If (phaseExists(i)) Then
                If (Double.IsNaN(phaseFractions(i))) Then throwError("MissingData", "Missing phase fractions for phase fraction basis conversion", "N/A", "phaseFractionToMass")
                If (phaseFractionMassBasis(i)) Then throwError("NotSupported", "Inconsistent phase fraction bases for phase fraction basis conversion; not supported", "N/A", "phaseFractionToMass")
                For j = 0 To simulator.compIDs.Length - 1
                    If (Double.IsNaN(phaseComposition(i, j))) Then throwError("MissingData", "Missing phase composition for phase fraction basis conversion", "N/A", "phaseFractionToMass")
                    composition(j) = phaseComposition(i, j)
                Next j
                If (phaseCompositionMassBasis(i)) Then
                    'convert to mole
                    Dim tot1 As Double = 0
                    For j = 0 To simulator.compIDs.Length - 1
                        composition(i) /= simulator.molWeights(i)
                        tot1 += composition(i)
                    Next j
                    tot1 = 1.0 / tot1
                    For j = 0 To simulator.compIDs.Length - 1
                        composition(i) *= tot1
                    Next j
                End If
                Dim phaseMW As Double = 0
                For j = 0 To simulator.compIDs.Length - 1
                    phaseMW += composition(j) * simulator.molWeights(i)
                Next j
                If (i = phaseIndex) Then myMW = phaseMW
                tot += phaseFractions(i) * phaseMW
            End If
        Next i
        If (tot <= 0) Then throwError("InvalidData", "total mass phase fraction is zero or negative", "N/A", "phaseFractionToMass")
        phaseFractionToMass = fraction * myMW / tot
    End Function

    'called by MO to set itself current
    Private Sub SetCurrentMO()
        'setting an MO on a PP can be expensive, as the PP may respons with querying and resolving
        ' compound etc, and remember that until a new MO is set current
        If (simulator.currentMO IsNot Me) Then
            'proceed
            Try
                simulator.ppMOContext.SetMaterial(Me)
            Catch ex As Exception
                throwError("FailedToSetMO", "Failed to set current MO on PP: " + MainWnd.ErrorFromCOObject(simulator.ppMOContext, ex), "N/A", "SetCurrentMO", ex)
            End Try
            simulator.currentMO = Me
        End If
    End Sub

End Class


