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


'this class allows us to look up single phase property definitions

Public Class propTable
    'class for storing raw property data
    Private Class PropTableEntry
        'useful things to know about a property
        Public name As String 'property name
        Public basisOrder As Integer 'order for mass/mole conversions, e.g. -1 for J/mol and 1 for mol/m3
        Public order As Integer '0=scalar, 1=vector, 2=matrix
        Sub New(ByVal name As String, ByVal basisOrder As Integer, ByVal order As Integer)
            Me.name = name
            Me.basisOrder = basisOrder
            Me.order = order
        End Sub
    End Class

    Public Enum DerivativeType
        none = 0
        temperature = 1
        pressure = 2
        moles = 3
        molFraction = 4
    End Enum

    Public Class PropInfo
        Sub New(ByVal fullName As String, ByVal basisOrder As Integer, ByVal deriv As DerivativeType, ByVal dataCount As Integer)
            _dataCount = dataCount
            _fullName = fullName
            _basisOrder = basisOrder
            _deriv = deriv
        End Sub
        'members
        Dim _dataCount As Integer  'number of values
        Dim _fullName As String    'full property name
        Dim _basisOrder As Integer 'order for mass/mole conversions, e.g. -1 for J/mol and 1 for mol/m3
        Dim _deriv As DerivativeType 'type of derivative
        'access
        Public ReadOnly Property FullName() As String
            Get
                Return _fullName
            End Get
        End Property
        Public ReadOnly Property DataCount() As Integer
            Get
                Return _dataCount
            End Get
        End Property
        Public ReadOnly Property BasisOrder() As Integer
            Get
                Return _basisOrder
            End Get
        End Property
        Public ReadOnly Property Derivative() As DerivativeType
            Get
                Return _deriv
            End Get
        End Property
    End Class

    'data members
    Private props() As PropTableEntry
    Private propCount As Integer
    Private table As Hashtable
    Private compCount As Integer
    'constructor
    Sub New(ByVal simulator As MainWnd)
        compCount = 0
        Dim i As Integer
        propCount = 38
        ReDim props(propCount - 1)
        props(0) = New PropTableEntry("Activity", 0, True)
        props(1) = New PropTableEntry("activityCoefficient", 0, True)
        props(2) = New PropTableEntry("Compressibility", 0, False)
        props(3) = New PropTableEntry("compressibilityFactor", 0, False)
        props(4) = New PropTableEntry("Density", 1, False)
        props(5) = New PropTableEntry("diffusionCoefficient", 0, False)
        props(6) = New PropTableEntry("dissociationConstant", 0, False)
        props(7) = New PropTableEntry("Enthalpy", -1, False)
        props(8) = New PropTableEntry("enthalpyF", -1, False)
        props(9) = New PropTableEntry("enthalpyNF", -1, False)
        props(10) = New PropTableEntry("Entropy1", -1, False)
        props(11) = New PropTableEntry("entropyF", -1, False)
        props(12) = New PropTableEntry("entropyNF", -1, False)
        props(13) = New PropTableEntry("excessEnthalpy", -1, False)
        props(14) = New PropTableEntry("excessEntropy", -1, False)
        props(15) = New PropTableEntry("excessGibbsEnergy", -1, False)
        props(16) = New PropTableEntry("excessHelmholtzEnergy", -1, False)
        props(17) = New PropTableEntry("excessInternalEnergy", -1, False)
        props(18) = New PropTableEntry("excessVolume", -1, False)
        props(19) = New PropTableEntry("fugacity", 0, True)
        props(20) = New PropTableEntry("fugacityCoefficient", 0, True)
        props(21) = New PropTableEntry("gibbsEnergy", -1, False)
        props(22) = New PropTableEntry("heatCapacityCp", -1, False)
        props(23) = New PropTableEntry("heatCapacityCv", -1, False)
        props(24) = New PropTableEntry("helmholtzEnergy", -1, False)
        props(25) = New PropTableEntry("internalEnergy", -1, False)
        props(26) = New PropTableEntry("jouleThomsonCoefficient", 0, False)
        props(27) = New PropTableEntry("logFugacity", 0, True)
        props(28) = New PropTableEntry("logFugacityCoefficient", 0, True)
        props(29) = New PropTableEntry("meanActivityCoefficient", 0, False)
        props(30) = New PropTableEntry("molecularWeight", 0, False)
        props(31) = New PropTableEntry("osmoticCoefficient", 0, False)
        props(32) = New PropTableEntry("pH", 0, False)
        props(33) = New PropTableEntry("pOH", 0, False)
        props(34) = New PropTableEntry("speedOfSound", 0, False)
        props(35) = New PropTableEntry("thermalConductivity", 0, False)
        props(36) = New PropTableEntry("viscosity", 0, False)
        props(37) = New PropTableEntry("volume", -1, False)
        'make case insensitive hash table for looking properties up
        table = New Hashtable(New CaseInsentiveComparer)
        For i = 0 To propCount - 1
            table.Item(props(i).name) = props(i)
        Next i
    End Sub

    Public Sub SetCompCount(ByVal count As Integer)
        compCount = count
    End Sub

    Public Function GetPropInfo(ByVal fullName As String, ByRef errorReturn As String) As PropInfo
        'check for derivative
        Dim index As Integer
        Dim propName As String
        Dim deriv As DerivativeType
        index = fullName.IndexOf(".")
        If (index >= 0) Then
            Dim derivName As String
            derivName = fullName.Substring(index + 1)
            If (MaterialObject.SameString(derivName, "Dtemperature")) Then
                deriv = DerivativeType.temperature
            ElseIf (MaterialObject.SameString(derivName, "Dpressure")) Then
                deriv = DerivativeType.pressure
            ElseIf (MaterialObject.SameString(derivName, "Dmoles")) Then
                deriv = DerivativeType.moles
            ElseIf (MaterialObject.SameString(derivName, "DmolFraction")) Then
                deriv = DerivativeType.molFraction
            Else
                'unknown derivative type
                errorReturn = "Unknown derivative: " + derivName
                Return Nothing
            End If
            propName = Left(fullName, index)
        Else
            'no derivative
            propName = fullName
            deriv = DerivativeType.none
        End If
        'look up property
        Dim entry As PropTableEntry = table.Item(propName)
        If (entry Is Nothing) Then
            errorReturn = "Unknown/unsupported property: " + propName
            Return Nothing
        End If
        Dim dataCount As Integer
        If (entry.order = 0) Then
            dataCount = 1
        ElseIf (entry.order = 1) Then
            dataCount = compCount
        ElseIf (entry.order = 2) Then
            dataCount = compCount * compCount
        Else
            errorReturn = "Internal error, invalid order in tables for: " + propName
            Return Nothing
        End If
        If (deriv = DerivativeType.molFraction) Or (deriv = DerivativeType.molFraction) Then dataCount *= compCount
        Return New PropInfo(fullName, entry.basisOrder, deriv, dataCount)
    End Function

End Class
