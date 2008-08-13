VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssDepVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pName As String
Private pID As Long
Private pDB As nssDatabase
Private pRegion As nssRegion
Private pAbbrev As String
Private pBCF As Double
Private ptDist As Double
Private pVariance As Double
Private pStations As Double
Private pCoeffDet As Double
Private pStdErr As Double
Private pEstErr As Double
Private pPreErr As Double
Private pStdErrPlus As Double
Private pEquivYears As Double
Private pConstant As Double
Private pMAD As Double
Private pExpDA As Double
Private pIsNew As Boolean
Private pUnits As nssUnits
Private pComponents As FastCollection

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get ID() As Long
  ID = pID
End Property
Public Property Let ID(ByVal newValue As Long)
  pID = newValue
End Property

Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get Abbrev() As String
  Abbrev = pAbbrev
End Property
Public Property Let Abbrev(ByVal newValue As String)
  pAbbrev = newValue
End Property

Public Property Get BCF() As Double
  BCF = pBCF
End Property
Public Property Let BCF(ByVal newValue As Double)
  pBCF = newValue
End Property

Public Property Get tDist() As Double
  tDist = ptDist
End Property
Public Property Let tDist(ByVal newValue As Double)
  ptDist = newValue
End Property

Public Property Get Variance() As Double
  Variance = pVariance
End Property
Public Property Let Variance(ByVal newValue As Double)
  pVariance = newValue
End Property

Public Property Get Stations() As Double
  Stations = pStations
End Property
Public Property Let Stations(ByVal newValue As Double)
  pStations = newValue
End Property

Public Property Get CoeffDet() As Double
  CoeffDet = pCoeffDet
End Property
Public Property Let CoeffDet(ByVal newValue As Double)
  pCoeffDet = newValue
End Property

Public Property Get StdErr() As Double
  StdErr = pStdErr
End Property
Public Property Let StdErr(ByVal newValue As Double)
  pStdErr = newValue
End Property

Public Property Get EstErr() As Double
  EstErr = pEstErr
End Property
Public Property Let EstErr(ByVal newValue As Double)
  pEstErr = newValue
End Property

Public Property Get PreErr() As Double
  PreErr = pPreErr
End Property
Public Property Let PreErr(ByVal newValue As Double)
  pPreErr = newValue
End Property

Public Property Get stdErrPlus() As Double
  stdErrPlus = pStdErrPlus
End Property
Public Property Let stdErrPlus(ByVal newValue As Double)
  pStdErrPlus = newValue
End Property

Public Property Get EquivYears() As Double
  EquivYears = pEquivYears
End Property
Public Property Let EquivYears(ByVal newValue As Double)
  pEquivYears = newValue
End Property

Public Property Get Constant() As Double
  Constant = pConstant
End Property
Public Property Let Constant(ByVal newValue As Double)
  pConstant = newValue
End Property

Public Property Get MAD() As Double
  MAD = pMAD
End Property
Public Property Let MAD(ByVal newValue As Double)
  pMAD = newValue
End Property

Public Property Get ExpDA() As Double
  ExpDA = pExpDA
End Property
Public Property Let ExpDA(ByVal newValue As Double)
  pExpDA = newValue
End Property

Public Property Get IsNew() As Double
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Double)
  pIsNew = newValue
End Property

Public Property Get Region() As nssRegion
  Set Region = pRegion
End Property
Public Property Set Region(newValue As nssRegion)
  Set pRegion = newValue
End Property

Private Sub Class_Initialize()
  pName = "Not Assigned"
End Sub

Public Property Set Components(ByRef newValue As FastCollection)
  Set pComponents = newValue
End Property
Public Property Get Components() As FastCollection
  'ID must be set for this return period before this will work
  If pComponents Is Nothing Then PopulateComponents
  Set Components = pComponents
End Property

Public Property Set Units(ByRef newValue As nssUnits)
  Set pUnits = newValue
End Property
Public Property Get Units() As nssUnits
  Set Units = pUnits
End Property

Private Sub PopulateComponents()
  Dim newComp As nssComponent
  Dim comprec As Recordset
  Dim sql$
  
  Set pComponents = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read components" & vbCr & "DB not set."
  Else
    sql = "SELECT Components.* FROM [Components] " & _
          "WHERE DepVarID=" & pID & " ORDER BY ExponentParmID, ParmID;"
    Set comprec = pDB.DB.OpenRecordset(sql)
    While Not comprec.EOF
      Set newComp = New nssComponent
      With newComp
        Set .DB = pDB
        Set .DepVar = Me
        If Not IsNull(comprec.Fields("DepVarID")) Then .DepVarID = comprec.Fields("DepVarID")
        If Not IsNull(comprec.Fields("ParmID")) Then .ParmID = comprec.Fields("ParmID")
        If Not IsNull(comprec.Fields("BaseMod")) Then .BaseMod = comprec.Fields("BaseMod")
        If Not IsNull(comprec.Fields("BaseCoeff")) Then .BaseCoeff = comprec.Fields("BaseCoeff")
        If Not IsNull(comprec.Fields("BaseExp")) Then .BaseExp = comprec.Fields("BaseExp")
        If Not IsNull(comprec.Fields("ExponentParmID")) Then .ExpID = comprec.Fields("ExponentParmID")
        If Not IsNull(comprec.Fields("ExpMod")) Then .ExpMod = comprec.Fields("ExpMod")
        If Not IsNull(comprec.Fields("ExpExp")) Then .ExpExp = comprec.Fields("ExpExp")
      End With
      pComponents.Add newComp, CStr(comprec("DepVarID") & " " & _
                               comprec("ParmID") & " " & comprec("ExponentParmID"))
      Set newComp = Nothing
      comprec.MoveNext
    Wend
    comprec.Close
  End If
End Sub

Public Function Add(IsReturn As Boolean, ByVal Region As nssRegion, _
    ByVal Name As String, Optional ByVal StdErr As String, _
    Optional ByVal EstErr As String, Optional ByVal PreErr As String, _
    Optional ByVal eqyrs As String, Optional ByVal Constant As String, _
    Optional ByVal BCF As String, Optional ByVal tDist As String, _
    Optional ByVal Variance As String, _
    Optional ByVal ExpDA As String) As Long
    'OPTIONAL ByVal Stations As String, OPTIONAL ByVal CoeffDet As String, OPTIONAL ByVal MAD As String) As Boolean
  Dim myRec As Recordset
  Dim ID&
  Dim StatLabelID&
  
  If Name = "find" Then
    Set myRec = Region.DB.DB.OpenRecordset("lfStatList", dbOpenDynaset)
    myRec.FindFirst "Stat='" & Name & "'"
    If myRec.NoMatch Then
      Name = "unknown"
    Else
      Name = myRec("Stat")
    End If
  End If
  StatLabelID = GetLabelID(Name, Region.DB)
  Set myRec = Region.DB.DB.OpenRecordset("DepVars", dbOpenDynaset)
  With myRec
    'check to see if parameter with this name already exists for this region
    '.FindFirst "RegionID=" & Region.ID & " AND StatLabelCode='" & Name & "'"
    .FindFirst "RegionID=" & Region.ID & " AND StatisticLabelID=" & StatLabelID
    If .NoMatch Then  'OK to add
      'find an unused ID
      If IsReturn Then
        ID = 1
      Else
        ID = 10001
      End If
      .FindFirst "DepVarID=" & ID
      While Not .NoMatch
        ID = !DepVarID + 1
        .FindFirst "DepVarID=" & ID
      Wend
      .AddNew
      !DepVarID = ID
      !RegionID = Region.ID
      '!statLabelCode = Name
      !StatisticLabelID = StatLabelID
      If IsNumeric(StdErr) Then !StdErr = StdErr
      If IsNumeric(EstErr) Then !EstErr = EstErr
      If IsNumeric(PreErr) Then !PreErr = PreErr
      If IsNumeric(eqyrs) Then !EquivYears = eqyrs
      If IsNumeric(Constant) Then !Constant = Constant
      If IsNumeric(BCF) Then !BCF = BCF
      If IsNumeric(tDist) Then !T = tDist
      If IsNumeric(Variance) Then !Variance = Variance
      If IsNumeric(ExpDA) Then !ExpDA = ExpDA
'      if isnumeric(Stations) then !Stas = Stations
'      if isnumeric(CoeffDet) then !CoeffDet = CoeffDet
'      if isnumeric(MAD) then !MAD = MAD
      .Update
      Add = ID
    Else
      ssMessageBox "'" & Name & "' already exists for " & Region.Name & _
            "." & vbCrLf & "No entry was made."
      Add = -1
    End If
  End With
End Function

Public Sub Delete()
  Dim myRec As Recordset
  
  Set myRec = Me.Region.DB.DB.OpenRecordset("DepVars", dbOpenDynaset)
  With myRec
    .FindFirst "DepVarID=" & Me.ID
    If Not .NoMatch Then
      .Delete
    End If
  End With
End Sub

Public Sub Edit(ByVal Name As String, _
    Optional ByVal StdErr As String, Optional ByVal EstErr As String, _
    Optional ByVal PreErr As String, Optional ByVal eqyrs As String, _
    Optional ByVal Constant As String, Optional ByVal BCF As String, _
    Optional ByVal tDist As String, Optional ByVal Variance As String, _
    Optional ByVal ExpDA As String)
    'ByVal Stations As String, ByVal CoeffDet As String, ByVal MAD As String) As Boolean
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("DepVars", dbOpenDynaset)
  With myRec
    .MoveLast
    .FindFirst "DepVarID=" & Me.ID
    If Not .NoMatch Then
      .Edit
      !Name = Name
      If IsNumeric(StdErr) Then !StdErr = StdErr
      If IsNumeric(EstErr) Then !EstErr = EstErr
      If IsNumeric(PreErr) Then !PreErr = PreErr
      If IsNumeric(eqyrs) Then !EquivYears = eqyrs
      If IsNumeric(Constant) Then !Constant = Constant
      If IsNumeric(BCF) Then !BCF = BCF
      If IsNumeric(tDist) Then !T = tDist
      If IsNumeric(Variance) Then !Variance = Variance
      If IsNumeric(ExpDA) Then !ExpDA = ExpDA
'      if isnumeric(Stations) then !Stas = Stations
'      if isnumeric(CoeffDet) then !CoeffDet = CoeffDet
'      if isnumeric(MAD) then !MAD = MAD
      .Update
    End If
  End With
  myRec.Close
End Sub

Public Function PopulateMatrix() As Variant
  Dim covRec As Recordset
  Dim myDim&, row&, col&
  Dim sql$
  Dim covArray() As String
  
  On Error GoTo 0
  If pDB Is Nothing Then
    ssMessageBox "Cannot read covariance matrix" & vbCr & "DB not set."
  Else
    myDim = Me.Components.Count + 1
    ReDim covArray(1 To myDim, 1 To myDim)
    sql = "SELECT * FROM [Covariance] " & _
          "WHERE DepVarID=" & Me.ID & _
          " ORDER BY Row, Col;"
    Set covRec = pDB.DB.OpenRecordset(sql)
    With covRec
      If .RecordCount > 0 Then
        .MoveLast
        If !row <> myDim Then GoTo DBProblem
        If !col <> myDim Then GoTo DBProblem
        .MoveFirst
      End If
      While Not .EOF
        row = !row
        col = !col
        If Not IsNull(!Value) Then covArray(row, col) = CStr(!Value)
        .MoveNext
      Wend
    End With
    covRec.Close
    PopulateMatrix = covArray
  End If
  Exit Function

DBProblem:
  ssMessageBox "The dimensions of the covariance matrix on the database are " & _
      "not the expected size.", vbExclamation, "Covariance Matrix Problem"

End Function

Public Sub AddMatrix(ByVal Region As nssRegion, _
    ByVal DepVarID As Long, ByRef covArray() As String)
  Dim covRec As Recordset
  Dim myDim&, row&, col&
  
  On Error GoTo 0
  If Region.DB Is Nothing Then
    ssMessageBox "Cannot access covariance matrix" & vbCr & "DB not set."
  Else
    myDim = UBound(covArray, 2)
    Set covRec = Region.DB.DB.OpenRecordset("Covariance", dbOpenDynaset)
    With covRec
      For row = 1 To myDim
        For col = 1 To myDim
          .AddNew
          !DepVarID = DepVarID
          !row = row
          !col = col
          If IsNumeric(covArray(row, col)) Then !Value = covArray(row, col)
          .Update
        Next col
      Next row
    End With
    covRec.Close
  End If
End Sub

Public Function ClearOldComponents() As Boolean
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("Components", dbOpenDynaset)
  With myRec
    While Not .EOF
      .FindFirst "DepVarID=" & Me.ID
      If Not .NoMatch Then
        .Delete
        ClearOldComponents = True
      Else
        GoTo X
      End If
    Wend
  End With
X:
  myRec.Close
  Set myRec = pDB.DB.OpenRecordset("Covariance", dbOpenDynaset)
  With myRec
    While Not .EOF
      .FindFirst "DepVarID=" & Me.ID
      If Not .NoMatch Then
        .Delete
        ClearOldComponents = True
      Else
        Exit Function
      End If
    Wend
  End With
  myRec.Close
End Function

