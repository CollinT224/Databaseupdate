VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fNumbersReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ensure FindSpreedSheetSize2 Exists in the modules folder
Private name As String
Private pfNumber As String
Private prNumber As String
Private prValue As Double
Private paCompare() As Variant
Private arrayCompare() As Variant
Private pstorageList As New ArrayList
Private pvalueList As New ArrayList
Private ppercentageList As New ArrayList
Private pgenericList As New ArrayList



Public Property Get fNumber() As String
fNumber = pfNumber
End Property

Public Property Let fNumber(ByVal sfNumber As String)
pfNumber = sfNumber
End Property

Public Property Get rNumber() As String
rNumber = prNumber
End Property
Public Property Let rNumber(ByVal srNumber As String)
prNumber = srNumber
End Property

Public Property Get rValue() As Double
rValue = prValue
End Property
Public Property Let rValue(ByVal srValue As Double)
prValue = srValue
End Property

Public Property Get parrayCompare() As Variant
arrayCompare = paCompare
End Property
'Public Property Let parrayCompare(spaArray As Variant)
'paCompare = spaArray
'End Property
Public Property Let parrayCompare(spaArray As Variant)
paCompare = spaArray
End Property

Public Property Get storageList() As ArrayList
Set storageList = pstorageList
End Property
Public Property Get valueList() As ArrayList
Set valueList = pvalueList
End Property
Public Property Get percentageList() As ArrayList
Set percentageList = ppercentageList
End Property
Public Property Get genericList() As ArrayList
Set genericList = pgenericList
End Property

Private Sub Class_Initialize()
Set pstorageList = New ArrayList
Set pvalueList = New ArrayList
Set pgenericList = New ArrayList
Set ppercentageList = New ArrayList
End Sub




Public Sub Hello2()
Debug.Print pfNumber
Debug.Print prNumber
Debug.Print prValue
End Sub

Function addStorage(Value As String)
    pstorageList.Add Value
End Function
Function addValue(Value As Double)
    pvalueList.Add Value
End Function
Function addPercentage(Value As Double)
    ppercentageList.Add Value
End Function
Function addGeneric(Value As String)
    pgenericList.Add Value
End Function

Public Sub ArrayCombine(ByVal position As Double)
Dim positionsize As Double
positionsize = position * 2
Dim paCompare(1) As Variant
paCompare(0) = prNumber
paCompare(1) = prValue
Debug.Print paCompare(0); " STORED IN ARRAY"
Debug.Print paCompare(1); " STORED IN ARRAY"
Debug.Print position; "Position HERE"
parrayCompare = paCompare
End Sub
'2md attempt

'Public Sub ArrayCombine(ByVal position As Double)
'Dim positionsize As Double
'positionsize = position * 2
'Dim paCompare(1, 1) As Variant
'paCompare(0, 0) = prNumber
'paCompare(0, 1) = prValue
'Debug.Print paCompare(0, 0); " STORED IN ARRAY"
'Debug.Print paCompare(0, 1); " STORED IN ARRAY"
'Debug.Print position; "Position HERE"
'parrayCompare = paCompare
'End Sub

'working

'Public Sub ArrayCombine(ByVal position As Double)
'Dim positionsize As Double
'positionsize = position * 2
'Dim paCompare(3, 1) As Variant
'paCompare((position - 1), 0) = prNumber
'paCompare((position - 1), 1) = prValue
'Debug.Print paCompare(position - 1, 0); " STORED IN ARRAY"
'Debug.Print paCompare(position - 1, 1); " STORED IN ARRAY"
'Debug.Print position; "Position HERE"
'parrayCompare = paCompare
'End Sub
Function ArrayPresent()
Dim compArray() As Variant

compArray = paCompare()
'Debug.Print compArray(0)
ArrayPresent = compArray

End Function

