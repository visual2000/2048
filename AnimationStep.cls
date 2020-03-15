VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnimationStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarstartX As Integer 'local copy
Private mvarstartY As Integer 'local copy
Private mvarendX As Integer 'local copy
Private mvarendY As Integer 'local copy

Public amIaMerge As Boolean


Public Property Let endY(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.endY = 5
    mvarendY = vData
End Property


Public Property Get endY() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.endY
    endY = mvarendY
End Property

Public Function printMe() As String
    printMe = CStr(mvarstartX) + "x" + CStr(mvarstartY) + "->" + CStr(mvarendX) + "x" + CStr(mvarendY)
End Function

Public Property Let endX(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.endX = 5
    mvarendX = vData
End Property


Public Property Get endX() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.endX
    endX = mvarendX
End Property



Public Property Let startY(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.startY = 5
    mvarstartY = vData
End Property


Public Property Get startY() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.startY
    startY = mvarstartY
End Property



Public Property Let startX(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.startX = 5
    mvarstartX = vData
End Property


Public Property Get startX() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.startX
    startX = mvarstartX
End Property



Private Sub Class_Initialize()
    amIaMerge = False
End Sub
