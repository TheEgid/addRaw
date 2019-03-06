VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmSproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Класс Спецпродукт
Option Explicit

Private mstrFeature As String, _
        mstrName_fas As String, _
        mstrName_raw As String, _
        mstrKod_fas As String, _
        mdblWeight_fas As Double, _
        mstrKod_raw As String, _
        mdblWeight_raw As Double, _
        mstrFab_raw As String, _
        mstrMix As String

'СВОЙСТВА -

Property Let Feature(itput As String)
    mstrFeature = itput
End Property

Property Get Feature() As String
    Feature = mstrFeature
End Property

Property Let Name_fas(itput As String)
    mstrName_fas = itput
End Property

Property Get Name_fas() As String
    Name_fas = mstrName_fas
End Property

Property Let Name_raw(itput As String)
    mstrName_raw = itput
End Property

Property Get Name_raw() As String
    Name_raw = mstrName_raw
End Property

Property Let Kod_fas(itput As String)
    mstrKod_fas = itput
End Property

Property Get Kod_fas() As String
    Kod_fas = mstrKod_fas
End Property

Property Let Weight_fas(itput As Double)
    mdblWeight_fas = itput
End Property

Property Get Weight_fas() As Double
    Weight_fas = mdblWeight_fas
End Property

Property Let Kod_raw(itput As String)
    mstrKod_raw = itput
End Property

Property Get Kod_raw() As String
    Kod_raw = mstrKod_raw
End Property

Property Let Weight_raw(itput As Double)
    mdblWeight_raw = itput
End Property

Property Get Weight_raw() As Double
    Weight_raw = mdblWeight_raw
End Property

Property Let Fab_raw(itput As String)
    mstrFab_raw = itput
End Property

Property Get Fab_raw() As String
    Fab_raw = mstrFab_raw
End Property

Property Let Mix(itput As String)
    mstrMix = itput
End Property

Property Get Mix() As String
    Mix = mstrMix
End Property

'МЕТОДЫ -

Function Calculate_qty_raw(qty As Integer) As Double
    Calculate_qty_raw = WorksheetFunction.RoundUp((qty * mdblWeight_fas) / mdblWeight_raw, 0)
End Function

Function Get_long_kod() As String
    Get_long_kod = mstrKod_raw & "-0" & mdblWeight_raw * 1000
End Function
