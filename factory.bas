Attribute VB_Name = "factory"

Public Function Create_FormGraphics(ByRef Parent As Object, _
                                 ByRef DisplaySurface As MSForms.Image, _
                                 Optional Color1 = &H77AADD, _
                                 Optional Color2 = &HDDAA77) As cFormGraphics
    Set Create_FormGraphics = New cFormGraphics
    Create_FormGraphics.InitiateProperties Parent:=Parent, _
                                        DisplaySurface:=DisplaySurface, _
                                        Color1:=Color1, _
                                        Color2:=Color2
End Function
