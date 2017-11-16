Attribute VB_Name = "Enumerations"
Public Enum DeviceCap
    LOGPIXELSX = 88 ' Logical pixels inch in X
    LOGPIXELSY = 90 ' Logical pixels inch in Y
End Enum

Public Enum TernaryRasterOperations
    SRCCOPY = &HCC0020
    SRCPAINT = &HEE0086
    SRCAND = &H8800C6
    SRCINVERT = &H660046
    SRCERASE = &H440328
    NOTSRCCOPY = &H330008
    NOTSRCERASE = &H1100A6
    MERGECOPY = &HC000CA
    MERGEPAINT = &HBB0226
    PATCOPY = &HF00021
    PATPAINT = &HFB0A09
    PATINVERT = &H5A0049
    DSTINVERT = &H550009
    BLACKNESS = &H42
    WHITENESS = &HFF0062
End Enum

Public Type PICTDESC     'For use when creating OLE pictureobject
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As LongPtr
   hPalOrXYExt As LongPtr
End Type

Public Type GUID          'For use when creating OLE pictureobject
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7)  As Byte
End Type

Public Type POINT
    x As Long
    y As Long
End Type

Public Type RECT
    topleft As POINT
    btmRight As POINT
End Type

Public Enum StockObject
    WHITE_BRUSH = &H0
    LTGRAY_BRUSH = &H1
    GRAY_BRUSH = &H2
    DKGRAY_BRUSH = &H3
    BLACK_BRUSH = &H4
    NULL_BRUSH = &H5
    HOLLOW_BRUSH = &H5
    WHITE_PEN = &H6
    BLACK_PEN = &H7
    NULL_PEN = &H8
    OEM_FIXED_FONT = &HA
    ANSI_FIXED_FONT = &HB
    ANSI_VAR_FONT = &HC
    SYSTEM_FONT = &HD
    DEVICE_DEFAULT_FONT = &HE
    DEFAULT_PALETTE = &HF
    SYSTEM_FIXED_FONT = &H10
    DEFAULT_GUI_FONT = &H11
    DC_BRUSH = &H12
    DC_PEN = &H13
End Enum
Public Enum PictureTypeConstants
    PICTYPE_UNINITIALIZED = -1
    PICTYPE_NONE = 0
    PICTYPE_BITMAP = 1
    PICTYPE_METAFILE = 2
    PICTYPE_ICON = 3
    PICTYPE_ENHMETAFILE = 4
End Enum

