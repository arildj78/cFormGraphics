VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hdcDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Boolean
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal fnObject As Long) As LongPtr
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr

'Drawing methods
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, ByRef lpRect As Any, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function Polygon Lib "gdi32" (ByVal hdc As LongPtr, ByVal lpPoints As LongPtr, ByVal nCount As Long) As Boolean


'Memory management methods
Private Declare PtrSafe Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean
Private Declare PtrSafe Function HeapDestroy Lib "kernel32" (ByVal hHeap As LongPtr) As Boolean






Private pParent As Object
Private pDisplaySurface As MSForms.Image
Private hDcDrawing As LongPtr
Private bmpDrawingOrig As LongPtr
Private bmpDrawing As LongPtr
Private BrushDrawing1 As LongPtr
Private BrushDrawing2 As LongPtr
Private BrushDrawingOrig As LongPtr
Private PenDrawingOrig As LongPtr

Private pxWidth As Long
Private pxHeight As Long
Private GlobalRect As RECT
Private LocalRect As RECT

Private hHeap As LongPtr

Private hWndForm As LongPtr
Private hDcDisplaySurface As LongPtr

Private Triangle(0 To 2) As POINT


Private result As Long 'result from various operations used for debugging




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Properties
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private pColor1 As Long
Private pColor2 As Long

''''''''''''''''''''''
' Color1 color property
''''''''''''''''''''''
Public Property Get Color1() As Double
    Color1 = pColor1
End Property
Public Property Let Color1(Value As Double)
    pColor1 = Value
    RedrawControl
End Property
''''''''''''''''''''''
' Color2 color property
''''''''''''''''''''''
Public Property Get Color2() As Double
    Color2 = pColor2
End Property
Public Property Let Color2(Value As Double)
    pColor2 = Value
    RedrawControl
End Property








Public Sub InitiateProperties(ByRef Parent As Object, _
                              ByRef DisplaySurface As MSForms.Image, _
                              Optional Color1 = &H77AADD, _
                              Optional Color2 = &HDDAA77)
    
    
    hHeap = HeapCreate(0, 0, 0)  ' Create a heap memory object to be used for allocating memory to usercontrol images
    
    
    'Set connection between me and my display surface
    Set pParent = Parent
    Set pDisplaySurface = DisplaySurface

    pColor1 = Color1
    pColor2 = Color2

    Find_Window_handle
    hDcDisplaySurface = GetDC(hWndForm)
    hDcDrawing = CreateCompatibleDC(0)            ' Create DeviceContext that matches the screen that image1 is shown on
    
    PenDrawingOrig = SelectObject(hDcDrawing, GetStockObject(StockObject.NULL_PEN))      'Set Pen to NULL_PEN
    
    DefineDimensions
    
    bmpDrawing = CreateCompatibleBitmap(hDcDisplaySurface, pxWidth, pxHeight)  'This line creates a blank bitmap for doublebuffered drawing compatible with pDisplaySurface
    Set pDisplaySurface.Picture = HandleToPicture(bmpDrawing, PICTYPE_BITMAP)

    BrushDrawing1 = CreateSolidBrush(pColor1)
    BrushDrawing2 = CreateSolidBrush(pColor2)
    

    RedrawControl
End Sub




Private Sub Class_Terminate()
    'Clean up
    HeapDestroy hHeap                  ' Heap memory object
    
    SelectObject hDcDrawing, PenDrawingOrig     'Set original Pen before destroying

    DeleteObject BrushDrawing1
    DeleteObject BrushDrawing2
    
    DeleteObject bmpDrawing
    
    ReleaseDC hWndForm, hDcDisplaySurface
    DeleteDC hDcDrawing
    
End Sub




Public Sub RedrawControl()
    'Select resources
    bmpDrawingOrig = SelectObject(hDcDrawing, bmpDrawing)
    
    
    'Redraw the control
   
    BrushDrawingOrig = SelectObject(hDcDrawing, BrushDrawing2)
    
    result = FillRect(hDcDrawing, LocalRect, BrushDrawing1)
    result = Polygon(hDcDrawing, VarPtr(Triangle(0)), 3)
    
    SelectObject hDcDrawing, BrushDrawingOrig
    
    
    
    result = BitBlt(hDcDisplaySurface, GlobalRect.topleft.x, GlobalRect.topleft.y, pxWidth, pxHeight, hDcDrawing, 0, 0, SRCCOPY)
    
    'Release resources
    SelectObject hDcDrawing, bmpDrawingOrig
End Sub

Private Sub DefineDimensions()
Dim xPixPrPt As Double
Dim yPixPrPt As Double

    
    'Set Scope rectangle for use when invalidating paint area
    xPixPrPt = GetDeviceCaps(hDcDrawing, LOGPIXELSX) / 72
    yPixPrPt = GetDeviceCaps(hDcDrawing, LOGPIXELSY) / 72
    GlobalRect.topleft.x = pDisplaySurface.Left * xPixPrPt
    GlobalRect.topleft.y = pDisplaySurface.Top * yPixPrPt
    GlobalRect.btmRight.x = (pDisplaySurface.Left + pDisplaySurface.Width) * xPixPrPt
    GlobalRect.btmRight.y = (pDisplaySurface.Top + pDisplaySurface.Height) * yPixPrPt
    
    pxWidth = pDisplaySurface.Width * xPixPrPt
    pxHeight = pDisplaySurface.Height * yPixPrPt
  
    LocalRect.topleft.x = 0
    LocalRect.topleft.y = 0
    LocalRect.btmRight.x = pxWidth
    LocalRect.btmRight.y = pxHeight
    
    Triangle(0).x = 0
    Triangle(0).y = pxHeight
    Triangle(1).x = pxWidth
    Triangle(1).y = 0
    Triangle(2).x = pxWidth
    Triangle(2).y = pxHeight
    
End Sub

Private Sub Find_Window_handle()
Dim s As String
Dim pS As LongPtr
Dim strLength As Long
    pS = HeapAlloc(hHeap, 0, 255)

    hWndForm = FindWindow("ThunderDFrame", pParent.Caption)
    strLength = GetWindowText(hWndForm, pS, 255)
    
    s = Mem_ReadString(pS, strLength)
    
    HeapFree hHeap, 0, pS
End Sub


''------------------------------------------------------
' Procedure : HandleToPicture
' Purpose   : Creates a StdPicture object to wrap a GDI
'             image handle
' Credit    : Eduardo A. Morcillo @ https://www.mvps.org/emorcillo/en/code/vb6/loadimagegdip.shtml
'------------------------------------------------------
'
Private Function HandleToPicture( _
   ByVal hGDIHandle As LongPtr, _
   ByVal ObjectType As PictureTypeConstants, _
   Optional ByVal hPal As LongPtr = 0) As StdPicture
Dim tPictDesc As PICTDESC
Dim IID_IPicture As GUID
Dim oPicture As IPicture
    
   ' Initialize the PICTDESC structure
   With tPictDesc
      .cbSizeOfStruct = Len(tPictDesc)
      .picType = ObjectType
      .hgdiObj = hGDIHandle
      .hPalOrXYExt = hPal
   End With
    
   ' Initialize the IPicture interface ID
   With IID_IPicture
      .Data1 = &H7BF80981
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(3) = &HAA
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
    
   ' Create the object
   OleCreatePictureIndirect tPictDesc, IID_IPicture, _
                            True, oPicture
    
   ' Return the picture object
   Set HandleToPicture = oPicture
        
End Function
