Class FormGraphics for MS Forms 2.0
--------------------
Base class for drawing custom UI elements in Visual Basic for Applications (VBA).
All drawing is done by calls to Windows API.

Theory of operation
After creating an instance of the object it is given a an empty Image usercontrol
to draw on.


Example project
* Create an empty Excel Workbook and open VBA
* Create a new UserForm
* Put an Image control on the form and name it Image1
* Put an CommandButton control on the form and name it CommandButton1
* Copy the following code to UserForm1 code (rightclick UserForm1 and select View Code)
* Import the following files to the project
    - cFormGraphics.cls
    - factory.bas
    - memory.bas
    - enumerations.bas

    Option Explicit
    
    Public MyGraphic As cFormGraphics
    
    Private Sub UserForm_Initialize()
        Set MyGraphic = factory.Create_FormGraphics(Me, Image1)
    End Sub
    
    Private Sub CommandButton1_Click()
        MyGraphic.Color1 = RGB(255, 0, 0)
        MyGraphic.Color2 = RGB(0, 0, 255)
    End Sub

Open UserForm1 by doubleclicking it in the project tree and run it by pressing F5.
You should now see that the image you put on the form have is divided into two colored
triangles. By pressing the CommandButton the colors should change.

