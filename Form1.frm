VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1920
      Picture         =   "Form1.frx":0363
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const SC_CLOSE As Long = &HF060&

Private Const GWL_WNDPROC As Long = (-4&)

Private Const MF_BYCOMMAND As Long = &H0&
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_CHECKED As Long = &H8&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_BITMAP = &H4&



Private Sub Form_Load()
    Dim hMenu As Long, hID As Long
    hMenu = GetSystemMenu(Me.hWnd, 0)
    'add a item in first pos
    InsertMenu hMenu, &H0, MF_BYPOSITION, IDM.a, "First Menu"
   
    'add a checked item before close item
    InsertMenu hMenu, SC_CLOSE, MF_BYCOMMAND + MF_CHECKED, IDM.b, "Option &Before Close"
    ''add separator after close item
    InsertMenu hMenu, SC_CLOSE, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    ''add item (after the last item)
    InsertMenu hMenu, &HFFFFFFFF, 0&, IDM.c, "&New Close Button"
    ''add a disabled item
    InsertMenu hMenu, &HFFFFFFFF, MF_GRAYED, IDM.d, "&Option MaRiO"
    'add a separator
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, IDM.e, "&Last Option"
    
    'refresh menu
    DrawMenuBar hMenu
    
    'draw icon in pos 13
    hID& = GetMenuItemID(hMenu&, 13)
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Picture1.Picture, Picture1.Picture
    
    'draw icon in first item
    hID& = GetMenuItemID(hMenu&, 0)
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Picture2.Picture, Picture2.Picture
    
    'delete the Close item
    DeleteMenu hMenu, SC_CLOSE, MF_BYCOMMAND
    'subclass
  
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub
