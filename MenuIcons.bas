Attribute VB_Name = "MenuIcons"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetMenu Lib "user32" _
    (ByVal hWnd As Long) As Long


Declare Function GetSubMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long


Declare Function GetMenuItemID Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long


Declare Function SetMenuItemBitmaps Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) As Long
    Public Const MF_BITMAP = &H4&


Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    End Type


Declare Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As Long) As Long


Declare Function GetMenuItemInfo Lib "user32" _
    Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
    ByVal un As Long, ByVal b As Boolean, _
    lpMenuItemInfo As MENUITEMINFO) As Boolean
    Public Const MIIM_ID = &H2
    Public Const MIIM_TYPE = &H10
    Public Const MFT_STRING = &H0&

