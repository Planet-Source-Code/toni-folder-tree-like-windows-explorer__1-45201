VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change Icon Size To 32"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox Pi1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   600
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8493
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Im16 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1560
      Picture         =   "Form1.frx":0000
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "Form1.frx":058A
      Top             =   4920
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_dwLrgIconWidth As Long
Dim m_dwLrgIconHeight As Long
Dim m_dwSmIconWidth As Long
Dim m_dwSmIconHeight As Long

Public Enum IcSize
            Icons16
            Icons32
End Enum
Dim m_IconsSize As IcSize

Private Sub Command1_Click()
If m_IconsSize = Icons32 Then
    Command1.Caption = "Change Icon Size To 32"
    m_IconsSize = Icons16
Else
    Command1.Caption = "Change Icon Size To 16"
    m_IconsSize = Icons32
End If
ChangeIconSize
FillDrivers
End Sub

Private Sub Form_Load()
ChangeIconSize
FillDrivers

End Sub
Private Sub ChangeIconSize()
Tree.Nodes.Clear
Im16.ListImages.Clear

If m_IconsSize = Icons16 Then
    Im16.ImageHeight = 16
    Im16.ImageWidth = 16
    Pi1.Width = 240
    Pi1.Height = 240
    Set Pi1.Picture = Image2.Picture
Else
    Im16.ImageHeight = 32
    Im16.ImageWidth = 32
    Pi1.Width = 480
    Pi1.Height = 480
    Set Pi1.Picture = Image1.Picture
End If

Im16.ListImages.Add 1, "Start Icon", Pi1.Picture
Im16.ListImages.Add 2, "Start Icon1", Pi1.Picture
Set Tree.ImageList = Im16
Tree.Indentation = 19 * Screen.TwipsPerPixelX
Tree.ImageList = Im16
Tree.Refresh

End Sub
Private Sub FillDrivers()
Dim Nod As Node, Nod1 As Node
Dim DRV As String
Dim Index

Set Nod = Tree.Nodes.Add(, , , "My Computer", "Start Icon", "Start Icon")
Nod.Expanded = True

For X = 67 To 85
 DRV = Chr(X) + ":"
    If DriveExist(DRV) Then
        GetIcon DRV + "\", False, Index
        Set Nod1 = Tree.Nodes.Add(Nod.Index, tvwChild, , MakeNames(GetDriveName(DRV) + " (" + DRV + ")"), Index, Index)
        Nod1.Sorted = True
        If HasFolderInside(DRV) = True Then
            Tree.Nodes.Add Nod1.Index, tvwChild, , "Dummy"
        End If
    End If
Next

End Sub

Public Function HasFolderInside(Path As String) As Boolean
On Error GoTo HASFErr
Dim Path1 As String
Path1 = Path
    Dim Fso, F, F1, FC, M, ArtH, Att, HasF As Boolean
    If Len(Path1) < 3 Then Path1 = Path1 + "\"
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set F = Fso.GetFolder(Path1)
    Set FC = F.SubFolders
    If FC.Count > 0 Then
        HasFolderInside = True
    Else
        HasFolderInside = False
    End If
Exit Function
HASFErr:
        HasFolderInside = False
End Function

Private Sub Form_Resize()
Tree.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
Command1.Move Tree.Width, 0, Me.ScaleWidth - Tree.Width
End Sub

Public Function DriveExist(DRV) As Boolean
    Dim Fso, msg
    Set Fso = CreateObject("Scripting.FileSystemObject")
    DriveExist = Fso.DriveExists(DRV)
End Function

Public Function GetDriveName(DRV)
On Error GoTo NotDrive
    Dim Fso, D, s
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set D = Fso.GetDrive(Fso.GetDriveName(DRV))
    If D.DriveType = "Remote" Then
        s = D.ShareName
    ElseIf D.IsReady Then
      s = D.VolumeName
    End If
    GetDriveName = s
Exit Function
NotDrive:
GetDriveName = ""
End Function


Private Function CheckImageKey(Key As String, Optional Index As Long) As Boolean
Dim D As Long
For D = 1 To Im16.ListImages.Count
    If LCase(Key) = LCase(Im16.ListImages(D).Key) Then
        Index = D
        CheckImageKey = True
        Exit Function
    End If
Next
Index = 0
CheckImageKey = False
End Function

Private Sub GetIcon(sFilePath As String, OpenIcon As Boolean, Index)
Dim Inde As Long, SysIn As Long, VBart As VbFileAttribute
Dim charc As String
charc = "Icon"

Index = 0
If OpenIcon = True Then
    If m_IconsSize = Icons16 Then
        GetIconSize sFilePath, m_dwSmIconWidth, m_dwSmIconHeight, SHGFI_SMALLICON Or SHGFI_OPENICON, Inde
    Else
        GetIconSize sFilePath, m_dwLrgIconWidth, m_dwLrgIconHeight, SHGFI_OPENICON, Inde
    End If
    If CheckImageKey(charc + Str(Inde), SysIn) = True Then
        Index = SysIn
    Else
        Set Pi1.Picture = Nothing
        If m_IconsSize = Icons16 Then
            ShowFileIcon sFilePath, SHGFI_SMALLICON, Pi1, True
        Else
            ShowFileIcon sFilePath, SHGFI_LARGEICON, Pi1, True
        End If
        
        Set Pi1.Picture = Pi1.Image
        
        Im16.ListImages.Add Im16.ListImages.Count, charc + Str(Inde), Pi1.Picture
        Index = Im16.ListImages.Count - 1
        
    End If
Else
    If m_IconsSize = Icons16 Then
        GetIconSize sFilePath, m_dwSmIconWidth, m_dwSmIconHeight, SHGFI_SMALLICON, Inde
    Else
        GetIconSize sFilePath, m_dwLrgIconWidth, m_dwLrgIconHeight, SHGFI_LARGEICON, Inde
    End If
    If CheckImageKey(charc + Str(Inde), SysIn) = True Then
        Index = SysIn
    Else
        Set Pi1.Picture = Nothing
        If m_IconsSize = Icons16 Then
            ShowFileIcon sFilePath, SHGFI_SMALLICON, Pi1, False
        Else
            ShowFileIcon sFilePath, SHGFI_LARGEICON, Pi1, False
        End If
        
        Set Pi1.Picture = Pi1.Image
        
        Im16.ListImages.Add Im16.ListImages.Count, charc + Str(Inde), Pi1.Picture
        
        Index = Im16.ListImages.Count - 1
    End If
End If

End Sub

Private Sub GetIconSize(sFilePath As String, _
                                    dwWidth As Long, _
                                    dwHeight As Long, _
                                    dwFlags As Long, SysIndex As Long)
  Dim shfi As SHFILEINFO, hSysImgLst As Long
  
  hSysImgLst = SHGetFileInfo(ByVal sFilePath, 0&, shfi, Len(shfi), _
                                            SHGFI_SYSICONINDEX Or dwFlags)
   SysIndex = shfi.iIcon
  ImageList_GetIconSize hSysImgLst, dwWidth, dwHeight
End Sub

Private Sub ShowFileIcon(sFilePath As String, _
                                      uFlags As Long, _
                                      objPB As PictureBox, OpenIcon As Boolean)
  Dim shfi As SHFILEINFO
  
  objPB.Cls   ' clear prev icon
  If OpenIcon = True Then
    SHGetFileInfo ByVal sFilePath, 0&, shfi, Len(shfi), SHGFI_ICON Or SHGFI_OPENICON Or uFlags
  Else
    SHGetFileInfo ByVal sFilePath, 0&, shfi, Len(shfi), SHGFI_ICON Or uFlags
  End If
  ' DrawIconEx() will shrink (or stretch) the
  ' icon per it's cxWidth & cyWidth params
  If uFlags And SHGFI_SMALLICON Then
    DrawIconEx objPB.hDC, 0, 0, shfi.Hicon, _
                      m_dwSmIconWidth, m_dwSmIconHeight, 0, 0, DI_NORMAL
  Else
    DrawIconEx objPB.hDC, 0, 0, shfi.Hicon, _
                      m_dwLrgIconWidth, m_dwLrgIconHeight, 0, 0, DI_NORMAL
  End If
  objPB.Refresh
  
  ' Clean up! -16x16 icons = 380 bytes, 32x32 icons = 1184 bytes
  DestroyIcon shfi.Hicon
End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
Dim NodPath As String, Index1, Index2
Dim Fso, F, F1, FC
Dim Nod As Node
'Debug.Print Node.FullPath, Node.Child.Text
NodPath = GetPath(Node)
If Node.Child.Text = "Dummy" Then
    Tree.Nodes.Remove (Node.Child.Index)
    
        Set Fso = CreateObject("Scripting.FileSystemObject")
        Set F = Fso.GetFolder(NodPath + "\")
        Set FC = F.SubFolders
        For Each F1 In FC
            GetIcon (NodPath + "\" + F1.Name), False, Index1 ' Close Icon
            GetIcon (NodPath + "\" + F1.Name), True, Index2 'Open Icon
            Set Nod = Tree.Nodes.Add(Node.Index, tvwChild, , MakeNames(F1.Name), Index1, Index2)
            Nod.Sorted = True
            
            If HasFolderInside(NodPath + "\" + F1.Name) = True Then
                If UCase(F1.Name) <> "RECYCLED" Then ' Only English Versions
                    Tree.Nodes.Add Nod.Index, tvwChild, , "Dummy"
                End If
            End If
        Next
End If
End Sub

Private Function GetPath(Nod As Node) As String
On Error GoTo NotPath
    GetPath = Mid(Nod.FullPath, InStr(1, Nod.FullPath, ":", vbTextCompare) - 1, Len(Nod.FullPath))
    If Mid(GetPath, 3, 1) = ")" Then
        GetPath = Left(GetPath, 2) + Mid(GetPath, 4, Len(GetPath))
    End If
Exit Function
NotPath:
GetPath = ""
End Function

Private Function MakeNames(Tex As String) As String
Dim X, t
                t = UCase(Left(Tex, 1)) + LCase(Right(Tex, Len(Tex) - 1))
                
                For X = 1 To Len(t)
                    If Mid(t, X, 1) = " " Or Mid(t, X, 1) = "-" Or Mid(t, X, 1) = "_" Or Mid(t, X, 1) = "," _
                    Or Mid(t, X, 1) = "." Or Mid(t, X, 1) = "(" Or Mid(t, X, 1) = ")" Or Mid(t, X, 1) = "[" _
                    Or Mid(t, X, 1) = "]" Or Mid(t, X, 1) = "\" Or Mid(t, X, 1) = "/" Or Mid(t, X, 1) = "=" _
                    Or Mid(t, X, 1) = "+" Then
                        t = Left(t, X) + UCase(Mid(t, X + 1, 1)) + Mid(t, X + 2, Len(t))
                    End If
                Next
    MakeNames = t
End Function

