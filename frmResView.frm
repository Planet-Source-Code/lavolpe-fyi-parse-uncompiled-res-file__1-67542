VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB IDE Resource Parser"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5010
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboItem 
      Height          =   315
      Left            =   4155
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1590
      Width           =   2220
   End
   Begin VB.ComboBox cboGroup 
      Height          =   315
      Left            =   4155
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1110
      Width           =   2220
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3840
      Left            =   120
      ScaleHeight     =   3780
      ScaleWidth      =   3780
      TabIndex        =   1
      Top             =   165
      Width           =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select VB RES File"
      Height          =   435
      Left            =   4155
      TabIndex        =   0
      Top             =   2115
      Width           =   2220
   End
End
Attribute VB_Name = "frmResView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cR As New cRESreader

Private Sub cboGroup_Click()

    ' resource group item was selected

    If cboGroup.ListIndex = -1 Then Exit Sub
    
    Dim Index As Long, sName As String
    cboItem.Clear
    
    ' enumerate the child resources for the item selected
    sName = cR.ResourceID(cboGroup.Text, Index)
    Do Until sName = vbNullString
        cboItem.AddItem sName
        Index = Index + 1
        sName = cR.ResourceID(cboGroup.Text, Index)
    Loop

    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0

End Sub

Private Sub cboItem_Click()
    
    Dim tArray() As Byte
    
    Select Case cboGroup.Text
    Case "Bitmap", "Icon", "Cursor"
        Set Picture1.Picture = cR.ExtractResPicture(cboGroup.Text, cboItem.Text)
    Case Else
        Set Picture1.Picture = Nothing
        Picture1.CurrentX = 0: Picture1.CurrentY = 0
        If cR.ExtractResStream(cboGroup.Text, cboItem.Text, tArray()) = False Then
            Picture1.Print vbCrLf; Space$(5); "No data was extracted from the resource"
        Else
            Picture1.Print vbCrLf; Space$(5); UBound(tArray) + 1; " bytes were extracted from the array"
        End If
    End Select
    Picture1.Refresh

End Sub

Private Sub Command1_Click()
    
    With CommonDialog1
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .CancelError = True
        .Filter = "VB Resource Files|*.res"
    End With
    On Error GoTo ExitRoutine
    CommonDialog1.ShowOpen
    
    Picture1.Cls
    Set Picture1.Picture = Nothing
    
    cR.ScanResources CommonDialog1.FileName
    
    cboGroup.Clear
    cboItem.Clear
    
    Dim Index As Long, sName As String
    
    sName = cR.ResourceSection(Index)
    Do Until sName = vbNullString
        cboGroup.AddItem sName
        Index = Index + 1
        sName = cR.ResourceSection(Index)
    Loop
    If cboGroup.ListCount > 0 Then cboGroup.ListIndex = 0
    
ExitRoutine:
End Sub
