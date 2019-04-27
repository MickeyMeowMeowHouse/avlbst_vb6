VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_DB As New clsAVLBST

Private Sub Form_Load()
Dim I As Long
For I = -99 To 99
    m_DB.Insert I, "Data of " & I
Next

Dim n As clsAVLNode
Set n = m_DB.Search(50)
Debug.Assert (n Is Nothing) = False
Debug.Print n.Userdata
End
End Sub
