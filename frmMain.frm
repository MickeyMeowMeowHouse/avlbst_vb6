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
   StartUpPosition =   3  '窗口缺省
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
'测试插入功能
For I = -9999 To 9999
    m_DB.Insert I, "Data of " & I
Next

'测试克隆功能
Dim DB2 As clsAVLBST
Set DB2 = m_DB.Clone

'测试移除功能
For I = 0 To 9999
    DB2.Remove -9999 + I
Next

'测试清理碎片功能
Debug.Print "Defrag: "; DB2.Defrag

'测试查询数据功能（核心功能）
Dim n As clsAVLNode
Set n = DB2.Search(50)
Debug.Assert (n Is Nothing) = False
Debug.Print n.Userdata
End
End Sub
