VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVLBST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RootNode As clsAVLNode

Private Function New_Node(ByVal Key As Long, Optional Userdata) As clsAVLNode
Set New_Node = New clsAVLNode
New_Node.Key = Key
New_Node.Userdata = Userdata
End Function

'找最大
Private Function Max(ByVal A As Long, ByVal B As Long) As Long
If A > B Then Max = A Else Max = B
End Function

'取节点高度
Private Function Height(Node As clsAVLNode)
If Node Is Nothing Then Exit Function
Height = Node.Height
End Function

'左旋
Private Function RoL(x As clsAVLNode) As clsAVLNode
Dim y As clsAVLNode
Dim T2 As clsAVLNode

Set y = x.Right
Set T2 = y.Left

Set y.Left = x
Set x.Right = T2

x.Height = Max(Height(x.Left), Height(x.Right)) + 1
y.Height = Max(Height(y.Left), Height(y.Right)) + 1

Set RoL = y
End Function

'右旋
Private Function RoR(y As clsAVLNode) As clsAVLNode
Dim x As clsAVLNode
Dim T2 As clsAVLNode

Set x = y.Left
Set T2 = x.Right

Set x.Right = y
Set y.Left = T2

y.Height = Max(Height(y.Left), Height(y.Right)) + 1
x.Height = Max(Height(x.Left), Height(x.Right)) + 1

Set RoR = x
End Function

'取平衡值
Private Function Get_Balance(n As clsAVLNode) As Long
If n Is Nothing Then Exit Function
Get_Balance = Height(n.Left) - Height(n.Right)
End Function

'插入后保持平衡
Private Function Keep_Balance_Insert(n As clsAVLNode, ByVal Key As Long) As clsAVLNode
Dim Balance As Long
Balance = Get_Balance(n)

'Left Left Case
If Balance > 1 Then
    If Key < n.Left.Key Then
        Set Keep_Balance_Insert = RoR(n)
        Exit Function
    End If
End If

'Right Right Case
If Balance < -1 Then
    If Key > n.Right.Key Then
        Set Keep_Balance_Insert = RoL(n)
        Exit Function
    End If
End If

'Left Right Case
If Balance > 1 Then
    If Key > n.Left.Key Then
        Set n.Left = RoL(n.Left)
        Set Keep_Balance_Insert = RoR(n)
        Exit Function
    End If
End If

'Right Left Case
If Balance < -1 Then
    If Key < n.Right.Key Then
        Set n.Right = RoR(n.Right)
        Set Keep_Balance_Insert = RoL(n)
        Exit Function
    End If
End If

Set Keep_Balance_Insert = n
End Function

'递归插入
Private Function Insert_Recursive(n As clsAVLNode, ByVal Key As Long, Optional Userdata) As clsAVLNode
'如果到了树梢
If n Is Nothing Then
    Set Insert_Recursive = New_Node(Key, Userdata) '创建节点并返回
    Exit Function
End If

'否则根据标签找到正确的树梢
If Key < n.Key Then
    Set n.Left = Insert_Recursive(n.Left, Key, Userdata)
ElseIf Key > n.Key Then
    Set n.Right = Insert_Recursive(n.Right, Key, Userdata)
Else '命中已有
    n.Userdata = Userdata
    Set Insert_Recursive = n
    Exit Function
End If

n.Height = Max(Height(n.Left), Height(n.Right)) + 1
Set Insert_Recursive = Keep_Balance_Insert(n, Key)
End Function

'封装为普通的插入过程
Sub Insert(ByVal Key As Long, Optional Userdata)
Set m_RootNode = Insert_Recursive(m_RootNode, Key, Userdata)
End Sub

'查找对应数据，返回节点。数据在节点的Userdata里
Function Search(ByVal Key As Long) As clsAVLNode
Dim n As clsAVLNode
Set n = m_RootNode
Do Until n Is Nothing
    If n.Key = Key Then
        Set Search = n
        Exit Function
    ElseIf Key < n.Key Then
        Set n = n.Left
    Else
        Set n = n.Right
    End If
Loop
End Function

'找到标签值最大的节点
Function Find_Max_Key_Node() As clsAVLNode
Dim n As clsAVLNode
Set n = m_RootNode
If n Is Nothing Then Exit Function
Do Until n.Right Is Nothing
    Set n = n.Right
Loop
Set Find_Max_Key_Node = n
End Function

'找到标签值最小的节点
Function Find_Min_Key_Node() As clsAVLNode
Dim n As clsAVLNode
Set n = m_RootNode
If n Is Nothing Then Exit Function
Do Until n.Left Is Nothing
    Set n = n.Left
Loop
Set Find_Min_Key_Node = n
End Function

'在移除节点后保持平衡
Private Function Keep_Balance_Remove(n As clsAVLNode) As clsAVLNode
Dim Balance As Long
Balance = Get_Balance(n)

'Left Left Case
If Balance > 1 Then
    If Get_Balance(n.Left) >= 0 Then
        Set Keep_Balance_Remove = RoR(n)
        Exit Function
    End If
End If

'Left Right Case
If Balance > 1 Then
    If Get_Balance(n.Left) < 0 Then
        Set n.Left = RoL(n.Left)
        Set Keep_Balance_Remove = RoR(n)
        Exit Function
    End If
End If

'Right Right Case
If Balance < -1 Then
    If Get_Balance(n.Right) <= 0 Then
        Set Keep_Balance_Remove = RoL(n)
        Exit Function
    End If
End If

'Right Left Case
If Balance < -1 Then
    If Get_Balance(n.Right) > 0 Then
        Set n.Right = RoR(n.Right)
        Set Keep_Balance_Remove = RoL(n)
        Exit Function
    End If
End If

Set Keep_Balance_Remove = n
End Function

'递归移除节点
Private Function Remove_Recursive(r As clsAVLNode, ByVal Key As Long) As clsAVLNode
If r Is Nothing Then Exit Function
If Key < r.Key Then
    Set r.Left = Remove_Recursive(r.Left, Key)
ElseIf Key > r.Key Then
    Set r.Right = Remove_Recursive(r.Right, Key)
Else
    Dim Temp As clsAVLNode
    If r.Left Is Nothing Or r.Right Is Nothing Then
        If r.Left Is Nothing Then Set Temp = r.Right Else Set Temp = r.Left
        If Temp Is Nothing Then
            Set Temp = r
            Set r = Nothing
        Else
            Set r = Temp
        End If
        Set Temp = Nothing
    Else
        Set Temp = Find_Min_Key_Node(r.Right)
        r.Key = Temp.Key
        r.Userdata = Temp.Userdata
        r.Right = Remove_Recursive(r.Right, Temp.Key)
    End If
End If

If r Is Nothing Then Exit Function
r.Height = Max(Height(r.Left), Height(r.Right)) + 1

Set Remove_Recursive = Keep_Balance_Remove(r)
End Function

'封装为普通过程
Sub Remove(ByVal Key As Long)
Set m_RootNode = Remove_Recursive(m_RootNode, Key)
End Sub


