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

Public m_RootNode As clsAVLNode

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
If m_RootNode Is Nothing Then Exit Function
Set Find_Max_Key_Node = m_RootNode.Find_Max_Key_Node
End Function

'找到标签值最小的节点
Function Find_Min_Key_Node() As clsAVLNode
If m_RootNode Is Nothing Then Exit Function
Set Find_Min_Key_Node = m_RootNode.Find_Min_Key_Node
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
        Set Temp = r.Right.Find_Min_Key_Node
        r.Key = Temp.Key
        r.Userdata = Temp.Userdata
        Set r.Right = Remove_Recursive(r.Right, Temp.Key)
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

Private Function CloneNodeRecursive(n As clsAVLNode) As clsAVLNode
If (n Is Nothing) = False Then
    Set CloneNodeRecursive = New clsAVLNode
    CloneNodeRecursive.Key = n.Key
    CloneNodeRecursive.Userdata = n.Userdata
    CloneNodeRecursive.Height = n.Height
    Set CloneNodeRecursive.Left = CloneNodeRecursive(n.Left)
    Set CloneNodeRecursive.Right = CloneNodeRecursive(n.Right)
End If
End Function

'克隆AVL二叉树
Function Clone() As clsAVLBST
Set Clone = New clsAVLBST
Set Clone.m_RootNode = CloneNodeRecursive(m_RootNode)
End Function

Private Function DefragNodeRecursive(n As clsAVLNode) As Long
If n Is Nothing Then Exit Function

Dim NewNode As clsAVLNode

If (n.Left Is Nothing) = False Then
    Set NewNode = New clsAVLNode
    If ObjPtr(NewNode) < ObjPtr(n.Left) Then
        NewNode.Key = n.Left.Key
        NewNode.Height = n.Left.Height
        NewNode.Userdata = n.Left.Userdata
        Set NewNode.Left = n.Left.Left
        Set NewNode.Right = n.Left.Right
        Set n.Left = NewNode
        DefragNodeRecursive = 1
    Else
        Set NewNode = n.Left
    End If
    DefragNodeRecursive = DefragNodeRecursive + DefragNodeRecursive(NewNode.Left)
    DefragNodeRecursive = DefragNodeRecursive + DefragNodeRecursive(NewNode.Right)
End If

If (n.Right Is Nothing) = False Then
    Set NewNode = New clsAVLNode
    If ObjPtr(NewNode) < ObjPtr(n.Right) Then
        NewNode.Key = n.Right.Key
        NewNode.Height = n.Right.Height
        NewNode.Userdata = n.Right.Userdata
        Set NewNode.Left = n.Right.Left
        Set NewNode.Right = n.Right.Right
        Set n.Right = NewNode
        DefragNodeRecursive = 1
    Else
        Set NewNode = n.Right
    End If
    DefragNodeRecursive = DefragNodeRecursive + DefragNodeRecursive(NewNode.Right)
    DefragNodeRecursive = DefragNodeRecursive + DefragNodeRecursive(NewNode.Right)
End If
End Function

'对整棵树进行内存碎片整理过程，用于集中释放的小块内存区域来提升内存分配性能
Function Defrag() As Long
Defrag = DefragNodeRecursive(m_RootNode)
End Function
