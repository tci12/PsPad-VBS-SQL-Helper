Class TSet
Dim left
Dim right
Dim parent
Dim value
Dim i

Property Let parentNode(p)
  Set parent = p  
End Property

Property Get leftNode()
  leftNode = left
End Property

Property Get rightNode()
  rightNode = right
End Property

Property Get getValue()
  getValue = value
End Property

Sub init
 i = False
End Sub

Sub insert(ByVal p)

    If Not Me.i Then
      value = p
      i = True
      Set left = Nothing
      Set right = Nothing
      Exit Sub
    End If
    
    If p = value Then
      Exit Sub
    End If
    
    If p > value Then
      If right is Nothing Then
        Set right = new TSet
        right.init
        right.parentNode = Me
      End If
      right.insert(p)      
    Else
      If left is Nothing Then
        Set left = new TSet
        left.init
        left.parentNode = Me
      End If
      left.insert(p)
    End If
    
End Sub

Sub clearAll()
  Me.i = False
  Set value = Nothing
  Set left = Nothing
  Set right = Nothing  
End Sub


Function getValues()       
  If Me.i Then
    Dim leftA,rightA
    Dim size
    size = 0
    
    If Not Me.left is Nothing Then
     leftA = left.getValues
     size = size + (UBound(leftA)+1)
    End If
    If Not Me.right is Nothing Then
      rightA = right.getValues
      size = size + (UBound(rightA)+1)
    End If
    Redim tmp(size)
    Dim x
    Dim ip
    i = 0  
    If Not Me.left is Nothing Then
      ip = x
      For x = ip to ip+Ubound(leftA)
        tmp(x) = leftA(x-ip)
      Next
    End If
    tmp(x) = value
    x = x+1
    If Not Me.right is Nothing Then
      ip = x
      For x = ip to ip+Ubound(RightA)
       tmp(x) = rightA(x-ip)
      Next
    End If                             
    getValues = tmp 
  Else          
    Set getValues = Nothing
  End If

End Function

End Class
