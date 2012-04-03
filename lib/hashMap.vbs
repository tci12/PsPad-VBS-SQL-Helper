Class HashMap
Dim map

Sub init
  set map = CreateObject("Scripting.Dictionary")
End sub

Sub put(ByVal key, ByVal value)

  If (map.Exists(key)) Then
    map.remove(key)
  End If
  map.add key,value   
  
End Sub

Function getItem(ByRef key)

  getItem = map.Item(key)

End Function

Function getKeys()

  getKeys = map.Keys

End Function

Sub remove(ByRef key)

  map.remove(key)

End Sub

Sub clear()

  map.RemoveAll

End Sub

End Class