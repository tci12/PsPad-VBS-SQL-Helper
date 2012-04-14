Class ArrayList

 

                Private length

                Private data_arr()

               

                Private Sub Class_Initialize()

                               length = 0

                End Sub

               

                Public Property Get arrayLength()

                               arrayLength = length

                End Property

               

                'Přidá na konec dynamického pole záznam

                Public Function add(data)

                               add = addToIndex(data,length)                              

                End Function

               

                'Přidá záznam na požadovaný index. Pokud je index mimo

                'délku pole, nebo je zadaná záporná hodnota, vrací

                'se False hodnota. Pokud je záměr dát záznam nakonec,

                'je záznam přidán, jinak je cílové pole přepsáno

                Public Function addToIndex(data, ByVal index)

                               If index > length OR index < 0 Then

                                               addToIndex = False

                                               Exit Function

                               End If

                               If length = 0 Then

                                               ReDim data_arr (1)

                                               length = 1

                               Else

                                               If index = length Then

                                                               'Přidávám na konec => musím o jednu zvětšit

                                                               length = length + 1

                                                               ReDim Preserve data_arr(length)

                                               End If

                               End If

                               data_arr(index) = data

                               addToIndex = True

                End Function

               

                'Vrátí data na požadovaném indexu. Pokud je index

                'větší než délka-1, nebo je záporný, tak vrátí

                'hodnotu Empty

                Public Function getData(index)

                               If index > length - 1 OR index < 0 Then

                                               getData = Empty

                               Else

                                               getData = data_arr(index)

                               End If                   

                End Function

               

                'Vrátí a vymaže data na požadovaném indexu. Pokud je index

                'větší než délka-1, nebo je záporný, tak vrátí

                'hodnotu Empty             

                Public Function remove(index)

                               If index > length - 1 OR index < 0 Then

                                               getData = Empty

                               Else

                                               Dim i

                                               remove = data_arr(index)

                                               For i = (index + 1) To UBound(data_arr)

                                                               data_arr(i-1) = data_arr(i)

                                               Next

                                               length = length - 1

                                               ReDim Preserve data_arr(length)

                               End If   

                End Function

               

                'Destruktor, uklízí po sobě

                Private Sub Class_Terminate()

                               Dim i

                               If length > 0 Then

                                               data_arr = Nothing                                       

                               End If

                               length = Nothing            

                              

                End Sub

               

End Class
