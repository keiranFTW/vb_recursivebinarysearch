Module Search
    Dim List() As String = {"Adam", "Bond", "Clark", "Farmer", "Ford", "Grant", "Jones", "King", "Miller", "Nelson", "Peters", "Smith", "Thornes", "Williams", "Young"}
    Dim Last As Integer = List.Length - 1
    Dim First As Integer = 0
    Dim Midpoint As Integer
    Dim SearchFailed As Boolean
    Dim ItemFound
    Dim Cycles As Integer = 0
    Dim ItemSought
    Sub BinarySearch(ByRef List() As String, ByRef First As Integer, ByRef Last As Integer, ByVal ItemSought As String)
        Cycles = Cycles + 1
        ItemFound = False
        SearchFailed = False
        Midpoint = (First + Last) / 2
        If List(Midpoint) = ItemSought Then
            ItemFound = True
        Else
            If First >= Last Then
                SearchFailed = True
            Else
                If List(Midpoint) > ItemSought Then
                    BinarySearch(List, First, Midpoint - 1, ItemSought)
                Else
                    BinarySearch(List, Midpoint + 1, Last, ItemSought)
                End If
            End If
        End If
    End Sub
    Sub Main()
        BinarySearch(List, First, Last, Console.ReadLine())
        Console.WriteLine("Times loop cycled: " & Cycles)
        If ItemFound Then
            Console.WriteLine("Item found")
        Else
            Console.WriteLine(ItemSought & " is at position " & Midpoint + 1)
        End If
        If SearchFailed Then
            Console.WriteLine("Search failed")
        End If
        Console.Read()
    End Sub

End Module
