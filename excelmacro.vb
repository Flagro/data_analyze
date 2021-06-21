Public Function Contains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key)
    Contains = (Err.Number = 0)
    Err.Clear
End Function
Public Function ArrLen(ByRef Arr As Variant) As Integer
    Let ArrLen = UBound(Arr) - LBound(Arr) + 1
End Function
Function GetAdjacencyList(ByVal MacroSheet As Worksheet) As Collection
    Dim AdjacencyList As Collection
    Set AdjacencyList = New Collection
    Dim CurRow As Integer
    Let CurRow = 1
    Dim CurrV As String
    Dim AdjactedV As String
    Do While MacroSheet.Cells(CurRow, 1).Text <> ""
        Let CurrV = MacroSheet.Cells(CurRow, 1).Text
        Let AdjactedV = MacroSheet.Cells(CurRow, 2).Text
        Dim AdjactedList As Variant
        Dim AdjactedListLen As Integer
        Dim NewAdjactedList() As String
        
        If Contains(AdjacencyList, CurrV) Then
            Let AdjactedList = AdjacencyList.Item(CurrV)
            Let AdjactedListLen = UBound(AdjactedList) - LBound(AdjactedList)
            ReDim Preserve AdjactedList(AdjactedListLen + 1)
            Let AdjactedList(AdjactedListLen + 1) = AdjactedV
            AdjacencyList.Remove CurrV
            AdjacencyList.Add Item:=AdjactedList, key:=CurrV
        Else
            ReDim NewAdjactedList(0) As String
            Let NewAdjactedList(0) = AdjactedV
            AdjacencyList.Add Item:=NewAdjactedList, key:=CurrV
        End If

        If Contains(AdjacencyList, AdjactedV) Then
            Let AdjactedList = AdjacencyList.Item(AdjactedV)
            Let AdjactedListLen = UBound(AdjactedList) - LBound(AdjactedList)
            ReDim Preserve AdjactedList(AdjactedListLen + 1)
            Let AdjactedList(AdjactedListLen + 1) = CurrV
            AdjacencyList.Remove AdjactedV
            AdjacencyList.Add Item:=AdjactedList, key:=AdjactedV
        Else
            ReDim NewAdjactedList(0) As String
            Let NewAdjactedList(0) = CurrV
            AdjacencyList.Add Item:=NewAdjactedList, key:=AdjactedV
        End If
        CurRow = CurRow + 1
    Loop
    Set GetAdjacencyList = AdjacencyList
End Function
Sub DFSColoring(ByVal CurrV As String, ByRef AdjacencyList As Collection, ByVal CurrColor As String, ByRef ColorList As Collection)
    ColorList.Add Item:=CurrColor, key:=CurrV
    Dim AdjactedV As Variant
    For Each AdjactedV In AdjacencyList.Item(CurrV)
        If Not Contains(ColorList, AdjactedV) Then
            Call DFSColoring(AdjactedV, AdjacencyList, CurrColor, ColorList)
        End If
    Next
End Sub
Function GetColorList(ByRef AdjacencyList As Collection, ByVal MacroSheet As Worksheet) As Collection
    Dim CurRow As Integer
    Let CurRow = 1
    Dim ColorList As Collection
    Set ColorList = New Collection
    Dim CurrColor As Integer
    Let CurrColor = 1
    Do While MacroSheet.Cells(CurRow, 1).Text <> ""
        Let CurrV = MacroSheet.Cells(CurRow, 1).Text
        If Not Contains(ColorList, CurrV) Then
            Call DFSColoring(CurrV, AdjacencyList, CurrColor, ColorList)
            CurrColor = CurrColor + 1
        End If
        CurRow = CurRow + 1
    Loop
    Set GetColorList = ColorList
End Function
Function FirstType(ByVal CurrColor As Integer, ByRef FirstV As Variant, ByRef SecondV As Variant, ByRef AdjacencyList As Collection, ByRef ColorList As Collection) As Boolean
    Dim Ans As Boolean
    Let Ans = True
    Dim ArrEl As Variant
    For Each ArrEl In FirstV
        If ColorList.Item(ArrEl) = CurrColor And ArrLen(AdjacencyList.Item(ArrEl)) <> 1 Then
            Ans = False
        End If
    Next
    For Each ArrEl In SecondV
        If ColorList.Item(ArrEl) = CurrColor And ArrLen(AdjacencyList.Item(ArrEl)) <> 1 Then
            Ans = False
        End If
    Next
    FirstType = Ans
End Function
Function SecondType(ByVal CurrColor As Integer, ByRef FirstV As Variant, ByRef SecondV As Variant, ByRef AdjacencyList As Collection, ByRef ColorList As Collection) As Boolean
    Dim Ans As Boolean
    Let Ans = True
    Dim ArrEl As Variant
    For Each ArrEl In SecondV
        If ColorList.Item(ArrEl) = CurrColor And ArrLen(AdjacencyList.Item(ArrEl)) <> 1 Then
            Ans = False
        End If
    Next
    SecondType = Ans
End Function
Function ThirdType(ByVal CurrColor As Integer, ByRef FirstV As Variant, ByRef SecondV As Variant, ByRef AdjacencyList As Collection, ByRef ColorList As Collection) As Boolean
    Dim Ans As Boolean
    Let Ans = True
    Dim ArrEl As Variant
    For Each ArrEl In FirstV
        If ColorList.Item(ArrEl) = CurrColor And ArrLen(AdjacencyList.Item(ArrEl)) <> 1 Then
            Ans = False
        End If
    Next
    ThirdType = Ans
End Function
Function GetComponentTypes(ByRef FirstV As Variant, ByRef SecondV As Variant, ByRef AdjacencyList As Collection, ByRef ColorList As Collection) As Collection
    Dim ComponentTypes As Collection
    Set ComponentTypes = New Collection
    Dim ColorCnt As Integer
    Let ColorCnt = 0
    Dim ColorListEl As Variant
    For Each ColorListEl In ColorList
        If ColorListEl > ColorCnt Then
            ColorCnt = ColorListEl
        End If
    Next
    Dim ColorNum As Integer
    For ColorNum = 1 To ColorCnt
        Dim Ans As String
        If FirstType(ColorNum, FirstV, SecondV, AdjacencyList, ColorList) Then
            Let Ans = "1_1"
        ElseIf SecondType(ColorNum, FirstV, SecondV, AdjacencyList, ColorList) Then
            Let Ans = "1_*"
        ElseIf ThirdType(ColorNum, FirstV, SecondV, AdjacencyList, ColorList) Then
            Let Ans = "*_1"
        Else
            Let Ans = "*_*"
        End If
        ComponentTypes.Add Item:=Ans, key:=CStr(ColorNum)
    Next
    Set GetComponentTypes = ComponentTypes
End Function
Sub GraphAnalyzeMacro()
'
' GraphAnalyzeMacro Macro
'

'
    Dim MacroSheet As Worksheet
    Set MacroSheet = ActiveSheet
    
    Dim FirstV() As String
    Dim SecondV() As String
    
    Dim VCnt As Integer
    Let VCnt = 0
    
    Dim CurRow As Integer
    Let CurRow = 1
    
    Do While MacroSheet.Cells(CurRow, 1).Text <> ""
        ReDim Preserve FirstV(VCnt)
        ReDim Preserve SecondV(VCnt)
        Let FirstV(VCnt) = MacroSheet.Cells(CurRow, 1).Text
        Let SecondV(VCnt) = MacroSheet.Cells(CurRow, 2).Text
        VCnt = VCnt + 1
        CurRow = CurRow + 1
    Loop

    Dim AdjacencyList As Collection
    Set AdjacencyList = GetAdjacencyList(MacroSheet)
    
    Dim ColorList As Collection
    Set ColorList = GetColorList(AdjacencyList, MacroSheet)
    
    Let CurRow = 1
    Do While MacroSheet.Cells(CurRow, 1).Text <> ""
        Let CurrV = MacroSheet.Cells(CurRow, 1).Text
        Let MacroSheet.Cells(CurRow, 4).NumberFormat = "0"
        Let MacroSheet.Cells(CurRow, 3).Value = ColorList.Item(CurrV)
        CurRow = CurRow + 1
    Loop
    
    Dim ComponentTypes As Collection
    Set ComponentTypes = GetComponentTypes(FirstV, SecondV, AdjacencyList, ColorList)
    
    Let CurRow = 1
    Dim CurrColor As Integer
    Do While MacroSheet.Cells(CurRow, 1).Text <> ""
        Let CurrColor = MacroSheet.Cells(CurRow, 3).Text
        Let MacroSheet.Cells(CurRow, 4).NumberFormat = "@"
        Let MacroSheet.Cells(CurRow, 4).Value = ComponentTypes.Item(CurrColor)
        CurRow = CurRow + 1
    Loop
End Sub
