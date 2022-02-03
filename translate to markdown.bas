Attribute VB_Name = "Module1"
Sub translate()
'create a variable called 'a' and a variable called 'b' that is the type range
Dim cell As Range
Dim b As Range

'this sets the b variable to be what the range of cells that the user selects
'this means that the range b will be the parts of the spreadsheet that the user wants to translate to markdown
Set b = Application.Selection
'MsgBox ("are there any columns to exclude") I think I actually want to use a userform

'variables to keep track of what is going on

'rows
Dim r As Integer
'columns
Dim c As Integer
'all of the columns
Dim tot As Integer
'what the column width is
Dim curW As Integer

'to calculate the total number of columns perform the count operation on the columns within the range that the user chooses
tot = b.Columns.Count

'create variable cW for the column widths that takes in a number
'set a limit of 50 for column count, so that it doesn't keep going forever
Dim cW(50) As String

'this sets the initial width for each column to 0
For i = 0 To tot
    
    cW(i) = 0
Next i

' a nested for loop to go through all of the rows in the range that the user chooses
' the loop will set the column width variable to be equal to the length of each cell
' This will be important for the conversion to markdown because I'll use this to say how many dashes to use to represent the columns
For Each Row In b.Rows
    c = 0
    
    For Each cell In Row.Cells
    
        curW = Len(cell.Value)
        
        If (curW > cW(c)) Then
            cW(c) = curW
        End If
        
        c = c + 1
        
    Next cell
Next Row

'create the variable for the markdown column syntax
Dim colLine As String

'set the number of rows counted to be 0
r = 0

'Nested for loop to go through each row/cell again
'this is also where I should have the checks for hyperlink. if it has a hyperlink, I will need to format it differently to match markdown
'I also need to find a way to let the user choose if they want the hyperlinks included
For Each Row In b.Rows
    c = 0
    colLine = "|"
    For Each cell In Row.Cells
    curW = cW(c)
    Dim xtra As Integer
    colLine = colLine & " "
    colLine = colLine & cell.Value
    xtra = curW - Len(cell.Value)
    
    For j = 0 To xtra
        colLine = colLine & " "
    Next j
    
    colLine = colLine & " |"
    c = c + 1
    
Next cell

Debug.Print colLine


If (r = 0) Then
    colLine = "|"
    c = 0
    
    For j = 0 To (tot - 1)
    
        colLine = colLine
        curW = cW(c)
        colLine = colLine & "-"
        
        For k = 0 To curW
            colLine = colLine & "-"
        Next k
        
        colLine = colLine & "-|"
        c = c + 1
    Next j
    Debug.Print colLine
    End If
    
    r = r + 1
    Next Row



End Sub
