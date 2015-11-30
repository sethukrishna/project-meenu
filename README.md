# project-meenu
how to connect to db?

1)add references
2)
Option Explicit

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public productID As Integer
Public productName As String
Public topCell As Range

Sub Main()
    ' Delete any previous results.
    Range("B1") = ""
    Set topCell = Worksheets("Orders").Range("A3")
    With topCell
        Range(.Offset(1, 0), .Offset(1, 4).End(xlDown)).ClearContents
    End With
    
    On Error GoTo ErrorLabel
    
    ' Open connection to database.
    With cn
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\Sales Orders.mdb"
        .Provider = "Microsoft Jet 4.0 OLE DB Provider"
        .Open
    End With
    
    Call GetProductList     'call sub to get product list for DDL
    Call GetOrderInfo       'call sub to get order info for the selected product
    
    ' Close the connection.
    cn.Close
    
    
    Range("A2").Select
    
    Exit Sub
    
ErrorLabel:
    MsgBox "Database Error: " & Err.Number & ": " & vbCrLf & vbCrLf & Error(Err.Number)
      
    'Quit the program; this should close
    'open connection and record set
    End
    
End Sub

Sub GetProductList()
    Dim SQL As String
    
    ' Import product info and use it to populate the list box.
    ' After frmProducts is unloaded, we will know the productID
    ' and productName of the selected product.
    SQL = "SELECT ProductID, ProductName FROM Products"
    'Run the SQL statement with the opened connection and
    'a corresponding recordset should be created
    rs.Open SQL, cn, adOpenStatic, adLockOptimistic
    frmProducts.Show   'initialize the form's DDL so that it is populated by recordset
    rs.Close
End Sub

Sub GetOrderInfo()
    Dim SQL As String
    Dim rowCount As Integer
        
    Range("B1") = productName
    
    ' Define SQL statement to get order info for selected product from DDL.
    SQL = "SELECT O.OrderID, O.OrderDate, L.QuantityOrdered, " _
        & "L.QuotedPrice, L.QuantityOrdered * L.QuotedPrice AS ExtendedPrice " _
        & "FROM Orders O INNER JOIN LineItems L ON O.OrderID = L.OrderID " _
        & "WHERE L.ProductID =" & productID & " " _
        & "ORDER BY O.OrderDate, O.OrderID"
    
    ' Run the query and use results to fill Orders sheet.
    With rs
        .Open SQL, cn
        rowCount = 0
        Do While Not .EOF
            rowCount = rowCount + 1
            topCell.Offset(rowCount, 0) = .Fields("OrderID")
            topCell.Offset(rowCount, 1) = .Fields("OrderDate")
            topCell.Offset(rowCount, 2) = .Fields("QuotedPrice")
            topCell.Offset(rowCount, 3) = .Fields("QuantityOrdered")
            topCell.Offset(rowCount, 4) = .Fields("ExtendedPrice")
            .MoveNext
        Loop
        .Close
    End With
End Sub

3) code inside forms

Private Sub btnCancel_Click()
    Unload Me
    End
End Sub

Private Sub btnOK_Click()
    productID = lbProducts.List(lbProducts.ListIndex, 0)
    productName = lbProducts.List(lbProducts.ListIndex, 1)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim rowCount As Integer
    Dim productArray(100, 2) As Variant ' 2-Dim array; assume no more than 100 products.
     
    
    'MsgBox rs.RecordCount
     
    ' Populate the two-column list box with items from the recordset.
    rowCount = 0
    With rs
        Do Until .EOF
            productArray(rowCount, 0) = .Fields("ProductID")
            productArray(rowCount, 1) = .Fields("ProductName")
            rowCount = rowCount + 1
            .MoveNext
        Loop
    End With
    lbProducts.List = productArray
    lbProducts.ListIndex = 0
End Sub
