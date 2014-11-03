VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type InvoiceData
    OrderID As String
    OrderDate As Date
    CustID As String
    CustInfo As String
    ProdInfo As Variant
End Type

Private m_Data As InvoiceData

Public Sub GetData(sOrderID As Variant, sConn As Variant)

    Dim oConn As Object, oRS As Object
    
    'Connect to the Northwind database
    Set oConn = CreateObject("ADODB.Connection")
    oConn.Open sConn
    
    'Obtain the Customer ID and Order Date
    Set oRS = CreateObject("ADODB.Recordset")
    oRS.Open "Select [OrderDate], [CustomerID] from Orders where " & _
             "[OrderID]=" & sOrderID, oConn, 3 'adOpenStatic=3
    m_Data.OrderID = sOrderID
    m_Data.OrderDate = CDate(oRS.Fields("OrderDate").Value)
    m_Data.CustID = oRS.Fields("CustomerID").Value
    oRS.Close
    
    'Obtain Customer Information
    Set oRS = CreateObject("ADODB.Recordset")
    oRS.Open "Select * from Customers Where CustomerID='" & _
             m_Data.CustID & "'", oConn, 3 'adOpenStatic=3
    m_Data.CustInfo = oRS.Fields("CompanyName").Value & vbCrLf & _
                      oRS.Fields("City") & " "
    If Not (IsNull(oRS.Fields("Region"))) Then
       m_Data.CustInfo = m_Data.CustInfo & oRS.Fields("Region").Value & " "
    End If
    m_Data.CustInfo = m_Data.CustInfo & oRS.Fields("PostalCode").Value & _
                      vbCrLf & oRS.Fields("Country").Value
    oRS.Close
    
    'Obtain Product information
    Set oRS = CreateObject("ADODB.Recordset")
    oRS.Open "Select ProductName, Quantity, [Order Details].UnitPrice," & _
             "Discount from Products Inner Join [Order Details] on " & _
             "Products.ProductID = [Order Details].ProductID " & _
             "Where OrderID = " & sOrderID, oConn, 3 'adOpenStatic=3
    m_Data.ProdInfo = oRS.GetRows
    oRS.Close
    
    'Close the connection to the database
    oConn.Close

End Sub

Public Sub SendData(oXML As Variant)

    'Extract the information from the DOMDocument object oXML and store
    'it in the private member variable m_Data.

    m_Data.OrderID = oXML.getElementsByTagName("OrderID").Item(0).Text
    m_Data.OrderDate = oXML.getElementsByTagName("OrderDate").Item(0).Text
    m_Data.CustID = oXML.getElementsByTagName("CustID").Item(0).Text
    m_Data.CustInfo = oXML.getElementsByTagName("CustInfo").Item(0).Text
    
    Dim oItems As Object, oItem As Object
    Set oItems = oXML.getElementsByTagName("Items").Item(0)
    ReDim vArray(0 To 3, 0 To oItems.childNodes.Length - 1) As Variant
    Dim i As Integer
    For i = 0 To UBound(vArray, 2)
        Set oItem = oItems.childNodes(i)
        vArray(0, i) = oItem.getAttribute("Desc")
        vArray(1, i) = oItem.getAttribute("Qty")
        vArray(2, i) = oItem.getAttribute("Price")
        vArray(3, i) = oItem.getAttribute("Disc")
    Next
    m_Data.ProdInfo = vArray

End Sub

Public Sub MakeInvoice(sTemplate As Variant, Optional bSave As Variant)
    
    Dim oWord As Object
    Dim oDoc As Object
    Dim oTable As Object
    
    If IsMissing(bSave) Then bSave = False

    'Open the document as read-only
    Set oWord = CreateObject("Word.Application")
    Set oDoc = oWord.Documents.Open(sTemplate, , True)
    
    'Fill in the bookmarks
    oDoc.Bookmarks("Customer_Info").Range.Text = m_Data.CustInfo
    oDoc.Bookmarks("Customer_ID").Range.Text = m_Data.CustID
    oDoc.Bookmarks("Order_ID").Range.Text = m_Data.OrderID
    oDoc.Bookmarks("Order_Date").Range.Text = m_Data.OrderDate
    
    'Fill in the table with the product information
    '** Note the table starts out with three rows -- the first row
    '   contains headers for the table, the second row is for
    '   the first set of product data and the third row contains a total.
    '   New rows are added for additional products before the "total row".

    Set oTable = oDoc.Tables(1)
    Dim r As Integer, c As Integer
    For r = 1 To UBound(m_Data.ProdInfo, 2) + 1
        If r > 1 Then oTable.Rows.Add (oTable.Rows(oTable.Rows.Count))
        For c = 1 To 4
            oTable.Cell(r + 1, c).Range.Text = _
               m_Data.ProdInfo(c - 1, r - 1)
        Next
        oTable.Cell(r + 1, 5).Formula _
            "=(B" & r + 1 & "*C" & r + 1 & ")*(1-D" & r + 1 & ")", _
            "#,##0.00"
    Next

    'Update the field for the grand total and protect the document
    oTable.Cell(oTable.Rows.Count, 5).Range.Fields.Update
    oDoc.Protect 1 'wdAllowOnlyComments=1
        
    If bSave Then
        'Save the document as "c:\invoice.doc" and quit Word
        Dim nResult As Long
        nResult = MsgBox("Are you sure you wish to create the document" & _
             " ""c:\invoice.doc""? If this document already exists, " & _
             "it will be replaced", vbYesNo, "AutomateWord")
        If nResult = vbYes Then oDoc.SaveAs "c:\invoice.doc"
        oDoc.Close False
        oWord.Quit
    Else
        'Make Word visible
        oWord.Visible = True
    End If
    
End Sub


