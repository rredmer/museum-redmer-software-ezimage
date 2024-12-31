VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
'  System......: RSC EZ-IMAGE(r)
'
'  Module......: clsXML.cls
'
'  Subroutine..: (General)(Declarations)
'
'  Description.: Loads and Saves data to XML file.
'
'  (c) 1997-2000 Redmer Software Company.  All Rights Reserved.
'----------------------------------------------------------------------------
Option Explicit



Public Function LoadXMLTableData(XMLTableElement As MSXML.IXMLDOMElement) As MSXML
    Dim rsOut As ADODB.Recordset
    
Dim rs As New Recordset
Dim rs2 As New Recordset
Dim c As New Connection
Dim s As New Stream

' Query the Titles table.
c.Open "provider=sqloledb;data source=mydb;initial catalog=pubs;user id=sa;password="
rs.CursorLocation = adUseClient
rs.Open "select * from titles", c, adOpenStatic

' Save to the file in the XML format. Note that if you don’t specify
' adPersistXML, a binary format (ADTG) will be used by default.
rs.Save "titles.sav", adPersistXML

' Save the Recordset into the ADO Stream object.
rs.Save s, adPersistXML
rs.Close
c.Close

Set rs = Nothing

' Reopen the file.
rs.Open "titles.sav", , , , adCmdFile
' Open the Stream back into a Recordset.
rs2.Open s



End Function
