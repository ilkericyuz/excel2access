# excel2access
A Microsoft Excel class module which lets you to connect and manipulate Microsoft Access databases.

You can interact with a MS Access database without any MS SQL knowledge. 
DatabaseUtility class generates SQL inside of it.
    
Example of selecting data from database:

    ' Let's say you have a MS Access database which has a table called "TestTable":
    
    ' ID  text_field  binary_field  number_field
    ' 1   text-1      Yes           10
    ' 2   text-2      No            20
    ' 3   text-3      Yes           30
    ' 4   text-4      No            40
    ' 5   text-5      Yes           50
  
  
    Dim db As New DatabaseUtility
    Dim criteriaFields As New Collection, criteriaValues As New Collection, _
            selectionFields As New Collection
    Dim results As New ADODB.Recordset
    Dim dbName as String, dbPath as String
    
    dbName = "NameOfTheMsAccessDatabase.accdb"
    dbPath = "/Path/to/MS/Access/Database/"
    
    ' Establish connection with database
    Call db.ConnectToDB(dbName, dbPath)
    
    ' Enter the fields you want to select in a collection
    selectionFields.Add "text_field"
    selectionFields.Add "binary_field"
    selectionFields.Add "number_field"
    
    ' Enter criterias for your query in a collection
    criteriaFields.Add "binary_field"
    operators.Add "="
    criteriaValues.Add "No"
        
    criteriaFields.Add "number_field"
    operators.Add ">"
    criteriaValues.Add "20"
    
    ' Get the results in a ADODB.Recordset object
    Set results = db.SelectRecords("TestTable", selectionFields, operators, criteriaFields, criteriaValues)
    
    Call db.disconnectFromDB
    Set db = Nothing
    
    ' results will have
    ' ID  text_field  binary_field  number_field
    ' 3   text-3      Yes           30
    ' 5   text-5      Yes           50
    

You can also execute your own query

    ExecuteSql(theSql As String) as ADODB.Recordset
    

Get names of all tables in a database file in a  collection:

    GetTableNames() As Collection


Insert records:
    
    InsertRecord(theTable As String, theSetFields As Collection, theSetValues As Collection)


Insert a single record:

    InsertRecord(theTable As String, theSetFields As Collection, theSetValues As Collection)
    

Select multiple records:

    SelectRecords(theTable As String, _
        Optional theSelectionFields As Collection, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theDistinct As Boolean = False, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderBy As String = "", Optional theAsc As Boolean = False) As ADODB.Recordset
    

Get multiple (or all) values in a single field:

    SelectField(theTable As String, _
        theSelectionField As String, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theDistinct As Boolean = False, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderBy As String = "", Optional theAsc As Boolean = False, _
        Optional isBlankIncluded As Boolean = True) As Collection
    
    
Update records:

    UpdateRecords(theTable As String, _
        theSetFields As Collection, _
        theSetValues As Collection, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection)
        
        
Select a single cell in a field

     SelectFieldCell(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theOrderBy As String = "", _
        Optional theAscending As Boolean) As Variant
        
        
Get sum of the cells of a number field according to criteria

    SelectFieldSum(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderByFieldName As String = "", _
        Optional theAscending As Boolean) As Double
        
    
Get count of the cells of a number field according to criteria

    SelectFieldCount(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderByFieldName As String = "", _
        Optional theAscending As Boolean) As Double
        
        
Delete records

    DeleteRecords(theTable As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection)
        
        
        
