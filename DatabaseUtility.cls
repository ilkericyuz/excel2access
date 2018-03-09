VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private DAC As DataAccessClass
Private cmd As New ADODB.Command
Private conn As New ADODB.connection
Public rs As New ADODB.Recordset
Public pTable As String
Public pField As String
Dim pDateString As String
Dim pfieldDataTypes As Dictionary
Dim pRecordsAffected As Double
Dim pDbPath As String, pDbName As String


' ==================================================
' LETTERS
' ==================================================
Public Property Let table(theTable As String)
    pTable = theTable
End Property

Public Property Let field(theFieldName As String)
    pField = theFieldName
End Property

Public Property Let dateString(theValue As String)
    pDateString = theValue
End Property

Public Property Let dbName(theValue As String)
    pDbName = theValue
End Property

Public Property Let dbPath(theValue As String)
    pDbPath = theValue
End Property

Public Property Let recordsAffected(theValue As Double)
    pRecordsAffected = recordsAffected
End Property


' ==================================================
' GETTERS
' ==================================================
Public Property Get table() As String
    table = pTable
End Property

Public Property Get field() As String
    field = pField
End Property

Public Property Get dateString() As String
    dateString = pDateString
End Property

Public Property Get dbName() As String
    dbName = pDbName
End Property

Public Property Get dbPath() As String
    dbPath = pDbPath
End Property

Public Property Get fieldDataTypes() As Dictionary
    Set fieldDataTypes = pfieldDataTypes
End Property

Public Property Get recordsAffected() As Double
    recordsAffected = pRecordsAffected
End Property

Public Property Get commandText() As String
    commandText = cmd.commandText
End Property

' ==================================================
' INITIALIZE AND TERMINATE
' ==================================================
Private Sub class_initialize()
    ''' Sets DataAccessClass object, its command and connection
        
    Set DAC = New DataAccessClass
    Set cmd = DAC.cmd
    Set conn = DAC.conn
    dateString = ""
    
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    Set DAC = Nothing
End Sub


' ==================================================
' CONNECTION FUNCTIONS
' ==================================================
Public Sub ConnectToDB(Optional theDbName As String = "", Optional theDbPath As String = "")
    ''' Sets connection to database according to given parameters or attributes
    ''' (dbName, dbPath) of the class.
    '''
    ''' @input:
    '''     theDbName:      Name of the MS Access database file with extension.
    '''                     If not given, dbName is used.
    '''     theDbPath:      Path of the MS Access database folder with backslash at the end.
    '''                     If not given, dbPath is used.
    '''
    ''' @output:
    '''     None
    
    ' Check db name
    If theDbName = "" Then theDbName = dbName
    
    ' Check db path
    If theDbPath = "" Then theDbPath = dbPath
    If Right(theDbPath, 1) <> "\" Then theDbPath = theDbPath & "\"

    Call DAC.ConnectToDB(theDbName, theDbPath)
    
End Sub


Public Sub DisconnectFromDB()
    ''' Disconnects from the connected database and resets variables
    '''
    ''' @input:
    '''     None
    '''
    ''' @output:
    '''     None
    
    On Error Resume Next
    Call DAC.DisconnectFromDB
    
End Sub


' ==================================================
' INTERFACE FUNCTIONS
' ==================================================
Public Sub ExecuteSql(theSql As String)
    ''' Executes sql on connected database.
    ''' If sql fetches records, they are saved into ADODB.Recordset (rs).
    ''' If sql changes records in database, number of affected rows can be acquired from recordsAffected.
    '''
    ''' Example:
    '''     ' let "db" be an object of DatabaseUtility
    '''     db.ExecutreSql([sql])
    '''     Do While Not db.rs.EOF
    '''         [variable] = db.rs.fields([field name]).Value
    '''         db.rs.MoveNext
    '''     Loop
    '''
    ''' @input:
    '''     theSql: sql to be executed
    '''
    ''' @output:
    '''     None
    
    cmd.commandText = theSql
    
    ' executes the sql and number of affected rows are given to recordsAffected
    Set rs = cmd.Execute(pRecordsAffected)

End Sub


Function GetTableNames() As Collection
    ''' Returns table names as strings of the connected database in a collection
    '''
    ''' @input:
    '''     None
    '''
    ''' @output:
    '''     GetTableNames:  Collection of strings. It contains all table names in the
    '''             connected database.
    
    Dim tableNames As Collection

    Set rs = conn.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, Empty, "Table"))

    Set tableNames = New Collection
    
    With rs
        Do While Not .EOF
            If .fields("TABLE_TYPE") <> "VIEW" Then
               tableNames.Add .fields("TABLE_NAME").Value
            End If
            .MoveNext
        Loop
    End With
    
    Set GetTableNames = tableNames

Sortie:
    Set tableNames = Nothing
    
End Function


Sub InsertRecord(theTable As String, theSetFields As Collection, theSetValues As Collection)
    ''' Inserts a new row into a table of connected database.
    ''' Field names and their respective values must be passed in two collections with respective order.
    ''' Table name is optional. If it is not passed, table attribute of the class is used.
    '''
    ''' @input:
    '''     theTable:       String of table name.
    '''     theSetFields:   Collection of field names. It can be Nothing if all fields will be filled
    '''             and theSetValues consist of all field values with order.
    '''     theSetValues:   Collection of field values.
    '''
    ''' @output:
    '''     None
    
    cmd.commandText = InsertSql(theTable, theSetFields, theSetValues)
'    Debug.Print cmd.commandText
    cmd.Execute
    
End Sub


Function SelectRecords(theTable As String, _
        Optional theSelectionFields As Collection, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theDistinct As Boolean = False, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderBy As String = "", Optional theAsc As Boolean = False) As ADODB.Recordset
    ''' Returns set of selected records according to the given criteria.
    ''' Contents of criteria fields and criteria values must be in the same order respectively.
    '''
    ''' @input:
    '''     theTable:           String of table name.
    '''     theCriteriaFields:  Collection of database field names which has criteria.
    '''     theCriteriaValues:  Collection of criteria values.
    '''     theSelectionFields: Collection of database field names which wanted to be returned.
    '''             This parameter is optional. If it is not sent, all fields in table are returned.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, relationship between
    '''             criteria fields and criteria values are taken as "=".
    '''     theDistinct:        Boolean value which indicates if returned result values should be distinct or not.
    '''             This parameter is optional. If it is not sent, returned results will not be distinct.
    '''     theLimitBy:         Number of desired records. Top records are returned. It is advised to be used with theOrderBy.
    '''             This parameter is optional. If not sent, all selected records are returned.
    '''     theOrderBy:         String of database field which can be used to sort result collection.
    '''             This parameter is optional. If it is not sent, no sorting occurs.
    '''     theAsc:             Boolean value which indicates if sorting should be ascending or not.
    '''             This parameter is optional. If it is not sent, sorting will be descending.
    ''' @output:
    '''     SelectRecords:      ADODB recordset of values in a database table.

    cmd.commandText = SelectSql(theTable, theSelectionFields, theCriteriaFields, theCriteriaValues, _
            theOperators, theDistinct, theLimitBy, theOrderBy, theAsc)
    
    Set rs = cmd.Execute
    
    Set SelectRecords = rs

Sortie:
    Exit Function

End Function


Function SelectField(theTable As String, _
        theSelectionField As String, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theDistinct As Boolean = False, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderBy As String = "", Optional theAsc As Boolean = False, _
        Optional isBlankIncluded As Boolean = True) As Collection
    ''' Returns single database field according to the given criteria
    ''' Contents of criteria fields and criteria values must be in the same order respectively.
    '''
    ''' @input:
    '''     theTable:           String of table name.
    '''     theSelectionField:  String of database field name in which data is wanted to be returned.
    '''             This parameter is optional. If it is not sent, field attribute of the class is used.
    '''     theCriteriaFields:  Collection of database field names which has criteria.
    '''             This parameter is optional.
    '''     theCriteriaValues:  Collection of criteria values.
    '''             This parameter is optional.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''     theDistinct:        Boolean value which indicates if returned result values should be distinct or not.
    '''             This parameter is optional. If it is not sent, returned results will not be distinct.
    '''     theLimitBy:         Number of desired records. Top records are returned. It is advised to be used with theOrderBy.
    '''             This parameter is optional. If not sent, all selected records are returned.
    '''     theOrderBy:         String of database field which can be used to sort result collection.
    '''             This parameter is optional. If it is not sent, no sorting occurs.
    '''     theAsc:             Boolean value which indicates if sorting should be ascending or not.
    '''             This parameter is optional. If it is not sent, sorting will be descending.
    '''     isBlankIncluded:    Boolean value which indicates if blanks values should be included in the result or not.
    '''             This parameter is optional. If it is not sent, blank values will be included.
    ''' @output:
    '''     SelectField:        Collection of values in a database table field.
    
    Dim selectionFields As New Collection, _
            results As New Collection

    
    selectionFields.Add theSelectionField

    cmd.commandText = SelectSql(theTable, selectionFields, theCriteriaFields, theCriteriaValues, _
            theOperators, theDistinct, theLimitBy, theOrderBy, theAsc)
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        rs.MoveFirst
        
        If isBlankIncluded Then
            Do Until rs.EOF
                results.Add rs.fields(0).Value
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                If Not IsNull(rs.fields(0)) And rs.fields(0) <> "" And rs.fields(0) <> 0 Then
                    results.Add rs.fields(0).Value
                End If
                rs.MoveNext
            Loop
        End If

    End If
    
    Set SelectField = results

Sortie:
    Set results = Nothing
    Set selectionFields = Nothing

End Function


Sub UpdateRecords(theTable As String, _
        theSetFields As Collection, _
        theSetValues As Collection, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection)
    ''' Updates records in database according to the given parameters.
    '''
    ''' @input:
    '''     theTable:           String of table name.
    '''     theSetFields:       Collection of field names which will be updated.
    '''     theSetValues:       Collection of field values which will be updated to.
    '''     theCriteriaFields:  Collection of criteria field names.
    '''     theCriteriaValues:  Collection of criteria field values.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''
    ''' @output:
    '''     None
    
    cmd.commandText = UpdateSql(theTable, theSetFields, theSetValues, theCriteriaFields, _
            theCriteriaValues, theOperators)
    cmd.Execute

End Sub


Function SelectFieldCell(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theOrderBy As String = "", _
        Optional theAscending As Boolean) As Variant
    ''' Returns single cell value from database table according to the given parameters.
    ''' if query returns multiple records, top record is returned.
    '''
    ''' @input:
    '''     theTable:           String of database table name.
    '''     theSelectionField:  String of database field name.
    '''     theCriteriaFields:  Collection of database criteria field names.
    '''     theCriteriaValues:  Collection of database criteria field values.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''     theOrderBy:         String of database field which can be used to sort result collection.
    '''             This parameter is optional. If it is not sent, no sorting occurs.
    '''     theAsc:             Boolean value which indicates if sorting should be ascending or not.
    '''             This parameter is optional. If it is not sent, sorting will be descending.
    '''
    ''' @output:
    '''     SelectFieldCell:    Variant of database cell value.

    Dim cellValue As Variant, i As Double
    Dim theSelectionFields As New Collection
    
    theSelectionFields.Add theSelectionField

    cmd.commandText = SelectSql(theTable, theSelectionFields, theCriteriaFields, theCriteriaValues, _
            theOperators, False, 1, theOrderBy, theAscending)
    
    Set rs = cmd.Execute

    If rs.EOF = True Then
        cellValue = "NOTINDB"
'    ElseIf IsNull(rs.Fields(0).Value) Then
'        cellValue = ""
    Else
        cellValue = rs.fields(0).Value
    End If
    
    SelectFieldCell = cellValue
    
Sortie:
    Set theSelectionFields = Nothing

End Function


Function SelectFieldSum(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderByFieldName As String = "", _
        Optional theAscending As Boolean) As Double
    ''' Returns sum of the selected field values according to the given parameters.
    '''
    ''' @input:
    '''     theTable:           String of database table name.
    '''     theSelectionField:  String of database field name which is wanted to be summed.
    '''             The field must be a summable field.
    '''     theCriteriaFields:  Collection of database criteria field names.
    '''     theCriteriaValues:  Collection of database criteria field values.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''     theLimitBy:         Number of desired records. Top records are returned. It is advised to be used with theOrderBy.
    '''             This parameter is optional. If not sent, all selected records are returned.
    '''     theOrderBy:         String of database field which can be used to sort result collection.
    '''             This parameter is optional. If it is not sent, no sorting occurs.
    '''     theAsc:             Boolean value which indicates if sorting should be ascending or not.
    '''             This parameter is optional. If it is not sent, sorting will be descending.
    '''
    ''' @output:
    '''     SelectFieldSum:     Double of sum of database field values.
    
    Dim sumValue As Double, i As Double
    Dim selectionFields As New Collection

    selectionFields.Add "SUM(" & theSelectionField & ")"

    cmd.commandText = SelectSql(theTable, selectionFields, theCriteriaFields, theCriteriaValues, _
            theOperators, , theLimitBy, theOrderByFieldName, theAscending)

    Set rs = cmd.Execute


    If rs.EOF = True Then
        sumValue = 0
    ElseIf rs.fields(0) = Null Then
        sumValue = 0
    Else
        sumValue = rs.fields(0).Value
    End If
    
    SelectFieldSum = sumValue

Sortie:
    Set selectionFields = Nothing
    
End Function


Function SelectFieldCount(theTable As String, _
        theSelectionField As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderByFieldName As String = "", _
        Optional theAscending As Boolean) As Double
    ''' Returns count of the selected field values according to the given parameters.
    '''
    ''' @input:
    '''     theTable:           String of database table name.
    '''     theSelectionField:  String of database field name which is wanted to be counted.
    '''             The field must be a countable field.
    '''     theCriteriaFields:  Collection of database criteria field names.
    '''     theCriteriaValues:  Collection of database criteria field values.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''     theLimitBy:         Number of desired records. Top records are returned. It is advised to be used with theOrderBy.
    '''             This parameter is optional. If not sent, all selected records are returned.
    '''     theOrderBy:         String of database field which can be used to sort result collection.
    '''             This parameter is optional. If it is not sent, no sorting occurs.
    '''     theAsc:             Boolean value which indicates if sorting should be ascending or not.
    '''             This parameter is optional. If it is not sent, sorting will be descending.
    '''
    ''' @output:
    '''     SelectFieldCount:     Double of count of database field values.
    
    Dim countValue As Double, i As Double
    Dim selectionFields As New Collection

    selectionFields.Add "COUNT(" & theSelectionField & ")"

    cmd.commandText = SelectSql(theTable, selectionFields, theCriteriaFields, theCriteriaValues, _
            theOperators, , theLimitBy, theOrderByFieldName, theAscending)

    Set rs = cmd.Execute


    If rs.EOF = True Then
        countValue = 0
    ElseIf rs.fields(0) = Null Then
        countValue = 0
    Else
        countValue = rs.fields(0).Value
    End If
    
    SelectFieldCount = countValue

Sortie:
    Set selectionFields = Nothing

End Function


Sub DeleteRecords(theTable As String, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection)
    ''' Deletes records in database table according to the given parameters.
    '''
    ''' @input:
    '''     theTable:           String of database table name.
    '''     theCriteriaFields:  Collection of database criteria field names.
    '''     theCriteriaValues:  Collection of database criteria field values.
    '''     theOperators:       Collection of operators which define relationship between
    '''             criteria fields and criteria values. Such as; "=", ">", "<", "<=", ">=".
    '''             This parameter is optional. If it is not sent, releationship between
    '''             criteria fields and criteria values are taken as "=".
    '''
    ''' @output:
    '''     None

    cmd.commandText = DeleteSql(theTable, theCriteriaFields, theCriteriaValues, theOperators)
    
    cmd.Execute
    
End Sub


' ==================================================
' SQL GENERATOR FUNCTIONS
' ==================================================
Private Function DeleteSql(theTable As String, _
        Optional theCriteriaFields As Collection, _
        Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection) As String
    ''' Generates delete sql query according to the parameters.
    '''
    ''' @input:
    '''     theTable:           String of the table name.
    '''     theCriteriaFields:  Collection of criteria fields.
    '''                         This parameter is optional. If not sent, delete query select query does not have a criteria.
    '''     theCriteriaValues:  Collection of criteria values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''     theOperators:       Collection of operators for criteria fields and values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''
    ''' @output:
    '''     DeleteSql:          String of delete SQL query.
    
    Dim SQLString As String, _
            fieldsString As String, _
            criteriaString As String, _
            tableString As String

    tableString = theTable

    criteriaString = ""
    fieldsString = "*"

    If Not theCriteriaFields Is Nothing Then
        criteriaString = CriteriaSqlString(theTable, theCriteriaFields, theCriteriaValues, theOperators)
    End If
    
    SQLString = "DELETE "
    
    SQLString = SQLString & fieldsString & _
            " FROM " & tableString
    
    If Not criteriaString = "" Then
        SQLString = SQLString & " WHERE (" & criteriaString & ")"
    End If
    
        
     DeleteSql = SQLString & ";"
    
End Function


Private Function SelectSql(theTable As String, _
        Optional theSelectionFields As Collection, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection, _
        Optional theDistinct As Boolean = False, _
        Optional theLimitBy As Double = 0, _
        Optional theOrderByFieldName As String = "", _
        Optional theAscending As Boolean = True) As String
    ''' Generates select sql query according to the parameters.
    '''
    ''' @input:
    '''     theTable:           String of the table name.
    '''     theSelectionField:  Collection of fields to select.
    '''                         This parameter is opetional. If not sent, select query includes a star(*), e.g. SELECT * ...
    '''     theCriteriaFields:  Collection of criteria fields.
    '''                         This parameter is optional. If not sent, select query select query does not have a criteria.
    '''     theCriteriaValues:  Collection of criteria values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''     theOperators:       Collection of operators for criteria fields and values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''     theDistinct:        Boolean which indicates if distinct values should be returned or not.
    '''                         This parameter is optional. If not sent, returned query does not have DISTINCT clause.
    '''     theLimitBy:         Number of values which should be returned.
    '''                         This parameter is optional. If not sent, returned query does not have LIMIT BY clause.
    '''     theOrderByFieldName:String of field name to order returned records.
    '''                         This parameter is optional. If not sent, returned query does not have ORDER BY clause.
    '''     theAscending:       Boolean which indicates if ordered records should be ascending or not.
    '''                         This parameter is optional. If not sent and theOrderByFieldName is sent,
    '''                         returned query contains ASCENDING clause.
    '''
    ''' @output:
    '''     SelectSql:          String of select SQL query.
    
    Dim SQLString As String, _
            fieldsString As String, _
            criteriaString As String, _
            length As Double, _
            i As Double
    
    fieldsString = ""
    criteriaString = ""
    
    If Not theSelectionFields Is Nothing Then
        length = theSelectionFields.Count
        For i = 1 To length
            fieldsString = fieldsString & CStr(theSelectionFields(i))
            If i <> length Then
                fieldsString = fieldsString & ", "
            End If
        Next i
    Else:
        fieldsString = "*"
    End If
        
    If Not theCriteriaFields Is Nothing Then
        criteriaString = CriteriaSqlString(theTable, theCriteriaFields, theCriteriaValues, theOperators)
    End If
    
    If theDistinct Then
        SQLString = "SELECT DISTINCT "
    Else
        SQLString = "SELECT "
    End If
    
    'This kind of select distinct sql works ridiculously fast
    If theDistinct Then
        Dim distinctSelectionString As String
        
        distinctSelectionString = CStr(theSelectionFields(1))

        If Not theCriteriaFields Is Nothing Then
            length = theCriteriaFields.Count

            For i = 1 To length
                distinctSelectionString = distinctSelectionString + ", " + CStr(theCriteriaFields(i))
            Next i
        End If
        theTable = "(SELECT DISTINCT " + distinctSelectionString + " FROM " + theTable + ")"
    End If
    
    
    If Not theLimitBy = 0 Then
        SQLString = SQLString & "TOP " & CStr(theLimitBy) & " " & fieldsString & _
                " FROM " & theTable
    Else
        SQLString = SQLString & fieldsString & _
                " FROM " & theTable
    End If
    
    If Not criteriaString = "" Then
        SQLString = SQLString & " WHERE (" & criteriaString & ")"
    End If
    
    If Not theOrderByFieldName = "" Then
        If theAscending Then
            SQLString = SQLString & " ORDER BY " & theOrderByFieldName & " ASC"
        Else
            SQLString = SQLString & " ORDER BY " & theOrderByFieldName & " DESC"
        End If
    End If
    
    SelectSql = SQLString & ";"
    
End Function


Private Function InsertSql(theTable As String, _
        theFields As Collection, _
        theValues As Collection) As String
    ''' Generates an insert query according to the parameters.
    ''' All parameters must be sent because they are needed to find out if quote is necessary
    '''         for each field.
    '''
    ''' @input:
    '''     theTable:   String of the table name to generate insert query.
    '''     theFields:  Collection of field names to insert.
    '''     theValues:  Collection of values to generate insert query.
    '''
    ''' @output:
    '''     InsertSql:  String of insert query.
        
    Dim SQLString As String, _
        fieldsString As String, _
        valuesString As String, _
        criteriaString As String, _
        length As Double, _
        i As Double
    
    fieldsString = ""
    valuesString = ""

    For i = 1 To theFields.Count
        fieldsString = fieldsString & CStr(theFields(i))
        If i <> theFields.Count Then
            fieldsString = fieldsString & ", "
        End If
    Next i

    For i = 1 To theValues.Count
        If IsNull(theValues(i)) Then
                    
            valuesString = valuesString & FixSQLValueString(theTable, _
                        theFields(i), "")
            If i <> theValues.Count Then
                valuesString = valuesString & ", "
            End If
            
        Else
                    
            valuesString = valuesString & FixSQLValueString(theTable, _
                        theFields(i), CStr(theValues(i)))
            If i <> theValues.Count Then
                valuesString = valuesString & ", "
            End If
            
        End If
    Next i
    
    If fieldsString = "" Then
        SQLString = "INSERT INTO " & theTable & _
                    " VALUES (" & valuesString & ");"
    Else
        SQLString = "INSERT INTO " & theTable & _
                    " (" & fieldsString & ") " & _
                    "VALUES (" & valuesString & ");"
    End If
    
'    MsgBox SQLString
    InsertSql = SQLString
    
End Function


Private Function UpdateSql(theTable As String, _
        theSetFields As Collection, _
        theSetValues As Collection, _
        theCriteriaFields As Collection, _
        theCriteriaValues As Collection, _
        Optional theOperators As Collection) As String
    ''' Generates an update query according to the parameters.
    ''' All parameters must be sent because they are needed to find out if quote is necessary
    '''         for each field.
    '''
    ''' @input:
    '''     theTable:           String of the table name to generate update query.
    '''     theSetFields:       Collection of field names to update.
    '''     theSetValues:       Collection of field values to update.
    '''     theCriteriaFields:  Collection of criteria field names.
    '''     theCriteriaValues:  Collection of criteria field values.
    '''     theOperators:       Collection of operators for criteria fields and values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''
    ''' @output:
    '''     UpdateSql:          String of update query.
        
    Dim SQLString As String, _
            criteriaString As String, _
            setString As String, orCriteriaString As String, _
            length As Double, orLength As Double, _
            i As Double, j As Double

    criteriaString = ""
    setString = ""

    length = theSetValues.Count
    
    For i = 1 To length
        
        setString = setString & CStr(theSetFields(i)) _
                & "=" & FixSQLValueString(theTable, theSetFields(i), CStr(theSetValues(i)))
        If i <> length Then
            setString = setString & ", "
        End If

    Next i
    
    criteriaString = CriteriaSqlString(theTable, theCriteriaFields, theCriteriaValues, theOperators)
    
    SQLString = "UPDATE " & table & _
                " SET " & setString & " " & _
                "WHERE (" & criteriaString & ");"
            
    UpdateSql = SQLString
    
End Function


' =================================================
' HELPER FUNCTIONS FOR SQL GENERATORS
' =================================================
Private Function CriteriaSqlString(theTable As String, _
        Optional theCriteriaFields As Collection, Optional theCriteriaValues As Collection, _
        Optional theOperators As Collection) As String
    ''' Generates criteria part of an sql query according to the parameters.
    '''
    ''' @input:
    '''     theTable:           String of the table name.
    '''     theCriteriaFields:  Collection of criteria fields.
    '''                         This parameter is optional. If not sent, select query select query does not have a criteria.
    '''     theCriteriaValues:  Collection of criteria values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''     theOperators:       Collection of operators for criteria fields and values.
    '''                         This parameter is optional. It must be sent if theCriteriaFields is sent.
    '''
    ''' @output:
    '''     None
    
    Dim criteriaString As String, _
            orCriteriaString As String, _
            length As Double, _
            i As Double, _
            j As Double, _
            orLength As Double
    
    criteriaString = ""
    orCriteriaString = ""
        
    If Not theCriteriaFields Is Nothing Then
    
        length = theCriteriaFields.Count
        
        If theOperators Is Nothing Then
            Set theOperators = New Collection
            For i = 1 To length
                theOperators.Add "="
            Next i
        End If
    
        For i = 1 To length
        
            'CHECK IF theCriteriaValues(i) IS A COLLECTION. IF IT IS, WHICH MEANS IT WILL BE "OR" STATEMENT
            If IsObject(theCriteriaValues(i)) Then
            
                orLength = theCriteriaValues(i).Count
                orCriteriaString = ""
                
                For j = 1 To orLength
                    '------------------------------------
                    'NULL QUERY
                    If theCriteriaValues(i)(j) = "NULL" Or theCriteriaValues(i)(j) = """""" _
                            Or IsNull(theCriteriaValues(i)(j)) Then
                        If theOperators(i) = "=" Then
                            orCriteriaString = orCriteriaString & "(" & CStr(theCriteriaFields(i)) _
                                    & " IS NULL"
                            'Check if criteria field requires quotes, add second "OR" argument to the criteria string
                            If isQuoteNecessary(theTable, CStr(theCriteriaFields(i))) Then
                                orCriteriaString = orCriteriaString & " OR " & CStr(theCriteriaFields(i)) & "='')"
                            Else
                                orCriteriaString = orCriteriaString & ")"
                            End If
                            
                            If j <> orLength Then
                                orCriteriaString = orCriteriaString & " OR "
                            End If
                        ElseIf theOperators(i) = "<>" Then
                            orCriteriaString = orCriteriaString & "(" & CStr(theCriteriaFields(i)) _
                                    & " IS NOT NULL"
                            'Check if criteria field requires quotes, add second "OR" argument to the criteria string
                            If isQuoteNecessary(theTable, CStr(theCriteriaFields(i))) Then
                                orCriteriaString = orCriteriaString & " OR " & CStr(theCriteriaFields(i)) & "<>'')"
                            Else
                                orCriteriaString = orCriteriaString & ")"
                            End If
                            
                            If j <> orLength Then
                                orCriteriaString = orCriteriaString & " OR "
                            End If
                        End If
                    '------------------------------------
                    'SEARCH QUERY
                    ElseIf Left(theCriteriaValues(i)(j), 1) = "*" Then
                        If theOperators(i) = "=" Then
                            orCriteriaString = orCriteriaString & "(" & CStr(theCriteriaFields(i)) _
                                    & " LIKE " & "'%" & Replace(CStr(theCriteriaValues(i)(j)), "*", "") & "%')"
                            If j <> orLength Then
                                orCriteriaString = orCriteriaString & " OR "
                            End If
                        ElseIf theOperators(i) = "<>" Then
                            orCriteriaString = orCriteriaString & "(" & CStr(theCriteriaFields(i)) _
                                    & " NOT LIKE " & "'%" & Replace(CStr(theCriteriaValues(i)(j)), "*", "") & "%')"
                            If j <> orLength Then
                                orCriteriaString = orCriteriaString & " OR "
                            End If
                        End If
                    Else
                        orCriteriaString = orCriteriaString & CStr(theCriteriaFields(i)) _
                                & theOperators(i) & FixSQLValueString(theTable, _
                                        theCriteriaFields(i), CStr(theCriteriaValues(i)(j)), True)
                        If j <> orLength Then
                            orCriteriaString = orCriteriaString & " OR "
                        End If
                    End If
                Next j
                
                criteriaString = criteriaString & " (" & orCriteriaString & ") "
                If i <> length Then
                    criteriaString = criteriaString & " AND "
                End If
                
                
            Else
            
            
                '------------------------------------
                'NULL QUERY
                If theCriteriaValues(i) = "NULL" Or theCriteriaValues(i) = """""" _
                        Or IsNull(theCriteriaValues(i)) Then
                    If theOperators(i) = "=" Then
                        criteriaString = criteriaString & "(" & CStr(theCriteriaFields(i)) _
                                & " IS NULL"
                        'Check if criteria field requires quotes, add second "OR" argument to the criteria string
                        If isQuoteNecessary(theTable, CStr(theCriteriaFields(i))) Then
                            criteriaString = criteriaString & " OR " & CStr(theCriteriaFields(i)) & "='')"
                        Else
                            criteriaString = criteriaString & ")"
                        End If
                        
                        If i <> length Then
                            criteriaString = criteriaString & " AND "
                        End If
                    ElseIf theOperators(i) = "<>" Then
                        criteriaString = criteriaString & "(" & CStr(theCriteriaFields(i)) _
                                & " IS NOT NULL"
                        'Check if criteria field requires quotes, add second "OR" argument to the criteria string
                        If isQuoteNecessary(theTable, CStr(theCriteriaFields(i))) And _
                            GetFieldDataType(theTable, CStr(theCriteriaFields(i))) <> 7 And _
                            GetFieldDataType(theTable, CStr(theCriteriaFields(i))) <> 133 And _
                            GetFieldDataType(theTable, CStr(theCriteriaFields(i))) <> 135 Then
                            criteriaString = criteriaString & " OR " & CStr(theCriteriaFields(i)) & "<>'')"
                        Else
                            criteriaString = criteriaString & ")"
                        End If
                        
                        If i <> length Then
                            criteriaString = criteriaString & " AND "
                        End If
                    End If
                '------------------------------------
                'SEARCH QUERY
                ElseIf Left(theCriteriaValues(i), 1) = "*" Then
                    If theOperators(i) = "=" Then
                        criteriaString = criteriaString & "(" & CStr(theCriteriaFields(i)) _
                                & " LIKE " & "'%" & Replace(CStr(theCriteriaValues(i)), "*", "") & "%')"
                        If i <> length Then
                            criteriaString = criteriaString & " AND "
                        End If
                    ElseIf theOperators(i) = "<>" Then
                        criteriaString = criteriaString & "(" & CStr(theCriteriaFields(i)) _
                                & " NOT LIKE " & "'%" & Replace(CStr(theCriteriaValues(i)), "*", "") & "%')"
                        If i <> length Then
                            criteriaString = criteriaString & " AND "
                        End If
                    End If
                Else
                    criteriaString = criteriaString & CStr(theCriteriaFields(i)) _
                            & theOperators(i) & FixSQLValueString(theTable, _
                                    theCriteriaFields(i), CStr(theCriteriaValues(i)), True)
                    If i <> length Then
                        criteriaString = criteriaString & " AND "
                    End If
                End If
                
            End If
        Next i
    End If
    
    CriteriaSqlString = criteriaString
    
End Function


Private Function FixSQLValueString(theTable As String, _
        theField As String, theValue As String, _
        Optional isCriteria As Boolean = False) As String
    ''' Fixes value string part of an sql expression.
    ''' e.g. Let's say sql is "UPDATE [table] SET [field 1]=[value 1] WHERE [field 2] = [value 2];"
    ''' it can fix [value 1] and [value 2] in correct format.
    '''
    ''' @input:
    '''     theTable:       String of the table name to generate update query.
    '''     theField:       String of table field name.
    '''     theValue:       String of field value. numbers and dates should be converted into string before sending here.
    '''     isCriteria:     Boolean which indicates if theValue is in criteria part of the query or not.
    '''             This field is optional. If not sent, it will be taken as False, which means that value string is
    '''             not in criteria field of the query.
    '''
    ''' @output:
    '''     FixSQLValueString:  String of value string in corrected format

    Dim result As String
       
    If isQuoteNecessary(theTable, theField) Then

        If theValue = "" Or theValue = "-" Then
            result = "NULL"
            GoTo Ausgang
        End If

        If isDate(theValue) Then 'TO CONVERT AND ADD # TO DATE IN CRITERIA
            If isCriteria Then
                result = FixCriteriaDate(theValue)
                GoTo Ausgang
            End If
        End If
        result = "'" & WorksheetFunction.Trim(Replace(theValue, "'", "", , , vbBinaryCompare)) & "'"
        
    Else
        If theValue = "" Or theValue = "-" Or theValue = " " Then theValue = "NULL" 'special treatment for null values
        theValue = Replace(Replace(theValue, " ", ""), ",", ".", , , vbBinaryCompare)
        If IsNumeric(theValue) Then
            'check if theValue is a numeric value
            result = theValue
        Else
            'try to fix theValue further
            theValue = Replace(Replace(theValue, "-", "", , , vbBinaryCompare), "+", "", , , vbBinaryCompare)
            If IsNumeric(theValue) Then
                result = theValue
            ElseIf theValue = "NULL" Then
                result = theValue
            ElseIf theValue = "Yes" Or theValue = "No" Or theValue = "True" Or theValue = "False" Then
                result = theValue
            Else
                result = ""
            End If
        End If
    
    End If

Ausgang:
    FixSQLValueString = result
    
End Function


Private Function FixCriteriaDate(theDateString) As String
    ''' In MS SQL, when a date string is in criteria part of the query, it must be in the format of:
    ''' "#yyyy-mm-dd#". This function converts dates from "dd.mm.yyyy" or "dd/mm/yyyy" formats into that format.
    '''
    ''' @input:
    '''     theDateString:      String of date in "dd.mm.yyyy" or "dd/mm/yyyy" formats
    '''
    ''' @output:
    '''     FixCriteriaDate:    String of date in "#yyyy-mm-dd#" format.
    
    Dim resultString() As String, _
        result As String
        
    If InStr(theDateString, "/") > 0 Then
        resultString = Split(theDateString, "/")
    ElseIf InStr(theDateString, ".") > 0 Then
        resultString = Split(theDateString, ".")
    End If
        
    FixCriteriaDate = "#" & resultString(2) & "-" & resultString(1) & "-" & resultString(0) & "#"
    
End Function


Public Sub Fetch_FieldDataTypes(theTable As String)
    ''' Fetches data types of the fields for a given table into pFieldDataTypes dictionary object.
    ''' Field types can be got from fieldDataTypes public get property
    '''
    ''' @input:
    '''     theTable:   Table name, which fields will be fetched.
    '''
    ''' @output:
    '''     None
        
    Set pfieldDataTypes = New Dictionary
    
    Dim i As Double, fieldName As Variant

    ' add each field name and its type into dictionary
    For Each fieldName In GetFieldNames(theTable)
    
        pfieldDataTypes.Add fieldName, GetFieldDataType(theTable, CStr(fieldName))
        
    Next fieldName
    
    pfieldDataTypes.Add 1, theTable
    

End Sub


Function GetFieldNames(Optional theTable As String = "") As Collection
    ''' Returns field names of a table in a collection
    '''
    ''' @input:
    '''     theTable:       Table name, which fields will be fetched.
    '''
    ''' @output:
    '''     GetFieldNames:  Collection of strings which are field names.

    Dim fieldNames As Collection, i As Double
    Set fieldNames = New Collection
    
    If theTable = "" Then theTable = table
    
    cmd.commandText = "SELECT TOP 1 * FROM " & theTable & ";"
    
    Set rs = cmd.Execute '("SELECT CustomerName FROM LoanTable WHERE RiskNumber=4204676036780;")

    For i = 0 To rs.fields.Count - 1
    
        fieldNames.Add rs.fields(i).Name
    
    Next i

    Set GetFieldNames = fieldNames
    Set fieldNames = Nothing

End Function


Function GetFieldDataType(theTable As String, theField As String) As Double
    ''' Returns datatype of a field as its double representation
    ''' Runs a SELECT query for a given field and returns the type of the field.
    '''
    ''' For field type codes:
    ''' http://www.w3schools.com/asp/prop_para_type.asp
    '''
    ''' @input:
    '''     theTable:           String of the table name to generate update query.
    '''     theField:           String of table field name.
    '''
    ''' @output:
    '''     GetFieldDataType:   Double of field data type
    
    Dim sqlStr As String
    
    sqlStr = "SELECT TOP 1 " & theField & " FROM " & theTable & ";"
    cmd.commandText = sqlStr
    
    Set rs = cmd.Execute

    GetFieldDataType = Int(rs.fields(0).Type)

End Function


Private Function isQuoteNecessary(theTable As String, theField As String) As Boolean
    ''' Checks if quote is necessary in sql string for a given field.
    ''' Uses GetFieldDataType function to do that. That function runs a SELECT query for a given field
    ''' and returns the type of the field. This function looks up quote necessary data type codes and
    ''' returns true or false.
    '''
    ''' For field type codes:
    ''' http://www.w3schools.com/asp/prop_para_type.asp
    '''
    ''' @input:
    '''     theTable:           String of the table name to generate update query.
    '''     theField:           String of table field name.
    '''
    ''' @output:
    '''     isQuoteNecessary:   Boolean value indicates if quotes (") are necessary for a field.
    
    On Error GoTo ErrorHandler
    
    Dim fieldCode As Double, i As Double
    
    ' if pFieldDataTypes collection object exists,
    If Not pfieldDataTypes Is Nothing Then
    
        ' and if table name is correct
        If pfieldDataTypes(1) = theTable Then
                    
            On Error GoTo NotInpfieldDataTypes
            
            ' and if field name is found in fieldNames collection of the pFieldDataTypes collection
            fieldCode = pfieldDataTypes(theField)
            GoTo Found

NotInpfieldDataTypes:

            ' if table doesn't exist in fieldNames collection
            fieldCode = GetFieldDataType(theTable, theField)
            
        ' or if table name is not correct
        Else
            fieldCode = GetFieldDataType(theTable, theField)
        End If
        
    ' if pFieldDataTypes collection object does not exist
    Else
        fieldCode = GetFieldDataType(theTable, theField)
    End If
    
Found:
    If fieldCode = 7 _
        Or fieldCode = 8 _
        Or fieldCode = 129 _
        Or fieldCode = 130 _
        Or fieldCode = 132 _
        Or fieldCode = 133 _
        Or fieldCode = 134 _
        Or fieldCode = 135 _
        Or fieldCode = 200 _
        Or fieldCode = 201 _
        Or fieldCode = 202 _
        Or fieldCode = 203 Then
        
        isQuoteNecessary = True
    Else
        isQuoteNecessary = False
    End If
    
Sortie:
    Exit Function
    
ErrorHandler:
    isQuoteNecessary = False
    Resume Sortie
    
End Function

