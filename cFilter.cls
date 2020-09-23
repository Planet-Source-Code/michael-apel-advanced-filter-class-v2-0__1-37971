VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0

Private Const cdblVersion As Double = 2

Private Const cerrBase As Long = vbObjectError + 21300 'change as needed
Private Const cstrWaitMsg As String = "Aquiring data..."
Private Const cstrSeekMsg As String = "Seeking..."

'##############################################################################
'#                                                                            #
'#                       Version 2.00 final 2002-08-15                        #
'#                                                                            #
'##############################################################################
'#                        (C)2002 BEG Bürkle GmbH+Co.                         #
'#                                                                            #
'#              all code by Michael Apel unless otherwise stated              #
'#              This module may be distributed free of charge in              #
'#              unmodified form. This code comes with absolutely              #
'#              no warrenty! Use at your own risk!                            #
'#                                                                            #
'#              Contact: apel@beg-buerkle.de                                  #
'##############################################################################
'#                                                                            #
'#  This module allows you to dynamicall apply filters on a data form.        #
'#  You can add and remove filter one by one or in groups while keeping       #
'#  the current record if it is still in the record group. It should perform  #
'#  resonably fast on Jet and show good performance on ODBC linked tables to  #
'#  a database server compared to a FindFirst approach.                       #
'#                                                                            #
'#  The basic idea is to use a SELECT COUNT instead of FindFirst to find the  #
'#  record again. FindFirst has to look at every value in the field while     #
'#  SELECT COUNT can be executed directly on the server.                      #
'#                                                                            #
'#  Requirements:                                                             #
'#                                                                            #
'#  Access 2000 (when replacing the enums it should run with down to '95 as   #
'#  well although it's untested), a bound Form with at least one control      #
'#  bound to a key (usually the primary) in the recordsource and for each     #
'#  field you want to order by.                                               #
'#  The idea behind that is to determine the current value of the field in    #
'#  the record the form is currently showing while still having the           #
'#  possibilty to check for the no-data situation. Those controls don't have  #
'#  to be visible.                                                            #
'#                                                                            #
'#  Notes:                                                                    #
'#                                                                            #
'#  Functions having the vRefresh parameter will not fire an Requery if the   #
'#  Filter is not active yet even if the parameter is set to true.            #
'#  No error will be generated as well, this is mainly due to implementation  #
'#  compatibility reasons and I didn't feel like that being a bad thing.      #
'#  Fields you filter on and order by should usually be indexed (d'uh) but    #
'#  it's not madatory.                                                        #
'#  You can only place one order by as multiple would require more selects    #
'#  to find the correct record. Ordering to other fields than the defined     #
'#  key will result in lower performance.                                     #
'#                                                                            #
'#  ToDo (in order of priority):                                              #
'#                                                                            #
'#  - Method to fast search for records                                       #
'#  - Better error handling                                                   #
'#  - Better docs                                                             #
'#  - Method to remove all filters that won't allow to show a record with     #
'#    a specified field value                                                 #
'#                                                                            #
'##############################################################################

'Basic usage example:
'##############################################################################
'Dim MyFilter As cFilter
'
'Private Sub Form_Open(Cancel As Integer)
'   Set MyFilter = New cFilter
'   If MyFilter.Version >= 2 Then
'
'       MyFilter.RecordSource = "SELECT * FROM MyTable"
'       MyFilter.Index = "MyKey"
'       Set MyFilter.Form = Me.Form
'       Set MyFilter.IndexControl = Me.MyKeyControl
'       MyFilter.Active = True
'   Else
'       Set MyFilter = Nothing
'   End If
'
'End Sub
'
'Private Sub OnFilterClick()
'   MyFilter.Add "MyFilterField", fltEqual, "FilterValue", True, "This is a nice filter"
'End Sub
'
'Private Sub OnNoFilterClick()
'   MyFilter.Remove "MyFilterField", "*", True
'End Sub
'
'Private Sub OnNoNoFilterClick()
'   MyFilter.RemoveAll True
'End If
'
'Private Sub OnOrderClick()
'   MyFilter.OrderBy "MyOrderField", Me.MyOrderControl, fltAsc
'End Sub
'
'Private Sub OnRequeryClick()
'   MyFilter.Requery True
'End If
'##############################################################################

'##############################################################################
'#                                                                            #
'#  Change history:                                                           #
'#                                                                            #
'#  1.00:           First production release                                  #
'#  1.01:           Lots of bugfixes (too many to list)                       #
'#  1.02:           More bugfixes (to consider first stable release)          #
'#  1.03:           Some speed improvements                                   #
'#  1.50:           Rewrite of the module as class                            #
'#  1.60:           Added support for indexes with other types than long      #
'#                  Ascending and descending order possible                   #
'#                  Enabled switching of most parameters when active          #
'#                  e.g. you can apply all filters set on one form on another #
'#                  with same datasource by just changing the form property   #
'#                  this is still to consider experimental and unsupported    #
'#                  should be 100% compatible with code written for 1.50      #
'#  1.61:           No longer neccessary to specify index data type           #
'#                  You should no longer set this as it is now ignored and    #
'#                  will probably be removed in the next version              #
'#                  Be sure to pass the correct data types!                   #
'#                  Added some fancy progress bars                            #
'#  1.65:           Fix: problems with non-numeric and floating-point indexes #
'#                  Requery did not behave correctly when the form was in a   #
'#                  state without data at the time of the call                #
'#                  Removing multiple filters at once was not alway possible  #
'#                  also improved the speed of the remove by scanning the     #
'#                  filter list from back to front and placing the redim      #
'#                  outside of the loop (probably barely noticable)           #
'#                  Multiple identical Filters will not be added any longer   #
'#                  Removed index type enum and property                      #
'#                  Slightly improved error handling                          #
'#  2.00:           Finally sorting for any field possible!                   #
'#                  Old Order property is still implemented but will be       #
'#                  removed in future versions!                               #
'#                                                                            #
'##############################################################################

'####################### public enums used in parameters ######################
Public Enum fltOrder

    fltAsc
    fltDesc

End Enum

Public Enum fltCompare

    fltEqual
    fltLike
    fltLess
    fltGreater
    fltLessOrEqual
    fltGreaterOrEqual
    fltIs
    
End Enum


'####################### private variables ####################################
Private Enum errError

    errInternal
    errRSNoData
    errRSInvalidSyntax
    errNotActive
    errNoData

End Enum

'fieldnames used in the filters
Private Fields() As String

'criterias applied to the fields
Private Criterias() As Variant

'logical compare operators
Private Compares() As fltCompare

'descriptions of the filters
Private Names() As String

Private MyDb As DAO.Database

'form to apply the filters on
Private frmForm As Form

'name of the field with a unique Index in the Data source to be used
Private strIndex As String

'control on the form with the controlsource of the indexed field
Private ctrlIndex As Control

'order ascendung or descending
Private enmOrder As fltOrder
Private enmOrderByOrder As fltOrder

'name of the field to order by
Private strOrderBy As String

'control on the form with the controlsource of the order by field
Private ctrlOrderBy As Control

'string containing the raw recordsource (without the filters)
Private strRecordSource As String

'count of current applied filters
Private lngFilterCount As Long

Private isActive As Boolean

'##############################################################################
'#  MoveToIndex:    move to the record with the specified key value           #
'#  THIS WORKS ONLY IF THE FORM IS SORTED ASCENDING ON THE KEY!               #
'#  Will be replaced with a generic search function in future versions        #
'##############################################################################
Public Function IndexSeek( _
    ByVal MoveToIndex As Variant) _
    As Long
    
    Dim rst As DAO.Recordset
    Dim oldIndex As Variant
    Dim strSQL As String
    
    Dim rs As String
    SysCmd acSysCmdInitMeter, cstrSeekMsg, 10
    
    rs = BuildRecordSource
    If lngFilterCount = 0 Then
        strSQL = "SELECT COUNT(*) FROM (" & rs & " WHERE " & strIndex & _
            " <= " & EscapeSQLParameter(MoveToIndex) & _
            SQLOrder() & ");"
    Else
        strSQL = "SELECT COUNT(*) FROM (" & rs & " AND " & strIndex & _
            " <= " & EscapeSQLParameter(MoveToIndex) & _
             SQLOrder() & ");"
    End If
    SysCmd acSysCmdUpdateMeter, 1
    Set rst = MyDb.OpenRecordset(strSQL, dbOpenForwardOnly)
    SysCmd acSysCmdUpdateMeter, 5
    If rst.Fields(0) > 0 Then
        IndexSeek = rst.Fields(0)
        oldIndex = ctrlIndex
        If ctrlIndex <> MoveToIndex Then
            DoCmd.GoToRecord acDataForm, frmForm.Name, acGoTo, rst.Fields(0)
        End If
        If ctrlIndex <> MoveToIndex Then
            IndexSeek oldIndex
            IndexSeek = -1
        End If
    Else
        IndexSeek = -1
    End If
    
    rst.Close
    Set rst = Nothing
    SysCmd acSysCmdRemoveMeter
End Function

Private Function strFilterType( _
    ByVal fltType As fltCompare) _
    As String

    Select Case fltType
    
        Case fltEqual
            strFilterType = " = "
        Case fltLike
            strFilterType = " LIKE "
        Case fltLess
            strFilterType = " < "
        Case fltGreater
            strFilterType = " > "
        Case fltLessOrEqual
            strFilterType = " <= "
        Case fltGreaterOrEqual
            strFilterType = " >= "
        Case fltIs
            strFilterType = " IS "
        Case Else
            Err.Raise 600, "fltType", "Compare '" & fltType & _
            "' isn't specified"
            
    End Select
    
End Function

Private Function BuildRecordSource() As String

    Dim i As Long
    
    BuildRecordSource = strRecordSource
    If lngFilterCount > 0 Then
        BuildRecordSource = BuildRecordSource & " WHERE"
    End If
    For i = 0 To lngFilterCount - 1
        If i > 0 Then BuildRecordSource = BuildRecordSource & " AND"
        
        BuildRecordSource = BuildRecordSource & " " & Fields(i) & _
            " " & strFilterType(Compares(i)) & " " & _
            EscapeSQLParameter(Criterias(i))
    Next i
End Function

Private Sub Class_Initialize()
    Set MyDb = CurrentDb()

    enmOrder = fltAsc
    enmOrderByOrder = fltAsc
    
    lngFilterCount = 0
    isActive = False
End Sub

Private Sub Class_Terminate()
On Error GoTo destroyErrorHandler
    Set frmForm = Nothing
    Set ctrlIndex = Nothing
    MyDb.Close
    Set MyDb = Nothing
    isActive = False
    Exit Sub
destroyErrorHandler:
    Exit Sub
End Sub

Public Property Get Form() As Form
    Set Form = frmForm
End Property

Public Property Set Form( _
    ByRef NewForm As Form)
   
    Set frmForm = NewForm
    If isActive Then
        Requery
    End If
End Property

Public Property Get Index() As String
    Index = strIndex
End Property

Public Property Let Index( _
    ByVal vNewIndex As String)
    
    strIndex = vNewIndex
    strOrderBy = vNewIndex
    If isActive Then
       Requery
    End If
End Property

Public Property Get IndexControl() As Control
    Set IndexControl = ctrlIndex
End Property

Public Property Set IndexControl( _
    ByRef NewControl As Control)

    Set ctrlIndex = NewControl
    Set ctrlOrderBy = NewControl
    
    If isActive Then
       Requery
    End If
End Property

Public Property Get RecordSource() As String
    RecordSource = strRecordSource & ";"
End Property

Public Property Let RecordSource( _
    ByRef pNewRecordSource As String)
    
    Dim lngPos As Long
    
    lngPos = InStr(1, pNewRecordSource, ";")
    
    If lngPos > 0 Then
        strRecordSource = Left$(pNewRecordSource, lngPos - 1)
    Else
        strRecordSource = pNewRecordSource
    End If
    
    If isActive Then
        RemoveAll
    End If
End Property

Public Property Get Active() As Boolean
    Active = isActive
End Property

Public Property Get Count() As Long
     Count = lngFilterCount
End Property

Public Property Get Description( _
    Optional vstrSeperator As String = ", " _
    ) As String

    Dim i As Long
    Dim FirstFilter As Boolean
    Dim strDescription As String
    
    FirstFilter = True
    
    For i = 0 To lngFilterCount - 1
        If Names(i) <> "" Then
            If FirstFilter Then
                strDescription = Names(i)
                FirstFilter = False
            Else
                 strDescription = strDescription & vstrSeperator & Names(i)
            End If
        End If
    Next i
    
    Description = strDescription
    
End Property

Public Property Let Active( _
    ByVal vNewActiveState As Boolean)
    
    On Error GoTo LetActiveError
            
    Dim rst As DAO.Recordset
    Dim strSQL As String
      
    If vNewActiveState = False Then
        isActive = False
    Else
        
        If Not isActive Then
                   
            'check for valid recordsource
            strSQL = BuildRecordSource & SQLOrder
            
            Set rst = MyDb.OpenRecordset(strSQL, dbOpenDynaset)
        
            If rst.EOF Then
                Err.Raise cerrBase + errRSNoData
            Else
                isActive = True
            End If
            
        End If
    End If
    rst.Close
    
LetActiveExit:

    Set rst = Nothing
    Exit Property

LetActiveError:
       
    isActive = False
    
    Select Case Err.Number
        Case vbObjectError + errRSNoData
            Err.Raise cerrBase + errRSNoData, "cFilter::Active", _
            "The recordsource you supplied does not return any data!"
        Case Else
            Err.Raise cerrBase + errRSInvalidSyntax, "cFilter::Active", _
            "The recordsource you supplied has an invalid syntax!" & _
                vbNewLine _
                & "SELECT clause may contain joins but no WHERE, ORDER BY " & _
                "or aggregate functions"
    End Select
    Resume LetActiveExit
End Property
Public Function Add( _
    ByVal vField As String, _
    ByVal vCompare As fltCompare, _
    ByVal vCriteria As Variant, _
    Optional ByVal vRefresh As Boolean = True, _
    Optional ByVal vName As String = "") _
    As Long
'##############################################################################
'#  vField:         Field in your Recordsource the Filter will be applied on  #
'#  vCompare:       Compare mode for the Filter                               #
'#  vCriteria:      Value for the Criteria                                    #
'#                                                                            #
'#  vRefresh:       If set to false the Filter is stored but not yet applied  #
'# (optional)       Notice: Checking for Filter to supply data is disabled    #
'#                  if false. Use the Requery method to apply manually.       #
'#                  You can use this to add multiple filters without the      #
'#                  need to requery the form after each add call.             #
'#                  If you add invalid filters the Requery method will fail   #
'#                  and you will have to remove all filters as you can't      #
'#                  determine which is the one that failed!                   #
'#                                                                            #
'#  vName:          Name for the applied filter. You can query all applied    #                                                                                    #
'# (optional)       filters with the Description property.                    #
'#                                                                            #
'# Return value:    >0 on Success                                             #
'##############################################################################
On Error GoTo AddError

    
    Dim lngSuccess As Long
    Dim i As Long
    Dim j As Long

    For i = 0 To lngFilterCount - 1
        If Fields(i) = vField _
            And Criterias(i) = vCriteria _
            And Compares(i) = vCompare _
            And Names(i) = vName Then
            
            Add = 1
            GoTo AddExit
        End If
    Next i
    'add one element
    ReDim Preserve Fields(lngFilterCount + 1)
    ReDim Preserve Criterias(lngFilterCount + 1)
    ReDim Preserve Compares(lngFilterCount + 1)
    ReDim Preserve Names(lngFilterCount + 1)
             
    Fields(lngFilterCount) = vField
    Criterias(lngFilterCount) = vCriteria
    Compares(lngFilterCount) = vCompare
    Names(lngFilterCount) = vName
    
    
    lngFilterCount = lngFilterCount + 1
    
    If vRefresh And isActive Then
    lngSuccess = Requery()
        If lngSuccess < 1 Then
        'just added filter does not retrieve any data -> remove it again
            Remove vField, vCriteria
        End If
    End If
           
    Add = lngSuccess
           
AddExit:

    Exit Function
    
AddError:
    Err.Raise cerrBase + errInternal, "cFiler::Add", _
        Err.Description
    Resume AddExit
End Function

Public Sub Remove( _
    ByVal vField As String, _
    ByVal vCriteria As String, _
    Optional ByVal vRefresh As Boolean = True)
'##############################################################################
'#  vField:         Field to delete                                           #
'#  vCriteria:      Value for the Criteria, can contain wildcards to remove   #
'#                  multiple Filters at once. Check the Count property before #
'#                  and after to get the number of removed Filters.           #
'#                                                                            #
'#  vRefresh:       If set to false the Filter is stored but not yet applied  #
'# (optional)       Notice: Checking for Filter to supply data is disabled    #
'#                  if set to false. Use the Requery method to apply manualy  #
'#                                                                            #
'##############################################################################

On Error GoTo RemoveError

    Dim i As Long
    Dim j As Long
    Dim lngRemoveCount As Long

    lngRemoveCount = 0

    For i = (lngFilterCount - 1) To 0 Step -1

        If Fields(i) = vField And Criterias(i) Like vCriteria Then

            For j = i To lngFilterCount - 2
            'move up following elements
                Fields(j) = Fields(j + 1)
                Criterias(j) = Criterias(j + 1)
                Names(j) = Names(j + 1)
                Compares(j) = Compares(j + 1)
            Next j

            'correct array sizes
            lngFilterCount = lngFilterCount - 1
            lngRemoveCount = lngRemoveCount + 1
            
        End If

    Next i

    ReDim Preserve Criterias(lngFilterCount)
    ReDim Preserve Fields(lngFilterCount)
    ReDim Preserve Names(lngFilterCount)
    ReDim Preserve Compares(lngFilterCount)

    If vRefresh And lngRemoveCount > 0 Then Requery

RemoveExit:
    
    Exit Sub
 
RemoveError:
    Err.Raise cerrBase + errInternal, "cFiler::Remove", _
        Err.Description
        
    Exit Sub
End Sub

Public Sub RemoveAll( _
    Optional ByVal vRefresh As Boolean = True)
    
    On Error GoTo RemoveAllError

    ReDim FilterCriterias(0)
    ReDim Filters(0)
    ReDim FilterTypes(0)
    lngFilterCount = 0

    If vRefresh Then Requery

RemoveAllExit:

    Exit Sub
    
RemoveAllError:

    Err.Raise cerrBase + errInternal, "cFiler::RemoveAll", _
        Err.Description
    Exit Sub

End Sub

Public Function Requery(Optional ByVal forced As Boolean = False) As Long
'##############################################################################
'#                                                                            #
'#  Applies all Filters remaining the current record if possible.             #
'#  Returns >0 when successful                                                #
'#  This is the function doing all the dirty work ;)                          #
'#                                                                            #
'##############################################################################

On Error GoTo RequeryError

Dim rst As DAO.Recordset
Dim strSQL As String
Dim strOldRecordSource As String
Dim strNewRecordSource As String
Dim varIndex As Variant
Dim varField As Variant
    
Dim strOrder As String
    
    strOrder = SQLOrder()
    
    SysCmd acSysCmdInitMeter, cstrWaitMsg, 20
    
    strNewRecordSource = BuildRecordSource
    SysCmd acSysCmdUpdateMeter, 1
    strOldRecordSource = frmForm.RecordSource
    If (strOldRecordSource <> strNewRecordSource & strOrder) Or forced Then
       
        
        If HasValue(ctrlIndex) Then
            varIndex = ctrlIndex.Value
            varField = ctrlOrderBy.Value
            
            SysCmd acSysCmdUpdateMeter, 2
            frmForm.RecordSource = strNewRecordSource & strOrder
                
            SysCmd acSysCmdUpdateMeter, 10
            If frmForm.RecordsetClone.RecordCount < 1 Then
                Requery = -1
                SysCmd acSysCmdUpdateMeter, 20
                GoTo RequeryExit
            End If
            If ctrlIndex <> varIndex Then
            'move to the old record if applicable
                
                '### calculate count of records to jump forward in the form
                If strIndex <> strOrderBy Then
                    'number of records to move _from_
                    'the first the order by field
                    strSQL = _
                        "(SELECT COUNT(*) FROM (" & _
                        strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & _
                        strIndex & IIf(enmOrder = fltDesc, ">=", "<=") & _
                        EscapeSQLParameter(varIndex) & " AND " & _
                        strIndex & IIf(enmOrder = fltDesc, "<=", ">=") & _
                        "(SELECT " & IIf(enmOrder = fltDesc, "Max", "Min") & _
                        "(" & strIndex & ") FROM (" & strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & strOrderBy & _
                        IIf(IsNull(varField), _
                            " IS NULL ", _
                            "=" & EscapeSQLParameter(varField)) & _
                        ")) " & _
                        "AND " & strOrderBy & _
                        IIf(IsNull(varField), _
                            " IS NULL ", _
                            "=" & _
                            EscapeSQLParameter(varField)) & _
                        ")) "
                    
                    'number of records to move _to_
                    'the first the order by field
                    If Not IsNull(varField) And IsNull(ctrlOrderBy) Then
                        strSQL = "SELECT COUNT (*) + " & strSQL & _
                        "FROM (" & strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & _
                        strOrderBy & _
                        IIf(enmOrderByOrder = fltDesc, ">", "<") & _
                        EscapeSQLParameter(varField) & " " & _
                        IIf(enmOrderByOrder = fltDesc, "AND ", "OR ") & _
                        strOrderBy & " IS NULL" & ")"
                    ElseIf IsNull(varField) And Not IsNull(ctrlOrderBy) Then
                        strSQL = "SELECT COUNT (*) + " & strSQL & _
                        "FROM (" & strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & _
                        strOrderBy & _
                        IIf(enmOrderByOrder = fltDesc, "<=", ">=") & _
                        EscapeSQLParameter(ctrlOrderBy) & ")"
                    ElseIf IsNull(varField) And IsNull(ctrlOrderBy) Then
                        'strSQL = strSQL
                    Else
                        strSQL = "SELECT COUNT (*) + " & strSQL & _
                        "FROM (" & strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & _
                        strOrderBy & _
                        IIf(enmOrderByOrder = fltDesc, "<=", ">=") & _
                        EscapeSQLParameter(ctrlOrderBy) & " AND " & _
                        strOrderBy & _
                        IIf(enmOrderByOrder = fltDesc, ">", "<") & _
                        EscapeSQLParameter(varField) & ")"
                    End If
                Else
                    'only ordered by key
                    strSQL = _
                        "SELECT COUNT(*) FROM (" & _
                        strNewRecordSource & " " & _
                        IIf(lngFilterCount = 0, "WHERE ", "AND ") & _
                        strIndex & IIf(enmOrder = fltDesc, ">=", "<=") & _
                        EscapeSQLParameter(varIndex) & " AND " & _
                        strIndex & IIf(enmOrder = fltDesc, "<=", ">=") & _
                        EscapeSQLParameter(ctrlIndex) & ")"
                End If
                
                SysCmd acSysCmdUpdateMeter, 11
                               
                Set rst = MyDb.OpenRecordset(strSQL, dbOpenSnapshot)
                               
                If rst.Fields(0) > 0 Then
                    DoCmd.GoToRecord acDataForm, frmForm.Name, _
                        acGoTo, rst.Fields(0)
                End If
                SysCmd acSysCmdUpdateMeter, 20
                Requery = 1
                
                rst.Close
            Else
            'current record was already the old one
                Requery = 1
            End If
            
        Else
        'the form currently has no valid data
            frmForm.RecordSource = strNewRecordSource & SQLOrder()
            
            SysCmd acSysCmdUpdateMeter, 10
            
            If HasValue(ctrlIndex) Then
                Requery = 1
            Else
            'the form STILL has no valid data
                Requery = -1
                frmForm.RecordSource = strOldRecordSource
            End If
        End If
    Else
        'no need to change recordsource
        Requery = 1
    End If
    SysCmd acSysCmdRemoveMeter
RequeryExit:
    Set rst = Nothing
    Exit Function
    
RequeryError:
    SysCmd acSysCmdRemoveMeter
    Err.Raise cerrBase + errInternal, "cFilter::Requery", _
        Err.Description
    
    Resume RequeryExit
End Function
Private Function HasValue( _
    obj As Object) _
    As Boolean

On Error GoTo isDataErr
    HasValue = False
    
    Dim dummy As Variant
    dummy = obj.Value
    If IsNull(dummy) Then Exit Function
    If IsEmpty(dummy) Then Exit Function
    
    HasValue = True
    Exit Function
isDataErr:
    HasValue = False
    Exit Function
End Function
Private Function SQLOrder() As String
    Dim strOrderDir As String
    Dim strOrderByDir As String

    If enmOrder = fltDesc Then
        strOrderDir = " DESC "
    Else
        strOrderDir = " ASC "
    End If
    
    If enmOrderByOrder = fltDesc Then
        strOrderByDir = " DESC "
    Else
        strOrderByDir = " ASC "
    End If

    If strOrderBy <> strIndex Then
        SQLOrder = _
            " ORDER BY " & strOrderBy & strOrderByDir & _
            ", " & strIndex & strOrderDir
    Else
        SQLOrder = " ORDER BY " & strIndex & strOrderDir
    End If
    
   
End Function

Public Property Get Order() As fltOrder
    Order = enmOrder
End Property

Public Property Let Order( _
    ByVal vNewOrder As fltOrder)
    
    enmOrder = fltAsc
If isActive Then
    Requery
End If

End Property

Public Property Get Version()
    Version = cdblVersion
End Property

Public Function OrderBy( _
    ByVal vstrOrderByField, _
    ByRef pctrlOrderByControl, _
    ByVal OrderAscDesc As fltOrder, _
    Optional ByVal vRefresh As Boolean = True)
    
    Dim strOldField As String
    Dim ctrlOld As Control
    Dim OldOrder As fltOrder
    
    Dim lngSuccess As Long
    
    strOldField = strOrderBy
    Set ctrlOld = ctrlOrderBy
    OldOrder = enmOrderByOrder
    
    strOrderBy = vstrOrderByField
    Set ctrlOrderBy = pctrlOrderByControl
    enmOrderByOrder = OrderAscDesc
    
    If strOrderBy = strIndex Then
        enmOrder = OrderAscDesc
    End If
    
    If vRefresh And Active Then
        lngSuccess = Requery()
        OrderBy = lngSuccess
        If Not lngSuccess > 0 Then
            strOrderBy = strOldField
            Set ctrlOrderBy = ctrlOld
            enmOrderByOrder = OldOrder
            Requery
        End If
    End If
End Function

'Utility functions
'##############################################################################
Public Function HasValue(obj As Object) As Boolean
'Does the object have a "value"?

On Error GoTo isDataErr
    HasValue = False
    
    Dim dummy As Variant
    dummy = obj.Value
    If IsNull(dummy) Then Exit Function
    If IsEmpty(dummy) Then Exit Function
    
    HasValue = True
    Exit Function
isDataErr:
    HasValue = False
    Exit Function
End Function

Public Function EscapeSQLParameter( _
    ByVal vParameter As Variant) _
    As String
' This function converts the usual vb data types in
' string compatible with jet sql. You should cast the
' parameter explicitely to the data type you want it
' to have (e.g. when you read out a date from an edit
' box you need to cast it to date with cdate(myfield)
' as it's otherwise string type which will result in
' wrong handling
    Dim strBuffer As String
    Dim lngSeperatorPos As Long
    
    Dim parType As VbVarType
    
    parType = VarType(vParameter)
    
    Select Case parType
    
        Case vbNull
            strBuffer = "NULL"
        Case vbObject
            Err.Raise 443 'object has no default value
        Case vbBoolean
            If vParameter Then
                strBuffer = "true"
            Else
                strBuffer = "false"
            End If
        Case vbByte, vbInteger, vbLong
            strBuffer = CStr(vParameter)
        Case vbCurrency, vbDecimal, vbSingle, vbDouble
            strBuffer = Str(vParameter) 'str always uses . as decimal point character
        Case vbDate
            'converting to double seems to be the
            'safest way for jet workspaces
            strBuffer = EscapeSQLParameter((CDbl(vParameter)))
        Case vbString
            strBuffer = "'" & EscapeSQLString(vParameter) & "'"
        Case Else
            Err.Raise 13 'type mismatch
    End Select
    
    EscapeSQLParameter = strBuffer
End Function

Public Function EscapeSQLString( _
    ByVal Text As Variant) As String
'Escapes ' character in Strings
'by John DeHope

Dim x

If IsNull(Text) Then
    EscapeSQLString = ""
Else
    x = 1
    Do While x <= Len(Text)
        
        If Mid(Text, x, 1) = "'" Then
            Text = Left(Text, x - 1) & "''" & Right(Text, Len(Text) - x)
            x = x + 2
        Else
            x = x + 1
        End If
        
    Loop
    EscapeSQLString = Text
End If
End Function

