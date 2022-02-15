Attribute VB_Name = "InspectWorkbookUtilsTest"
'@TestModule
'@Folder "InspectWorkbookUtilsProject.Tests"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'WhitelistedIdentifiers i, j
Option Explicit
Option Private Module

'@Ignore VariableNotUsed
Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

Private Const DebugObjectSheetName As String = "Sheet1"
Private Const DummyCodeString As String = _
"Private Sub Dummy()" & vbCrLf & _
"End Sub"
Private Const SkipTaskPaneAddInsTest As Boolean = True

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Function SetScreenUpdating(ByVal ScreenUpdating As Boolean) As Boolean
    Dim currentScreenUpdating As Boolean
    
    currentScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = ScreenUpdating
    
    SetScreenUpdating = currentScreenUpdating
End Function

Private Function CreateNewWorkbook(ByVal SheetsCount As Long) As Workbook
    Dim workbook_ As Workbook
    
    Set workbook_ = Application.Workbooks.Add
    Dim i As Long
    Dim currentDisplayAlerts As Boolean
    With workbook_
        If .Worksheets.Count <= SheetsCount Then
            For i = .Worksheets.Count + 1 To SheetsCount
                Call .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
            Next
        Else
            For i = .Worksheets.Count To SheetsCount + 1 Step -1
                currentDisplayAlerts = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Call .Worksheets(i).Delete
                Application.DisplayAlerts = currentDisplayAlerts
            Next
        End If
    End With
    
    Set CreateNewWorkbook = workbook_
End Function

Private Function CreateDummyImage(ByVal OutputPath As String) As String
    Dim workbook_ As Workbook
    Set workbook_ = Workbooks.Add
    Dim chart_ As chart
    Set chart_ = workbook_.Worksheets(1).ChartObjects.Add(0, 0, 1, 1).chart
    Call chart_.Export(OutputPath)
    Call chart_.Parent.Delete
    Call workbook_.Close(False)
    CreateDummyImage = OutputPath
End Function

Private Function DictionaryEquals(ByVal One As Object, ByVal Another As Object) As Boolean
    If One Is Nothing Or Another Is Nothing Then
        DictionaryEquals = False
        Exit Function
    End If
    
    If TypeName(One) <> "Dictionary" Or TypeName(Another) <> "Dictionary" Then
        DictionaryEquals = False
        Exit Function
    End If
    
    If One Is Another Then
        DictionaryEquals = True
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In One.Keys
        If Not Another.Exists(key) Then
            DictionaryEquals = False
            Exit Function
        End If
    Next
    For Each key In Another.Keys
        If Not One.Exists(key) Then
            DictionaryEquals = False
            Exit Function
        End If
    Next
    
    For Each key In One.Keys
        If One(key) <> Another(key) Then
            DictionaryEquals = False
            Exit Function
        End If
    Next
    
    DictionaryEquals = True
End Function

Private Function DictionariesContains(ByVal Dictionaries As Collection, ByVal Dictionary As Object) As Boolean
    If Dictionaries Is Nothing Or Dictionary Is Nothing Then
        DictionariesContains = False
        Exit Function
    End If
    
    If TypeName(Dictionary) <> "Dictionary" Then
        DictionariesContains = False
        Exit Function
    End If
    
    Dim dict As Object
    For Each dict In Dictionaries
        If TypeName(dict) = "Dictionary" Then
            If DictionaryEquals(Dictionary, dict) Then
                DictionariesContains = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function DictionariesEquals(ByVal One As Collection, ByVal Another As Collection) As Boolean
    If One Is Nothing Or Another Is Nothing Then
        DictionariesEquals = False
        Exit Function
    End If
    
    If One Is Another Then
        DictionariesEquals = True
        Exit Function
    End If
    
    Dim dict As Object
    For Each dict In One
        If Not DictionariesContains(Another, dict) Then
            DictionariesEquals = False
            Exit Function
        End If
    Next
    For Each dict In Another
        If Not DictionariesContains(One, dict) Then
            DictionariesEquals = False
            Exit Function
        End If
    Next
    DictionariesEquals = True
End Function

Public Function GetParentWorkbook(ByVal TargetObject As Object) As Workbook
    If TypeOf TargetObject Is Workbook Then
        Set GetParentWorkbook = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentWorkbook = GetParentWorkbook(TargetObject.Parent)
End Function

Public Function GetParentWorksheet(ByVal TargetObject As Object) As Worksheet
    If TypeOf TargetObject Is Worksheet Then
        Set GetParentWorksheet = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Workbook Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentWorksheet = GetParentWorksheet(TargetObject.Parent)
End Function

Public Function GetParentRange(ByVal TargetObject As Object) As Range
    If TypeOf TargetObject Is Range Then
        Set GetParentRange = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Worksheet Or _
            TypeOf TargetObject Is Workbook Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentRange = GetParentRange(TargetObject.Parent)
End Function

Private Function GetWorkbookLocation(ByVal TargetWorkbook As Workbook) As String
    Dim externalAddressWithoutRange As String
    externalAddressWithoutRange = GetWorksheetLocation(TargetWorkbook.Worksheets(1))
    
    GetWorkbookLocation = Left$(externalAddressWithoutRange, InStrRev(externalAddressWithoutRange, "]"))
End Function

Private Function GetWorksheetLocation(ByVal TargetWorksheet As Worksheet) As String
    Dim externalAddress_ As String
    externalAddress_ = TargetWorksheet.Cells(1, 1).Address(External:=True)
    
    GetWorksheetLocation = Left$(externalAddress_, InStrRev(externalAddress_, "!") - 1)
End Function

'' = Initialize =

'@TestMethod("Initialize")
Public Sub Initialize_CorrectCall_Succeeded()
    Dim instance As InspectWorkbookUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectWorkbookUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisWorkbook)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_CallTwice_RaiseError()
    Dim instance As InspectWorkbookUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectWorkbookUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisWorkbook)
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisWorkbook)
    
    Assert.Fail
    
    GoTo FINALLY
CATCH:

    Assert.AreEqual instance.InvalidOperationError, Err.Number
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_TargetWorkbookIsNothing_RaiseError()
    Dim instance As InspectWorkbookUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectWorkbookUtils
    Call instance.Initialize(Nothing)
    
    Assert.Fail
    
    GoTo FINALLY
CATCH:
    Assert.AreEqual instance.InvalidArgumentError, Err.Number
    
    Resume FINALLY
FINALLY:
End Sub

'' = For all Office documents =
'' == Comments ==
Private Function DecolateWorksheetForCommentsTest(ByVal TargetWorksheet As Worksheet, ByRef visibleCommentsCount As Long, ByRef hiddenCommentsCount As Long) As Worksheet
    visibleCommentsCount = 0
    hiddenCommentsCount = 0
    
    Dim i As Long
    Dim newComment As Comment
    For i = 1 To 5
        Set newComment = TargetWorksheet.Cells(i, i).AddComment("Comment" & i)
        If i Mod 2 = 1 Then
            newComment.Visible = True: visibleCommentsCount = visibleCommentsCount + 1
        Else
            hiddenCommentsCount = hiddenCommentsCount + 1
        End If
    Next
    
    Set DecolateWorksheetForCommentsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForCommentsTest(ByVal Target As Workbook, ByRef visibleCommentsCount As Long, ByRef hiddenCommentsCount As Long) As Workbook
    visibleCommentsCount = 0
    hiddenCommentsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim visibleCommentsCountInWorksheet As Long
    Dim hiddenCommentsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForCommentsTest(worksheet_, visibleCommentsCountInWorksheet, hiddenCommentsCountInWorksheet): visibleCommentsCount = visibleCommentsCount + visibleCommentsCountInWorksheet: hiddenCommentsCount = hiddenCommentsCount + hiddenCommentsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForCommentsTest = Target
End Function

'@TestMethod("Comments")
Public Sub ListCommentsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    
    Dim visibleCommentsCountInWorksheet As Long
    Dim hiddenCommentsCountInWorksheet As Long
    Call DecolateWorksheetForCommentsTest(testWorksheet, visibleCommentsCountInWorksheet, hiddenCommentsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedComments As Collection
    Set listedComments = instance.ListCommentsInWorksheet(testWorksheet)
    
    Assert.AreEqual visibleCommentsCountInWorksheet + hiddenCommentsCountInWorksheet, listedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub ListComments_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    On Error GoTo CATCH
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments()
    
    Assert.AreEqual visibleCommentsCount + hiddenCommentsCount, listedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub VisualizeCommentsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    
    Dim visibleCommentsCountInWorksheet As Long
    Dim hiddenCommentsCountInWorksheet As Long
    Call DecolateWorksheetForCommentsTest(testWorksheet, visibleCommentsCountInWorksheet, hiddenCommentsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim visualizedCommentsInWorksheet As Collection
    Set visualizedCommentsInWorksheet = instance.VisualizeCommentsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenCommentsCountInWorksheet, visualizedCommentsInWorksheet.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub VisualizeComments_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    On Error GoTo CATCH
    
    Dim visualizedComments As Collection
    Set visualizedComments = instance.VisualizeComments
    
    Assert.AreEqual hiddenCommentsCount, visualizedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub DeleteCommentsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    
    Dim visibleCommentsCountInWorksheet As Long
    Dim hiddenCommentsCountInWorksheet As Long
    Call DecolateWorksheetForCommentsTest(testWorksheet, visibleCommentsCountInWorksheet, hiddenCommentsCountInWorksheet)
    
    Dim listedComments As Collection
    Set listedComments = instance.ListCommentsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim comment_ As Comment
    Dim dict As Object
    For Each comment_ In listedComments
        Set dict = CreateObject("Scripting.Dictionary")
        With comment_
            dict("Type") = TypeName(comment_)
            dict("Location") = GetParentRange(comment_).Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedCommentsInWorksheet As Collection
    Set deletedCommentsInWorksheet = instance.DeleteCommentsInWorksheet(testWorksheet)
    
    Assert.AreEqual visibleCommentsCountInWorksheet + hiddenCommentsCountInWorksheet, deletedCommentsInWorksheet.Count
    Assert.AreEqual visibleCommentsCount + hiddenCommentsCount, instance.ListComments.Count
    Assert.IsTrue DictionariesEquals(deletedCommentsInWorksheet, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub DeleteCommentsTest_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim comment_ As Comment
    Dim dict As Object
    For Each comment_ In listedComments
        Set dict = CreateObject("Scripting.Dictionary")
        With comment_
            dict("Type") = TypeName(comment_)
            dict("Location") = GetParentRange(comment_).Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedComments As Collection
    Set deletedComments = instance.DeleteComments
    
    Assert.AreEqual visibleCommentsCount + hiddenCommentsCount, deletedComments.Count
    Assert.AreEqual 0&, instance.ListComments.Count
    Assert.IsTrue DictionariesEquals(deletedComments, expectedDictionaries)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub RemoveCommentsTest_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleCommentsCount As Long
    Dim hiddenCommentsCount As Long
    Call DecolateWorkbookForCommentsTest(instance.Target, visibleCommentsCount, hiddenCommentsCount)
    
    On Error GoTo CATCH
    
    instance.RemoveComments
    
    Assert.AreEqual 0&, instance.ListComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@TestMethod("DocumentProperties")
Public Sub RemoveDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateWorkbookForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    instance.RemoveDocumentProperties
    
    Assert.AreEqual 0&, instance.ListCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' ==== Built-in Document Properties ====
Private Function DecolateWorkbookForBuiltInDocumentProperties(ByVal Target As Workbook, ByRef builtInDocumentPropertiesCount As Long) As Workbook
    builtInDocumentPropertiesCount = 0
    
    Dim property As DocumentProperty
    For Each property In Target.builtInDocumentProperties
        With property
            Select Case .Type
                Case msoPropertyTypeNumber
                    On Error Resume Next
                    .value = 1
                    On Error GoTo 0
                Case msoPropertyTypeBoolean
                    On Error Resume Next
                    .value = True
                    On Error GoTo 0
                Case msoPropertyTypeDate
                    On Error Resume Next
                    .value = Now
                    On Error GoTo 0
                Case msoPropertyTypeString
                    On Error Resume Next
                    .value = "Dummy"
                    On Error GoTo 0
                Case msoPropertyTypeFloat
                    On Error Resume Next
                    .value = 1.5
                    On Error GoTo 0
            End Select
        End With
        builtInDocumentPropertiesCount = builtInDocumentPropertiesCount + 1
    Next
    
    Set DecolateWorkbookForBuiltInDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateWorkbookForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedBuiltInDocumentProperties As Collection
    Set listedBuiltInDocumentProperties = instance.ListBuiltInDocumentProperties
    
    Assert.AreEqual builtInDocumentPropertiesCount, listedBuiltInDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateWorkbookForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim clearedDocumentProperties As Collection
    Set clearedDocumentProperties = instance.ClearBuiltInDocumentProperties
    
    Assert.AreEqual builtInDocumentPropertiesCount, clearedDocumentProperties.Count
    Assert.AreEqual builtInDocumentPropertiesCount, clearedDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' ==== Custom Document Properties ====
Private Function DecolateWorkbookForCustomDocumentProperties(ByVal Target As Workbook, ByRef customDocumentPropertiesCount As Long) As Workbook
    customDocumentPropertiesCount = 0
    
    Dim i As Long
    Dim value(1 To 5) As Variant
    value(msoPropertyTypeNumber) = 1
    value(msoPropertyTypeBoolean) = True
    value(msoPropertyTypeDate) = Now
    value(msoPropertyTypeString) = "string"
    value(msoPropertyTypeFloat) = 1.5
    For i = 1 To 5
        Call Target.customDocumentProperties.Add( _
            Name:="Custom" & i, _
            LinkToContent:=False, _
            Type:=i, _
            value:=value(i) _
        )
        customDocumentPropertiesCount = customDocumentPropertiesCount + 1
    Next
    
    Set DecolateWorkbookForCustomDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListCustomDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateWorkbookForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomDocumentProperties As Collection
    Set listedCustomDocumentProperties = instance.ListCustomDocumentProperties
    
    Assert.AreEqual customDocumentPropertiesCount, listedCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearCustomDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateWorkbookForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim clearedDocumentProperties As Collection
    Set clearedDocumentProperties = instance.ClearCustomDocumentProperties
    
    Assert.AreEqual customDocumentPropertiesCount, clearedDocumentProperties.Count
    Assert.AreEqual customDocumentPropertiesCount, clearedDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub DeleteCustomDocumentProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateWorkbookForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
    Dim listedCustomDocumentProperties As Collection
    Set listedCustomDocumentProperties = instance.ListCustomDocumentProperties
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim docProp As DocumentProperty
    Dim dict As Object
    For Each docProp In listedCustomDocumentProperties
        Set dict = CreateObject("Scripting.Dictionary")
        With docProp
            dict("Type") = TypeName(docProp)
            dict("Name") = .Name
            dict("Location") = GetWorkbookLocation(GetParentWorkbook(docProp))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedDocumentProperties As Collection
    Set deletedDocumentProperties = instance.DeleteCustomDocumentProperties
    
    Assert.AreEqual customDocumentPropertiesCount, deletedDocumentProperties.Count
    Assert.AreEqual 0&, instance.ListCustomDocumentProperties.Count
    Assert.IsTrue DictionariesEquals(deletedDocumentProperties, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Personal Information ===
'@TestMethod("PersonalInformation")
Public Sub RemovePersonalInformation_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    instance.RemovePersonalInformation
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Printer Path ===
'@TestMethod("RemovePrinterPaths")
Public Sub RemovePrinterPaths_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    instance.RemovePrinterPaths
    
    GoTo FINALLY
CATCH:
    Assert.Fail
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Content Add-ins ==
Private Function DecolateWorksheetForContentAddInsTest(ByVal TargetWorksheet As Worksheet, ByRef contentAddInsCount As Long) As Worksheet
    contentAddInsCount = 0
    
    Dim app As Shape
    Dim shape_ As Shape
    For Each shape_ In ThisWorkbook.Worksheets(DebugObjectSheetName).Shapes
        If shape_.Type = msoContentApp Then
            Set app = shape_
            Exit For
        End If
    Next
    
    If app Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 327, Source:="DecolateWorksheetForInkTest", Description:="For Ink testing, ink object must be in """ & DebugObjectSheetName & """ worksheet_.")
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 2
        Call Application.Wait(DateAdd("s", 3, Now))
        Call app.Copy
        Call Application.Wait(DateAdd("s", 3, Now))
        Call TargetWorksheet.Paste: contentAddInsCount = contentAddInsCount + 1
    Next
    
    Set DecolateWorksheetForContentAddInsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForContentAddInsTest(ByVal Target As Workbook, ByRef contentAddInsCount As Long) As Workbook
    contentAddInsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim contentAddInsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForContentAddInsTest(worksheet_, contentAddInsCountInWorksheet): contentAddInsCount = contentAddInsCount + contentAddInsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForContentAddInsTest = Target
End Function

'@TestMethod("ContentAddIns")
Public Sub ListContentAddInsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim contentAddInsCountInWorksheet As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorksheetForContentAddInsTest(testWorksheet, contentAddInsCountInWorksheet)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddInsInWorksheet(testWorksheet)
    
    Assert.AreEqual contentAddInsCountInWorksheet, listedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub ListContentAddIns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddIns
    
    Assert.AreEqual contentAddInsCount, listedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub ConvertContentAddInsToImagesInWorksheet_CollectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim contentAddInsCountInWorksheet As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorksheetForContentAddInsTest(testWorksheet, contentAddInsCountInWorksheet)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim convertedContentAddIns As Collection
    Set convertedContentAddIns = instance.ConvertContentAddInsToImagesInWorksheet(testWorksheet)
    
    Assert.AreEqual contentAddInsCountInWorksheet, convertedContentAddIns.Count
    Assert.AreEqual contentAddInsCount, instance.ListContentAddIns.Count
    Assert.AreEqual contentAddInsCountInWorksheet, testWorksheet.Pictures.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub ConvertContentAddInsToImages_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim convertedContentAddIns As Collection
    Set convertedContentAddIns = instance.ConvertContentAddInsToImages
    
    Dim worksheet_ As Worksheet
    Dim pictureCount As Long
    For Each worksheet_ In instance.Target.Worksheets
        pictureCount = pictureCount + worksheet_.Pictures.Count
    Next
    
    Assert.AreEqual contentAddInsCount, convertedContentAddIns.Count
    Assert.AreEqual 0&, instance.ListContentAddIns.Count
    Assert.AreEqual pictureCount, convertedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub DeleteContentAddInsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim contentAddInsCountInWorksheet As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorksheetForContentAddInsTest(testWorksheet, contentAddInsCountInWorksheet)
    On Error GoTo 0
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddInsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedContentAddIns
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedContentAddIns As Collection
    Set deletedContentAddIns = instance.DeleteContentAddInsInWorksheet(testWorksheet)
    
    Assert.AreEqual contentAddInsCountInWorksheet, deletedContentAddIns.Count
    Assert.AreEqual contentAddInsCount, instance.ListContentAddIns.Count
    Assert.IsTrue DictionariesEquals(deletedContentAddIns, expectedDictionaries)
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub DeleteContentAddIns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddIns
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedContentAddIns
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedContentAddIns As Collection
    Set deletedContentAddIns = instance.DeleteContentAddIns
    
    Assert.AreEqual contentAddInsCount, deletedContentAddIns.Count
    Assert.AreEqual 0&, instance.ListContentAddIns.Count
    Assert.IsTrue DictionariesEquals(deletedContentAddIns, expectedDictionaries)
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ContentAddIns")
Public Sub RemoveContentAddIns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateWorkbookForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    instance.RemoveContentAddIns
    
    Assert.AreEqual 0&, instance.ListContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Task Pane Add-ins ==
'@TestMethod("TaskPaneAddIns")
Public Sub RemoveTaskPaneAddIns_CorrectCall_Successed()
    If SkipTaskPaneAddInsTest Then
        Assert.Inconclusive "TaskPaneAddIns test was skipped."
        Exit Sub
    End If
    
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Beep
    '@Ignore StopKeyword
    Stop ' Add TaskPaneAddIns manually
    instance.RemoveTaskPaneAddIns
    '@Ignore StopKeyword
    Stop ' Check TaskPaneAddIns deletion by visual confirmation
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Embedded Documents ==
Private Function DecolateWorksheetForEmbeddedDocumentsTest(ByVal TargetWorksheet As Worksheet, ByVal DummyFilePath As String, ByRef EmbeddedDocumentsCount As Long) As Worksheet
    EmbeddedDocumentsCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call TargetWorksheet.OLEObjects.Add( _
            fileName:=DummyFilePath, _
            link:=False, _
            DisplayAsIcon:=False _
        )
        EmbeddedDocumentsCount = EmbeddedDocumentsCount + 1
    Next
    
    Set DecolateWorksheetForEmbeddedDocumentsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForEmbeddedDocumentsTest(ByVal Target As Workbook, ByVal DummyFilePath As String, ByRef EmbeddedDocumentsCount As Long) As Workbook
    EmbeddedDocumentsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim EmbeddedDocumentsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForEmbeddedDocumentsTest(worksheet_, DummyFilePath, EmbeddedDocumentsCountInWorksheet): EmbeddedDocumentsCount = EmbeddedDocumentsCount + EmbeddedDocumentsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForEmbeddedDocumentsTest = Target
End Function

'@TestMethod("EmbeddedDocuments")
Public Sub ListEmbeddedDocumentsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateWorkbookForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim EmbeddedDocumentsCountInWorksheet As Long
    Call DecolateWorksheetForEmbeddedDocumentsTest(testWorksheet, dummyFile, EmbeddedDocumentsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Call instance.ListEmbeddedDocumentsInWorksheet(testWorksheet)
    
    Assert.AreEqual EmbeddedDocumentsCountInWorksheet, instance.ListEmbeddedDocumentsInWorksheet(testWorksheet).Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub ListEmbeddedDocuments_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateWorkbookForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    On Error GoTo CATCH
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocuments
    
    Assert.AreEqual EmbeddedDocumentsCount, listedEmbeddedDocuments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub DeleteEmbeddedDocumentsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateWorkbookForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim EmbeddedDocumentsCountInWorksheet As Long
    Call DecolateWorksheetForEmbeddedDocumentsTest(testWorksheet, dummyFile, EmbeddedDocumentsCountInWorksheet)
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocumentsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim oleObj As OLEObject
    Dim dict As Object
    For Each oleObj In listedEmbeddedDocuments
        Set dict = CreateObject("Scripting.Dictionary")
        With oleObj
            dict("Type") = TypeName(oleObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(oleObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("OLEObject.OLEType") = .OLEType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedEmbeddedDocuments As Collection
    Set deletedEmbeddedDocuments = instance.DeleteEmbeddedDocumentsInWorksheet(testWorksheet)
    
    Assert.AreEqual EmbeddedDocumentsCountInWorksheet, deletedEmbeddedDocuments.Count
    Assert.AreEqual EmbeddedDocumentsCount, instance.ListEmbeddedDocuments.Count
    Assert.IsTrue DictionariesEquals(deletedEmbeddedDocuments, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub DeleteEmbeddedDocuments_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateWorkbookForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocuments
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim oleObj As OLEObject
    Dim dict As Object
    For Each oleObj In listedEmbeddedDocuments
        Set dict = CreateObject("Scripting.Dictionary")
        With oleObj
            dict("Type") = TypeName(oleObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(oleObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("OLEObject.OLEType") = .OLEType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedEmbeddedDocuments As Collection
    Set deletedEmbeddedDocuments = instance.DeleteEmbeddedDocuments
    
    Assert.AreEqual EmbeddedDocumentsCount, deletedEmbeddedDocuments.Count
    Assert.AreEqual 0&, instance.ListEmbeddedDocuments.Count
    Assert.IsTrue DictionariesEquals(deletedEmbeddedDocuments, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Macros, Forms, And ActiveX Controls ==
'' === Macros ===
Private Function DecolateWorkbookForMacrosTest(ByVal Target As Workbook, ByRef macrosCount As Long) As Workbook
    macrosCount = 0
    
    Dim i As Long
    For i = 1 To Target.VBProject.VBComponents.Count
        With Target.VBProject.VBComponents(i)
            If .Type = 100 Then
                If .Name = "ThisWorkbook" Or i Mod 2 = 0 Then
                    Call .CodeModule.AddFromString(DummyCodeString)
                    macrosCount = macrosCount + 1
                Else
                    Call .CodeModule.DeleteLines(1, .CodeModule.CountOfLines)
                End If
            End If
        End With
    Next
    
    Dim j As Long
    For j = 1 To 3
        Call Target.VBProject.VBComponents.Add(j): macrosCount = macrosCount + 1
    Next
    
    Set DecolateWorkbookForMacrosTest = Target
End Function

'@TestMethod("Macros")
Public Sub ListMacros_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim macrosCount As Long
    Call DecolateWorkbookForMacrosTest(instance.Target, macrosCount)
    
    On Error GoTo CATCH
    
    Dim listedMacros As Collection
    Set listedMacros = instance.ListMacros
    
    Assert.AreEqual macrosCount, listedMacros.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Macros")
Public Sub DeleteMacros_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim macrosCount As Long
    Call DecolateWorkbookForMacrosTest(instance.Target, macrosCount)
    
    Dim listedMacros As Collection
    Set listedMacros = instance.ListMacros
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim macro As Object
    Dim dict As Object
    For Each macro In listedMacros
        Set dict = CreateObject("Scripting.Dictionary")
        With macro
            dict("Type") = TypeName(macro)
            dict("Name") = .Name
            dict("Location") = GetWorkbookLocation(instance.Target)
        End With
        Select Case macro.Type
            Case 1 ' vbext_ct_StdModule
                dict("VBComponent.ComponentType") = "vbext_ct_StdModule"
            Case 2 ' vbext_ct_ClassModule
                dict("VBComponent.ComponentType") = "vbext_ct_ClassModule"
            Case 3 ' vbext_ct_MSForm
                dict("VBComponent.ComponentType") = "vbext_ct_MSForm"
            Case 11 ' vbext_ct_ActiveXDesigner
                dict("VBComponent.ComponentType") = "vbext_ct_ActiveXDesigner"
            Case 100 ' vbext_ct_Document
                dict("VBComponent.ComponentType") = "vbext_ct_Document"
        End Select
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim removedMacros As Collection
    Set removedMacros = instance.DeleteMacros
    
    Assert.AreEqual macrosCount, removedMacros.Count
    Assert.AreEqual 0&, instance.ListMacros.Count
    Assert.IsTrue DictionariesEquals(removedMacros, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Forms ===
Private Function DecolateWorksheetForFormsTest(ByVal TargetWorksheet As Worksheet, ByRef formsCount As Long) As Worksheet
    formsCount = 0
    
    Dim i As Long
    For i = 0 To 9
        If i <> 3 Then
            Call TargetWorksheet.Shapes.AddFormControl(i, 50 * i, 10, 20, 20): formsCount = formsCount + 1
        End If
    Next
    
    Set DecolateWorksheetForFormsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForFormsTest(ByVal Target As Workbook, ByRef formsCount As Long) As Workbook
    formsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim formsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForFormsTest(worksheet_, formsCountInWorksheet): formsCount = formsCount + formsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForFormsTest = Target
End Function

'@TestMethod("Forms")
Public Sub ListFormsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim formsCount As Long
    Call DecolateWorkbookForFormsTest(instance.Target, formsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim formsCountInWorksheet As Long
    Call DecolateWorksheetForFormsTest(testWorksheet, formsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedForms As Collection
    Set listedForms = instance.ListFormsInWorksheet(testWorksheet)
    
    Assert.AreEqual formsCountInWorksheet, listedForms.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Forms")
Public Sub ListForms_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim formsCount As Long
    Call DecolateWorkbookForFormsTest(instance.Target, formsCount)
    
    On Error GoTo CATCH
    
    Dim listedForms As Collection
    Set listedForms = instance.ListForms
    
    Assert.AreEqual formsCount, listedForms.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Forms")
Public Sub DeleteFormsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim formsCount As Long
    Call DecolateWorkbookForFormsTest(instance.Target, formsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim formsCountInWorksheet As Long
    Call DecolateWorksheetForFormsTest(testWorksheet, formsCountInWorksheet)
    
    Dim listedForms As Collection
    Set listedForms = instance.ListFormsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedForms
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedForms As Collection
    Set deletedForms = instance.DeleteFormsInWorksheet(testWorksheet)
    
    Assert.AreEqual formsCountInWorksheet, deletedForms.Count
    Assert.AreEqual formsCount, instance.ListForms.Count
    Assert.IsTrue DictionariesEquals(deletedForms, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Forms")
Public Sub DeleteForms_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim formsCount As Long
    Call DecolateWorkbookForFormsTest(instance.Target, formsCount)
    
    Dim listedForms As Collection
    Set listedForms = instance.ListForms
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedForms
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedForms As Collection
    Set deletedForms = instance.DeleteForms
    
    Assert.AreEqual formsCount, deletedForms.Count
    Assert.AreEqual 0&, instance.ListForms.Count
    Assert.IsTrue DictionariesEquals(deletedForms, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === ActiveX Controls ===
Private Function DecolateWorksheetForActiveXControlsTest(ByVal TargetWorksheet As Worksheet, ByRef activeXControlsCount As Long) As Worksheet
    activeXControlsCount = 0
    
    Dim i As Long
    Dim classType As String
    For i = 1 To 13
        Select Case i
            Case 1
              classType = "Forms.CommandButton.1"
            Case 2
              classType = "Forms.ComboBox.1"
            Case 3
              classType = "Forms.CheckBox.1"
            Case 4
              classType = "Forms.ListBox.1"
            Case 5
              classType = "Forms.TextBox.1"
            Case 6
              classType = "Forms.ScrollBar.1"
            Case 7
              classType = "Forms.SpinButton.1"
            Case 8
              classType = "Forms.OptionButton.1"
            Case 9
              classType = "Forms.Label.1"
            Case 10
              classType = "Forms.Image.1"
            Case 11
              classType = "Forms.ToggleButton.1"
        End Select
        Call TargetWorksheet.OLEObjects.Add(classType:=classType, Left:=50 * i, Top:=10, Width:=20, Height:=20): activeXControlsCount = activeXControlsCount + 1
    Next
    
    Set DecolateWorksheetForActiveXControlsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForActiveXControlsTest(ByVal Target As Workbook, ByRef activeXControlsCount As Long) As Workbook
    activeXControlsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim activeXControlsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForActiveXControlsTest(worksheet_, activeXControlsCountInWorksheet): activeXControlsCount = activeXControlsCount + activeXControlsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForActiveXControlsTest = Target
End Function

'@TestMethod("ActiveXControls")
Public Sub ListActiveXControlsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim activeXControlsCount As Long
    Call DecolateWorkbookForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim activeXControlsCountInWorksheet As Long
    Call DecolateWorksheetForActiveXControlsTest(testWorksheet, activeXControlsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Call instance.ListActiveXControlsInWorksheet(testWorksheet)
    
    Assert.AreEqual activeXControlsCountInWorksheet, instance.ListActiveXControlsInWorksheet(testWorksheet).Count
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ActiveXControls")
Public Sub ListActiveXControls_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim activeXControlsCount As Long
    Call DecolateWorkbookForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    On Error GoTo CATCH
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Assert.AreEqual activeXControlsCount, listedActiveXControls.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ActiveXControls")
Public Sub DeleteActiveXControlsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim activeXControlsCount As Long
    Call DecolateWorkbookForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim activeXControlsCountInWorksheet As Long
    Call DecolateWorksheetForActiveXControlsTest(testWorksheet, activeXControlsCountInWorksheet)
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControlsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim oleObj As OLEObject
    Dim dict As Object
    For Each oleObj In listedActiveXControls
        Set dict = CreateObject("Scripting.Dictionary")
        With oleObj
            dict("Type") = TypeName(oleObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(oleObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("OLEObject.OLEType") = .OLEType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedActiveXControls As Collection
    Set deletedActiveXControls = instance.DeleteActiveXControlsInWorksheet(testWorksheet)
    
    Assert.AreEqual activeXControlsCountInWorksheet, deletedActiveXControls.Count
    Assert.AreEqual activeXControlsCount, instance.ListActiveXControls.Count
    Assert.IsTrue DictionariesEquals(deletedActiveXControls, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ActiveXControls")
Public Sub DeleteActiveXControls_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim activeXControlsCount As Long
    Call DecolateWorkbookForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim oleObj As OLEObject
    Dim dict As Object
    For Each oleObj In listedActiveXControls
        Set dict = CreateObject("Scripting.Dictionary")
        With oleObj
            dict("Type") = TypeName(oleObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(oleObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("OLEObject.OLEType") = .OLEType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedActiveXControls As Collection
    Set deletedActiveXControls = instance.DeleteActiveXControls
    
    Assert.AreEqual activeXControlsCount, deletedActiveXControls.Count
    Assert.AreEqual 0&, instance.ListActiveXControls.Count
    Assert.IsTrue DictionariesEquals(deletedActiveXControls, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Headers and Footers ==
'' === Headers and Footers ===
'@TestMethod("HeadersAndFooters")
Public Sub InspectHeadersAndFooters_NoHeaderAndFooter_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub InspectHeadersAndFooters_WithHeaderAndFooter_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount), dummyImage, containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub FixHeadersAndFooters_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount), dummyImage, containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixHeadersAndFooters
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Headers ===
Private Function DecolateWorksheetForHeadersTest(ByVal TargetWorksheet As Worksheet, ByVal imagePath As String, ByRef containsHeadersSheetsCount As Long) As Worksheet
    containsHeadersSheetsCount = 0
    
    With TargetWorksheet.PageSetup
        .LeftHeader = "test"
        .CenterHeader = "test"
        .RightHeader = "test"
        .LeftHeaderPicture.fileName = imagePath
        .CenterHeaderPicture.fileName = imagePath
        .RightHeaderPicture.fileName = imagePath
    End With
    containsHeadersSheetsCount = containsHeadersSheetsCount + 1
    
    Set DecolateWorksheetForHeadersTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForHeadersTest(ByVal Target As Workbook, ByVal imagePath As String, ByRef containsHeadersSheetsCount As Long) As Workbook
    containsHeadersSheetsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim headersTestCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForHeadersTest(worksheet_, imagePath, headersTestCountInWorksheet): containsHeadersSheetsCount = containsHeadersSheetsCount + headersTestCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForHeadersTest = Target
End Function

'@TestMethod("HeadersAndFooters")
Public Sub ContainsHeadersInWorksheet_NoHeader_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    
    On Error GoTo CATCH
    
    Assert.IsFalse instance.ContainsHeadersInWorksheet(testWorksheet)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ContainsHeadersInWorksheet_WithHeader_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Dim instance As InspectWorkbookUtils
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\ContainsHeadersInWorksheet_WithHeader_Successed.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Dim pattern As Long
    For pattern = 1 To 6
        Set testWorkbook = CreateNewWorkbook(3)
        Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
        
        Call DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount)
    
        Set testWorksheet = instance.Target.Worksheets.Add
        With testWorksheet.PageSetup
            Select Case pattern
                Case 1
                    .LeftHeader = "test"
                Case 2
                    .CenterHeader = "test"
                Case 3
                    .RightHeader = "test"
                Case 4
                    .LeftHeaderPicture.fileName = dummyImage
                Case 5
                    .CenterHeaderPicture.fileName = dummyImage
                Case 6
                    .RightHeaderPicture.fileName = dummyImage
            End Select
        End With
        
        On Error GoTo CATCH
        
        Assert.IsTrue instance.ContainsHeadersInWorksheet(testWorksheet)
        
        GoTo FINALLY
CATCH:
        Assert.Fail Err.Description
        
        Resume FINALLY
FINALLY:
        Call testWorkbook.Close(False)
    Next
    Kill dummyImage
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ListSheetsContainsHeaders_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Dim listedSheetsContainsHeaders As Collection
    Set listedSheetsContainsHeaders = instance.ListSheetsContainsHeaders
    
    Assert.AreEqual containsHeadersSheetsCount, listedSheetsContainsHeaders.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ClearHeadersInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount)
    
    Dim containsHeadersSheetsCountInWorksheet As Long
    Set testWorksheet = instance.Target.Worksheets.Add
    Call DecolateWorksheetForHeadersTest(testWorksheet, dummyImage, containsHeadersSheetsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim cleared As Boolean
    cleared = instance.ClearHeadersInWorksheet(testWorksheet)
    
    Assert.IsTrue cleared
    Assert.IsFalse instance.ContainsHeadersInWorksheet(testWorksheet)
    Assert.IsTrue instance.ListSheetsContainsHeaders.Count > 0
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ClearHeaders_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateWorkbookForHeadersTest(instance.Target, dummyImage, containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Dim clearedSheets As Collection
    Set clearedSheets = instance.ClearHeaders
    
    Assert.AreEqual containsHeadersSheetsCount, clearedSheets.Count
    Assert.AreEqual 0&, instance.ListSheetsContainsHeaders.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Footers ===
Private Function DecolateWorksheetForFootersTest(ByVal TargetWorksheet As Worksheet, ByVal imagePath As String, ByRef containsFootersSheetsCount As Long) As Worksheet
    containsFootersSheetsCount = 0
    
    With TargetWorksheet.PageSetup
        .LeftFooter = "test"
        .CenterFooter = "test"
        .RightFooter = "test"
        .LeftFooterPicture.fileName = imagePath
        .CenterFooterPicture.fileName = imagePath
        .RightFooterPicture.fileName = imagePath
    End With
    containsFootersSheetsCount = containsFootersSheetsCount + 1
    
    Set DecolateWorksheetForFootersTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForFootersTest(ByVal Target As Workbook, ByVal imagePath As String, ByRef containsFootersSheetsCount As Long) As Workbook
    containsFootersSheetsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim FootersTestCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForFootersTest(worksheet_, imagePath, FootersTestCountInWorksheet): containsFootersSheetsCount = containsFootersSheetsCount + FootersTestCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForFootersTest = Target
End Function

'@TestMethod("HeadersAndFooters")
Public Sub ContainsFooters_NoFooter_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsFootersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(instance.Target, dummyImage, containsFootersSheetsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    
    On Error GoTo CATCH
    
    Assert.IsFalse instance.ContainsFootersInWorksheet(testWorksheet)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ContainsFootersInWorksheet_WithFooter_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Dim instance As InspectWorkbookUtils
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\ContainsFootersInWorksheet_WithFooter_Successed.xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsFootersSheetsCount As Long
    Dim pattern As Long
    For pattern = 1 To 6
        Set testWorkbook = CreateNewWorkbook(3)
        Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
        
        Call DecolateWorkbookForFootersTest(instance.Target, dummyImage, containsFootersSheetsCount)
    
        Set testWorksheet = instance.Target.Worksheets.Add
        With testWorksheet.PageSetup
            Select Case pattern
                Case 1
                    .LeftFooter = "test"
                Case 2
                    .CenterFooter = "test"
                Case 3
                    .RightFooter = "test"
                Case 4
                    .LeftFooterPicture.fileName = dummyImage
                Case 5
                    .CenterFooterPicture.fileName = dummyImage
                Case 6
                    .RightFooterPicture.fileName = dummyImage
            End Select
        End With
        
        On Error GoTo CATCH
        
        Assert.IsTrue instance.ContainsFootersInWorksheet(testWorksheet)
        
        GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
        Call testWorkbook.Close(False)
    Next
    Kill dummyImage
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ListSheetsContainsFooters_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsFootersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(instance.Target, dummyImage, containsFootersSheetsCount)
    
    On Error GoTo CATCH
    
    Dim listedSheetsContainsFooters As Collection
    Set listedSheetsContainsFooters = instance.ListSheetsContainsFooters
    
    Assert.AreEqual containsFootersSheetsCount, listedSheetsContainsFooters.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ClearFootersInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsFootersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(instance.Target, dummyImage, containsFootersSheetsCount)
    
    Dim containsFootersSheetsCountInWorksheet As Long
    Set testWorksheet = instance.Target.Worksheets.Add
    Call DecolateWorksheetForFootersTest(testWorksheet, dummyImage, containsFootersSheetsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim cleared As Boolean
    cleared = instance.ClearFootersInWorksheet(testWorksheet)
    
    Assert.IsTrue cleared
    Assert.IsFalse instance.ContainsFootersInWorksheet(testWorksheet)
    Assert.IsTrue instance.ListSheetsContainsFooters.Count > 0
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ClearFooters_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyImage As String
    dummyImage = fso.GetSpecialFolder(2) & "\" & ThisWorkbook.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyImage)
    
    Dim containsFootersSheetsCount As Long
    Call DecolateWorkbookForFootersTest(instance.Target, dummyImage, containsFootersSheetsCount)
    
    On Error GoTo CATCH
    
    Dim clearedSheets As Collection
    Set clearedSheets = instance.ClearFooters
    
    Assert.AreEqual containsFootersSheetsCount, clearedSheets.Count
    Assert.AreEqual 0&, instance.ListSheetsContainsFooters.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyImage
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Invisible Content ==
Private Function DecolateWorksheetForInvisibleContentTest(ByVal TargetWorksheet As Worksheet, ByRef invisibleContentCount As Long) As Worksheet
    invisibleContentCount = 0
    
    Dim i As Long
    Dim shape_ As Shape
    For i = 1 To 5
        Set shape_ = TargetWorksheet.Shapes.AddShape(msoShapeRound1Rectangle, 10 * i, 10 * i, 10, 10)
        If i Mod 2 <> 0 Then
            shape_.Visible = False: invisibleContentCount = invisibleContentCount + 1
        End If
    Next
    
    Set DecolateWorksheetForInvisibleContentTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForInvisibleContentTest(ByVal Target As Workbook, ByRef invisibleContentCount As Long) As Workbook
    invisibleContentCount = 0
    
    Dim worksheet_ As Worksheet
    Dim invisibleContentCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForInvisibleContentTest(worksheet_, invisibleContentCountInWorksheet): invisibleContentCount = invisibleContentCount + invisibleContentCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForInvisibleContentTest = Target
End Function

'@TestMethod("InvisibleContent")
Public Sub ListInvisibleContentInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testWorksheet = testWorkbook.Worksheets.Add
    Dim invisibleContentCountInWorksheet As Long
    Call DecolateWorksheetForInvisibleContentTest(testWorksheet, invisibleContentCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContentInWorksheet(testWorksheet)
    
    Assert.AreEqual invisibleContentCountInWorksheet, listedInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub ListInvisibleContent_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContent
    
    Assert.AreEqual invisibleContentCount, listedInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub VisualizeInvisibleContentInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testWorksheet = testWorkbook.Worksheets.Add
    Dim invisibleContentCountInWorksheet As Long
    Call DecolateWorksheetForInvisibleContentTest(testWorksheet, invisibleContentCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim visualizedInvisibleContent As Collection
    Set visualizedInvisibleContent = instance.VisualizeInvisibleContentInWorksheet(testWorksheet)
    
    Assert.AreEqual invisibleContentCountInWorksheet, visualizedInvisibleContent.Count
    Assert.AreEqual invisibleContentCount, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub VisualizeInvisibleContent_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Dim visualizedInvisibleContent As Collection
    Set visualizedInvisibleContent = instance.VisualizeInvisibleContent
    
    Assert.AreEqual invisibleContentCount, visualizedInvisibleContent.Count
    Assert.AreEqual 0&, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub DeleteInvisibleContentInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testWorksheet = testWorkbook.Worksheets.Add
    Dim invisibleContentCountInWorksheet As Long
    Call DecolateWorksheetForInvisibleContentTest(testWorksheet, invisibleContentCountInWorksheet)
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContentInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedInvisibleContent
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedInvisibleContent As Collection
    Set deletedInvisibleContent = instance.DeleteInvisibleContentInWorksheet(testWorksheet)
    
    Assert.AreEqual invisibleContentCountInWorksheet, deletedInvisibleContent.Count
    Assert.AreEqual invisibleContentCount, instance.ListInvisibleContent.Count
    Assert.IsTrue DictionariesEquals(deletedInvisibleContent, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub DeleteInvisibleContent_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContent
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedInvisibleContent
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedInvisibleContent As Collection
    Set deletedInvisibleContent = instance.DeleteInvisibleContent
    
    Assert.AreEqual invisibleContentCount, deletedInvisibleContent.Count
    Assert.AreEqual 0&, instance.ListInvisibleContent.Count
    Assert.IsTrue DictionariesEquals(deletedInvisibleContent, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_NoInvisibleContent_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_WithInvisibleContent_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub FixInvisibleContent_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim invisibleContentCount As Long
    Call DecolateWorkbookForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixInvisibleContent
    Assert.AreEqual 0&, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' = For Excel documents =
'' == Data Model ==
Private Function DecolateWorkbookForDataModelsTest(ByVal Target As Workbook, ByRef dataModelsCount As Long) As Workbook
    dataModelsCount = 0
    
    Dim sourceSheet  As Worksheet
    Set sourceSheet = Target.Worksheets(1)
    Dim source_ As ListObject
    Dim sourceData_ As Variant
    Dim dataRange As Range
    If sourceSheet.ListObjects.Count = 0 Then
        sourceData_ = SampleTableData
        With sourceSheet
            Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
        End With
        dataRange = sourceData_
        
        Set source_ = sourceSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=dataRange, _
            XlListObjectHasHeaders:=xlYes _
        )
        source_.Name = "Table1"
    Else
        Set source_ = sourceSheet.ListObjects(1)
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call Target.Connections.Add2( _
            Name:="hoge", _
            Description:=vbNullString, _
            ConnectionString:="WORKSHEET;" & Target.FullName, _
            CommandText:=Target.Name & "!" & source_.Name _
        ): dataModelsCount = dataModelsCount + 1
    Next
    
    Set DecolateWorkbookForDataModelsTest = Target
End Function

'@TestMethod("DataModels")
Public Sub ListDataModels_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim dataModelsCount As Long
    Call DecolateWorkbookForDataModelsTest(instance.Target, dataModelsCount)
    
    On Error GoTo CATCH
    
    Dim listedDataModels As Collection
    Set listedDataModels = instance.ListDataModels
    
    Assert.AreEqual dataModelsCount, listedDataModels.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DataModels")
Public Sub RemoveDataModels_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim dataModelsCount As Long
    Call DecolateWorkbookForDataModelsTest(instance.Target, dataModelsCount)
    
    On Error GoTo CATCH
    
    instance.RemoveDataModels
    
    Assert.AreEqual 0&, instance.ListDataModels.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == PivotTables, PivotCharts, Cube Formulas, Slices, and Timelines ==
'' === PivotTables, PivotCharts, Cube Formulas, Slices, and Timelines ===
Private Function SampleTableData() As Variant
    Dim sourceData(1 To 4, 1 To 3) As Variant
    
    sourceData(1, 1) = "StringColumn"
    sourceData(1, 2) = "DateColumn"
    sourceData(1, 3) = "NumberColumn"
    sourceData(2, 1) = "AAA"
    sourceData(2, 2) = "2000/01/01"
    sourceData(2, 3) = "1.5"
    sourceData(3, 1) = "BBB"
    sourceData(3, 2) = "2000/02/29"
    sourceData(3, 3) = "3"
    sourceData(4, 1) = "CCC"
    sourceData(4, 2) = "2000/12/31"
    sourceData(4, 3) = "4.5"
    
    SampleTableData = sourceData
End Function

'' === PivotTables ===
Private Function DecolateWorksheetForPivotTablesTest(ByVal TargetWorksheet As Worksheet, ByRef PivotTablesCount As Long) As Worksheet
    PivotTablesCount = 0
    
    Dim sourceData_ As Variant
    sourceData_ = SampleTableData
    
    Dim dataRange As Range
    With TargetWorksheet
        Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
    End With
    dataRange = sourceData_
    
    Dim Target As Workbook
    Set Target = GetParentWorkbook(TargetWorksheet)
    
    Dim conn As WorkbookConnection
    Dim connIndex As Long
    connIndex = Target.Connections.Count
    If connIndex = 0 Then
        connIndex = 1
    End If
    Set conn = Target.Connections.Add2( _
        Name:="Connections_" & connIndex, _
        Description:=vbNullString, _
        ConnectionString:="WORKSHEET;[" & Target.Name & "]" & TargetWorksheet.Name, _
        CommandText:=TargetWorksheet.Name & "!" & dataRange.Address, _
        lCmdType:=xlCmdExcel, _
        CreateModelConnection:=True, _
        ImportRelationships:=False _
    )
    
    Dim pivCache As PivotCache
    Set pivCache = Target.PivotCaches.Create( _
        SourceType:=xlExternal, _
        sourceData:=conn, _
        Version:=xlPivotTableVersion15 _
    )
    Call pivCache.CreatePivotTable( _
        TableDestination:=TargetWorksheet.Cells(1, 1 + UBound(sourceData_, 2)), _
        TableName:=Join(Array("PivotTable", TargetWorksheet.Name, 1), "-"), _
        DefaultVersion:=xlPivotTableVersion15 _
    )
    PivotTablesCount = PivotTablesCount + 1
    
    Set pivCache = Target.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=TargetWorksheet.Name & "!" & dataRange.Address, _
        Version:=xlPivotTableVersion15 _
    )
    Dim i As Long
    For i = PivotTablesCount + 1 To PivotTablesCount + 2
        Call pivCache.CreatePivotTable( _
            TableDestination:=TargetWorksheet.Cells(1, 1 + UBound(sourceData_, 2) * i), _
            TableName:=Join(Array("PivotTable", TargetWorksheet.Name, i), "-"), _
            DefaultVersion:=xlPivotTableVersion15 _
        )
        PivotTablesCount = PivotTablesCount + 1
    Next
    
    Set DecolateWorksheetForPivotTablesTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForPivotTablesTest(ByVal Target As Workbook, ByRef PivotTablesCount As Long) As Workbook
    PivotTablesCount = 0
    
    Dim worksheet_ As Worksheet
    Dim pivotTablesCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForPivotTablesTest(worksheet_, pivotTablesCountInWorksheet): PivotTablesCount = PivotTablesCount + pivotTablesCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForPivotTablesTest = Target
End Function

'@TestMethod("PivotTables")
Public Sub ListPivotTablesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotTablesCount As Long
    Call DecolateWorkbookForPivotTablesTest(instance.Target, PivotTablesCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim pivotTablesCountInWorksheet As Long
    Call DecolateWorksheetForPivotTablesTest(testWorksheet, pivotTablesCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedPivotTables As Collection
    Set listedPivotTables = instance.ListPivotTablesInWorksheet(testWorksheet)
    
    Assert.AreEqual pivotTablesCountInWorksheet, listedPivotTables.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotTables")
Public Sub ListPivotTables_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotTablesCount As Long
    Call DecolateWorkbookForPivotTablesTest(instance.Target, PivotTablesCount)
    
    On Error GoTo CATCH
    
    Dim listedPivotTables As Collection
    Set listedPivotTables = instance.ListPivotTables
    
    Assert.AreEqual PivotTablesCount, listedPivotTables.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotTables")
Public Sub DeletePivotTablesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotTablesCount As Long
    Call DecolateWorkbookForPivotTablesTest(instance.Target, PivotTablesCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim pivotTablesCountInWorksheet As Long
    Call DecolateWorksheetForPivotTablesTest(testWorksheet, pivotTablesCountInWorksheet)
    
    Dim listedPivotTables As Collection
    Set listedPivotTables = instance.ListPivotTablesInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim pivTable As PivotTable
    Dim dict As Object
    For Each pivTable In listedPivotTables
        Set dict = CreateObject("Scripting.Dictionary")
        With pivTable
            dict("Type") = TypeName(pivTable)
            dict("Name") = .Name
            dict("Location") = .TableRange2.Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedPivotTables As Collection
    Set deletedPivotTables = instance.DeletePivotTablesInWorksheet(testWorksheet)
    
    Assert.AreEqual pivotTablesCountInWorksheet, deletedPivotTables.Count
    Assert.IsTrue instance.ListPivotTables.Count > 0
    Assert.IsTrue DictionariesEquals(deletedPivotTables, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotTables")
Public Sub DeletePivotTables_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotTablesCount As Long
    Call DecolateWorkbookForPivotTablesTest(instance.Target, PivotTablesCount)
    
    Dim listedPivotTables As Collection
    Set listedPivotTables = instance.ListPivotTables
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim pivTable As PivotTable
    Dim dict As Object
    For Each pivTable In listedPivotTables
        Set dict = CreateObject("Scripting.Dictionary")
        With pivTable
            dict("Type") = TypeName(pivTable)
            dict("Name") = .Name
            dict("Location") = .TableRange2.Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedPivotTables As Collection
    Set deletedPivotTables = instance.DeletePivotTables
    
    Assert.AreEqual PivotTablesCount, deletedPivotTables.Count
    Assert.AreEqual 0&, instance.ListPivotTables.Count
    Assert.IsTrue DictionariesEquals(deletedPivotTables, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === PivotCharts ===
Private Function DecolateWorksheetForPivotChartsTest(ByVal TargetWorksheet As Worksheet, ByRef PivotChartsCount As Long) As Worksheet
    PivotChartsCount = 0
    
    Dim sourceData_ As Variant
    sourceData_ = SampleTableData
    
    Dim dataRange As Range
    With TargetWorksheet
        Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
    End With
    dataRange = sourceData_
    
    Dim Target As Workbook
    Set Target = GetParentWorkbook(TargetWorksheet)
    
    Dim conn As WorkbookConnection
    Dim connIndex As Long
    connIndex = Target.Connections.Count
    If connIndex = 0 Then
        connIndex = 1
    End If
    Set conn = Target.Connections.Add2( _
        Name:="Connections_" & connIndex, _
        Description:=vbNullString, _
        ConnectionString:="WORKSHEET;[" & Target.Name & "]" & TargetWorksheet.Name, _
        CommandText:=TargetWorksheet.Name & "!" & dataRange.Address, _
        lCmdType:=xlCmdExcel, _
        CreateModelConnection:=True, _
        ImportRelationships:=False _
    )
    
    Dim pivCache As PivotCache
    Set pivCache = Target.PivotCaches.Create( _
        SourceType:=xlExternal, _
        sourceData:=conn, _
        Version:=xlPivotTableVersion15 _
    )
    
    Dim pivChart As chart
    Set pivChart = pivCache.CreatePivotChart( _
            ChartDestination:=TargetWorksheet, _
            XlChartType:=xlLine _
        ).chart
    PivotChartsCount = PivotChartsCount + 1
    
    Set pivCache = Target.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=TargetWorksheet.Name & "!" & dataRange.Address, _
        Version:=xlPivotTableVersion15 _
    )
    Call pivCache.CreatePivotTable( _
        TableDestination:=TargetWorksheet.Cells(1, 1 + UBound(sourceData_, 2)), _
        TableName:=Join(Array("PivotTable", TargetWorksheet.Name, 1), "-"), _
        DefaultVersion:=xlPivotTableVersion15 _
    )
    
    Set pivChart = TargetWorksheet.Shapes.AddChart2( _
        Style:=xlPie, _
        XlChartType:=xlColumnClustered _
    ).chart
    Call pivChart.SetSourceData(Source:=dataRange)
    PivotChartsCount = PivotChartsCount + 1
    
    Set DecolateWorksheetForPivotChartsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForPivotChartsTest(ByVal Target As Workbook, ByRef PivotChartsCount As Long) As Workbook
    PivotChartsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim PivotChartsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForPivotChartsTest(worksheet_, PivotChartsCountInWorksheet): PivotChartsCount = PivotChartsCount + PivotChartsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForPivotChartsTest = Target
End Function

'@TestMethod("PivotCharts")
Public Sub ListPivotChartsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotChartsCount As Long
    Call DecolateWorkbookForPivotChartsTest(instance.Target, PivotChartsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim PivotChartsCountInWorksheet As Long
    Call DecolateWorksheetForPivotChartsTest(testWorksheet, PivotChartsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedPivotCharts As Collection
    Set listedPivotCharts = instance.ListPivotChartsInWorksheet(testWorksheet)
    
    Assert.AreEqual PivotChartsCountInWorksheet, listedPivotCharts.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotCharts")
Public Sub ListPivotCharts_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotChartsCount As Long
    Call DecolateWorkbookForPivotChartsTest(instance.Target, PivotChartsCount)
    
    On Error GoTo CATCH
    
    Dim listedPivotCharts As Collection
    Set listedPivotCharts = instance.ListPivotCharts
    
    Assert.AreEqual PivotChartsCount, listedPivotCharts.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotCharts")
Public Sub DeletePivotChartsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotChartsCount As Long
    Call DecolateWorkbookForPivotChartsTest(instance.Target, PivotChartsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim PivotChartsCountInWorksheet As Long
    Call DecolateWorksheetForPivotChartsTest(testWorksheet, PivotChartsCountInWorksheet)
    
    Dim listedPivotCharts As Collection
    Set listedPivotCharts = instance.ListPivotChartsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim chartObj As ChartObject
    Dim dict As Object
    For Each chartObj In listedPivotCharts
        Set dict = CreateObject("Scripting.Dictionary")
        With chartObj
            dict("Type") = TypeName(chartObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(chartObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedPivotCharts As Collection
    Set deletedPivotCharts = instance.DeletePivotChartsInWorksheet(testWorksheet)
    
    Assert.AreEqual PivotChartsCountInWorksheet, deletedPivotCharts.Count
    Assert.IsTrue instance.ListPivotCharts.Count > 0
    Assert.IsTrue DictionariesEquals(deletedPivotCharts, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("PivotCharts")
Public Sub DeletePivotCharts_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim PivotChartsCount As Long
    Call DecolateWorkbookForPivotChartsTest(instance.Target, PivotChartsCount)
    
    Dim listedPivotCharts As Collection
    Set listedPivotCharts = instance.ListPivotCharts
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim chartObj As ChartObject
    Dim dict As Object
    For Each chartObj In listedPivotCharts
        Set dict = CreateObject("Scripting.Dictionary")
        With chartObj
            dict("Type") = TypeName(chartObj)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(chartObj).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedPivotCharts As Collection
    Set deletedPivotCharts = instance.DeletePivotCharts
    
    Assert.AreEqual PivotChartsCount, deletedPivotCharts.Count
    Assert.AreEqual 0&, instance.ListPivotCharts.Count
    Assert.IsTrue DictionariesEquals(deletedPivotCharts, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Cube Formulas ===
Private Function DecolateWorksheetForCubeFormulasTest(ByVal TargetWorksheet As Worksheet, ByRef CubeFormulasCount As Long) As Worksheet
    CubeFormulasCount = 0
    
    Dim testPattern As Collection: Set testPattern = New Collection
    ' CUBEKPIMEMBER
    Call testPattern.Add("=CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=   CuBeKpImEmBeR(""foo"", ""bar"", 1)")
    Call testPattern.Add("=ISERR(CUBEKPIMEMBER(""foo"", ""bar"", 1))")
    Call testPattern.Add("=ISERR(   CuBeKpImEmBeR(""foo"", ""bar"", 1))")
    Call testPattern.Add("=AND(TRUE,CUBEKPIMEMBER(""foo"", ""bar"", 1))")
    Call testPattern.Add("=AND(TRUE,   CuBeKpImEmBeR(""foo"", ""bar"", 1))")
    Call testPattern.Add("=1+CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1   +     CuBeKpImEmBeR(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1-CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1*CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1/CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1^CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1=CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1<CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1>CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=""1""&CUBEKPIMEMBER(""foo"", ""bar"", 1)")
    ' CUBEMEMBER
    Call testPattern.Add("=CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=   CuBeMeMbEr(""foo"", ""bar"")")
    Call testPattern.Add("=ISERR(CUBEMEMBER(""foo"", ""bar""))")
    Call testPattern.Add("=ISERR(   CuBeMeMbEr(""foo"", ""bar""))")
    Call testPattern.Add("=AND(TRUE,CUBEMEMBER(""foo"", ""bar""))")
    Call testPattern.Add("=AND(TRUE,   CuBeMeMbEr(""foo"", ""bar""))")
    Call testPattern.Add("=1+CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1   +     CuBeMeMbEr(""foo"", ""bar"")")
    Call testPattern.Add("=1-CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1*CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1/CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1^CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1=CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1<CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=1>CUBEMEMBER(""foo"", ""bar"")")
    Call testPattern.Add("=""1""&CUBEMEMBER(""foo"", ""bar"")")
    ' CUBEMEMBERPROPERTY
    Call testPattern.Add("=CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=   CuBeMeMbErPrOpErTy(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=ISERR(CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=ISERR(   CuBeMeMbErPrOpErTy(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=AND(TRUE,CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=AND(TRUE,   CuBeMeMbErPrOpErTy(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=1+CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1   +     CuBeMeMbErPrOpErTy(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1-CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1*CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1/CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1^CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1=CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1<CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1>CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=""1""&CUBEMEMBERPROPERTY(""foo"", ""bar"", ""baz"")")
    ' CUBERANKEDMEMBER
    Call testPattern.Add("=CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=   CuBeRaNkEdMeMbEr(""foo"", ""bar"", 1)")
    Call testPattern.Add("=ISERR(CUBERANKEDMEMBER(""foo"", ""bar"", 1))")
    Call testPattern.Add("=ISERR(   CuBeRaNkEdMeMbEr(""foo"", ""bar"", 1))")
    Call testPattern.Add("=AND(TRUE,CUBERANKEDMEMBER(""foo"", ""bar"", 1))")
    Call testPattern.Add("=AND(TRUE,   CuBeRaNkEdMeMbEr(""foo"", ""bar"", 1))")
    Call testPattern.Add("=1+CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1   +     CuBeRaNkEdMeMbEr(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1-CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1*CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1/CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1^CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1=CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1<CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=1>CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    Call testPattern.Add("=""1""&CUBERANKEDMEMBER(""foo"", ""bar"", 1)")
    ' CUBESET
    Call testPattern.Add("=CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=   CuBeSeT(""foo"", ""bar"")")
    Call testPattern.Add("=ISERR(CUBESET(""foo"", ""bar""))")
    Call testPattern.Add("=ISERR(   CuBeSeT(""foo"", ""bar""))")
    Call testPattern.Add("=AND(TRUE,CUBESET(""foo"", ""bar""))")
    Call testPattern.Add("=AND(TRUE,   CuBeSeT(""foo"", ""bar""))")
    Call testPattern.Add("=1+CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1   +     CuBeSeT(""foo"", ""bar"")")
    Call testPattern.Add("=1-CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1*CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1/CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1^CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1=CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1<CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=1>CUBESET(""foo"", ""bar"")")
    Call testPattern.Add("=""1""&CUBESET(""foo"", ""bar"")")
    ' CUBESETCOUNT
    Call testPattern.Add("=   CuBeSeTcOuNt(""foo"")")
    Call testPattern.Add("=ISERR(CUBESETCOUNT(""foo""))")
    Call testPattern.Add("=ISERR(   CuBeSeTcOuNt(""foo""))")
    Call testPattern.Add("=AND(TRUE,CUBESETCOUNT(""foo""))")
    Call testPattern.Add("=AND(TRUE,   CuBeSeTcOuNt(""foo""))")
    Call testPattern.Add("=1+CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1   +     CuBeSeTcOuNt(""foo"")")
    Call testPattern.Add("=1-CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1*CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1/CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1^CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1=CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1<CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=1>CUBESETCOUNT(""foo"")")
    Call testPattern.Add("=""1""&CUBESETCOUNT(""foo"")")
    ' CUBEVALUE
    Call testPattern.Add("=CUBEVALUE(""foo"")")
    Call testPattern.Add("=   CUBEVALUE(""foo"")")
    Call testPattern.Add("=ISERR(CUBEVALUE(""foo""))")
    Call testPattern.Add("=ISERR(   CUBEVALUE(""foo""))")
    Call testPattern.Add("=AND(TRUE,CUBEVALUE(""foo""))")
    Call testPattern.Add("=AND(TRUE,   CUBEVALUE(""foo""))")
    Call testPattern.Add("=1+CUBEVALUE(""foo"")")
    Call testPattern.Add("=1   +     CUBEVALUE(""foo"")")
    Call testPattern.Add("=1-CUBEVALUE(""foo"")")
    Call testPattern.Add("=1*CUBEVALUE(""foo"")")
    Call testPattern.Add("=1/CUBEVALUE(""foo"")")
    Call testPattern.Add("=1^CUBEVALUE(""foo"")")
    Call testPattern.Add("=1=CUBEVALUE(""foo"")")
    Call testPattern.Add("=1<CUBEVALUE(""foo"")")
    Call testPattern.Add("=1>CUBEVALUE(""foo"")")
    Call testPattern.Add("=""1""&CUBEVALUE(""foo"")")
    
    Dim i As Long
    Dim range_ As Range
    For i = 1 To testPattern.Count
        Set range_ = TargetWorksheet.Cells(i, i)
        range_.formula = testPattern(i): CubeFormulasCount = CubeFormulasCount + 1
    Next
    
    Set DecolateWorksheetForCubeFormulasTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForCubeFormulasTest(ByVal Target As Workbook, ByRef CubeFormulasCount As Long) As Workbook
    CubeFormulasCount = 0
    
    Dim worksheet_ As Worksheet
    Dim CubeFormulasCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForCubeFormulasTest(worksheet_, CubeFormulasCountInWorksheet): CubeFormulasCount = CubeFormulasCount + CubeFormulasCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForCubeFormulasTest = Target
End Function

'@TestMethod("CubeFormulas")
Public Sub ListCubeFormulasInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim CubeFormulasCount As Long
    Call DecolateWorkbookForCubeFormulasTest(instance.Target, CubeFormulasCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim CubeFormulasCountInWorksheet As Long
    Call DecolateWorksheetForCubeFormulasTest(testWorksheet, CubeFormulasCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedCubeFormulasCell As Collection
    Set listedCubeFormulasCell = instance.ListCubeFormulasInWorksheet(testWorksheet)
    
    Assert.AreEqual CubeFormulasCountInWorksheet, listedCubeFormulasCell.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CubeFormulas")
Public Sub ListCubeFormulas_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim CubeFormulasCount As Long
    Call DecolateWorkbookForCubeFormulasTest(instance.Target, CubeFormulasCount)
    
    On Error GoTo CATCH
    
    Dim listedCubeFormulasCell As Collection
    Set listedCubeFormulasCell = instance.ListCubeFormulas
    
    Assert.AreEqual CubeFormulasCount, listedCubeFormulasCell.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Slicers ===
Private Function DecolateWorksheetForSlicersTest(ByVal TargetWorksheet As Worksheet, ByRef SlicersCount As Long) As Worksheet
    SlicersCount = 0
    
    Dim sourceData_ As Variant
    sourceData_ = SampleTableData
    
    Dim dataRange As Range
    With TargetWorksheet
        Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
    End With
    dataRange = sourceData_
    
    Dim Target As Workbook
    Set Target = GetParentWorkbook(TargetWorksheet)
    
    Dim conn As WorkbookConnection
    Dim connIndex As Long
    connIndex = Target.Connections.Count
    If connIndex = 0 Then
        connIndex = 1
    End If
    Set conn = Target.Connections.Add2( _
        Name:="Connections_" & connIndex, _
        Description:=vbNullString, _
        ConnectionString:="WORKSHEET;[" & Target.Name & "]" & TargetWorksheet.Name, _
        CommandText:=TargetWorksheet.Name & "!" & dataRange.Address, _
        lCmdType:=xlCmdExcel, _
        CreateModelConnection:=True, _
        ImportRelationships:=False _
    )
    
    Dim modelTab As ModelTable
    
    Dim i As Long
    For Each modelTab In conn.ModelTables
        For i = 1 To UBound(sourceData_, 2)
            Call Target.SlicerCaches.Add2( _
                Source:="ThisWorkbookDataModel", _
                SourceField:="[" & modelTab.Name & "].[" & dataRange.Cells(1, i).value & "]", _
                Name:=Join(Array("Slicer", TargetWorksheet.Name, Replace(modelTab.Name, " ", vbNullString), i), "_"), _
                SlicerCacheType:=xlSlicer _
            ): SlicersCount = SlicersCount + 1
        Next
    Next
    
    Set DecolateWorksheetForSlicersTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForSlicersTest(ByVal Target As Workbook, ByRef SlicersCount As Long) As Workbook
    SlicersCount = 0
    
    Dim worksheet_ As Worksheet
    Dim SlicersCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForSlicersTest(worksheet_, SlicersCountInWorksheet): SlicersCount = SlicersCount + SlicersCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForSlicersTest = Target
End Function

'@TestMethod("Slicers")
Public Sub ListSlicersTest_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim SlicersCount As Long
    Call DecolateWorkbookForSlicersTest(instance.Target, SlicersCount)
    
    On Error GoTo CATCH
    
    Dim listedSlicers As Collection
    Set listedSlicers = instance.ListSlicers
    
    Assert.AreEqual SlicersCount, listedSlicers.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Slicers")
Public Sub DeleteSlicers_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim SlicersCount As Long
    Call DecolateWorkbookForSlicersTest(instance.Target, SlicersCount)
    
    Dim listedSlicers As Collection
    Set listedSlicers = instance.ListSlicers
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim slicerCache_ As SlicerCache
    Dim dict As Object
    For Each slicerCache_ In listedSlicers
        Set dict = CreateObject("Scripting.Dictionary")
        With slicerCache_
            dict("Type") = TypeName(slicerCache_)
            dict("Name") = .Name
            dict("Location") = GetWorkbookLocation(GetParentWorkbook(slicerCache_))
            dict("SlicerCache.SlicerCacheType") = .SlicerCacheType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedSlicers As Collection
    Set deletedSlicers = instance.DeleteSlicers
    
    Assert.AreEqual SlicersCount, deletedSlicers.Count
    Assert.AreEqual 0&, instance.ListSlicers.Count
    Assert.IsTrue DictionariesEquals(deletedSlicers, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Timelines (cache) ===
Private Function DecolateWorksheetForTimelinesTest(ByVal TargetWorksheet As Worksheet, ByRef TimelinesCount As Long) As Worksheet
    TimelinesCount = 0
    
    Dim sourceData_ As Variant
    sourceData_ = SampleTableData
    
    Dim dataRange As Range
    With TargetWorksheet
        Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
    End With
    dataRange = sourceData_
    
    Dim Target As Workbook
    Set Target = GetParentWorkbook(TargetWorksheet)
    
    Dim conn As WorkbookConnection
    Dim connIndex As Long
    connIndex = Target.Connections.Count
    If connIndex = 0 Then
        connIndex = 1
    End If
    Set conn = Target.Connections.Add2( _
        Name:="Connections_" & connIndex, _
        Description:=vbNullString, _
        ConnectionString:="WORKSHEET;[" & Target.Name & "]" & TargetWorksheet.Name, _
        CommandText:=TargetWorksheet.Name & "!" & dataRange.Address, _
        lCmdType:=xlCmdExcel, _
        CreateModelConnection:=True, _
        ImportRelationships:=False _
    )
    
    Dim modelTab As ModelTable
    
    Dim i As Long
    For Each modelTab In conn.ModelTables
        For i = 1 To UBound(sourceData_, 2)
            If dataRange.Cells(1, i).value Like "Date*" Then
                Call Target.SlicerCaches.Add2( _
                    Source:="ThisWorkbookDataModel", _
                    SourceField:="[" & modelTab.Name & "].[" & dataRange.Cells(1, i).value & "]", _
                    Name:=Join(Array("Timeline", TargetWorksheet.Name, Replace(modelTab.Name, " ", vbNullString), i), "_"), _
                    SlicerCacheType:=xlTimeline _
                ): TimelinesCount = TimelinesCount + 1
            
            End If
        Next
    Next
    
    Set DecolateWorksheetForTimelinesTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForTimelinesTest(ByVal Target As Workbook, ByRef TimelinesCount As Long) As Workbook
    TimelinesCount = 0
    
    Dim worksheet_ As Worksheet
    Dim TimelinesCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForTimelinesTest(worksheet_, TimelinesCountInWorksheet): TimelinesCount = TimelinesCount + TimelinesCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForTimelinesTest = Target
End Function

'@TestMethod("Timelines")
Public Sub ListTimelines_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim TimelinesCount As Long
    Call DecolateWorkbookForTimelinesTest(instance.Target, TimelinesCount)
    
    On Error GoTo CATCH
    
    Dim listedTimelines As Collection
    Set listedTimelines = instance.ListTimelines
    
    Assert.AreEqual TimelinesCount, listedTimelines.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Timelines")
Public Sub DeleteTimelines_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim TimelinesCount As Long
    Call DecolateWorkbookForTimelinesTest(instance.Target, TimelinesCount)
    
    Dim listedTimelines As Collection
    Set listedTimelines = instance.ListTimelines
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim slicerCache_ As SlicerCache
    Dim dict As Object
    For Each slicerCache_ In listedTimelines
        Set dict = CreateObject("Scripting.Dictionary")
        With slicerCache_
            dict("Type") = TypeName(slicerCache_)
            dict("Name") = .Name
            dict("Location") = GetWorkbookLocation(GetParentWorkbook(slicerCache_))
            dict("SlicerCache.SlicerCacheType") = .SlicerCacheType
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedTimelines As Collection
    Set deletedTimelines = instance.DeleteTimelines
    
    Assert.AreEqual TimelinesCount, deletedTimelines.Count
    Assert.AreEqual 0&, instance.ListTimelines.Count
    Assert.IsTrue DictionariesEquals(deletedTimelines, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Links to Other Files ==
Private Function DecolateWorksheetForLinksToOtherFilesTest(ByVal TargetWorksheet As Worksheet, ByVal DummyWorkbooks As Collection, ByRef ExternalLinkCellsCount As Long) As Worksheet
    ExternalLinkCellsCount = 0
    
    Dim i As Long
    Dim j As Long
    For j = 1 To DummyWorkbooks.Count
        For i = 1 To 3
            TargetWorksheet.Cells(j, i).formula = "=" & DummyWorkbooks(j).Worksheets(1).Cells(j, i).Address(External:=True): ExternalLinkCellsCount = ExternalLinkCellsCount + 1
        Next
    Next
    
    Set DecolateWorksheetForLinksToOtherFilesTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForLinksToOtherFilesTest(ByVal Target As Workbook, ByVal DummyWorkbooks As Collection, ByRef ExternalLinkCellsCount As Long) As Workbook
    ExternalLinkCellsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim ExternalLinkCellsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForLinksToOtherFilesTest(worksheet_, DummyWorkbooks, ExternalLinkCellsCountInWorksheet): ExternalLinkCellsCount = ExternalLinkCellsCount + ExternalLinkCellsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForLinksToOtherFilesTest = Target
End Function

'@TestMethod("LinksToOtherFiles")
Public Sub ListLinksToOtherFiles_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    Dim dummyBooks As Collection: Set dummyBooks = New Collection
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    
    Dim LinksToOtherFilesCellsCount As Long
    Call DecolateWorkbookForLinksToOtherFilesTest(instance.Target, dummyBooks, LinksToOtherFilesCellsCount)
    
    On Error GoTo CATCH
    
    Dim listedLinksToOtherFiles As Collection
    Set listedLinksToOtherFiles = instance.ListLinksToOtherFiles
    
    Assert.AreEqual dummyBooks.Count, listedLinksToOtherFiles.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Dim dummyBook As Workbook
    For Each dummyBook In dummyBooks
        Call dummyBook.Close(False)
    Next
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("LinksToOtherFiles")
Public Sub ListExternalLinkCellsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    Dim dummyBooks As Collection: Set dummyBooks = New Collection
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    
    Dim LinksToOtherFilesCellsCount As Long
    Call DecolateWorkbookForLinksToOtherFilesTest(instance.Target, dummyBooks, LinksToOtherFilesCellsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim LinksToOtherFilesCellsCountInWorksheet As Long
    Call DecolateWorksheetForLinksToOtherFilesTest(testWorksheet, dummyBooks, LinksToOtherFilesCellsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedExternalLinkCells As Collection
    Set listedExternalLinkCells = instance.ListExternalLinkCellsInWorksheet(testWorksheet)
    
    Assert.AreEqual LinksToOtherFilesCellsCountInWorksheet, listedExternalLinkCells.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Dim dummyBook As Workbook
    For Each dummyBook In dummyBooks
        Call dummyBook.Close(False)
    Next
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("LinksToOtherFiles")
Public Sub ListExternalLinkCells_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    Dim dummyBooks As Collection: Set dummyBooks = New Collection
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    
    Dim LinksToOtherFilesCellsCount As Long
    Call DecolateWorkbookForLinksToOtherFilesTest(instance.Target, dummyBooks, LinksToOtherFilesCellsCount)
    
    On Error GoTo CATCH
    
    Dim listedExternalLinkCells As Collection
    Set listedExternalLinkCells = instance.ListExternalLinkCells
    
    Assert.AreEqual LinksToOtherFilesCellsCount, listedExternalLinkCells.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Dim dummyBook As Workbook
    For Each dummyBook In dummyBooks
        Call dummyBook.Close(False)
    Next
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("LinksToOtherFiles")
Public Sub BreakLinksToOtherFiles_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    Dim dummyBooks As Collection: Set dummyBooks = New Collection
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    Call dummyBooks.Add(Workbooks.Add)
    
    Dim LinksToOtherFilesCount As Long
    Call DecolateWorkbookForLinksToOtherFilesTest(instance.Target, dummyBooks, LinksToOtherFilesCount)
    
    On Error GoTo CATCH
    
    Dim brokenLinksToOtherFiles As Collection
    Set brokenLinksToOtherFiles = instance.BreakLinksToOtherFiles
    
    Assert.AreEqual dummyBooks.Count, brokenLinksToOtherFiles.Count
    Assert.AreEqual 0&, instance.ListLinksToOtherFiles.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Dim dummyBook As Workbook
    For Each dummyBook In dummyBooks
        Call dummyBook.Close(False)
    Next
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Real Time Data Functions ==
Private Function DecolateWorksheetForRealTimeDataFunctionsTest(ByVal TargetWorksheet As Worksheet, ByRef dataFunctionsCount As Long) As Worksheet
    dataFunctionsCount = 0
    
    Dim testPattern As Collection: Set testPattern = New Collection
    Call testPattern.Add("=RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=   rTd(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=ISERR(RTD(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=ISERR(   rTd(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=AND(TRUE,RTD(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=AND(TRUE,   RTD(""foo"", ""bar"", ""baz""))")
    Call testPattern.Add("=1+RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1   +     RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1-RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1*RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1/RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1^RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1=RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1<RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=1>RTD(""foo"", ""bar"", ""baz"")")
    Call testPattern.Add("=""1""&RTD(""foo"", ""bar"", ""baz"")")
    
    Dim i As Long
    Dim range_ As Range
    For i = 1 To testPattern.Count
        Set range_ = TargetWorksheet.Cells(i, i)
        range_.formula = testPattern(i): dataFunctionsCount = dataFunctionsCount + 1
    Next
    
    Set DecolateWorksheetForRealTimeDataFunctionsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForRealTimeDataFunctionsTest(ByVal Target As Workbook, ByRef dataFunctionsCount As Long) As Workbook
    dataFunctionsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim dataFunctionsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForRealTimeDataFunctionsTest(worksheet_, dataFunctionsCountInWorksheet): dataFunctionsCount = dataFunctionsCount + dataFunctionsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForRealTimeDataFunctionsTest = Target
End Function

'@TestMethod("RealTimeDataFunctions")
Public Sub ListRealTimeDataFunctionsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim dataFunctionsCount As Long
    Call DecolateWorkbookForRealTimeDataFunctionsTest(instance.Target, dataFunctionsCount)
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim dataFunctionsCountInWorksheet As Long
    Call DecolateWorksheetForRealTimeDataFunctionsTest(testWorksheet, dataFunctionsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedRealTimeDataFunctionsCell As Collection
    Set listedRealTimeDataFunctionsCell = instance.ListRealTimeDataFunctionsInWorksheet(testWorksheet)
    
    Assert.AreEqual dataFunctionsCountInWorksheet, listedRealTimeDataFunctionsCell.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("RealTimeDataFunctions")
Public Sub ListRealTimeDataFunctions_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim dataFunctionsCount As Long
    Call DecolateWorkbookForRealTimeDataFunctionsTest(instance.Target, dataFunctionsCount)
    
    On Error GoTo CATCH
    
    Dim listedRealTimeDataFunctionsCell As Collection
    Set listedRealTimeDataFunctionsCell = instance.ListRealTimeDataFunctions
    
    Assert.AreEqual dataFunctionsCount, listedRealTimeDataFunctionsCell.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Excel Surveys ==

'' == Defined Scenarios ==
Private Function DecolateWorksheetForDefinedScenariosTest(ByVal TargetWorksheet As Worksheet, ByRef definedScenariosCount As Long) As Worksheet
    definedScenariosCount = 0
    
    Dim i As Long
    For i = 1 To 3
        Call TargetWorksheet.Scenarios.Add( _
            Name:="Scrnario" & i, _
            ChangingCells:=TargetWorksheet.Cells(i, i), _
            Values:=Array(vbNullString), _
            Comment:=vbNullString, _
            Locked:=True, _
            Hidden:=False _
        )
        definedScenariosCount = definedScenariosCount + 1
    Next
    
    Set DecolateWorksheetForDefinedScenariosTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForDefinedScenariosTest(ByVal Target As Workbook, ByRef definedScenariosCount As Long) As Workbook
    definedScenariosCount = 0
    
    Dim worksheet_ As Worksheet
    Dim definedScenariosCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForDefinedScenariosTest(worksheet_, definedScenariosCountInWorksheet): definedScenariosCount = definedScenariosCount + definedScenariosCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForDefinedScenariosTest = Target
End Function

'@TestMethod("DefinedScenarios")
Public Sub ListDefinedScenariosInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim definedScenariosCount As Long
    Call DecolateWorkbookForDefinedScenariosTest(instance.Target, definedScenariosCount)
    
    Set testWorksheet = testWorkbook.Worksheets.Add
    Dim definedScenariosCountInWorksheet As Long
    Call DecolateWorksheetForDefinedScenariosTest(testWorksheet, definedScenariosCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedDefinedScenarios As Collection
    Set listedDefinedScenarios = instance.ListDefinedScenariosInWorksheet(testWorksheet)
    
    Assert.AreEqual definedScenariosCountInWorksheet, listedDefinedScenarios.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DefinedScenarios")
Public Sub ListDefinedScenarios_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim definedScenariosCount As Long
    Call DecolateWorkbookForDefinedScenariosTest(instance.Target, definedScenariosCount)
    
    On Error GoTo CATCH
    
    Dim listedDefinedScenarios As Collection
    Set listedDefinedScenarios = instance.ListDefinedScenarios
    
    Assert.AreEqual definedScenariosCount, listedDefinedScenarios.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DefinedScenarios")
Public Sub DeleteDefinedScenariosInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim definedScenariosCount As Long
    Call DecolateWorkbookForDefinedScenariosTest(instance.Target, definedScenariosCount)
    
    Set testWorksheet = testWorkbook.Worksheets.Add
    Dim definedScenariosCountInWorksheet As Long
    Call DecolateWorksheetForDefinedScenariosTest(testWorksheet, definedScenariosCountInWorksheet)
    
    Dim listedDefinedScenarios As Collection
    Set listedDefinedScenarios = instance.ListDefinedScenariosInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim scenario_ As Scenario
    Dim dict As Object
    For Each scenario_ In listedDefinedScenarios
        Set dict = CreateObject("Scripting.Dictionary")
        With scenario_
            dict("Type") = TypeName(scenario_)
            dict("Name") = .Name
            dict("Location") = GetWorksheetLocation(GetParentWorksheet(scenario_))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedDefinedScenarios As Collection
    Set deletedDefinedScenarios = instance.DeleteDefinedScenariosInWorksheet(testWorksheet)
    
    Assert.AreEqual definedScenariosCountInWorksheet, deletedDefinedScenarios.Count
    Assert.AreEqual definedScenariosCount, instance.ListDefinedScenarios.Count
    Assert.IsTrue DictionariesEquals(deletedDefinedScenarios, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DefinedScenarios")
Public Sub DeleteDefinedScenarios_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim definedScenariosCount As Long
    Call DecolateWorkbookForDefinedScenariosTest(instance.Target, definedScenariosCount)
    
    Dim listedDefinedScenarios As Collection
    Set listedDefinedScenarios = instance.ListDefinedScenarios
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim scenario_ As Scenario
    Dim dict As Object
    For Each scenario_ In listedDefinedScenarios
        Set dict = CreateObject("Scripting.Dictionary")
        With scenario_
            dict("Type") = TypeName(scenario_)
            dict("Name") = .Name
            dict("Location") = GetWorksheetLocation(GetParentWorksheet(scenario_))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedDefinedScenarios As Collection
    Set deletedDefinedScenarios = instance.DeleteDefinedScenarios
    
    Assert.AreEqual definedScenariosCount, deletedDefinedScenarios.Count
    Assert.AreEqual 0&, instance.ListDefinedScenarios.Count
    Assert.IsTrue DictionariesEquals(deletedDefinedScenarios, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Active Filters ==
Private Function DecolateWorksheetForActiveFiltersTest(ByVal TargetWorksheet As Worksheet, ByVal ActivateFilter As Boolean, ByRef ActiveFiltersCount As Long) As Worksheet
    ActiveFiltersCount = 0
    
    Dim sourceData_ As Variant
    sourceData_ = SampleTableData
    
    Dim dataRange As Range
    With TargetWorksheet
        Set dataRange = .Range(.Cells(1, 1), .Cells(UBound(sourceData_, 1), UBound(sourceData_, 2)))
    End With
    dataRange = sourceData_
    
    If ActivateFilter Then
        Call dataRange.AutoFilter(Field:=1, Criteria1:="=AAA"): ActiveFiltersCount = ActiveFiltersCount + 1
    Else
        Call dataRange.AutoFilter
    End If
    
    Set DecolateWorksheetForActiveFiltersTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForActiveFiltersTest(ByVal Target As Workbook, ByRef ActiveFiltersCount As Long) As Workbook
    ActiveFiltersCount = 0
    
    Dim worksheet_ As Worksheet
    Dim ActiveFiltersCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 3 <> 0 Then
            Call DecolateWorksheetForActiveFiltersTest(worksheet_, worksheet_.Index Mod 3 = 1, ActiveFiltersCountInWorksheet): ActiveFiltersCount = ActiveFiltersCount + ActiveFiltersCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForActiveFiltersTest = Target
End Function

'@TestMethod("ActiveFilters")
Public Sub ListActiveFilters_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim ActiveFiltersCount As Long
    Call DecolateWorkbookForActiveFiltersTest(instance.Target, ActiveFiltersCount)
    
    On Error GoTo CATCH
    
    Dim listedActiveFilter As Collection
    Set listedActiveFilter = instance.ListActiveFilters
    
    Assert.AreEqual ActiveFiltersCount, listedActiveFilter.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ActiveFilters")
Public Sub InactivateActiveFilters_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim ActiveFiltersCount As Long
    Call DecolateWorkbookForActiveFiltersTest(instance.Target, ActiveFiltersCount)
    
    On Error GoTo CATCH
    
    Dim inactivatedFilter As Collection
    Set inactivatedFilter = instance.InactivateActiveFilters
    
    Assert.AreEqual ActiveFiltersCount, inactivatedFilter.Count
    Assert.AreEqual 0&, instance.ListActiveFilters.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Custom Worksheet Properties ==
Private Function DecolateWorksheetForCustomProperiesTest(ByVal TargetWorksheet As Worksheet, ByRef customWorksheetPropertiesCount As Long) As Worksheet
    customWorksheetPropertiesCount = 0
    
    Dim i As Long
    For i = 1 To 3
        Call TargetWorksheet.CustomProperties.Add("prop" & i, i): customWorksheetPropertiesCount = customWorksheetPropertiesCount + 1
    Next
    
    Set DecolateWorksheetForCustomProperiesTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForCustomProperiesTest(ByVal Target As Workbook, ByRef customWorksheetPropertiesCount As Long) As Workbook
    customWorksheetPropertiesCount = 0
    
    Dim worksheet_ As Worksheet
    Dim customWorksheetPropertiesCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForCustomProperiesTest(worksheet_, customWorksheetPropertiesCountInWorksheet): customWorksheetPropertiesCount = customWorksheetPropertiesCount + customWorksheetPropertiesCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForCustomProperiesTest = Target
End Function

'@TestMethod("CustomWorksheetProperties")
Public Sub ListCustomWorksheetPropertiesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customWorksheetPropertiesCount As Long
    Call DecolateWorkbookForCustomProperiesTest(instance.Target, customWorksheetPropertiesCount)
    
    Dim customWorksheetPropertiesCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForCustomProperiesTest(instance.Target.Worksheets.Add, customWorksheetPropertiesCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedCustomWorksheetProperties As Collection
    Set listedCustomWorksheetProperties = instance.ListCustomWorksheetPropertiesInWorksheet(testWorksheet)
    
    Assert.AreEqual customWorksheetPropertiesCountInWorksheet, listedCustomWorksheetProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomWorksheetProperties")
Public Sub ListCustomWorksheetProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customWorksheetPropertiesCount As Long
    Call DecolateWorkbookForCustomProperiesTest(instance.Target, customWorksheetPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomWorksheetProperties As Collection
    Set listedCustomWorksheetProperties = instance.ListCustomWorksheetProperties
    
    Assert.AreEqual customWorksheetPropertiesCount, listedCustomWorksheetProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomWorksheetProperties")
Public Sub DeleteCustomWorksheetPropertiesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customWorksheetPropertiesCount As Long
    Call DecolateWorkbookForCustomProperiesTest(instance.Target, customWorksheetPropertiesCount)
    
    Dim customWorksheetPropertiesCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForCustomProperiesTest(instance.Target.Worksheets.Add, customWorksheetPropertiesCountInWorksheet)
    
    Dim listedCustomWorksheetProperties As Collection
    Set listedCustomWorksheetProperties = instance.ListCustomWorksheetPropertiesInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim customProp As CustomProperty
    Dim dict As Object
    For Each customProp In listedCustomWorksheetProperties
        Set dict = CreateObject("Scripting.Dictionary")
        With customProp
            dict("Type") = TypeName(customProp)
            dict("Name") = .Name
            dict("Location") = GetWorksheetLocation(GetParentWorksheet(customProp))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedCustomWorksheetProperties As Collection
    Set deletedCustomWorksheetProperties = instance.DeleteCustomWorksheetPropertiesInWorksheet(testWorksheet)
    
    Assert.AreEqual customWorksheetPropertiesCountInWorksheet, deletedCustomWorksheetProperties.Count
    Assert.IsTrue instance.ListCustomWorksheetProperties.Count > 0
    Assert.IsTrue DictionariesEquals(deletedCustomWorksheetProperties, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomWorksheetProperties")
Public Sub DeleteCustomWorksheetProperties_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customWorksheetPropertiesCount As Long
    Call DecolateWorkbookForCustomProperiesTest(instance.Target, customWorksheetPropertiesCount)
    
    Dim listedCustomWorksheetProperties As Collection
    Set listedCustomWorksheetProperties = instance.ListCustomWorksheetProperties
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim customProp As CustomProperty
    Dim dict As Object
    For Each customProp In listedCustomWorksheetProperties
        Set dict = CreateObject("Scripting.Dictionary")
        With customProp
            dict("Type") = TypeName(customProp)
            dict("Name") = .Name
            dict("Location") = GetWorksheetLocation(GetParentWorksheet(customProp))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedCustomWorksheetProperties As Collection
    Set deletedCustomWorksheetProperties = instance.DeleteCustomWorksheetProperties
    
    Assert.AreEqual customWorksheetPropertiesCount, deletedCustomWorksheetProperties.Count
    Assert.AreEqual 0&, instance.ListCustomWorksheetProperties.Count
    Assert.IsTrue DictionariesEquals(deletedCustomWorksheetProperties, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Hidden Names ==
Private Function DecolateWorksheetForHiddenNamesTest(ByVal TargetWorksheet As Worksheet, ByRef VisibleNamesCount As Long, ByRef HiddenNamesCount As Long) As Worksheet
    HiddenNamesCount = 0
    VisibleNamesCount = 0
    
    Dim i As Long
    Dim nameObject As Name
    For i = 1 To 9
        Set nameObject = TargetWorksheet.names.Add(Name:="NameIn" & TargetWorksheet.Name & "_" & i, RefersToR1C1:="=" & TargetWorksheet.Name & "!A" & i)
        If i Mod 2 = 0 Then
            nameObject.Visible = True
            VisibleNamesCount = VisibleNamesCount + 1
        Else
            nameObject.Visible = False
            HiddenNamesCount = HiddenNamesCount + 1
        End If
    Next
    
    Set DecolateWorksheetForHiddenNamesTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForHiddenNamesTest(ByVal Target As Workbook, ByRef visibleNamesInWorkbookCount As Long, ByRef hiddenNamesInWorkbookCount As Long, ByRef visibleNamesInWorksheetsCount As Long, ByRef hiddenNamesInWorksheetsCount As Long) As Workbook
    visibleNamesInWorkbookCount = 0
    hiddenNamesInWorkbookCount = 0
    visibleNamesInWorksheetsCount = 0
    hiddenNamesInWorksheetsCount = 0
    
    Dim i As Long
    Dim visibleCount As Long
    Dim hiddenCount As Long
    Dim nameObject As Name
    For i = 1 To 5
        Set nameObject = Target.names.Add(Name:="NameIn" & Target.Name & "_" & i, RefersToR1C1:="=" & Target.Worksheets(1).Name & "!A" & i)
        If i Mod 2 = 0 Then
            nameObject.Visible = True: visibleNamesInWorkbookCount = visibleNamesInWorkbookCount + 1
        Else
            nameObject.Visible = False: hiddenNamesInWorkbookCount = hiddenNamesInWorkbookCount + 1
        End If
    Next
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForHiddenNamesTest(worksheet_, visibleCount, hiddenCount): visibleNamesInWorksheetsCount = visibleNamesInWorksheetsCount + visibleCount: hiddenNamesInWorksheetsCount = hiddenNamesInWorksheetsCount + hiddenCount
        End If
    Next
    
    Set DecolateWorkbookForHiddenNamesTest = Target
End Function

'@TestMethod("HiddenNames")
Public Sub ListHiddenNamesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    Dim visibleNamesInTestWorksheetCount As Long
    Dim hiddenNamesInTestWorksheetCount As Long
    Set testWorksheet = DecolateWorksheetForHiddenNamesTest(instance.Target.Worksheets.Add, visibleNamesInTestWorksheetCount, hiddenNamesInTestWorksheetCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNamesInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenNamesInTestWorksheetCount, listedHiddenNames.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub ListHiddenNamesInWorkbook_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNamesInWorkbook()
    
    Assert.AreEqual hiddenNamesInWorkbookCount, listedHiddenNames.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub ListHiddenNames_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNames()
    
    Assert.AreEqual hiddenNamesInWorkbookCount + hiddenNamesInWorksheetsCount, listedHiddenNames.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub VisualizeHiddenNamesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    Dim visibleNamesInTestWorksheetCount As Long
    Dim hiddenNamesInTestWorksheetCount As Long
    Set testWorksheet = DecolateWorksheetForHiddenNamesTest(instance.Target.Worksheets.Add, visibleNamesInTestWorksheetCount, hiddenNamesInTestWorksheetCount)
    
    On Error GoTo CATCH
    
    Dim visualizedNames As Collection
    Set visualizedNames = instance.VisualizeHiddenNamesInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenNamesInTestWorksheetCount, visualizedNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + hiddenNamesInWorksheetsCount, instance.ListHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + visibleNamesInWorkbookCount + hiddenNamesInWorksheetsCount + visibleNamesInWorksheetsCount + hiddenNamesInTestWorksheetCount + visibleNamesInTestWorksheetCount, instance.Target.names.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub VisualizeHiddenNamesInWorkbook_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    On Error GoTo CATCH
    
    Dim visualizedNames As Collection
    Set visualizedNames = instance.VisualizeHiddenNamesInWorkbook
    
    Assert.AreEqual hiddenNamesInWorkbookCount, visualizedNames.Count
    Assert.AreEqual hiddenNamesInWorksheetsCount, instance.ListHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + visibleNamesInWorkbookCount + hiddenNamesInWorksheetsCount + visibleNamesInWorksheetsCount, instance.Target.names.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub VisualizeHiddenNames_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    On Error GoTo CATCH
    
    Dim visualizedNames As Collection
    Set visualizedNames = instance.VisualizeHiddenNames
    
    Assert.AreEqual hiddenNamesInWorkbookCount + hiddenNamesInWorksheetsCount, visualizedNames.Count
    Assert.AreEqual 0&, instance.ListHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + visibleNamesInWorkbookCount + hiddenNamesInWorksheetsCount + visibleNamesInWorksheetsCount, instance.Target.names.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub DeleteHiddenNamesInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    Dim visibleNamesInTestWorksheetCount As Long
    Dim hiddenNamesInTestWorksheetCount As Long
    Set testWorksheet = DecolateWorksheetForHiddenNamesTest(instance.Target.Worksheets.Add, visibleNamesInTestWorksheetCount, hiddenNamesInTestWorksheetCount)
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNamesInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim name_ As Name
    Dim dict As Object
    For Each name_ In listedHiddenNames
        Set dict = CreateObject("Scripting.Dictionary")
        With name_
            dict("Type") = TypeName(name_)
            dict("Name") = .Name
            If TypeOf .Parent Is Workbook Then
                dict("Location") = GetWorkbookLocation(GetParentWorkbook(name_))
            Else
                dict("Location") = GetWorksheetLocation(GetParentWorksheet(name_))
            End If
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenNames As Collection
    Set deletedHiddenNames = instance.DeleteHiddenNamesInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenNamesInTestWorksheetCount, deletedHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + hiddenNamesInWorksheetsCount, instance.ListHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorkbookCount + visibleNamesInWorkbookCount + hiddenNamesInWorksheetsCount + visibleNamesInWorksheetsCount + visibleNamesInTestWorksheetCount, instance.Target.names.Count
    Assert.IsTrue DictionariesEquals(deletedHiddenNames, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub DeleteHiddenNamesInWorkbook_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNamesInWorkbook
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim name_ As Name
    Dim dict As Object
    For Each name_ In listedHiddenNames
        Set dict = CreateObject("Scripting.Dictionary")
        With name_
            dict("Type") = TypeName(name_)
            dict("Name") = .Name
            If TypeOf .Parent Is Workbook Then
                dict("Location") = GetWorkbookLocation(GetParentWorkbook(name_))
            Else
                dict("Location") = GetWorksheetLocation(GetParentWorksheet(name_))
            End If
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenNames As Collection
    Set deletedHiddenNames = instance.DeleteHiddenNamesInWorkbook
    
    Assert.AreEqual hiddenNamesInWorkbookCount, deletedHiddenNames.Count
    Assert.AreEqual hiddenNamesInWorksheetsCount, instance.ListHiddenNames.Count
    Assert.AreEqual visibleNamesInWorkbookCount + hiddenNamesInWorksheetsCount + visibleNamesInWorksheetsCount, instance.Target.names.Count
    Assert.IsTrue DictionariesEquals(deletedHiddenNames, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenNames")
Public Sub DeleteHiddenNames_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim visibleNamesInWorkbookCount As Long
    Dim hiddenNamesInWorkbookCount As Long
    Dim visibleNamesInWorksheetsCount As Long
    Dim hiddenNamesInWorksheetsCount As Long
    Call DecolateWorkbookForHiddenNamesTest(instance.Target, visibleNamesInWorkbookCount, hiddenNamesInWorkbookCount, visibleNamesInWorksheetsCount, hiddenNamesInWorksheetsCount)
    
    Dim listedHiddenNames As Collection
    Set listedHiddenNames = instance.ListHiddenNames
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim name_ As Name
    Dim dict As Object
    For Each name_ In listedHiddenNames
        Set dict = CreateObject("Scripting.Dictionary")
        With name_
            dict("Type") = TypeName(name_)
            dict("Name") = .Name
            If TypeOf .Parent Is Workbook Then
                dict("Location") = GetWorkbookLocation(GetParentWorkbook(name_))
            Else
                dict("Location") = GetWorksheetLocation(GetParentWorksheet(name_))
            End If
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenNames As Collection
    Set deletedHiddenNames = instance.DeleteHiddenNames
    
    Assert.AreEqual hiddenNamesInWorkbookCount + hiddenNamesInWorksheetsCount, deletedHiddenNames.Count
    Assert.AreEqual 0&, instance.ListHiddenNames.Count
    Assert.AreEqual visibleNamesInWorkbookCount + visibleNamesInWorksheetsCount, instance.Target.names.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Ink ==
Private Function DecolateWorksheetForInkTest(ByVal TargetWorksheet As Worksheet, ByRef inkCount As Long) As Worksheet
    inkCount = 0
    
    Dim ink As Shape
    Dim shape_ As Shape
    For Each shape_ In ThisWorkbook.Worksheets(DebugObjectSheetName).Shapes
        If shape_.Type = msoInkComment Then
            Set ink = shape_
            Exit For
        End If
    Next
    
    If ink Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 327, Source:="DecolateWorksheetForInkTest", Description:="For Ink testing, ink object must be in """ & DebugObjectSheetName & """ worksheet_.")
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call ink.Copy: Call TargetWorksheet.Paste: inkCount = inkCount + 1
    Next
    
    Set DecolateWorksheetForInkTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForInkTest(ByVal Target As Workbook, ByRef inkCount As Long) As Workbook
    inkCount = 0
    
    Dim worksheet_ As Worksheet
    Dim inkCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForInkTest(worksheet_, inkCountInWorksheet): inkCount = inkCount + inkCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForInkTest = Target
End Function

'@TestMethod("Ink")
Public Sub ListInkInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorkbookForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim inkCountInWorksheet As Long
    
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorksheetForInkTest(testWorksheet, inkCountInWorksheet)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInkInWorksheet(testWorksheet)
    
    Assert.AreEqual inkCountInWorksheet, listedInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub ListInk_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorkbookForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInk
    
    Assert.AreEqual inkCount, listedInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub DeleteInkInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorkbookForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Set testWorksheet = instance.Target.Worksheets.Add
    Dim inkCountInWorksheet As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorksheetForInkTest(testWorksheet, inkCountInWorksheet)
    On Error GoTo 0
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInkInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedInk
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedInk As Collection
    Set deletedInk = instance.DeleteInkInWorksheet(testWorksheet)
    
    Assert.AreEqual inkCountInWorksheet, deletedInk.Count
    Assert.AreEqual inkCount, instance.ListInk.Count
    Assert.IsTrue DictionariesEquals(deletedInk, expectedDictionaries)
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub DeleteInk_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorkbookForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInk
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedInk
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetParentWorksheet(shape_).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
            dict("Shape.Type") = .Type
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedInk As Collection
    Set deletedInk = instance.DeleteInk
    
    Assert.AreEqual inkCount, deletedInk.Count
    Assert.AreEqual 0&, instance.ListInk.Count
    Assert.IsTrue DictionariesEquals(deletedInk, expectedDictionaries)
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub RemoveInk_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateWorkbookForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    instance.RemoveInk
    
    Assert.AreEqual 0&, instance.ListInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisWorkbook."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Custom XML Data ==
Private Function DecolateWorkbookForCustomXMLDataTest(ByVal Target As Workbook, ByRef CustomXMLDataCount As Long) As Workbook
    CustomXMLDataCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 5
        Call Target.CustomXMLParts.Add: CustomXMLDataCount = CustomXMLDataCount + 1
    Next
    
    Set DecolateWorkbookForCustomXMLDataTest = Target
End Function

'@TestMethod("CustomXMLData")
Public Sub ListCustomXMLData_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testWorkbook.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolateWorkbookForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomXMLData As Collection
    Set listedCustomXMLData = instance.ListCustomXMLData
    
    Assert.AreEqual testWorkbook.CustomXMLParts.Count - builtInPartsCount, listedCustomXMLData.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub DeleteCustomXMLData_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testWorkbook.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolateWorkbookForCustomXMLDataTest(instance.Target, customPartsCount)
    
    Dim listedCustomXMLData As Collection
    Set listedCustomXMLData = instance.ListCustomXMLData
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim xmlPart As CustomXMLPart
    Dim dict As Object
    For Each xmlPart In listedCustomXMLData
        Set dict = CreateObject("Scripting.Dictionary")
        With xmlPart
            dict("Type") = TypeName(xmlPart)
            dict("Id") = .Id
            dict("Location") = GetWorkbookLocation(instance.Target)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedCustomXMLData As Collection
    Set deletedCustomXMLData = instance.DeleteCustomXMLData
    
    Assert.AreEqual 0&, instance.ListCustomXMLData.Count
    Assert.AreEqual builtInPartsCount, instance.Target.CustomXMLParts.Count
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    Assert.IsTrue DictionariesEquals(deletedCustomXMLData, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_NoCustomXMLData_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_WithCustomXMLData_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customPartsCount As Long
    Call DecolateWorkbookForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub FixCustomXMLData_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim customPartsCount As Long
    Call DecolateWorkbookForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixCustomXMLData
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Long External References ==

'' == Hidden Rows and Columns ==
'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenRowsAndColumnsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    Dim hiddenRowsCountInWorksheet As Long: hiddenRowsCountInWorksheet = 0
    Dim hiddenColumnsCountInWorksheet As Long: hiddenColumnsCountInWorksheet = 0
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet), hiddenColumnsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedHiddenRowsAndColumns As Collection
    Set listedHiddenRowsAndColumns = instance.ListHiddenRowsAndColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet + hiddenColumnsCountInWorksheet, listedHiddenRowsAndColumns.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenRowsAndColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenRowsAndColumns As Collection
    Set listedHiddenRowsAndColumns = instance.ListHiddenRowsAndColumns()
    
    Assert.AreEqual hiddenRowsCount + hiddenColumnsCount, listedHiddenRowsAndColumns.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenRowsAndColumnsInWorksheet_CorrectCall_Succcessed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    Dim hiddenRowsCountInWorksheet As Long: hiddenRowsCountInWorksheet = 0
    Dim hiddenColumnsCountInWorksheet As Long: hiddenColumnsCountInWorksheet = 0
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet), hiddenColumnsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedRows As Collection
    Set visualizedRows = instance.VisualizeHiddenRowsAndColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet + hiddenColumnsCountInWorksheet, visualizedRows.Count
    Assert.AreEqual 0&, instance.ListHiddenRowsAndColumnsInWorksheet(testWorksheet).Count
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenRowsAndColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedRowsAndColumns As Collection
    Set visualizedRowsAndColumns = instance.VisualizeHiddenRowsAndColumns()
    
    Assert.AreEqual hiddenRowsCount + hiddenColumnsCount, visualizedRowsAndColumns.Count
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenRowsAndColumns
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenRowsAndColumnsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    Dim hiddenRowsCountInWorksheet As Long: hiddenRowsCountInWorksheet = 0
    Dim hiddenColumnsCountInWorksheet As Long: hiddenColumnsCountInWorksheet = 0
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet), hiddenColumnsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenRowsAndColumns As Collection
    Set listedHiddenRowsAndColumns = instance.ListHiddenRowsAndColumnsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenRowsAndColumns
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenRowsAndColumns As Collection
    Set deletedHiddenRowsAndColumns = instance.DeleteHiddenRowsAndColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet + hiddenColumnsCountInWorksheet, deletedHiddenRowsAndColumns.Count
    Assert.AreEqual 0&, instance.ListHiddenRowsAndColumnsInWorksheet(testWorksheet).Count
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenRowsAndColumns, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenRowsAndColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long: hiddenRowsCount = 0
    Dim hiddenColumnsCount As Long: hiddenColumnsCount = 0
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount), hiddenColumnsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenRowsAndColumns As Collection
    Set listedHiddenRowsAndColumns = instance.ListHiddenRowsAndColumns
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenRowsAndColumns
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenRowsAndColumns As Collection
    Set deletedHiddenRowsAndColumns = instance.DeleteHiddenRowsAndColumns()
    
    Assert.AreEqual hiddenRowsCount + hiddenColumnsCount, deletedHiddenRowsAndColumns.Count
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenRowsAndColumns
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenRowsAndColumns, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub InspectHiddenRowsAndColumns_NoHiddenRowsAndColumns_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenRowsAndColumns
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub InspectHiddenRowsAndColumns_WithHiddenRowsAndColumns()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsAndColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsAndColumnsCount), hiddenRowsAndColumnsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectHiddenRowsAndColumns
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub FixHiddenRowsAndColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsAndColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsAndColumnsCount), hiddenRowsAndColumnsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixHiddenRowsAndColumns
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenRowsAndColumns
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Hidden Rows ===
Private Function DecolateWorksheetForHiddenRowsTest(ByVal TargetWorksheet As Worksheet, ByRef hiddenRowsCount As Long) As Worksheet
    hiddenRowsCount = 0
    
    With TargetWorksheet
        .Range(.Cells(1, 1), .Cells(1, 1)).EntireRow.Hidden = True: hiddenRowsCount = hiddenRowsCount + 1
        .Range(.Cells(3, 1), .Cells(4, 1)).EntireRow.Hidden = True: hiddenRowsCount = hiddenRowsCount + 1
        .Range(.Cells(7, 1), .Cells(.Rows.Count, 1)).EntireRow.Hidden = True: hiddenRowsCount = hiddenRowsCount + 1
    End With
    
    Set DecolateWorksheetForHiddenRowsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForHiddenRowsTest(ByVal Target As Workbook, ByRef hiddenRowsCount As Long) As Workbook
    hiddenRowsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim hiddenRowsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForHiddenRowsTest(worksheet_, hiddenRowsCountInWorksheet): hiddenRowsCount = hiddenRowsCount + hiddenRowsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForHiddenRowsTest = Target
End Function

'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenRowsInWorksheet_CorreclCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    Dim hiddenRowsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedHiddenRows As Collection
    Set listedHiddenRows = instance.ListHiddenRowsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet, listedHiddenRows.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenRows_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenRowsCount As Collection
    Set listedHiddenRowsCount = instance.ListHiddenRows()
    
    Assert.AreEqual hiddenRowsCount, listedHiddenRowsCount.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenRowsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    Dim hiddenRowsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedRows As Collection
    Set visualizedRows = instance.VisualizeHiddenRowsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet, visualizedRows.Count
    Assert.AreEqual 0&, instance.ListHiddenRowsInWorksheet(testWorksheet).Count
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenRows_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedRows As Collection
    Set visualizedRows = instance.VisualizeHiddenRows()
    
    Assert.AreEqual hiddenRowsCount, visualizedRows.Count
    Assert.AreEqual 0&, instance.ListHiddenRows.Count
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenRowsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    Dim hiddenRowsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenRowsTest(instance.Target.Worksheets.Add, hiddenRowsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenRows As Collection
    Set listedHiddenRows = instance.ListHiddenRowsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenRows
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenRows As Collection
    Set deletedHiddenRows = instance.DeleteHiddenRowsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenRowsCountInWorksheet, deletedHiddenRows.Count
    Assert.AreEqual 0&, instance.ListHiddenRowsInWorksheet(testWorksheet).Count
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenRows, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenRows_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenRowsCount As Long
    Call DecolateWorkbookForHiddenRowsTest(instance.Target, hiddenRowsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenRows As Collection
    Set listedHiddenRows = instance.ListHiddenRows
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenRows
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenRows As Collection
    Set deletedHiddenRows = instance.DeleteHiddenRows()
    
    Assert.AreEqual hiddenRowsCount, deletedHiddenRows.Count
    Assert.AreEqual 0&, instance.ListHiddenRows.Count
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenRows, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Hidden Columns ===
Private Function DecolateWorksheetForHiddenColumnsTest(ByVal TargetWorksheet As Worksheet, ByRef hiddenColumnsCount As Long) As Worksheet
    hiddenColumnsCount = 0
    
    With TargetWorksheet
        .Range(.Cells(1, 1), .Cells(1, 1)).EntireColumn.Hidden = True: hiddenColumnsCount = hiddenColumnsCount + 1
        .Range(.Cells(1, 3), .Cells(1, 4)).EntireColumn.Hidden = True: hiddenColumnsCount = hiddenColumnsCount + 1
        .Range(.Cells(1, 7), .Cells(1, .Columns.Count)).EntireColumn.Hidden = True: hiddenColumnsCount = hiddenColumnsCount + 1
    End With
    
    Set DecolateWorksheetForHiddenColumnsTest = TargetWorksheet
End Function

Private Function DecolateWorkbookForHiddenColumnsTest(ByVal Target As Workbook, ByRef hiddenColumnsCount As Long) As Workbook
    hiddenColumnsCount = 0
    
    Dim worksheet_ As Worksheet
    Dim hiddenColumnsCountInWorksheet As Long
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Index Mod 2 = 1 Then
            Call DecolateWorksheetForHiddenColumnsTest(worksheet_, hiddenColumnsCountInWorksheet): hiddenColumnsCount = hiddenColumnsCount + hiddenColumnsCountInWorksheet
        End If
    Next
    
    Set DecolateWorkbookForHiddenColumnsTest = Target
End Function

'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenColumnsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    Dim hiddenColumnsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(instance.Target.Worksheets.Add, hiddenColumnsCountInWorksheet)
    
    On Error GoTo CATCH
    
    Dim listedHiddenColumns As Collection
    Set listedHiddenColumns = instance.ListHiddenColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenColumnsCountInWorksheet, listedHiddenColumns.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub ListHiddenColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenColumnsCount As Collection
    Set listedHiddenColumnsCount = instance.ListHiddenColumns()
    
    Assert.AreEqual hiddenColumnsCount, listedHiddenColumnsCount.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenColumnsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    Dim hiddenColumnsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(instance.Target.Worksheets.Add, hiddenColumnsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedColumns As Collection
    Set visualizedColumns = instance.VisualizeHiddenColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenColumnsCountInWorksheet, visualizedColumns.Count
    Assert.AreEqual 0&, instance.ListHiddenColumnsInWorksheet(testWorksheet).Count
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub VisualizeHiddenColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    On Error GoTo CATCH
    
    Dim visualizedColumns As Collection
    Set visualizedColumns = instance.VisualizeHiddenColumns()
    
    Assert.AreEqual hiddenColumnsCount, visualizedColumns.Count
    Assert.AreEqual 0&, instance.ListHiddenColumns.Count
    Assert.AreEqual "mark", testWorksheet.Range(markerAddress).value
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenColumnsInWorksheet_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    Dim hiddenColumnsCountInWorksheet As Long
    Set testWorksheet = DecolateWorksheetForHiddenColumnsTest(instance.Target.Worksheets.Add, hiddenColumnsCountInWorksheet)
    
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenColumns As Collection
    Set listedHiddenColumns = instance.ListHiddenColumnsInWorksheet(testWorksheet)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenColumns
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenColumns As Collection
    Set deletedHiddenColumns = instance.DeleteHiddenColumnsInWorksheet(testWorksheet)
    
    Assert.AreEqual hiddenColumnsCountInWorksheet, deletedHiddenColumns.Count
    Assert.AreEqual 0&, instance.ListHiddenColumnsInWorksheet(testWorksheet).Count
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenColumns, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenRowsAndColumns")
Public Sub DeleteHiddenColumns_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    Dim testWorksheet As Worksheet
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenColumnsCount As Long
    Call DecolateWorkbookForHiddenColumnsTest(instance.Target, hiddenColumnsCount)
    
    Set testWorksheet = instance.Target.Worksheets(1)
    Dim markerCell As Range
    Dim markerAddress As String
    Set markerCell = testWorksheet.Cells(1, 1)
    markerCell.value = "mark"
    markerAddress = markerCell.Address(External:=True)
    
    Dim listedHiddenColumns As Collection
    Set listedHiddenColumns = instance.ListHiddenColumns
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim range_ As Range
    Dim dict As Object
    For Each range_ In listedHiddenColumns
        Set dict = CreateObject("Scripting.Dictionary")
        With range_
            dict("Type") = TypeName(range_)
            dict("Location") = .Address(External:=True)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenColumns As Collection
    Set deletedHiddenColumns = instance.DeleteHiddenColumns()
    
    Assert.AreEqual hiddenColumnsCount, deletedHiddenColumns.Count
    Assert.AreEqual 0&, instance.ListHiddenColumns.Count
    Assert.AreNotEqual "mark", testWorksheet.Range(markerAddress).value
    Assert.IsTrue DictionariesEquals(deletedHiddenColumns, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Hidden Worksheets ==
Private Function DecolateWorkbookForHiddenWorksheets(ByVal Target As Workbook, ByRef HiddenWorksheetsCount As Long, ByRef VeryHiddenWorksheetsCount As Long) As Workbook
    HiddenWorksheetsCount = 0
    VeryHiddenWorksheetsCount = 0
    
    Dim i As Long
    Dim worksheet_ As Worksheet
    For i = 1 To 9
        Set worksheet_ = Target.Worksheets.Add
        If i Mod 2 = 1 Then
            worksheet_.Visible = xlSheetHidden
            HiddenWorksheetsCount = HiddenWorksheetsCount + 1
        Else
            worksheet_.Visible = xlSheetVeryHidden
            VeryHiddenWorksheetsCount = VeryHiddenWorksheetsCount + 1
        End If
    Next
    
    Set DecolateWorkbookForHiddenWorksheets = Target
End Function

'@TestMethod("HiddenWorksheets")
Public Sub ListHiddenWorksheets_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenSheetsCount As Long
    Dim veryHiddenSheetsCount As Long
    Call DecolateWorkbookForHiddenWorksheets(instance.Target, hiddenSheetsCount, veryHiddenSheetsCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenSheets As Collection
    Set listedHiddenSheets = instance.ListHiddenWorksheets
    
    Assert.AreEqual hiddenSheetsCount + veryHiddenSheetsCount, listedHiddenSheets.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenWorksheets")
Public Sub VisualizeHiddenWorksheets_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim defaultSheetCount As Long
    defaultSheetCount = instance.Target.Worksheets.Count
    
    Dim hiddenSheetsCount As Long
    Dim veryHiddenSheetsCount As Long
    Call DecolateWorkbookForHiddenWorksheets(instance.Target, hiddenSheetsCount, veryHiddenSheetsCount)
    
    On Error GoTo CATCH
    
    Dim visualizedSheet As Collection
    Set visualizedSheet = instance.VisualizeHiddenWorksheets
    
    Assert.AreEqual hiddenSheetsCount + veryHiddenSheetsCount, visualizedSheet.Count
    Assert.AreEqual defaultSheetCount + hiddenSheetsCount + veryHiddenSheetsCount, instance.Target.Worksheets.Count
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenWorksheets
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenWorksheets")
Public Sub DeleteHiddenWorksheets_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim defaultSheetCount As Long
    defaultSheetCount = instance.Target.Worksheets.Count
    
    Dim hiddenSheetsCount As Long
    Dim veryHiddenSheetsCount As Long
    Call DecolateWorkbookForHiddenWorksheets(instance.Target, hiddenSheetsCount, veryHiddenSheetsCount)
    
    Dim listedHiddenWorksheets As Collection
    Set listedHiddenWorksheets = instance.ListHiddenWorksheets
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim worksheet_ As Worksheet
    Dim dict As Object
    For Each worksheet_ In listedHiddenWorksheets
        Set dict = CreateObject("Scripting.Dictionary")
        With worksheet_
            dict("Type") = TypeName(worksheet_)
            dict("Name") = .Name
            dict("Location") = GetWorksheetLocation(worksheet_)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedHiddenWorksheets As Collection
    Set deletedHiddenWorksheets = instance.DeleteHiddenWorksheets
    
    Assert.AreEqual hiddenSheetsCount + veryHiddenSheetsCount, deletedHiddenWorksheets.Count
    Assert.AreEqual defaultSheetCount, instance.Target.Worksheets.Count
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenWorksheets
    Assert.IsTrue DictionariesEquals(deletedHiddenWorksheets, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenWorksheets")
Public Sub InspectHiddenWorksheets_NoHiddenWorksheet_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenWorksheets
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenWorksheets")
Public Sub InspectHiddenWorksheets_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenSheetsCount As Long
    Dim veryHiddenSheetsCount As Long
    Call DecolateWorkbookForHiddenWorksheets(instance.Target, hiddenSheetsCount, veryHiddenSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectHiddenWorksheets
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenWorksheets")
Public Sub FixHiddenWorksheets_CorrectCall_Successed()
    Dim testWorkbook As Workbook
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testWorkbook = CreateNewWorkbook(3)
    Dim instance As InspectWorkbookUtils: Set instance = New InspectWorkbookUtils: Call instance.Initialize(testWorkbook)
    
    Dim hiddenSheetsCount As Long
    Dim veryHiddenSheetsCount As Long
    Call DecolateWorkbookForHiddenWorksheets(instance.Target, hiddenSheetsCount, veryHiddenSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixHiddenWorksheets
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenWorksheets
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testWorkbook.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub
