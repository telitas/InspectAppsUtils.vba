Attribute VB_Name = "InspectDocumentUtilsTest"
'@TestModule
'@Folder "InspectDocumentUtilsProject.Tests"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'WhitelistedIdentifiers i, j
Option Explicit
Option Private Module

'@Ignore VariableNotUsed
Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

Private Const DummyCodeString As String = _
"Private Sub Dummy()" & vbCrLf & _
"End Sub"
Private Const SkipTaskPaneAddInsTest As Boolean = True
Private Const DummyImagePath As String = "C:\Path\To\Dummy.png"

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

Private Function CreateNewDocument() As Document
    Dim document_ As Document
    
    Set document_ = Application.Documents.Add
    
    Set CreateNewDocument = document_
End Function

Private Function CreateDummyPdf(ByVal OutputPath As String) As String
    Dim document_ As Document
    Set document_ = Documents.Add
    With document_.PageSetup
        .TopMargin = MillimetersToPoints(0)
        .BottomMargin = MillimetersToPoints(0)
        .LeftMargin = MillimetersToPoints(0)
        .RightMargin = MillimetersToPoints(0)
        .PageWidth = MillimetersToPoints(12.7)
        .PageHeight = MillimetersToPoints(7.2)
    End With
    Call document_.ExportAsFixedFormat( _
        OutputFileName:=OutputPath, _
        ExportFormat:=wdExportFormatPDF _
    )
    Call document_.Close(False)
    CreateDummyPdf = OutputPath
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

Private Function GetDocumentLocation(ByVal TargetDocument As Document) As String
    GetDocumentLocation = "'" & TargetDocument.Path & "[" & TargetDocument.Name & "]'"
End Function

Private Function GetRangeLocation(ByVal TargetRange As Range) As String
    GetRangeLocation = GetDocumentLocation(TargetRange.Parent) & "!Range(" & TargetRange.Start & "," & TargetRange.End & ")"
End Function

Private Function GetShapeLocation(ByVal TargetShape As Shape) As String
    GetShapeLocation = GetDocumentLocation(TargetShape.Parent) & "!Location(" & TargetShape.Left & "," & TargetShape.Top & ")"
End Function

'' = Initialize =

'@TestMethod("Initialize")
Public Sub Initialize_CorrectCall_Succeeded()
    Dim instance As InspectDocumentUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectDocumentUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisDocument)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_CallTwice_RaiseError()
    Dim instance As InspectDocumentUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectDocumentUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisDocument)
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(ThisDocument)
    
    Assert.Fail
    
    GoTo FINALLY
CATCH:

    Assert.AreEqual instance.InvalidOperationError, Err.Number
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_TargetBookIsNothing_RaiseError()
    Dim instance As InspectDocumentUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectDocumentUtils
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
Private Function DecolateDocumentForCommentsTest(ByVal Target As Document) As Document
    Dim i As Long
    Dim range_ As Range
    For i = 1 To 5
        Set range_ = Target.Paragraphs.Last.Range
        With range_
            .Text = "Text " & i
            Call .Comments.Add(range_, "Comment " & i)
        End With
        Call Target.Content.Paragraphs.Add
    Next
    Set DecolateDocumentForCommentsTest = Target
End Function

'@TestMethod("Comments")
Public Sub ListComments_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Call DecolateDocumentForCommentsTest(instance.Target)
    
    On Error GoTo CATCH
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments()
    
    Assert.AreEqual testDocument.Content.Comments.Count, listedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub DeleteCommentsTest_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Call DecolateDocumentForCommentsTest(instance.Target)
    Dim commentCount As Long
    commentCount = testDocument.Content.Comments.Count
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim comment_ As Comment
    Dim dict As Object
    For Each comment_ In listedComments
        Set dict = CreateObject("Scripting.Dictionary")
        With comment_
            dict("Type") = TypeName(comment_)
            dict("Location") = GetRangeLocation(comment_.Range)
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedComments As Collection
    Set deletedComments = instance.DeleteComments
    
    Assert.AreEqual commentCount, deletedComments.Count
    Assert.AreEqual 0&, instance.ListComments.Count
    Assert.IsTrue DictionariesEquals(deletedComments, expectedDictionaries)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Comments")
Public Sub RemoveCommentsTest_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Call DecolateDocumentForCommentsTest(instance.Target)
    
    On Error GoTo CATCH
    
    instance.RemoveComments
    
    Assert.AreEqual 0&, instance.ListComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Ink ==
Private Function DecolateDocumentForInkTest(ByVal TargetDocument As Document, ByRef inkCount As Long) As Document
    inkCount = 0
    
    Dim ink As Shape
    Dim shape_ As Shape
    For Each shape_ In ThisDocument.Shapes
        If shape_.Type = msoInkComment Then
            Set ink = shape_
            Exit For
        End If
    Next
    
    If ink Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 327, Source:="DecolateDocumentForInkTest", Description:="For Ink testing, ink object must be in document.")
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call ink.Select: Call Selection.Copy: Call TargetDocument.Content.Paste: inkCount = inkCount + 1
    Next
    
    Set DecolateDocumentForInkTest = TargetDocument
End Function

'@TestMethod("Ink")
Public Sub ListInk_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateDocumentForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInk
    
    Assert.AreEqual inkCount, listedInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisDocument."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub DeleteInk_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateDocumentForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInk
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Object
    Dim dict As Object
    For Each shape_ In listedInk
        Set dict = CreateObject("Scripting.Dictionary")
        With shape_
            dict("Type") = TypeName(shape_)
            dict("Name") = .Name
            dict("Location") = GetShapeLocation(shape_)
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
    Assert.Inconclusive "There is no Inks object on ThisDocument."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Ink")
Public Sub RemoveInk_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateDocumentForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    instance.RemoveInk
    
    Assert.AreEqual 0&, instance.ListInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on ThisDocument."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@TestMethod("DocumentProperties")
Public Sub RemoveDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateDocumentForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    instance.RemoveDocumentProperties
    
    Assert.AreEqual 0&, instance.ListCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' ==== Built-in Document Properties ====
Private Function DecolateDocumentForBuiltInDocumentProperties(ByVal Target As Document, ByRef builtInDocumentPropertiesCount As Long) As Document
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
    
    Set DecolateDocumentForBuiltInDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateDocumentForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedBuiltInDocumentProperties As Collection
    Set listedBuiltInDocumentProperties = instance.ListBuiltInDocumentProperties
    
    Assert.AreEqual builtInDocumentPropertiesCount, listedBuiltInDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolateDocumentForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' ==== Custom Document Properties ====
Private Function DecolateDocumentForCustomDocumentProperties(ByVal Target As Document, ByRef customDocumentPropertiesCount As Long) As Document
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
    
    Set DecolateDocumentForCustomDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListCustomDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateDocumentForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomDocumentProperties As Collection
    Set listedCustomDocumentProperties = instance.ListCustomDocumentProperties
    
    Assert.AreEqual customDocumentPropertiesCount, listedCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearCustomDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateDocumentForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("DocumentProperties")
Public Sub DeleteCustomDocumentProperties_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolateDocumentForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
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
            dict("Location") = GetDocumentLocation(.Parent)
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Personal Information ===
'@TestMethod("PersonalInformation")
Public Sub RemovePersonalInformation_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    instance.RemovePersonalInformation
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Task Pane Apps ==
'@TestMethod("TaskPaneAddIns")
Public Sub RemoveTaskPaneAddIns_CorrectCall_Successed()
    If SkipTaskPaneAddInsTest Then
        Assert.Inconclusive "TaskPaneAddIns test was skipped."
        Exit Sub
    End If
    
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Embedded Documents ==
Private Function DecolateDocumentForEmbeddedDocumentsTest(ByVal TargetDocument As Document, ByVal DummyPdfPath As String, ByRef EmbeddedDocumentsCount As Long) As Document
    EmbeddedDocumentsCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call TargetDocument.InlineShapes.AddOLEObject( _
            classType:="AcroExch.Document.DC", _
            fileName:=DummyPdfPath, _
            LinkToFile:=False, _
            DisplayAsIcon:=True _
        )
        EmbeddedDocumentsCount = EmbeddedDocumentsCount + 1
    Next
    Selection.Collapse
    Set DecolateDocumentForEmbeddedDocumentsTest = TargetDocument
End Function

'@TestMethod("EmbeddedDocuments")
Public Sub ListEmbeddedDocuments_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisDocument.Name & ".docx.pdf"
    Call CreateDummyPdf(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateDocumentForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub DeleteEmbeddedDocuments_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & ThisDocument.Name & ".docx.pdf"
    Call CreateDummyPdf(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolateDocumentForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocuments
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim inlineShape_ As InlineShape
    Dim dict As Object
    For Each inlineShape_ In listedEmbeddedDocuments
        Set dict = CreateObject("Scripting.Dictionary")
        With inlineShape_
            dict("Type") = TypeName(inlineShape_)
            dict("Location") = GetRangeLocation(.Range)
            dict("InlineShape.Type") = .Type
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Macros, Forms, And ActiveX Controls ==
'' === Macros ===
Private Function DecolateDocumentForMacrosTest(ByVal Target As Document, ByRef macrosCount As Long) As Document
    macrosCount = 0
    
    Dim i As Long
    For i = 1 To Target.VBProject.VBComponents.Count
        With Target.VBProject.VBComponents(i)
            If .Type = 100 Then
                If .Name = "ThisDocument" Or i Mod 2 = 0 Then
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
    
    Set DecolateDocumentForMacrosTest = Target
End Function

'@TestMethod("Macros")
Public Sub ListMacros_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim macrosCount As Long
    Call DecolateDocumentForMacrosTest(instance.Target, macrosCount)
    
    On Error GoTo CATCH
    
    Dim listedMacros As Collection
    Set listedMacros = instance.ListMacros
    
    Assert.AreEqual macrosCount, listedMacros.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Macros")
Public Sub DeleteMacros_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim macrosCount As Long
    Call DecolateDocumentForMacrosTest(instance.Target, macrosCount)
    
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
            dict("Location") = GetDocumentLocation(instance.Target)
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Forms ===
Private Function DecolateDocumentForFormsTest(ByVal TargetDocument As Document) As Document
    Dim i As Long
    For i = 1 To 3
        Call TargetDocument.FormFields.Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
    Next
    
    Set DecolateDocumentForFormsTest = TargetDocument
End Function

'@TestMethod("Forms")
Public Sub ListForms_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Call DecolateDocumentForFormsTest(instance.Target)
    Dim formsCount As Long
    formsCount = testDocument.FormFields.Count
    
    On Error GoTo CATCH
    
    Dim listedForms As Collection
    Set listedForms = instance.ListForms
    
    Assert.AreEqual formsCount, listedForms.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("Forms")
Public Sub DeleteForms_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Call DecolateDocumentForFormsTest(instance.Target)
    Dim formsCount As Long
    formsCount = testDocument.FormFields.Count
    
    Dim listedForms As Collection
    Set listedForms = instance.ListForms
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim form_ As FormField
    Dim dict As Object
    For Each form_ In listedForms
        Set dict = CreateObject("Scripting.Dictionary")
        With form_
            dict("Type") = TypeName(form_)
            dict("Name") = .Name
            dict("Location") = GetRangeLocation(.Range)
            dict("FormField.Type") = .Type
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === ActiveX Controls ===
Private Function DecolateDocumentForActiveXControlsTest(ByVal TargetDocument As Document, ByRef activeXControlsCount As Long) As Document
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
        Call TargetDocument.InlineShapes.AddOLEControl(classType:=classType): activeXControlsCount = activeXControlsCount + 1
    Next
    
    Set DecolateDocumentForActiveXControlsTest = TargetDocument
End Function

'@TestMethod("ActiveXControls")
Public Sub ListActiveXControls_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim activeXControlsCount As Long
    Call DecolateDocumentForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    On Error GoTo CATCH
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Assert.AreEqual activeXControlsCount, listedActiveXControls.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("ActiveXControls")
Public Sub DeleteActiveXControls_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim activeXControlsCount As Long
    Call DecolateDocumentForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim inlineShape_ As InlineShape
    Dim dict As Object
    For Each inlineShape_ In listedActiveXControls
        Set dict = CreateObject("Scripting.Dictionary")
        With inlineShape_
            dict("Type") = TypeName(inlineShape_)
            dict("Location") = GetRangeLocation(.Range)
            dict("InlineShape.Type") = .Type
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Custom XML Data ==
Private Function DecolateDocumentForCustomXMLDataTest(ByVal Target As Document, ByRef CustomXMLDataCount As Long) As Document
    CustomXMLDataCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 5
        Call Target.CustomXMLParts.Add: CustomXMLDataCount = CustomXMLDataCount + 1
    Next
    
    Set DecolateDocumentForCustomXMLDataTest = Target
End Function

'@TestMethod("CustomXMLData")
Public Sub ListCustomXMLData_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testDocument.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolateDocumentForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomXMLData As Collection
    Set listedCustomXMLData = instance.ListCustomXMLData
    
    Assert.AreEqual testDocument.CustomXMLParts.Count - builtInPartsCount, listedCustomXMLData.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub DeleteCustomXMLData_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testDocument.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolateDocumentForCustomXMLDataTest(instance.Target, customPartsCount)
    
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
            dict("Location") = GetDocumentLocation(.Parent.Parent)
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_NoCustomXMLData_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_WithCustomXMLData_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim customPartsCount As Long
    Call DecolateDocumentForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CustomXMLData")
Public Sub FixCustomXMLData_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim customPartsCount As Long
    Call DecolateDocumentForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixCustomXMLData
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Headers, Footers, and Watermarks ==
'' === Headers and Footers ===
'@TestMethod("HeadersAndFooters")
Public Sub InspectHeadersAndFooters_NoHeaderAndFooter_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub InspectHeadersAndFooters_WithHeaderAndFooter_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateDocumentForFootersTest(DecolateDocumentForHeadersTest(instance.Target, containsHeadersSheetsCount), containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub FixHeadersAndFooters_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim containsHeadersSheetsCount As Long
    Call DecolateDocumentForFootersTest(DecolateDocumentForHeadersTest(instance.Target, containsHeadersSheetsCount), containsHeadersSheetsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixHeadersAndFooters
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHeadersAndFooters
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Headers ===
Private Function DecolateSectionForHeadersTest(ByVal TargetSection As Section, ByRef headersCount As Long) As Section
    Dim header_ As HeaderFooter
    Dim index_ As WdHeaderFooterIndex
    Dim i As Long
    For i = 1 To 3
        Select Case i
            Case 1
                index_ = wdHeaderFooterFirstPage
            Case 2
                index_ = wdHeaderFooterPrimary
            Case 3
                index_ = wdHeaderFooterEvenPages
        End Select
        Set header_ = TargetSection.Headers(index_)
        With header_
            .LinkToPrevious = False
            Call .Shapes.AddPicture(DummyImagePath)
            .Range.Text = "Dummy " & TargetSection.Index & "-" & wdHeaderFooterFirstPage
            Call .Range.InlineShapes.AddPicture(DummyImagePath)
            If .Exists And Not .IsEmpty Then
                headersCount = headersCount + 1
            End If
        End With
    Next
    Set DecolateSectionForHeadersTest = TargetSection
End Function

Private Function DecolateDocumentForHeadersTest(ByVal TargetDocument As Document, ByRef headersCount As Long) As Document
    Dim section_ As Section
    Dim i As Long
    Dim j As Long
    Dim index_ As WdHeaderFooterIndex
    For i = 1 To 4
        TargetDocument.Sections.Add
    Next
    For i = 1 To TargetDocument.Sections.Count
        Set section_ = TargetDocument.Sections(i)
        Call DecolateSectionForHeadersTest(section_, headersCount)
        For j = 1 To 3
            Select Case j
                Case 1
                    index_ = wdHeaderFooterFirstPage
                Case 2
                    index_ = wdHeaderFooterPrimary
                Case 3
                    index_ = wdHeaderFooterEvenPages
            End Select
            With section_.Headers(index_)
                .LinkToPrevious = (i Mod 2 = 1)
                If .LinkToPrevious And (.Exists And Not .IsEmpty) Then
                    headersCount = headersCount - 1
                End If
            End With
        Next
    Next
    Set DecolateDocumentForHeadersTest = TargetDocument
End Function

'@TestMethod("HeadersAndHeaders")
Public Sub ListHeadersInSection_CorrectCall_Successed()
    If Dir(DummyImagePath) = vbNullString Then
        Assert.Inconclusive "Dummy image does not exist."
        Exit Sub
    End If
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim headersCount As Long
    Call DecolateDocumentForHeadersTest(instance.Target, headersCount)
    
    On Error GoTo CATCH
    
    Dim testSection As Section
    Set testSection = instance.Target.Sections(1)
    Dim i As Long
    Dim expectCount As Long
    Dim index_ As WdHeaderFooterIndex
    For i = 1 To 3
        Select Case i
            Case 1
                index_ = wdHeaderFooterFirstPage
            Case 2
                index_ = wdHeaderFooterPrimary
            Case 3
                index_ = wdHeaderFooterEvenPages
        End Select
        With testSection.Headers(index_)
            .LinkToPrevious = (i Mod 2 = 1)
            If .Exists And Not .IsEmpty Then
                expectCount = expectCount + 1
            End If
        End With
    Next
    
    Dim listedHeaders As Collection
    Set listedHeaders = instance.ListHeadersInSection(testSection)
    
    Assert.AreEqual expectCount, listedHeaders.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndHeaders")
Public Sub ListHeaders_CorrectCall_Successed()
    If Dir(DummyImagePath) = vbNullString Then
        Assert.Inconclusive "Dummy image does not exist."
        Exit Sub
    End If
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim headersCount As Long
    Call DecolateDocumentForHeadersTest(instance.Target, headersCount)
    
    On Error GoTo CATCH
    
    Dim listedHeaders As Collection
    Set listedHeaders = instance.ListHeaders
    
    Assert.AreEqual headersCount, listedHeaders.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' === Footers ===
Private Function DecolateSectionForFootersTest(ByVal TargetSection As Section, ByRef footersCount As Long) As Section
    Dim footer_ As HeaderFooter
    Dim index_ As WdHeaderFooterIndex
    Dim i As Long
    For i = 1 To 3
        Select Case i
            Case 1
                index_ = wdHeaderFooterFirstPage
            Case 2
                index_ = wdHeaderFooterPrimary
            Case 3
                index_ = wdHeaderFooterEvenPages
        End Select
        Set footer_ = TargetSection.Footers(index_)
        With footer_
            .LinkToPrevious = False
            Call .Shapes.AddPicture(DummyImagePath)
            .Range.Text = "Dummy " & TargetSection.Index & "-" & wdHeaderFooterFirstPage
            Call .Range.InlineShapes.AddPicture(DummyImagePath)
            If .Exists And Not .IsEmpty Then
                footersCount = footersCount + 1
            End If
        End With
    Next
    Set DecolateSectionForFootersTest = TargetSection
End Function

Private Function DecolateDocumentForFootersTest(ByVal TargetDocument As Document, ByRef footersCount As Long) As Document
    Dim section_ As Section
    Dim i As Long
    Dim j As Long
    Dim index_ As WdHeaderFooterIndex
    For i = 1 To 4
        TargetDocument.Sections.Add
    Next
    For i = 1 To TargetDocument.Sections.Count
        Set section_ = TargetDocument.Sections(i)
        Call DecolateSectionForFootersTest(section_, footersCount)
        For j = 1 To 3
            Select Case j
                Case 1
                    index_ = wdHeaderFooterFirstPage
                Case 2
                    index_ = wdHeaderFooterPrimary
                Case 3
                    index_ = wdHeaderFooterEvenPages
            End Select
            With section_.Footers(index_)
                .LinkToPrevious = (i Mod 2 = 1)
                If .LinkToPrevious And (.Exists And Not .IsEmpty) Then
                    footersCount = footersCount - 1
                End If
            End With
        Next
    Next
    Set DecolateDocumentForFootersTest = TargetDocument
End Function

'@TestMethod("HeadersAndFooters")
Public Sub ListFootersInSection_CorrectCall_Successed()
    If Dir(DummyImagePath) = vbNullString Then
        Assert.Inconclusive "Dummy image does not exist."
        Exit Sub
    End If
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim footersCount As Long
    Call DecolateDocumentForFootersTest(instance.Target, footersCount)
    
    On Error GoTo CATCH
    
    Dim testSection As Section
    Set testSection = instance.Target.Sections(1)
    Dim i As Long
    Dim expectCount As Long
    Dim index_ As WdHeaderFooterIndex
    For i = 1 To 3
        Select Case i
            Case 1
                index_ = wdHeaderFooterFirstPage
            Case 2
                index_ = wdHeaderFooterPrimary
            Case 3
                index_ = wdHeaderFooterEvenPages
        End Select
        With testSection.Footers(index_)
            .LinkToPrevious = (i Mod 2 = 1)
            If .Exists And Not .IsEmpty Then
                expectCount = expectCount + 1
            End If
        End With
    Next
    
    Dim listedFooters As Collection
    Set listedFooters = instance.ListFootersInSection(testSection)
    
    Assert.AreEqual expectCount, listedFooters.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HeadersAndFooters")
Public Sub ListFooters_CorrectCall_Successed()
    If Dir(DummyImagePath) = vbNullString Then
        Assert.Inconclusive "Dummy image does not exist."
        Exit Sub
    End If
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim footersCount As Long
    Call DecolateDocumentForFootersTest(instance.Target, footersCount)
    
    On Error GoTo CATCH
    
    Dim listedFooters As Collection
    Set listedFooters = instance.ListFooters
    
    Assert.AreEqual footersCount, listedFooters.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Invisible Content ==
Private Function DecolateDocumentForInvisibleContentTest(ByVal TargetDocument As Document, ByRef invisibleContentCount As Long) As Document
    invisibleContentCount = 0
    
    Dim i As Long
    Dim shape_ As Shape
    For i = 1 To 5
        Set shape_ = TargetDocument.Shapes.AddShape(msoShapeRound1Rectangle, 10 * i, 10 * i, 10, 10)
        If i Mod 2 <> 0 Then
            shape_.Visible = False: invisibleContentCount = invisibleContentCount + 1
        End If
    Next
    
    Set DecolateDocumentForInvisibleContentTest = TargetDocument
End Function

'@TestMethod("InvisibleContent")
Public Sub ListInvisibleContent_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim invisibleContentCount As Long
    Call DecolateDocumentForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContent
    
    Assert.AreEqual invisibleContentCount, listedInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub VisualizeInvisibleContent_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim invisibleContentCount As Long
    Call DecolateDocumentForInvisibleContentTest(instance.Target, invisibleContentCount)
    
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub DeleteInvisibleContent_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim invisibleContentCount As Long
    Call DecolateDocumentForInvisibleContentTest(instance.Target, invisibleContentCount)
    
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
            dict("Location") = GetShapeLocation(shape_)
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
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_NoInvisibleContent_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_WithInvisibleContent_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim invisibleContentCount As Long
    Call DecolateDocumentForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("InvisibleContent")
Public Sub FixInvisibleContent_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim invisibleContentCount As Long
    Call DecolateDocumentForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixInvisibleContent
    Assert.AreEqual 0&, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Collapsed Headings ==
Private Function DecolateDocumentForCollapsedHeadingsTest(ByVal TargetDocument As Document, ByRef collapsedParagraphsCount As Long) As Document
    collapsedParagraphsCount = 0
    
    Dim i As Long
    Dim paragraph_ As Paragraph
    Const paragraphsCount As Long = 6
    For i = 1 To paragraphsCount
        Call TargetDocument.Paragraphs.Add
    Next
    For i = 1 To paragraphsCount
        Set paragraph_ = TargetDocument.Paragraphs(i)
        paragraph_.Range.Text = "Paragraph " & i
        Call paragraph_.Range.InsertParagraphAfter
        If i Mod 2 = 1 Then
            paragraph_.Style = TargetDocument.Styles(wdStyleHeading1)
        End If
    Next
    For i = 1 To paragraphsCount
        Set paragraph_ = TargetDocument.Paragraphs(i)
        If i Mod 4 = 1 Then
            paragraph_.CollapsedState = True
            collapsedParagraphsCount = collapsedParagraphsCount + 1
        End If
    Next
    
    Set DecolateDocumentForCollapsedHeadingsTest = TargetDocument
End Function

'@TestMethod("CollapsedHeadings")
Public Sub ListCollapsedHeadings_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim collapsedParagraphsCount As Long
    Call DecolateDocumentForCollapsedHeadingsTest(instance.Target, collapsedParagraphsCount)
    
    On Error GoTo CATCH
    
    Dim listedCollapsedHeadings As Collection
    Set listedCollapsedHeadings = instance.ListCollapsedHeadings
    
    Assert.AreEqual collapsedParagraphsCount, listedCollapsedHeadings.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CollapsedHeadings")
Public Sub ExpandCollapsedHeadings_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim collapsedParagraphsCount As Long
    Call DecolateDocumentForCollapsedHeadingsTest(instance.Target, collapsedParagraphsCount)
    
    On Error GoTo CATCH
    
    Dim visualizedCollapsedHeadings As Collection
    Set visualizedCollapsedHeadings = instance.ExpandCollapsedHeadings
    
    Assert.AreEqual collapsedParagraphsCount, visualizedCollapsedHeadings.Count
    Assert.AreEqual 0&, instance.ListCollapsedHeadings.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CollapsedHeadings")
Public Sub InspectCollapsedHeadings_NoCollapsedHeadings_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCollapsedHeadings
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CollapsedHeadings")
Public Sub InspectCollapsedHeadings_WithCollapsedHeadings_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim collapsedParagraphsCount As Long
    Call DecolateDocumentForCollapsedHeadingsTest(instance.Target, collapsedParagraphsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectCollapsedHeadings
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("CollapsedHeadings")
Public Sub FixCollapsedHeadings_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim collapsedParagraphsCount As Long
    Call DecolateDocumentForCollapsedHeadingsTest(instance.Target, collapsedParagraphsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixCollapsedHeadings
    Assert.AreEqual 0&, instance.ListCollapsedHeadings.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'' == Hidden Text ==
Private Function DecolateDocumentForHiddenTextTest(ByVal TargetDocument As Document, ByRef hiddenTextCount As Long) As Document
    hiddenTextCount = 0
    
    Dim i As Long
    Dim paragraph_ As Paragraph
    Const paragraphsCount As Long = 6
    For i = 1 To paragraphsCount
        Call TargetDocument.Paragraphs.Add
    Next
    For i = 1 To paragraphsCount
        Set paragraph_ = TargetDocument.Paragraphs(i)
        paragraph_.Range.Text = "Paragraph " & i
        Call paragraph_.Range.InsertParagraphAfter
        If i Mod 2 = 1 Then
            paragraph_.Range.Font.Hidden = True
            hiddenTextCount = hiddenTextCount + 1
        End If
    Next
    
    Set DecolateDocumentForHiddenTextTest = TargetDocument
End Function

'@TestMethod("HiddenText")
Public Sub ListHiddenText_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim hiddenTextCount As Long
    Call DecolateDocumentForHiddenTextTest(instance.Target, hiddenTextCount)
    
    On Error GoTo CATCH
    
    Dim listedHiddenText As Collection
    Set listedHiddenText = instance.ListHiddenText
    
    Assert.AreEqual hiddenTextCount, listedHiddenText.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenText")
Public Sub VisualizeHiddenText_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim hiddenTextCount As Long
    Call DecolateDocumentForHiddenTextTest(instance.Target, hiddenTextCount)
    
    On Error GoTo CATCH
    
    Dim visualizedHiddenText As Collection
    Set visualizedHiddenText = instance.VisualizeHiddenText
    
    Assert.AreEqual hiddenTextCount, visualizedHiddenText.Count
    Assert.AreEqual 0&, instance.ListHiddenText.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenText")
Public Sub DeleteHiddenText_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim hiddenTextCount As Long
    Call DecolateDocumentForHiddenTextTest(instance.Target, hiddenTextCount)
    
    On Error GoTo CATCH
    
    Dim visualizedHiddenText As Collection
    Set visualizedHiddenText = instance.DeleteHiddenText
    
    Assert.AreEqual hiddenTextCount, visualizedHiddenText.Count
    Assert.AreEqual 0&, instance.ListHiddenText.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenText")
Public Sub InspectHiddenText_NoHiddenText_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectHiddenText
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenText")
Public Sub InspectHiddenText_WithHiddenText_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim hiddenTextCount As Long
    Call DecolateDocumentForHiddenTextTest(instance.Target, hiddenTextCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectHiddenText
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub

'@TestMethod("HiddenText")
Public Sub FixHiddenText_CorrectCall_Successed()
    Dim testDocument As Document
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = SetScreenUpdating(False)
    
    Set testDocument = CreateNewDocument
    Dim instance As InspectDocumentUtils: Set instance = New InspectDocumentUtils: Call instance.Initialize(testDocument)
    
    Dim hiddenTextCount As Long
    Call DecolateDocumentForHiddenTextTest(instance.Target, hiddenTextCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixHiddenText
    Assert.AreEqual 0&, instance.ListHiddenText.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testDocument.Close(False)
    Call SetScreenUpdating(currentScreenUpdating)
End Sub
