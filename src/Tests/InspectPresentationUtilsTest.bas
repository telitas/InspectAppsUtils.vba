Attribute VB_Name = "InspectPresentationUtilsTest"
'@TestModule
'@Folder "InspectPresentationUtilsProject.Tests"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'WhitelistedIdentifiers i, j
Option Explicit
Option Private Module

'@Ignore VariableNotUsed
Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

Private TesterPresentation As Presentation
Private Const DummyCodeString As String = _
"Private Sub Dummy()" & vbCrLf & _
"End Sub"

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
    If Presentations.Count > 1 Then
        Call Err.Raise(Number:=445, Description:="Close other presentations.")
    End If
    Set TesterPresentation = Presentations(1)
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Function CreateNewPresentation(ByVal SlidesCount As Long) As Presentation
    Dim presentation_ As Presentation
    
    Set presentation_ = Application.Presentations.Add
    Dim currentDisplayAlerts As Boolean
    Dim i As Long
    With presentation_
        If .Slides.Count <= SlidesCount Then
            For i = .Slides.Count + 1 To SlidesCount
                Call .Slides.Add(1, ppLayoutBlank)
            Next
        Else
            For i = .Slides.Count To SlidesCount + 1 Step -1
                currentDisplayAlerts = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Call .Slides(i).Delete
                Application.DisplayAlerts = currentDisplayAlerts
            Next
        End If
    End With
    
    Set CreateNewPresentation = presentation_
End Function

Private Function CreateDummyImage(ByVal OutputPath As String) As String
    Dim presentation_ As Presentation
    Set presentation_ = Presentations.Add
    Dim slide_ As Slide
    Set slide_ = presentation_.Slides.AddSlide(1, presentation_.SlideMaster.CustomLayouts(7))
    With presentation_.PageSetup
        .SlideWidth = 72
        .SlideHeight = 72
    End With
    Call slide_.Export(OutputPath, "jpg")
    Call presentation_.Close
    
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

Public Function GetParentPresentation(ByVal TargetObject As Object) As Presentation
    If TypeOf TargetObject Is Presentation Then
        Set GetParentPresentation = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentPresentation = GetParentPresentation(TargetObject.Parent)
End Function

Public Function GetParentSlide(ByVal TargetObject As Object) As Slide
    If TypeOf TargetObject Is Slide Then
        Set GetParentSlide = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Presentation Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentSlide = GetParentSlide(TargetObject.Parent)
End Function

Private Function GetPresentationLocation(ByVal TargetPresentation As Presentation) As String
    GetPresentationLocation = "'" & TargetPresentation.Path & "[" & TargetPresentation.Name & "]'"
End Function

Private Function GetSlideLocation(ByVal TargetSlide As Slide) As String
    GetSlideLocation = GetPresentationLocation(GetParentPresentation(TargetSlide)) & "!Slide(" & TargetSlide.SlideIndex & ")"
End Function

Private Function GetShapeLocation(ByVal TargetShape As Shape) As String
    GetShapeLocation = GetSlideLocation(GetParentSlide(TargetShape)) & "!Point(" & TargetShape.Left & ", " & TargetShape.Top & ")"
End Function

Sub Wait(ByVal WaitTime As Long)
    Dim start As Long
    start = Timer
    Do While Timer < start + WaitTime
        DoEvents
    Loop
End Sub

'' = Initialize =
'@TestMethod("Initialize")
Public Sub Initialize_CorrectCall_Succeeded()
    Dim instance As InspectPresentationUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectPresentationUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(TesterPresentation)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_CallTwice_RaiseError()
    Dim instance As InspectPresentationUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectPresentationUtils
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(TesterPresentation)
    '@Ignore ArgumentWithIncompatibleObjectType
    Call instance.Initialize(TesterPresentation)
    
    Assert.Fail
    
    GoTo FINALLY
CATCH:

    Assert.AreEqual instance.InvalidOperationError, Err.Number
    
    Resume FINALLY
FINALLY:
End Sub

'@TestMethod("Initialize")
Public Sub Initialize_TargetPresentationIsNothing_RaiseError()
    Dim instance As InspectPresentationUtils
    
    On Error GoTo CATCH
    
    Set instance = New InspectPresentationUtils
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
Private Function DecolateSheetForCommentsTest(ByVal TargetSlide As Slide, ByRef commentsCount As Long) As Slide
    commentsCount = 0
    Dim i As Long
    For i = 1 To 5
        Call TargetSlide.Comments.Add2(0, 0, "Test User", "TU", "Comment " & i, "Windows Live", "testuser"): commentsCount = commentsCount + 1
    Next
    
    Set DecolateSheetForCommentsTest = TargetSlide
End Function

Private Function DecolatePresentationForCommentsTest(ByVal Target As Presentation, ByRef commentsCount As Long) As Presentation
    commentsCount = 0
    Dim slide_ As Slide
    Dim commentsCountInSlide As Long
    
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForCommentsTest(slide_, commentsCountInSlide): commentsCount = commentsCount + commentsCountInSlide
        End If
    Next
    
    Set DecolatePresentationForCommentsTest = Target
End Function

'@TestMethod("Comments")
Public Sub ListCommentsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim commentsCount As Long
    Call DecolatePresentationForCommentsTest(instance.Target, commentsCount)
    
    Set testSlide = instance.Target.Slides.AddSlide(1, instance.Target.SlideMaster.CustomLayouts(7))
    
    Dim commentsCountInSlide As Long
    Call DecolateSheetForCommentsTest(testSlide, commentsCountInSlide)
    
    On Error GoTo CATCH
    
    Dim listedComments As Collection
    Set listedComments = instance.ListCommentsInSlide(testSlide)
    
    Assert.AreEqual commentsCountInSlide, listedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Comments")
Public Sub ListComments_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim commentsCount As Long
    Call DecolatePresentationForCommentsTest(instance.Target, commentsCount)
    
    On Error GoTo CATCH
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments()
    
    Assert.AreEqual commentsCount, listedComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Comments")
Public Sub DeleteCommentsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim commentsCount As Long
    Call DecolatePresentationForCommentsTest(instance.Target, commentsCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    
    Dim commentsCountInSlide As Long
    Call DecolateSheetForCommentsTest(testSlide, commentsCountInSlide)
    
    Dim listedComments As Collection
    Set listedComments = instance.ListCommentsInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim comment_ As Comment
    Dim dict As Object
    For Each comment_ In listedComments
        Set dict = CreateObject("Scripting.Dictionary")
        With comment_
            dict("Type") = TypeName(comment_)
            dict("Location") = GetSlideLocation(GetParentSlide(comment_))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedCommentsInSlide As Collection
    Set deletedCommentsInSlide = instance.DeleteCommentsInSlide(testSlide)
    
    Assert.AreEqual commentsCountInSlide, deletedCommentsInSlide.Count
    Assert.AreEqual commentsCount, instance.ListComments.Count
    Assert.IsTrue DictionariesEquals(deletedCommentsInSlide, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Comments")
Public Sub DeleteCommentsTest_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim commentsCount As Long
    Call DecolatePresentationForCommentsTest(instance.Target, commentsCount)
    
    Dim listedComments As Collection
    Set listedComments = instance.ListComments
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim comment_ As Comment
    Dim dict As Object
    For Each comment_ In listedComments
        Set dict = CreateObject("Scripting.Dictionary")
        With comment_
            dict("Type") = TypeName(comment_)
            dict("Location") = GetSlideLocation(GetParentSlide(comment_))
        End With
        Call expectedDictionaries.Add(dict)
    Next
    
    On Error GoTo CATCH
    
    Dim deletedComments As Collection
    Set deletedComments = instance.DeleteComments
    
    Assert.AreEqual commentsCount, deletedComments.Count
    Assert.AreEqual 0&, instance.ListComments.Count
    Assert.IsTrue DictionariesEquals(deletedComments, expectedDictionaries)
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Comments")
Public Sub RemoveCommentsTest_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim commentsCount As Long
    Call DecolatePresentationForCommentsTest(instance.Target, commentsCount)
    
    On Error GoTo CATCH
    
    instance.RemoveComments
    
    Assert.AreEqual 0&, instance.ListComments.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@TestMethod("DocumentProperties")
Public Sub RemoveDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolatePresentationForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    instance.RemoveDocumentProperties
    
    Assert.AreEqual 0&, instance.ListCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' ==== Built-in Document Properties ====
Private Function DecolatePresentationForBuiltInDocumentProperties(ByVal Target As Presentation, ByRef builtInDocumentPropertiesCount As Long) As Presentation
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
    
    Set DecolatePresentationForBuiltInDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolatePresentationForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedBuiltInDocumentProperties As Collection
    Set listedBuiltInDocumentProperties = instance.ListBuiltInDocumentProperties
    
    Assert.AreEqual builtInDocumentPropertiesCount, listedBuiltInDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearBuiltInDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim builtInDocumentPropertiesCount As Long
    Call DecolatePresentationForBuiltInDocumentProperties(instance.Target, builtInDocumentPropertiesCount)
    
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
    Call testPresentation.Close
End Sub

'' ==== Custom Document Properties ====
Private Function DecolatePresentationForCustomDocumentProperties(ByVal Target As Presentation, ByRef customDocumentPropertiesCount As Long) As Presentation
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
    
    Set DecolatePresentationForCustomDocumentProperties = Target
End Function

'@TestMethod("DocumentProperties")
Public Sub ListCustomDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolatePresentationForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomDocumentProperties As Collection
    Set listedCustomDocumentProperties = instance.ListCustomDocumentProperties
    
    Assert.AreEqual customDocumentPropertiesCount, listedCustomDocumentProperties.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("DocumentProperties")
Public Sub ClearCustomDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolatePresentationForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
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
    Call testPresentation.Close
End Sub

'@TestMethod("DocumentProperties")
Public Sub DeleteCustomDocumentProperties_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customDocumentPropertiesCount As Long
    Call DecolatePresentationForCustomDocumentProperties(instance.Target, customDocumentPropertiesCount)
    
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
            dict("Location") = GetPresentationLocation(GetParentPresentation(docProp))
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
    Call testPresentation.Close
End Sub

'' === Personal Information ===
'@TestMethod("PersonalInformation")
Public Sub RemovePersonalInformation_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    On Error GoTo CATCH
    
    instance.RemovePersonalInformation
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Content Add-ins ==
Private Function DecolateSheetForContentAddInsTest(ByVal TargetSlide As Slide, ByRef contentAddInsCount As Long) As Slide
    contentAddInsCount = 0
    
    Dim app As Shape
    Dim shape_ As Shape
    For Each shape_ In TesterPresentation.Slides(1).Shapes
        If shape_.Type = msoContentApp Then
            Set app = shape_
            Exit For
        End If
    Next
    
    If app Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 327, Source:="DecolateSheetForInkTest", Description:="For Ink testing, ink object must be in Slides(1).")
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 2
        Call Wait(1)
        Call app.Copy
        Call Wait(1)
        Call TargetSlide.Shapes.Paste: contentAddInsCount = contentAddInsCount + 1
    Next
    
    Set DecolateSheetForContentAddInsTest = TargetSlide
End Function

Private Function DecolatePresentationForContentAddInsTest(ByVal Target As Presentation, ByRef contentAddInsCount As Long) As Presentation
    contentAddInsCount = 0
    
    Dim slide_ As Slide
    Dim contentAddInsCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForContentAddInsTest(slide_, contentAddInsCountInSlide): contentAddInsCount = contentAddInsCount + contentAddInsCountInSlide
        End If
    Next
    
    Set DecolatePresentationForContentAddInsTest = Target
End Function

'@TestMethod("ContentAddIns")
Public Sub ListContentAddInsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim contentAddInsCountInSlide As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateSheetForContentAddInsTest(testSlide, contentAddInsCountInSlide)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddInsInSlide(testSlide)
    
    Assert.AreEqual contentAddInsCountInSlide, listedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ContentAddIns")
Public Sub ListContentAddIns_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddIns
    
    Assert.AreEqual contentAddInsCount, listedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ContentAddIns")
Public Sub ConvertContentAddInsToImagesInSlide_CollectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim contentAddInsCountInSlide As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateSheetForContentAddInsTest(testSlide, contentAddInsCountInSlide)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim convertedContentAddIns As Collection
    Set convertedContentAddIns = instance.ConvertContentAddInsToImagesInSlide(testSlide)
    
    Assert.AreEqual contentAddInsCountInSlide, convertedContentAddIns.Count
    Assert.AreEqual contentAddInsCount, instance.ListContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ContentAddIns")
Public Sub ConvertContentAddInsToImages_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim convertedContentAddIns As Collection
    Set convertedContentAddIns = instance.ConvertContentAddInsToImages
    
    Dim slide_ As Slide
    Dim shape_ As Shape
    Dim pictureCount As Long
    For Each slide_ In instance.Target.Slides
        For Each shape_ In slide_.Shapes
            If shape_.Type = msoPicture Then
                pictureCount = pictureCount + 1
            End If
        Next
    Next
    
    Assert.AreEqual contentAddInsCount, convertedContentAddIns.Count
    Assert.AreEqual 0&, instance.ListContentAddIns.Count
    Assert.AreEqual pictureCount, convertedContentAddIns.Count
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ContentAddIns")
Public Sub DeleteContentAddInsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
    On Error GoTo 0
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim contentAddInsCountInSlide As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolateSheetForContentAddInsTest(testSlide, contentAddInsCountInSlide)
    On Error GoTo 0
    
    Dim listedContentAddIns As Collection
    Set listedContentAddIns = instance.ListContentAddInsInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedContentAddIns
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
    
    Dim deletedContentAddIns As Collection
    Set deletedContentAddIns = instance.DeleteContentAddInsInSlide(testSlide)
    
    Assert.AreEqual contentAddInsCountInSlide, deletedContentAddIns.Count
    Assert.AreEqual contentAddInsCount, instance.ListContentAddIns.Count
    Assert.IsTrue DictionariesEquals(deletedContentAddIns, expectedDictionaries)
    
    GoTo FINALLY
NO_CONTENT_APPS_FOUND:
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ContentAddIns")
Public Sub DeleteContentAddIns_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim contentAddInsCount As Long
    On Error GoTo NO_CONTENT_APPS_FOUND
    Call DecolatePresentationForContentAddInsTest(instance.Target, contentAddInsCount)
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
            dict("Location") = GetShapeLocation(shape_)
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
    Assert.Inconclusive "There is no ContentAddIns object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Embedded Documents ==
Private Function DecolateSheetForEmbeddedDocumentsTest(ByVal TargetSlide As Slide, ByVal DummyFilePath As String, ByRef EmbeddedDocumentsCount As Long) As Slide
    EmbeddedDocumentsCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call TargetSlide.Shapes.AddOLEObject( _
            fileName:=DummyFilePath, _
            link:=False, _
            DisplayAsIcon:=False _
        )
        EmbeddedDocumentsCount = EmbeddedDocumentsCount + 1
    Next
    
    Set DecolateSheetForEmbeddedDocumentsTest = TargetSlide
End Function

Private Function DecolatePresentationForEmbeddedDocumentsTest(ByVal Target As Presentation, ByVal DummyFilePath As String, ByRef EmbeddedDocumentsCount As Long) As Presentation
    EmbeddedDocumentsCount = 0
    
    Dim slide_ As Slide
    Dim EmbeddedDocumentsCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForEmbeddedDocumentsTest(slide_, DummyFilePath, EmbeddedDocumentsCountInSlide): EmbeddedDocumentsCount = EmbeddedDocumentsCount + EmbeddedDocumentsCountInSlide
        End If
    Next
    
    Set DecolatePresentationForEmbeddedDocumentsTest = Target
End Function

'@TestMethod("EmbeddedDocuments")
Public Sub ListEmbeddedDocumentsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & TesterPresentation.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolatePresentationForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim EmbeddedDocumentsCountInSlide As Long
    Call DecolateSheetForEmbeddedDocumentsTest(testSlide, dummyFile, EmbeddedDocumentsCountInSlide)
    
    On Error GoTo CATCH
    
    Call instance.ListEmbeddedDocumentsInSlide(testSlide)
    
    Assert.AreEqual EmbeddedDocumentsCountInSlide, instance.ListEmbeddedDocumentsInSlide(testSlide).Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testPresentation.Close
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub ListEmbeddedDocuments_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & TesterPresentation.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolatePresentationForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
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
    Call testPresentation.Close
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub DeleteEmbeddedDocumentsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & TesterPresentation.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolatePresentationForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim EmbeddedDocumentsCountInSlide As Long
    Call DecolateSheetForEmbeddedDocumentsTest(testSlide, dummyFile, EmbeddedDocumentsCountInSlide)
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocumentsInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedEmbeddedDocuments
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
    
    Dim deletedEmbeddedDocuments As Collection
    Set deletedEmbeddedDocuments = instance.DeleteEmbeddedDocumentsInSlide(testSlide)
    
    Assert.AreEqual EmbeddedDocumentsCountInSlide, deletedEmbeddedDocuments.Count
    Assert.AreEqual EmbeddedDocumentsCount, instance.ListEmbeddedDocuments.Count
    Assert.IsTrue DictionariesEquals(deletedEmbeddedDocuments, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Kill dummyFile
    Call testPresentation.Close
End Sub

'@TestMethod("EmbeddedDocuments")
Public Sub DeleteEmbeddedDocuments_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dummyFile As String
    dummyFile = fso.GetSpecialFolder(2) & "\" & TesterPresentation.Name & ".xlsm.jpg"
    Call CreateDummyImage(dummyFile)
    
    Dim EmbeddedDocumentsCount As Long
    Call DecolatePresentationForEmbeddedDocumentsTest(instance.Target, dummyFile, EmbeddedDocumentsCount)
    
    Dim listedEmbeddedDocuments As Collection
    Set listedEmbeddedDocuments = instance.ListEmbeddedDocuments
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedEmbeddedDocuments
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
    Call testPresentation.Close
End Sub

'' == Macros, Forms, And ActiveX Controls ==
'' === Macros ===
Private Function DecolatePresentationForMacrosTest(ByVal Target As Presentation, ByRef macrosCount As Long) As Presentation
    macrosCount = 0
    
    Dim i As Long
    For i = 1 To Target.VBProject.VBComponents.Count
        With Target.VBProject.VBComponents(i)
            If .Type = 100 Then
                If .Name = "TesterPresentation" Or i Mod 2 = 0 Then
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
    
    Set DecolatePresentationForMacrosTest = Target
End Function

'@TestMethod("Macros")
Public Sub ListMacros_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim macrosCount As Long
    Call DecolatePresentationForMacrosTest(instance.Target, macrosCount)
    
    On Error GoTo CATCH
    
    Dim listedMacros As Collection
    Set listedMacros = instance.ListMacros
    
    Assert.AreEqual macrosCount, listedMacros.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Macros")
Public Sub DeleteMacros_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim macrosCount As Long
    Call DecolatePresentationForMacrosTest(instance.Target, macrosCount)
    
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
            dict("Location") = GetPresentationLocation(instance.Target)
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
    Call testPresentation.Close
End Sub

'' === ActiveX Controls ===
Private Function DecolateSheetForActiveXControlsTest(ByVal TargetSlide As Slide, ByRef activeXControlsCount As Long) As Slide
    activeXControlsCount = 0
    
    Dim i As Long
    Dim className_ As String
    For i = 1 To 13
        Select Case i
            Case 1
              className_ = "Forms.CommandButton.1"
            Case 2
              className_ = "Forms.ComboBox.1"
            Case 3
              className_ = "Forms.CheckBox.1"
            Case 4
              className_ = "Forms.ListBox.1"
            Case 5
              className_ = "Forms.TextBox.1"
            Case 6
              className_ = "Forms.ScrollBar.1"
            Case 7
              className_ = "Forms.SpinButton.1"
            Case 8
              className_ = "Forms.OptionButton.1"
            Case 9
              className_ = "Forms.Label.1"
            Case 10
              className_ = "Forms.Image.1"
            Case 11
              className_ = "Forms.ToggleButton.1"
        End Select
        Call TargetSlide.Shapes.AddOLEObject(ClassName:=className_, Left:=50 * i, Top:=10, Width:=20, Height:=20): activeXControlsCount = activeXControlsCount + 1
    Next
    
    Set DecolateSheetForActiveXControlsTest = TargetSlide
End Function

Private Function DecolatePresentationForActiveXControlsTest(ByVal Target As Presentation, ByRef activeXControlsCount As Long) As Presentation
    activeXControlsCount = 0
    
    Dim slide_ As Slide
    Dim activeXControlsCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForActiveXControlsTest(slide_, activeXControlsCountInSlide): activeXControlsCount = activeXControlsCount + activeXControlsCountInSlide
        End If
    Next
    
    Set DecolatePresentationForActiveXControlsTest = Target
End Function

'@TestMethod("ActiveXControls")
Public Sub ListActiveXControlsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim activeXControlsCount As Long
    Call DecolatePresentationForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim activeXControlsCountInSlide As Long
    Call DecolateSheetForActiveXControlsTest(testSlide, activeXControlsCountInSlide)
    
    On Error GoTo CATCH
    
    Call instance.ListActiveXControlsInSlide(testSlide)
    
    Assert.AreEqual activeXControlsCountInSlide, instance.ListActiveXControlsInSlide(testSlide).Count
        
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ActiveXControls")
Public Sub ListActiveXControls_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim activeXControlsCount As Long
    Call DecolatePresentationForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    On Error GoTo CATCH
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Assert.AreEqual activeXControlsCount, listedActiveXControls.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ActiveXControls")
Public Sub DeleteActiveXControlsInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim activeXControlsCount As Long
    Call DecolatePresentationForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim activeXControlsCountInSlide As Long
    Call DecolateSheetForActiveXControlsTest(testSlide, activeXControlsCountInSlide)
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControlsInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedActiveXControls
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
    
    Dim deletedActiveXControls As Collection
    Set deletedActiveXControls = instance.DeleteActiveXControlsInSlide(testSlide)
    
    Assert.AreEqual activeXControlsCountInSlide, deletedActiveXControls.Count
    Assert.AreEqual activeXControlsCount, instance.ListActiveXControls.Count
    Assert.IsTrue DictionariesEquals(deletedActiveXControls, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("ActiveXControls")
Public Sub DeleteActiveXControls_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim activeXControlsCount As Long
    Call DecolatePresentationForActiveXControlsTest(instance.Target, activeXControlsCount)
    
    Dim listedActiveXControls As Collection
    Set listedActiveXControls = instance.ListActiveXControls
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedActiveXControls
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
    Call testPresentation.Close
End Sub

'' == Ink ==
Private Function DecolateSheetForInkTest(ByVal TargetSlide As Slide, ByRef inkCount As Long) As Slide
    inkCount = 0
    
    Dim ink As Shape
    Dim shape_ As Shape
    For Each shape_ In TesterPresentation.Slides(1).Shapes
        If shape_.Type = msoInkComment Then
            Set ink = shape_
            Exit For
        End If
    Next
    
    If ink Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 327, Source:="DecolateSheetForInkTest", Description:="For Ink testing, ink object must be in Slides(1).")
    End If
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 3
        Call ink.Copy: Call TargetSlide.Shapes.Paste: inkCount = inkCount + 1
    Next
    
    Set DecolateSheetForInkTest = TargetSlide
End Function

Private Function DecolatePresentationForInkTest(ByVal Target As Presentation, ByRef inkCount As Long) As Presentation
    inkCount = 0
    
    Dim slide_ As Slide
    Dim inkCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForInkTest(slide_, inkCountInSlide): inkCount = inkCount + inkCountInSlide
        End If
    Next
    
    Set DecolatePresentationForInkTest = Target
End Function

'@TestMethod("Ink")
Public Sub ListInkInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolatePresentationForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim inkCountInSlide As Long
    
    On Error GoTo NO_INKS_FOUND
    Call DecolateSheetForInkTest(testSlide, inkCountInSlide)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInkInSlide(testSlide)
    
    Assert.AreEqual inkCountInSlide, listedInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Ink")
Public Sub ListInk_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolatePresentationForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInk
    
    Assert.AreEqual inkCount, listedInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Ink")
Public Sub DeleteInkInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolatePresentationForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim inkCountInSlide As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolateSheetForInkTest(testSlide, inkCountInSlide)
    On Error GoTo 0
    
    Dim listedInk As Collection
    Set listedInk = instance.ListInkInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
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
    Set deletedInk = instance.DeleteInkInSlide(testSlide)
    
    Assert.AreEqual inkCountInSlide, deletedInk.Count
    Assert.AreEqual inkCount, instance.ListInk.Count
    Assert.IsTrue DictionariesEquals(deletedInk, expectedDictionaries)
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Ink")
Public Sub DeleteInk_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolatePresentationForInkTest(instance.Target, inkCount)
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
    Assert.Inconclusive "There is no Inks object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("Ink")
Public Sub RemoveInk_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim inkCount As Long
    On Error GoTo NO_INKS_FOUND
    Call DecolatePresentationForInkTest(instance.Target, inkCount)
    On Error GoTo 0
    
    On Error GoTo CATCH
    
    instance.RemoveInk
    
    Assert.AreEqual 0&, instance.ListInk.Count
    
    GoTo FINALLY
NO_INKS_FOUND:
    Assert.Inconclusive "There is no Inks object on TesterPresentation."
    Resume FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Custom XML Data ==
Private Function DecolatePresentationForCustomXMLDataTest(ByVal Target As Presentation, ByRef CustomXMLDataCount As Long) As Presentation
    CustomXMLDataCount = 0
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 5
        Call Target.CustomXMLParts.Add: CustomXMLDataCount = CustomXMLDataCount + 1
    Next
    
    Set DecolatePresentationForCustomXMLDataTest = Target
End Function

'@TestMethod("CustomXMLData")
Public Sub ListCustomXMLData_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testPresentation.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolatePresentationForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Dim listedCustomXMLData As Collection
    Set listedCustomXMLData = instance.ListCustomXMLData
    
    Assert.AreEqual testPresentation.CustomXMLParts.Count - builtInPartsCount, listedCustomXMLData.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("CustomXMLData")
Public Sub DeleteCustomXMLData_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim builtInPartsCount As Long
    builtInPartsCount = testPresentation.CustomXMLParts.Count
    
    Dim customPartsCount As Long
    Call DecolatePresentationForCustomXMLDataTest(instance.Target, customPartsCount)
    
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
            dict("Location") = GetPresentationLocation(.Parent.Parent)
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
    Call testPresentation.Close
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_NoCustomXMLData_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("CustomXMLData")
Public Sub InspectCustomXMLData_WithCustomXMLData_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("CustomXMLData")
Public Sub FixCustomXMLData_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForCustomXMLDataTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixCustomXMLData
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectCustomXMLData
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Invisible On-Slide Content ==
Private Function DecolateSheetForInvisibleContentTest(ByVal TargetSlide As Slide, ByRef invisibleContentCount As Long) As Slide
    invisibleContentCount = 0
    
    Dim i As Long
    Dim shape_ As Shape
    For i = 1 To 5
        Set shape_ = TargetSlide.Shapes.AddShape(msoShapeRound1Rectangle, 10 * i, 10 * i, 10, 10)
        If i Mod 2 <> 0 Then
            shape_.Visible = False: invisibleContentCount = invisibleContentCount + 1
        End If
    Next
    
    Set DecolateSheetForInvisibleContentTest = TargetSlide
End Function

Private Function DecolatePresentationForInvisibleContentTest(ByVal Target As Presentation, ByRef invisibleContentCount As Long) As Presentation
    invisibleContentCount = 0
    
    Dim slide_ As Slide
    Dim invisibleContentCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForInvisibleContentTest(slide_, invisibleContentCountInSlide): invisibleContentCount = invisibleContentCount + invisibleContentCountInSlide
        End If
    Next
    
    Set DecolatePresentationForInvisibleContentTest = Target
End Function

'@TestMethod("InvisibleContent")
Public Sub ListInvisibleContentInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testSlide = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Dim invisibleContentCountInSlide As Long
    Call DecolateSheetForInvisibleContentTest(testSlide, invisibleContentCountInSlide)
    
    On Error GoTo CATCH
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContentInSlide(testSlide)
    
    Assert.AreEqual invisibleContentCountInSlide, listedInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub ListInvisibleContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContent
    
    Assert.AreEqual invisibleContentCount, listedInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub VisualizeInvisibleContentInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testSlide = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Dim invisibleContentCountInSlide As Long
    Call DecolateSheetForInvisibleContentTest(testSlide, invisibleContentCountInSlide)
    
    On Error GoTo CATCH
    
    Dim visualizedInvisibleContent As Collection
    Set visualizedInvisibleContent = instance.VisualizeInvisibleContentInSlide(testSlide)
    
    Assert.AreEqual invisibleContentCountInSlide, visualizedInvisibleContent.Count
    Assert.AreEqual invisibleContentCount, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub VisualizeInvisibleContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
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
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub DeleteInvisibleContentInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    Set testSlide = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Dim invisibleContentCountInSlide As Long
    Call DecolateSheetForInvisibleContentTest(testSlide, invisibleContentCountInSlide)
    
    Dim listedInvisibleContent As Collection
    Set listedInvisibleContent = instance.ListInvisibleContentInSlide(testSlide)
    
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
    Set deletedInvisibleContent = instance.DeleteInvisibleContentInSlide(testSlide)
    
    Assert.AreEqual invisibleContentCountInSlide, deletedInvisibleContent.Count
    Assert.AreEqual invisibleContentCount, instance.ListInvisibleContent.Count
    Assert.IsTrue DictionariesEquals(deletedInvisibleContent, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub DeleteInvisibleContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
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
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_NoInvisibleContent_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub InspectInvisibleContent_WithInvisibleContent_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectInvisibleContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("InvisibleContent")
Public Sub FixInvisibleContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim invisibleContentCount As Long
    Call DecolatePresentationForInvisibleContentTest(instance.Target, invisibleContentCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixInvisibleContent
    Assert.AreEqual 0&, instance.ListInvisibleContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Revision Tracking Data ==

'' == Off-slide Content ==
Private Function DecolateSheetForOffSlideContentTest(ByVal TargetSlide As Slide, ByRef OffSlideContentCount As Long) As Slide
    OffSlideContentCount = 0
    
    Dim pageSetup_ As PageSetup
    Set pageSetup_ = TargetSlide.Parent.PageSetup
    With TargetSlide.Shapes
        Call .AddShape(msoShapeRectangle, -1, -1, 1, 1)
        Call .AddShape(msoShapeRectangle, -1 - 0.1, -1 - 0.1, 1, 1): OffSlideContentCount = OffSlideContentCount + 1
        Call .AddShape(msoShapeRectangle, pageSetup_.SlideWidth, -1, 1, 1)
        Call .AddShape(msoShapeRectangle, pageSetup_.SlideWidth + 0.1, -1, 1, 1): OffSlideContentCount = OffSlideContentCount + 1
        Call .AddShape(msoShapeRectangle, -1, pageSetup_.SlideHeight, 1, 1)
        Call .AddShape(msoShapeRectangle, -1, pageSetup_.SlideHeight + 0.1, 1, 1): OffSlideContentCount = OffSlideContentCount + 1
        Call .AddShape(msoShapeRectangle, pageSetup_.SlideWidth, pageSetup_.SlideHeight, 1, 1)
        Call .AddShape(msoShapeRectangle, pageSetup_.SlideWidth + 0.1, pageSetup_.SlideHeight + 0.1, 1, 1): OffSlideContentCount = OffSlideContentCount + 1
    End With
    
    Set DecolateSheetForOffSlideContentTest = TargetSlide
End Function

Private Function DecolatePresentationForOffSlideContentTest(ByVal Target As Presentation, ByRef OffSlideContentCount As Long) As Presentation
    OffSlideContentCount = 0
    
    Dim slide_ As Slide
    Dim OffSlideContentCountInSlide As Long
    For Each slide_ In Target.Slides
        If slide_.SlideIndex Mod 2 = 1 Then
            Call DecolateSheetForOffSlideContentTest(slide_, OffSlideContentCountInSlide): OffSlideContentCount = OffSlideContentCount + OffSlideContentCountInSlide
        End If
    Next
    
    Set DecolatePresentationForOffSlideContentTest = Target
End Function

'@TestMethod("OffSlideContent")
Public Sub ListOffSlideContentInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim OffSlideContentCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, OffSlideContentCount)
    
    Set testSlide = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Dim OffSlideContentCountInSlide As Long
    Call DecolateSheetForOffSlideContentTest(testSlide, OffSlideContentCountInSlide)
    
    On Error GoTo CATCH
    
    Dim listedOffSlideContent As Collection
    Set listedOffSlideContent = instance.ListOffSlideContentInSlide(testSlide)
    
    Assert.AreEqual OffSlideContentCountInSlide, listedOffSlideContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("OffSlideContent")
Public Sub ListOffSlideContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim OffSlideContentCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, OffSlideContentCount)
    
    On Error GoTo CATCH
    
    Dim listedOffSlideContent As Collection
    Set listedOffSlideContent = instance.ListOffSlideContent
    
    Assert.AreEqual OffSlideContentCount, listedOffSlideContent.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("OffSlideContent")
Public Sub DeleteOffSlideContentInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim OffSlideContentCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, OffSlideContentCount)
    
    Set testSlide = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Dim OffSlideContentCountInSlide As Long
    Call DecolateSheetForOffSlideContentTest(testSlide, OffSlideContentCountInSlide)
    
    Dim listedOffSlideContent As Collection
    Set listedOffSlideContent = instance.ListOffSlideContentInSlide(testSlide)
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedOffSlideContent
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
    
    Dim deletedOffSlideContent As Collection
    Set deletedOffSlideContent = instance.DeleteOffSlideContentInSlide(testSlide)
    
    Assert.AreEqual OffSlideContentCountInSlide, deletedOffSlideContent.Count
    Assert.AreEqual OffSlideContentCount, instance.ListOffSlideContent.Count
    Assert.IsTrue DictionariesEquals(deletedOffSlideContent, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("OffSlideContent")
Public Sub DeleteOffSlideContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim OffSlideContentCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, OffSlideContentCount)
    
    Dim listedOffSlideContent As Collection
    Set listedOffSlideContent = instance.ListOffSlideContent
    
    Dim expectedDictionaries As Collection: Set expectedDictionaries = New Collection
    Dim shape_ As Shape
    Dim dict As Object
    For Each shape_ In listedOffSlideContent
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
    
    Dim deletedOffSlideContent As Collection
    Set deletedOffSlideContent = instance.DeleteOffSlideContent
    
    Assert.AreEqual OffSlideContentCount, deletedOffSlideContent.Count
    Assert.AreEqual 0&, instance.ListOffSlideContent.Count
    Assert.IsTrue DictionariesEquals(deletedOffSlideContent, expectedDictionaries)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("OffSlideContent")
Public Sub InspectOffSlideContent_WithOffSlideContent_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectOffSlideContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("OffSlideContent")
Public Sub FixOffSlideContent_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForOffSlideContentTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixOffSlideContent
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectOffSlideContent
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'' == Presentation Notes ==
Private Function DecolateSheetForPresentationNotesTest(ByVal TargetSlide As Slide, ByRef slidesContainNotesCount As Long) As Slide
    slidesContainNotesCount = 0
    
    Dim placeHolder_ As Shape
    Set placeHolder_ = TargetSlide.NotesPage.Shapes.Placeholders(2)
    placeHolder_.TextFrame.TextRange.Text = "Note " & TargetSlide.SlideIndex: slidesContainNotesCount = slidesContainNotesCount + 1
        
    Set DecolateSheetForPresentationNotesTest = TargetSlide
End Function

Private Function DecolatePresentationForPresentationNotesTest(ByVal TargetPresentation As Presentation, ByRef slidesContainNotesCount As Long) As Presentation
    slidesContainNotesCount = 0
    
    Dim slidesContainNotesCountInSlide As Long
    Dim slide_ As Slide
    Dim i As Long
    For i = 1 To TargetPresentation.Slides.Count
        Set slide_ = TargetPresentation.Slides(i)
        If i Mod 2 = 1 Then
            Call DecolateSheetForPresentationNotesTest(slide_, slidesContainNotesCountInSlide): slidesContainNotesCount = slidesContainNotesCount + slidesContainNotesCountInSlide
        End If
    Next
    
    Set DecolatePresentationForPresentationNotesTest = TargetPresentation
End Function

'@TestMethod("PresentationNotes")
Public Sub ContainsPresentationNoteInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim slidesContainNotesCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, slidesContainNotesCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim slidesContainNotesCountInSlide As Long
    Call DecolateSheetForPresentationNotesTest(testSlide, slidesContainNotesCountInSlide)
    
    Dim actual As Boolean
    On Error GoTo CATCH
    
    actual = instance.ContainsPresentationNoteInSlide(testSlide)
    
    Assert.IsTrue actual
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("PresentationNotes")
Public Sub ListSlidesContainPresentationNotes_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
        
    Dim slidesContainNotesCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, slidesContainNotesCount)
    
    On Error GoTo CATCH
    
    Dim listedSlidesContainNotes As Collection
    Set listedSlidesContainNotes = instance.ListSlidesContainPresentationNotes
    
    Assert.AreEqual slidesContainNotesCount, listedSlidesContainNotes.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("PresentationNotes")
Public Sub ClearPresentationNoteInSlide_CorrectCall_Successed()
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim slidesContainNotesCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, slidesContainNotesCount)
    
    Set testSlide = instance.Target.Slides.Add(1, ppLayoutBlank)
    Dim slidesContainNotesCountInSlide As Long
    Call DecolateSheetForPresentationNotesTest(testSlide, slidesContainNotesCountInSlide)
    
    Assert.IsTrue instance.ClearPresentationNoteInSlide(testSlide)
    
    GoTo FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("PresentationNotes")
Public Sub ClearPresentationNotes_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim slidesContainNotesCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, slidesContainNotesCount)
        
    On Error GoTo CATCH
    
    Dim clearedSlidesContainNotes As Collection
    Set clearedSlidesContainNotes = instance.ClearPresentationNotes
    
    Assert.AreEqual slidesContainNotesCount, clearedSlidesContainNotes.Count
    Assert.AreEqual 0&, instance.ListSlidesContainPresentationNotes.Count
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("PresentationNotes")
Public Sub InspectPresentationNotes_WithPresentationNotes_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusIssueFound, instance.InspectPresentationNotes
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("PresentationNotes")
Public Sub FixPresentationNotes_CorrectCall_Successed()
    Dim testPresentation As Presentation
    
    Set testPresentation = CreateNewPresentation(3)
    Dim instance As InspectPresentationUtils: Set instance = New InspectPresentationUtils: Call instance.Initialize(testPresentation)
    
    Dim customPartsCount As Long
    Call DecolatePresentationForPresentationNotesTest(instance.Target, customPartsCount)
    
    On Error GoTo CATCH
    
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.FixPresentationNotes
    Assert.AreEqual msoDocInspectorStatusDocOk, instance.InspectPresentationNotes
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    
    Resume FINALLY
FINALLY:
    Call testPresentation.Close
End Sub
