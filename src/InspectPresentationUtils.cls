VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InspectPresentationUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "_(Description-InspectPresentationUtils)"
'Copyright (c) 2022 telitas
'This file is released under the MIT License.
'See the LICENSE.txt file or https://opensource.org/licenses/mit-license.php for details.
'Package version: $PackageVersion
'@Folder "InspectPresentationUtilsProject"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'@ModuleDescription("_(Description-InspectPresentationUtils)")
Option Explicit

Private Const ModuleName As String = "InspectPresentationUtils"

Private Enum InspectorIndex
    CustomXMLData = 1
    InvisibleContent = 2
    OffSlideContent = 3
    PresentationNotes = 4
End Enum

'@VariableDescription "_(Description-Target_)"
Private Target_ As Presentation
Attribute Target_.VB_VarDescription = "_(Description-Target_)"

'@Description "_(Description-InvalidArgumentError)"
Public Property Get InvalidArgumentError() As Long
Attribute InvalidArgumentError.VB_Description = "_(Description-InvalidArgumentError)"
    InvalidArgumentError = vbObjectError + 5
End Property

'@Description "_(Description-InvalidOperationError)"
Public Property Get InvalidOperationError() As Long
Attribute InvalidOperationError.VB_Description = "_(Description-InvalidOperationError)"
    InvalidOperationError = vbObjectError + 425
End Property

'@Description "_(Description-Target)"
Public Property Get Target() As Presentation
Attribute Target.VB_Description = "_(Description-Target)"
    Set Target = Target_
End Property

'@Description "_(Description-Class_Initialize)"
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "_(Description-Class_Initialize)"
    Set Target_ = Nothing
End Sub

'@Description "_(Description-Initialize)"
Public Sub Initialize(ByVal Target As Presentation)
Attribute Initialize.VB_Description = "_(Description-Initialize)"
    Const ProcedureName As String = "Initialize"
    
    If Not Target_ Is Nothing Then
        Call RaiseError(Number:=InvalidOperationError, ProcedureName:=ProcedureName, Description:="This instance is Already initialized.")
    End If
    If Target Is Nothing Then
        Call RaiseErrorIfArgumentIsNothing(ProcedureName:=ProcedureName, TargetArgument:=Target, ArgumentName:="Target")
    End If
    
    Set Target_ = Target
End Sub

'@Description "_(Description-RaiseError)"
Private Sub RaiseError(ByVal Number As Long, ByVal ProcedureName As String, ByVal Description As String)
Attribute RaiseError.VB_Description = "_(Description-RaiseError)"
    Call Err.Raise(Number:=Number, Source:=Join(Array("Module=" & ModuleName, "Procedure=" & ProcedureName), ";"), Description:=Description)
End Sub

'@Description "_(Description-RaiseErrorIfNotInitialized)"
Private Sub RaiseErrorIfNotInitialized(ByVal ProcedureName As String)
Attribute RaiseErrorIfNotInitialized.VB_Description = "_(Description-RaiseErrorIfNotInitialized)"
    If Target Is Nothing Then
        Call RaiseError(Number:=InvalidOperationError, ProcedureName:=ProcedureName, Description:="This instance is not initialized.")
    End If
End Sub

'@Description "_(Description-RaiseErrorIfParentIsNotTargetPresentation)"
Private Sub RaiseErrorIfParentIsNotTargetPresentation(ByVal ProcedureName As String, ByVal TargetSlide As Slide)
Attribute RaiseErrorIfParentIsNotTargetPresentation.VB_Description = "_(Description-RaiseErrorIfParentIsNotTargetPresentation)"
    If Not TargetSlide.Parent Is Target Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="This slide is not a child of " & Target.Name & ".")
    End If
End Sub

'@Description "_(Description-RaiseErrorIfArgumentIsNothing)"
Private Sub RaiseErrorIfArgumentIsNothing(ByVal ProcedureName As String, ByVal TargetArgument As Object, ByVal ArgumentName As String)
Attribute RaiseErrorIfArgumentIsNothing.VB_Description = "_(Description-RaiseErrorIfArgumentIsNothing)"
    If TargetArgument Is Nothing Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="Argument """ & ArgumentName & """ must be not Nothing.")
    End If
End Sub

'@Description "_(Description-GetParentPresentation)"
Public Function GetParentPresentation(ByVal TargetObject As Object) As Presentation
Attribute GetParentPresentation.VB_Description = "_(Description-GetParentPresentation)"
    If TypeOf TargetObject Is Presentation Then
        Set GetParentPresentation = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    
    Set GetParentPresentation = GetParentPresentation(TargetObject.Parent)
End Function

'@Description "_(Description-GetParentSlide)"
Public Function GetParentSlide(ByVal TargetObject As Object) As Slide
Attribute GetParentSlide.VB_Description = "_(Description-GetParentSlide)"
    If TypeOf TargetObject Is Slide Then
        Set GetParentSlide = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Presentation Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    
    Set GetParentSlide = GetParentSlide(TargetObject.Parent)
End Function

'@Description "_(Description-GetPresentationLocation)"
Private Function GetPresentationLocation(ByVal TargetPresentation As Presentation) As String
Attribute GetPresentationLocation.VB_Description = "_(Description-GetPresentationLocation)"
    GetPresentationLocation = "'" & TargetPresentation.Path & "[" & TargetPresentation.Name & "]'"
End Function

'@Description "_(Description-GetSlideLocation)"
Private Function GetSlideLocation(ByVal TargetSlide As Slide) As String
Attribute GetSlideLocation.VB_Description = "_(Description-GetSlideLocation)"
    GetSlideLocation = GetPresentationLocation(GetParentPresentation(TargetSlide)) & "!Slide(" & TargetSlide.SlideIndex & ")"
End Function

'@Description "_(Description-GetShapeLocation)"
Private Function GetShapeLocation(ByVal TargetShape As Shape) As String
Attribute GetShapeLocation.VB_Description = "_(Description-GetShapeLocation)"
    GetShapeLocation = GetSlideLocation(GetParentSlide(TargetShape)) & "!Point(" & TargetShape.Left & ", " & TargetShape.Top & ")"
End Function

'' = For all Office documents =
'@Description "_(Description-ListShapesInSlide)"
Private Function ListShapesInSlide(ByVal TargetSlide As Slide, ByVal ShapeType As MsoShapeType) As Collection
Attribute ListShapesInSlide.VB_Description = "_(Description-ListShapesInSlide)"
    Dim listedShapes As Collection: Set listedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In TargetSlide.Shapes
        If shape_.Type = ShapeType Then
            Call listedShapes.Add(shape_)
        End If
    Next
    
    Set ListShapesInSlide = listedShapes
End Function

'@Description "_(Description-ListShapes)"
Private Function ListShapes(ByVal ShapeType As MsoShapeType) As Collection
Attribute ListShapes.VB_Description = "_(Description-ListShapes)"
    Dim listedShapes As Collection: Set listedShapes = New Collection
    
    Dim shape_ As Shape
    Dim slide_ As Slide
    For Each slide_ In Target.Slides
        For Each shape_ In ListShapesInSlide(slide_, ShapeType)
           Call listedShapes.Add(shape_)
        Next
    Next
    
    Set ListShapes = listedShapes
End Function

'@Description "_(Description-VisualizeShape)"
Private Function VisualizeShape(ByVal TargetShape As Shape) As Shape
Attribute VisualizeShape.VB_Description = "_(Description-VisualizeShape)"
    TargetShape.Visible = True
    
    Set VisualizeShape = TargetShape
End Function

'@Description "_(Description-DeleteShape)"
Private Function DeleteShape(ByVal TargetShape As Shape) As Object
Attribute DeleteShape.VB_Description = "_(Description-DeleteShape)"
    Dim deletedShapeInfo As Object: Set deletedShapeInfo = CreateObject("Scripting.Dictionary")
    
    With TargetShape
        deletedShapeInfo("Type") = TypeName(TargetShape)
        deletedShapeInfo("Name") = .Name
        deletedShapeInfo("Location") = GetShapeLocation(TargetShape)
        deletedShapeInfo("Shape.Type") = .Type
        Call .Delete
    End With
    
    Set DeleteShape = deletedShapeInfo
End Function

'@Description "_(Description-ConvertShapeToImage)"
Private Function ConvertShapeToImage(ByVal TargetShape As Shape) As Object
Attribute ConvertShapeToImage.VB_Description = "_(Description-ConvertShapeToImage)"
    Dim convertedImage As Object
    
    Set convertedImage = GetParentSlide(TargetShape).Shapes.PasteSpecial(Datatype:=ppPastePNG)
    With convertedImage
        .Left = TargetShape.Left
        .Top = TargetShape.Top
        .Name = TargetShape.Name
    End With
    Call TargetShape.Delete
    
    Set ConvertShapeToImage = convertedImage
End Function

'@Description "_(Description-ConvertShapesToImagesInSlide)"
Private Function ConvertShapesToImagesInSlide(ByVal TargetSlide As Slide, ByVal ShapeType As MsoShapeType) As Collection
Attribute ConvertShapesToImagesInSlide.VB_Description = "_(Description-ConvertShapesToImagesInSlide)"
    Const ProcedureName As String = "ConvertShapesToImagesInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim convertedShapes As Collection: Set convertedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListShapesInSlide(TargetSlide, ShapeType)
        Call convertedShapes.Add(ConvertShapeToImage(shape_))
    Next
    
    Set ConvertShapesToImagesInSlide = convertedShapes
End Function

'@Description "_(Description-ConvertShapesToImages)"
Private Function ConvertShapesToImages(ByVal ShapeType As MsoShapeType) As Collection
Attribute ConvertShapesToImages.VB_Description = "_(Description-ConvertShapesToImages)"
    Const ProcedureName As String = "ConvertShapesToImages"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim convertedShapes As Collection: Set convertedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListShapes(ShapeType)
        Call convertedShapes.Add(ConvertShapeToImage(shape_))
    Next
    
    Set ConvertShapesToImages = convertedShapes
End Function

'@Description "_(Description-DeleteComment)"
Private Function DeleteComment(ByVal TargetComment As Comment) As Object
Attribute DeleteComment.VB_Description = "_(Description-DeleteComment)"
    Dim deletedCommentInfo As Object: Set deletedCommentInfo = CreateObject("Scripting.Dictionary")
    
    With TargetComment
        deletedCommentInfo("Type") = TypeName(TargetComment)
        deletedCommentInfo("Location") = GetSlideLocation(GetParentSlide(TargetComment))
        Call .Delete
    End With
        
    Set DeleteComment = deletedCommentInfo
End Function

'@Description "_(Description-DeleteDocumentProperty)"
Private Function DeleteDocumentProperty(ByVal TargetDocumentProperty As DocumentProperty) As Object
Attribute DeleteDocumentProperty.VB_Description = "_(Description-DeleteDocumentProperty)"
    Const ProcedureName As String = "DeleteDocumentProperty"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedPropertyInfo As Object: Set deletedPropertyInfo = CreateObject("Scripting.Dictionary")
    
    With TargetDocumentProperty
        deletedPropertyInfo("Type") = TypeName(TargetDocumentProperty)
        deletedPropertyInfo("Name") = .Name
        deletedPropertyInfo("Location") = GetPresentationLocation(GetParentPresentation(TargetDocumentProperty))
        Call .Delete
    End With
    
    Set DeleteDocumentProperty = deletedPropertyInfo
End Function

'@Description "_(Description-IsDefaultDocumentPropertyValue)"
Private Function IsDefaultDocumentPropertyValue(ByVal TargetDocumentProperty As DocumentProperty) As Boolean
Attribute IsDefaultDocumentPropertyValue.VB_Description = "_(Description-IsDefaultDocumentPropertyValue)"
    Dim isDefaualt As Boolean
    
    Dim errorCatched As Boolean
    '@Ignore VariableNotUsed
    Dim dummy As Variant
    With TargetDocumentProperty
        Select Case .Type
            Case msoPropertyTypeNumber
                On Error GoTo ERROR_CATCHED:
                dummy = .value
                On Error GoTo 0
                isDefaualt = errorCatched
            Case msoPropertyTypeBoolean
                isDefaualt = .value = False
            Case msoPropertyTypeDate
                On Error GoTo ERROR_CATCHED:
                dummy = .value
                On Error GoTo 0
                isDefaualt = errorCatched
            Case msoPropertyTypeString
                isDefaualt = .value = vbNullString
            Case msoPropertyTypeFloat
                isDefaualt = .value = 0#
        End Select
    End With
    
    IsDefaultDocumentPropertyValue = isDefaualt
    Exit Function
    
ERROR_CATCHED:
    errorCatched = True
    Resume Next
End Function

'@Description "_(Description-ClearDocumentProperty)"
Private Function ClearDocumentProperty(ByVal TargetDocumentProperty As DocumentProperty) As DocumentProperty
Attribute ClearDocumentProperty.VB_Description = "_(Description-ClearDocumentProperty)"
    Const ProcedureName As String = "ClearDocumentProperty"
    
    '@Ignore VariableNotAssigned
    Dim numberDefault As Long
    '@Ignore VariableNotAssigned
    Dim booleanDefault As Boolean
    '@Ignore VariableNotAssigned
    Dim dateDefault As Date
    '@Ignore VariableNotAssigned
    Dim stringDefault As String
    '@Ignore VariableNotAssigned
    Dim floatDefault As Double
    Call RaiseErrorIfNotInitialized(ProcedureName)
    If Not IsDefaultDocumentPropertyValue(TargetDocumentProperty) Then
        With TargetDocumentProperty
            Select Case .Type
                Case msoPropertyTypeNumber
                    '@Ignore UnassignedVariableUsage
                    .value = numberDefault
                Case msoPropertyTypeBoolean
                    '@Ignore UnassignedVariableUsage
                    .value = booleanDefault
                Case msoPropertyTypeDate
                    '@Ignore UnassignedVariableUsage
                    .value = dateDefault
                Case msoPropertyTypeString
                    '@Ignore UnassignedVariableUsage
                    .value = stringDefault
                Case msoPropertyTypeFloat
                    '@Ignore UnassignedVariableUsage
                    .value = floatDefault
            End Select
        End With
    End If
    
    Set ClearDocumentProperty = TargetDocumentProperty
End Function

'@Description "_(Description-DeleteVBComponent)"
Private Function DeleteVBComponent(ByVal TargetVBComponent As Object) As Object
Attribute DeleteVBComponent.VB_Description = "_(Description-DeleteVBComponent)"
    Dim deletedMacroInfo As Object: Set deletedMacroInfo = CreateObject("Scripting.Dictionary")
    
    With TargetVBComponent
        deletedMacroInfo("Type") = TypeName(TargetVBComponent)
        deletedMacroInfo("Name") = .Name
        deletedMacroInfo("Location") = GetPresentationLocation(Target)
    End With
    
    Select Case TargetVBComponent.Type
        Case 1 ' vbext_ct_StdModule
            deletedMacroInfo("VBComponent.ComponentType") = "vbext_ct_StdModule"
            Call Target.VBProject.VBComponents.Remove(TargetVBComponent)
        Case 2 ' vbext_ct_ClassModule
            deletedMacroInfo("VBComponent.ComponentType") = "vbext_ct_ClassModule"
            Call Target.VBProject.VBComponents.Remove(TargetVBComponent)
        Case 3 ' vbext_ct_MSForm
            deletedMacroInfo("VBComponent.ComponentType") = "vbext_ct_MSForm"
            Call Target.VBProject.VBComponents.Remove(TargetVBComponent)
        Case 11 ' vbext_ct_ActiveXDesigner
            deletedMacroInfo("VBComponent.ComponentType") = "vbext_ct_ActiveXDesigner"
            Call TargetVBComponent.CodeModule.DeleteLines(1, TargetVBComponent.CodeModule.CountOfLines)
        Case 100 ' vbext_ct_Document
            deletedMacroInfo("VBComponent.ComponentType") = "vbext_ct_Document"
            Call TargetVBComponent.CodeModule.DeleteLines(1, TargetVBComponent.CodeModule.CountOfLines)
    End Select
    
    Set DeleteVBComponent = deletedMacroInfo
End Function

'@Description "_(Description-DeleteCustomXMLPart)"
Private Function DeleteCustomXMLPart(ByVal TargetCustomXMLPart As CustomXMLPart) As Object
Attribute DeleteCustomXMLPart.VB_Description = "_(Description-DeleteCustomXMLPart)"
    Dim deletedCustomXMLPartInfo As Object: Set deletedCustomXMLPartInfo = CreateObject("Scripting.Dictionary")
    
    With TargetCustomXMLPart
        deletedCustomXMLPartInfo("Type") = TypeName(TargetCustomXMLPart)
        deletedCustomXMLPartInfo("Id") = .Id
        deletedCustomXMLPartInfo("Location") = GetPresentationLocation(GetParentPresentation(TargetCustomXMLPart))
        Call .Delete
    End With
    
    Set DeleteCustomXMLPart = deletedCustomXMLPartInfo
End Function

'' == Comments ==
'@Description "_(Description-ListCommentsInSlide_)"
Private Function ListCommentsInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListCommentsInSlide_.VB_Description = "_(Description-ListCommentsInSlide_)"
    Dim listedComments As Collection: Set listedComments = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In TargetSlide.Comments
        Call listedComments.Add(comment_)
    Next
    
    Set ListCommentsInSlide_ = listedComments
End Function

'@Description "_(Description-ListCommentsInSlide)"
Public Function ListCommentsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListCommentsInSlide.VB_Description = "_(Description-ListCommentsInSlide)"
    Const ProcedureName As String = "ListCommentsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListCommentsInSlide = ListCommentsInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListComments_)"
Private Function ListComments_() As Collection
Attribute ListComments_.VB_Description = "_(Description-ListComments_)"
    Dim listedComments As Collection: Set listedComments = New Collection
    
    Dim slide_ As Slide
    Dim comment_ As Comment
    For Each slide_ In Target.Slides
        For Each comment_ In ListCommentsInSlide_(slide_)
            Call listedComments.Add(comment_)
        Next
    Next
    
    Set ListComments_ = listedComments
End Function

'@Description "_(Description-ListComments)"
Public Function ListComments() As Collection
Attribute ListComments.VB_Description = "_(Description-ListComments)"
    Const ProcedureName As String = "ListComments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListComments = ListComments_
End Function

'@Description "_(Description-DeleteCommentsInSlide)"
Public Function DeleteCommentsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteCommentsInSlide.VB_Description = "_(Description-DeleteCommentsInSlide)"
    Dim deletedCommentAddresses As Collection: Set deletedCommentAddresses = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In ListCommentsInSlide_(TargetSlide)
        Call deletedCommentAddresses.Add(DeleteComment(comment_))
    Next
    
    Set DeleteCommentsInSlide = deletedCommentAddresses
End Function

'@Description "_(Description-DeleteComments)"
Public Function DeleteComments() As Collection
Attribute DeleteComments.VB_Description = "_(Description-DeleteComments)"
    Dim deletedCommentAddresses As Collection: Set deletedCommentAddresses = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In ListComments_()
        Call deletedCommentAddresses.Add(DeleteComment(comment_))
    Next
    
    Set DeleteComments = deletedCommentAddresses
End Function

'@Description "_(Description-RemoveComments)"
Public Sub RemoveComments()
Attribute RemoveComments.VB_Description = "_(Description-RemoveComments)"
    Const ProcedureName As String = "RemoveComments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(ppRDIComments)
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@Description "_(Description-RemoveDocumentProperties)"
Public Sub RemoveDocumentProperties()
Attribute RemoveDocumentProperties.VB_Description = "_(Description-RemoveDocumentProperties)"
    Const ProcedureName As String = "RemoveDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(ppRDIDocumentProperties)
End Sub

'' ==== Built-in Document Properties ====
'@Description "_(Description-ListBuiltInDocumentProperties_)"
Private Function ListBuiltInDocumentProperties_() As Collection
Attribute ListBuiltInDocumentProperties_.VB_Description = "_(Description-ListBuiltInDocumentProperties_)"
    Dim listedBuiltInDocumentProperties As Collection: Set listedBuiltInDocumentProperties = New Collection
    
    Dim property As DocumentProperty
    For Each property In Target.builtInDocumentProperties
        Call listedBuiltInDocumentProperties.Add(property)
    Next
    
    Set ListBuiltInDocumentProperties_ = listedBuiltInDocumentProperties
End Function

'@Description "_(Description-ListBuiltInDocumentProperties)"
Public Function ListBuiltInDocumentProperties() As Collection
Attribute ListBuiltInDocumentProperties.VB_Description = "_(Description-ListBuiltInDocumentProperties)"
    Const ProcedureName As String = "ListBuiltInDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListBuiltInDocumentProperties = ListBuiltInDocumentProperties_
End Function

'@Description "_(Description-ClearBuiltInDocumentProperties)"
Public Function ClearBuiltInDocumentProperties() As Collection
Attribute ClearBuiltInDocumentProperties.VB_Description = "_(Description-ClearBuiltInDocumentProperties)"
    Const ProcedureName As String = "ClearBuiltInDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim clearedDocumentsPropertyNames As Collection: Set clearedDocumentsPropertyNames = New Collection
    
    Dim property As DocumentProperty
    For Each property In ListBuiltInDocumentProperties_
        Call clearedDocumentsPropertyNames.Add(ClearDocumentProperty(property))
    Next
    
    Set ClearBuiltInDocumentProperties = clearedDocumentsPropertyNames
End Function

'' ==== Custom Document Properties ====
'@Description "_(Description-ListCustomDocumentProperties_)"
Private Function ListCustomDocumentProperties_() As Collection
Attribute ListCustomDocumentProperties_.VB_Description = "_(Description-ListCustomDocumentProperties_)"
    Dim listedCustomDocumentProperties As Collection: Set listedCustomDocumentProperties = New Collection
    
    Dim property As DocumentProperty
    For Each property In Target.customDocumentProperties
        Call listedCustomDocumentProperties.Add(property)
    Next
    
    Set ListCustomDocumentProperties_ = listedCustomDocumentProperties
End Function

'@Description "_(Description-ListCustomDocumentProperties)"
Public Function ListCustomDocumentProperties() As Collection
Attribute ListCustomDocumentProperties.VB_Description = "_(Description-ListCustomDocumentProperties)"
    Const ProcedureName As String = "ListCustomDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListCustomDocumentProperties = ListCustomDocumentProperties_
End Function

'@Description "_(Description-ClearCustomDocumentProperties)"
Public Function ClearCustomDocumentProperties() As Collection
Attribute ClearCustomDocumentProperties.VB_Description = "_(Description-ClearCustomDocumentProperties)"
    Const ProcedureName As String = "ClearCustomDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim clearedDocumentsPropertyNames As Collection: Set clearedDocumentsPropertyNames = New Collection
    
    Dim property As DocumentProperty
    For Each property In ListCustomDocumentProperties_
        Call clearedDocumentsPropertyNames.Add(ClearDocumentProperty(property))
    Next
    
    Set ClearCustomDocumentProperties = clearedDocumentsPropertyNames
End Function

'@Description "_(Description-DeleteCustomDocumentProperties)"
Public Function DeleteCustomDocumentProperties() As Collection
Attribute DeleteCustomDocumentProperties.VB_Description = "_(Description-DeleteCustomDocumentProperties)"
    Const ProcedureName As String = "DeleteCustomDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedCustomPropertyInfos As Collection: Set deletedCustomPropertyInfos = New Collection
    
    Dim property As DocumentProperty
    For Each property In ListCustomDocumentProperties_
        Call deletedCustomPropertyInfos.Add(DeleteDocumentProperty(property))
    Next
    
    Set DeleteCustomDocumentProperties = deletedCustomPropertyInfos
End Function

'' === Personal Information ===
'@Description "_(Description-RemovePersonalInformation)"
Public Sub RemovePersonalInformation()
Attribute RemovePersonalInformation.VB_Description = "_(Description-RemovePersonalInformation)"
    Const ProcedureName As String = "RemovePersonalInformation"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(ppRDIRemovePersonalInformation)
End Sub

'' == Content Add-ins ==
'@Description "_(Description-ListContentAddInsInSlide_)"
Private Function ListContentAddInsInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListContentAddInsInSlide_.VB_Description = "_(Description-ListContentAddInsInSlide_)"
    Set ListContentAddInsInSlide_ = ListShapesInSlide(TargetSlide, msoContentApp)
End Function

'@Description "_(Description-ListContentAddInsInSlide)"
Public Function ListContentAddInsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListContentAddInsInSlide.VB_Description = "_(Description-ListContentAddInsInSlide)"
    Const ProcedureName As String = "ListContentAddInsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListContentAddInsInSlide = ListContentAddInsInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListContentAddIns_)"
Private Function ListContentAddIns_() As Collection
Attribute ListContentAddIns_.VB_Description = "_(Description-ListContentAddIns_)"
    Set ListContentAddIns_ = ListShapes(msoContentApp)
End Function

'@Description "_(Description-ListContentAddIns)"
Public Function ListContentAddIns() As Collection
Attribute ListContentAddIns.VB_Description = "_(Description-ListContentAddIns)"
    Const ProcedureName As String = "ListContentAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListContentAddIns = ListContentAddIns_
End Function

'@Description "_(Description-ConvertContentAddInsToImagesInSlide)"
Public Function ConvertContentAddInsToImagesInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ConvertContentAddInsToImagesInSlide.VB_Description = "_(Description-ConvertContentAddInsToImagesInSlide)"
    Const ProcedureName As String = "ConvertContentAddInsToImagesInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ConvertContentAddInsToImagesInSlide = ConvertShapesToImagesInSlide(TargetSlide, msoContentApp)
End Function

'@Description "_(Description-ConvertContentAddInsToImages)"
Public Function ConvertContentAddInsToImages() As Collection
Attribute ConvertContentAddInsToImages.VB_Description = "_(Description-ConvertContentAddInsToImages)"
    Const ProcedureName As String = "ConvertContentAddInsToImages"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ConvertContentAddInsToImages = ConvertShapesToImages(msoContentApp)
End Function

'@Description "_(Description-DeleteContentAddInsInSlide)"
Public Function DeleteContentAddInsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteContentAddInsInSlide.VB_Description = "_(Description-DeleteContentAddInsInSlide)"
    Const ProcedureName As String = "DeleteContentAddInsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedContentAppInfos As Collection: Set deletedContentAppInfos = New Collection
    
    Dim contentApp As Shape
    For Each contentApp In ListContentAddInsInSlide_(TargetSlide)
        Call deletedContentAppInfos.Add(DeleteShape(contentApp))
    Next
    
    Set DeleteContentAddInsInSlide = deletedContentAppInfos
End Function

'@Description "_(Description-DeleteContentAddIns)"
Public Function DeleteContentAddIns() As Collection
Attribute DeleteContentAddIns.VB_Description = "_(Description-DeleteContentAddIns)"
    Const ProcedureName As String = "DeleteContentAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedContentAppInfos As Collection: Set deletedContentAppInfos = New Collection
    
    Dim contentApp As Shape
    For Each contentApp In ListContentAddIns_
        Call deletedContentAppInfos.Add(DeleteShape(contentApp))
    Next
    
    Set DeleteContentAddIns = deletedContentAppInfos
End Function

'' == TaskPane Add-ins ==

'' == Embedded Documents ==
'@Description "_(Description-ListEmbeddedDocumentsInSlide_)"
Private Function ListEmbeddedDocumentsInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListEmbeddedDocumentsInSlide_.VB_Description = "_(Description-ListEmbeddedDocumentsInSlide_)"
    Set ListEmbeddedDocumentsInSlide_ = ListShapesInSlide(TargetSlide, msoEmbeddedOLEObject)
End Function

'@Description "_(Description-ListEmbeddedDocumentsInSlide)"
Public Function ListEmbeddedDocumentsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListEmbeddedDocumentsInSlide.VB_Description = "_(Description-ListEmbeddedDocumentsInSlide)"
    Const ProcedureName As String = "ListEmbeddedDocumentsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListEmbeddedDocumentsInSlide = ListEmbeddedDocumentsInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListEmbeddedDocuments_)"
Private Function ListEmbeddedDocuments_() As Collection
Attribute ListEmbeddedDocuments_.VB_Description = "_(Description-ListEmbeddedDocuments_)"
    Set ListEmbeddedDocuments_ = ListShapes(msoEmbeddedOLEObject)
End Function

'@Description "_(Description-ListEmbeddedDocuments)"
Public Function ListEmbeddedDocuments() As Collection
Attribute ListEmbeddedDocuments.VB_Description = "_(Description-ListEmbeddedDocuments)"
    Const ProcedureName As String = "ListEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListEmbeddedDocuments = ListEmbeddedDocuments_
End Function

'@Description "_(Description-DeleteEmbeddedDocumentsInSlide)"
Public Function DeleteEmbeddedDocumentsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteEmbeddedDocumentsInSlide.VB_Description = "_(Description-DeleteEmbeddedDocumentsInSlide)"
    Const ProcedureName As String = "DeleteEmbeddedDocumentsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedEmbeddedDocumentInfos As Collection: Set deletedEmbeddedDocumentInfos = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListEmbeddedDocumentsInSlide_(TargetSlide)
        Call deletedEmbeddedDocumentInfos.Add(DeleteShape(shape_))
    Next
    
    Set DeleteEmbeddedDocumentsInSlide = deletedEmbeddedDocumentInfos
End Function

'@Description "_(Description-DeleteEmbeddedDocuments)"
Public Function DeleteEmbeddedDocuments() As Collection
Attribute DeleteEmbeddedDocuments.VB_Description = "_(Description-DeleteEmbeddedDocuments)"
    Const ProcedureName As String = "DeleteEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedEmbeddedDocumentInfos As Collection: Set deletedEmbeddedDocumentInfos = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListEmbeddedDocuments_
        Call deletedEmbeddedDocumentInfos.Add(DeleteShape(shape_))
    Next
    
    Set DeleteEmbeddedDocuments = deletedEmbeddedDocumentInfos
End Function

'' == Macros, Forms, And ActiveX Controls ==
'' === Macros ===
'@Description "_(Description-ListMacros_)"
Private Function ListMacros_() As Collection
Attribute ListMacros_.VB_Description = "_(Description-ListMacros_)"
    Dim listedMacros As Collection: Set listedMacros = New Collection
    
    Dim component As Object
    For Each component In Target.VBProject.VBComponents
        With component
            Select Case .Type
                Case 1, 2, 3 'vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                    Call listedMacros.Add(component)
                Case Else
                    If .CodeModule.CountOfLines > 0 Then
                        Call listedMacros.Add(component)
                    End If
            End Select
        End With
    Next
    
    Set ListMacros_ = listedMacros
End Function

'@Description "_(Description-ListMacros)"
Public Function ListMacros() As Collection
Attribute ListMacros.VB_Description = "_(Description-ListMacros)"
    Const ProcedureName As String = "ListMacros"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListMacros = ListMacros_
End Function

'@Description "_(Description-DeleteMacros)"
Public Function DeleteMacros() As Collection
Attribute DeleteMacros.VB_Description = "_(Description-DeleteMacros)"
    Const ProcedureName As String = "DeleteMacros"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedMacroInfos As Collection: Set deletedMacroInfos = New Collection
    
    Dim component As Object
    For Each component In ListMacros_
        Call deletedMacroInfos.Add(DeleteVBComponent(component))
    Next
    
    Set DeleteMacros = deletedMacroInfos
End Function

'' === ActiveX Controls ===
'@Description "_(Description-ListActiveXControlsInSlide_)"
Private Function ListActiveXControlsInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListActiveXControlsInSlide_.VB_Description = "_(Description-ListActiveXControlsInSlide_)"
    Set ListActiveXControlsInSlide_ = ListShapesInSlide(TargetSlide, msoOLEControlObject)
End Function

'@Description "_(Description-ListActiveXControlsInSlide)"
Public Function ListActiveXControlsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListActiveXControlsInSlide.VB_Description = "_(Description-ListActiveXControlsInSlide)"
    Const ProcedureName As String = "ListActiveXControlsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListActiveXControlsInSlide = ListActiveXControlsInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListActiveXControls_)"
Private Function ListActiveXControls_() As Collection
Attribute ListActiveXControls_.VB_Description = "_(Description-ListActiveXControls_)"
    Set ListActiveXControls_ = ListShapes(msoOLEControlObject)
End Function

'@Description "_(Description-ListActiveXControls)"
Public Function ListActiveXControls() As Collection
Attribute ListActiveXControls.VB_Description = "_(Description-ListActiveXControls)"
    Const ProcedureName As String = "ListActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListActiveXControls = ListActiveXControls_
End Function

'@Description "_(Description-DeleteActiveXControlsInSlide)"
Public Function DeleteActiveXControlsInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteActiveXControlsInSlide.VB_Description = "_(Description-DeleteActiveXControlsInSlide)"
    Const ProcedureName As String = "DeleteActiveXControlsInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedActiveXControlInfos As Collection: Set deletedActiveXControlInfos = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListActiveXControlsInSlide_(TargetSlide)
        Call deletedActiveXControlInfos.Add(DeleteShape(shape_))
    Next
    
    Set DeleteActiveXControlsInSlide = deletedActiveXControlInfos
End Function

'@Description "_(Description-DeleteActiveXControls)"
Public Function DeleteActiveXControls() As Collection
Attribute DeleteActiveXControls.VB_Description = "_(Description-DeleteActiveXControls)"
    Const ProcedureName As String = "DeleteActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedActiveXControlInfos As Collection: Set deletedActiveXControlInfos = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListActiveXControls_
        Call deletedActiveXControlInfos.Add(DeleteShape(shape_))
    Next
    
    Set DeleteActiveXControls = deletedActiveXControlInfos
End Function

'' === Ink ===
'@Description "_(Description-ListInkInSlide_)"
Private Function ListInkInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListInkInSlide_.VB_Description = "_(Description-ListInkInSlide_)"
    Set ListInkInSlide_ = ListShapesInSlide(TargetSlide, msoInkComment)
End Function

'@Description "_(Description-ListInkInSlide)"
Public Function ListInkInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListInkInSlide.VB_Description = "_(Description-ListInkInSlide)"
    Const ProcedureName As String = "ListInkInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListInkInSlide = ListInkInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListInk_)"
Private Function ListInk_() As Collection
Attribute ListInk_.VB_Description = "_(Description-ListInk_)"
    Set ListInk_ = ListShapes(msoInkComment)
End Function

'@Description "_(Description-ListInk)"
Public Function ListInk() As Collection
Attribute ListInk.VB_Description = "_(Description-ListInk)"
    Const ProcedureName As String = "ListInk"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListInk = ListInk_
End Function

'@Description "_(Description-DeleteInkInSlide)"
Public Function DeleteInkInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteInkInSlide.VB_Description = "_(Description-DeleteInkInSlide)"
    Const ProcedureName As String = "DeleteInkInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedInkInfos As Collection: Set deletedInkInfos = New Collection
    
    Dim ink_ As Shape
    For Each ink_ In ListInkInSlide_(TargetSlide)
        Call deletedInkInfos.Add(DeleteShape(ink_))
    Next
    
    Set DeleteInkInSlide = deletedInkInfos
End Function

'@Description "_(Description-DeleteInk)"
Public Function DeleteInk() As Collection
Attribute DeleteInk.VB_Description = "_(Description-DeleteInk)"
    Const ProcedureName As String = "DeleteInk"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedInkInfos As Collection: Set deletedInkInfos = New Collection
    
    Dim ink_ As Shape
    For Each ink_ In ListInk_
        Call deletedInkInfos.Add(DeleteShape(ink_))
    Next
    
    Set DeleteInk = deletedInkInfos
End Function

'@Description "_(Description-RemoveInk)"
Public Sub RemoveInk()
Attribute RemoveInk.VB_Description = "_(Description-RemoveInk)"
    Const ProcedureName As String = "RemoveInk"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(ppRDIInkAnnotations)
End Sub

'' == Custom XML Data ==
'@Description "_(Description-ListCustomXMLData_)"
Private Function ListCustomXMLData_() As Collection
Attribute ListCustomXMLData_.VB_Description = "_(Description-ListCustomXMLData_)"
    Dim listedCustomXMLData As Collection: Set listedCustomXMLData = New Collection
    
    Dim xmlPart As CustomXMLPart
    For Each xmlPart In Target.CustomXMLParts
        If Not xmlPart.BuiltIn Then
            Call listedCustomXMLData.Add(xmlPart)
        End If
    Next
    
    Set ListCustomXMLData_ = listedCustomXMLData
End Function

'@Description "_(Description-ListCustomXMLData)"
Public Function ListCustomXMLData() As Collection
Attribute ListCustomXMLData.VB_Description = "_(Description-ListCustomXMLData)"
    Const ProcedureName As String = "ListCustomXMLData"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListCustomXMLData = ListCustomXMLData_
End Function

'@Description "_(Description-DeleteCustomXMLData)"
Public Function DeleteCustomXMLData() As Collection
Attribute DeleteCustomXMLData.VB_Description = "_(Description-DeleteCustomXMLData)"
    Const ProcedureName As String = "DeleteCustomXMLData"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedCustomXMLData As Collection: Set deletedCustomXMLData = New Collection
    
    Dim xmlPart As CustomXMLPart
    For Each xmlPart In ListCustomXMLData_
        Call deletedCustomXMLData.Add(DeleteCustomXMLPart(xmlPart))
    Next
    
    Set DeleteCustomXMLData = deletedCustomXMLData
End Function

'@Description "_(Description-InspectCustomXMLData)"
Public Function InspectCustomXMLData(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectCustomXMLData.VB_Description = "_(Description-InspectCustomXMLData)"
    Const ProcedureName As String = "InspectCustomXMLData"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.CustomXMLData)
    Call inspector.Inspect(status, Results)
    
    InspectCustomXMLData = status
End Function

'@Description "_(Description-FixCustomXMLData)"
Public Function FixCustomXMLData(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixCustomXMLData.VB_Description = "_(Description-FixCustomXMLData)"
    Const ProcedureName As String = "FixCustomXMLData"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.CustomXMLData)
    Call inspector.Fix(status, Results)
    
    FixCustomXMLData = status
End Function

'' == Invisible On-Slide Content ==
'@Description "_(Description-ListInvisibleContentInSlide_)"
Private Function ListInvisibleContentInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListInvisibleContentInSlide_.VB_Description = "_(Description-ListInvisibleContentInSlide_)"
    Dim listedInvisibleContent As Collection: Set listedInvisibleContent = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In TargetSlide.Shapes
        If shape_.Visible = msoFalse Then
            Call listedInvisibleContent.Add(shape_)
        End If
    Next
    
    Set ListInvisibleContentInSlide_ = listedInvisibleContent
End Function

'@Description "_(Description-ListInvisibleContentInSlide)"
Public Function ListInvisibleContentInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListInvisibleContentInSlide.VB_Description = "_(Description-ListInvisibleContentInSlide)"
    Const ProcedureName As String = "ListInvisibleContentInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListInvisibleContentInSlide = ListInvisibleContentInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListInvisibleContent_)"
Private Function ListInvisibleContent_() As Collection
Attribute ListInvisibleContent_.VB_Description = "_(Description-ListInvisibleContent_)"
    Const ProcedureName As String = "ListInvisibleContent_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedInvisibleContent As Collection: Set listedInvisibleContent = New Collection
    
    Dim slide_ As Slide
    Dim shape_ As Shape
    For Each slide_ In Target.Slides
        For Each shape_ In ListInvisibleContentInSlide_(slide_)
            Call listedInvisibleContent.Add(shape_)
        Next
    Next
    
    Set ListInvisibleContent_ = listedInvisibleContent
End Function

'@Description "_(Description-ListInvisibleContent)"
Public Function ListInvisibleContent() As Collection
Attribute ListInvisibleContent.VB_Description = "_(Description-ListInvisibleContent)"
    Const ProcedureName As String = "ListInvisibleContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListInvisibleContent = ListInvisibleContent_()
End Function

'@Description "_(Description-VisualizeInvisibleContentInSlide)"
Public Function VisualizeInvisibleContentInSlide(ByVal TargetSlide As Slide) As Collection
Attribute VisualizeInvisibleContentInSlide.VB_Description = "_(Description-VisualizeInvisibleContentInSlide)"
    Const ProcedureName As String = "VisualizeInvisibleContentInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim visualizedInvisibleContent As Collection: Set visualizedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContentInSlide_(TargetSlide)
        Call visualizedInvisibleContent.Add(VisualizeShape(content_))
    Next
    
    Set VisualizeInvisibleContentInSlide = visualizedInvisibleContent
End Function

'@Description "_(Description-VisualizeInvisibleContent)"
Public Function VisualizeInvisibleContent() As Collection
Attribute VisualizeInvisibleContent.VB_Description = "_(Description-VisualizeInvisibleContent)"
    Const ProcedureName As String = "VisualizeInvisibleContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedInvisibleContent As Collection: Set visualizedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContent_
        Call visualizedInvisibleContent.Add(VisualizeShape(content_))
    Next
    
    Set VisualizeInvisibleContent = visualizedInvisibleContent
End Function

'@Description "_(Description-DeleteInvisibleContentInSlide)"
Public Function DeleteInvisibleContentInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteInvisibleContentInSlide.VB_Description = "_(Description-DeleteInvisibleContentInSlide)"
    Const ProcedureName As String = "DeleteInvisibleContentInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedInvisibleContent As Collection: Set deletedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContentInSlide_(TargetSlide)
        Call deletedInvisibleContent.Add(DeleteShape(content_))
    Next
    
    Set DeleteInvisibleContentInSlide = deletedInvisibleContent
End Function

'@Description "_(Description-DeleteInvisibleContent)"
Public Function DeleteInvisibleContent() As Collection
Attribute DeleteInvisibleContent.VB_Description = "_(Description-DeleteInvisibleContent)"
    Const ProcedureName As String = "DeleteInvisibleContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedInvisibleContent As Collection: Set deletedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContent_
        Call deletedInvisibleContent.Add(DeleteShape(content_))
    Next
    
    Set DeleteInvisibleContent = deletedInvisibleContent
End Function

'@Description "_(Description-InspectInvisibleContent)"
Public Function InspectInvisibleContent(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectInvisibleContent.VB_Description = "_(Description-InspectInvisibleContent)"
    Const ProcedureName As String = "InspectInvisibleContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.InvisibleContent)
    Call inspector.Inspect(status, Results)
    
    InspectInvisibleContent = status
End Function

'@Description "_(Description-FixInvisibleContent)"
Public Function FixInvisibleContent(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixInvisibleContent.VB_Description = "_(Description-FixInvisibleContent)"
    Const ProcedureName As String = "FixInvisibleContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.InvisibleContent)
    Call inspector.Fix(status, Results)
    
    FixInvisibleContent = status
End Function

'' = For PowerPoint documents =

'' == Revision Tracking Data ==

'' == Off-slide Content ==
'@Description "_(Description-ListOffSlideContentInSlide_)"
Private Function ListOffSlideContentInSlide_(ByVal TargetSlide As Slide) As Collection
Attribute ListOffSlideContentInSlide_.VB_Description = "_(Description-ListOffSlideContentInSlide_)"
    Dim listedOffSlideContent As Collection: Set listedOffSlideContent = New Collection
    
    Dim pageSetup_ As PageSetup
    Set pageSetup_ = GetParentPresentation(TargetSlide).PageSetup
    Dim shape_ As Shape
    For Each shape_ In TargetSlide.Shapes
        If shape_.Left + shape_.Width < 0 Or _
            shape_.Left > pageSetup_.SlideWidth Or _
            shape_.Top + shape_.Height < 0 Or _
            shape_.Top > pageSetup_.SlideHeight Then
            Call listedOffSlideContent.Add(shape_)
        End If
    Next
    
    Set ListOffSlideContentInSlide_ = listedOffSlideContent
End Function

'@Description "_(Description-ListOffSlideContentInSlide)"
Public Function ListOffSlideContentInSlide(ByVal TargetSlide As Slide) As Collection
Attribute ListOffSlideContentInSlide.VB_Description = "_(Description-ListOffSlideContentInSlide)"
    Const ProcedureName As String = "ListOffSlideContentInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Set ListOffSlideContentInSlide = ListOffSlideContentInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListOffSlideContent_)"
Private Function ListOffSlideContent_() As Collection
Attribute ListOffSlideContent_.VB_Description = "_(Description-ListOffSlideContent_)"
    Const ProcedureName As String = "ListOffSlideContent_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedOffSlideContent As Collection: Set listedOffSlideContent = New Collection
    
    Dim slide_ As Slide
    Dim shape_ As Shape
    For Each slide_ In Target.Slides
        For Each shape_ In ListOffSlideContentInSlide_(slide_)
            Call listedOffSlideContent.Add(shape_)
        Next
    Next
    
    Set ListOffSlideContent_ = listedOffSlideContent
End Function

'@Description "_(Description-ListOffSlideContent)"
Public Function ListOffSlideContent() As Collection
Attribute ListOffSlideContent.VB_Description = "_(Description-ListOffSlideContent)"
    Const ProcedureName As String = "ListOffSlideContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListOffSlideContent = ListOffSlideContent_()
End Function

'@Description "_(Description-DeleteOffSlideContentInSlide)"
Public Function DeleteOffSlideContentInSlide(ByVal TargetSlide As Slide) As Collection
Attribute DeleteOffSlideContentInSlide.VB_Description = "_(Description-DeleteOffSlideContentInSlide)"
    Const ProcedureName As String = "DeleteOffSlideContentInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    Dim deletedOffSlideContent As Collection: Set deletedOffSlideContent = New Collection
    
    Dim content As Shape
    For Each content In ListOffSlideContentInSlide_(TargetSlide)
        Call deletedOffSlideContent.Add(DeleteShape(content))
    Next
    
    Set DeleteOffSlideContentInSlide = deletedOffSlideContent
End Function

'@Description "_(Description-DeleteOffSlideContent)"
Public Function DeleteOffSlideContent() As Collection
Attribute DeleteOffSlideContent.VB_Description = "_(Description-DeleteOffSlideContent)"
    Const ProcedureName As String = "DeleteOffSlideContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedOffSlideContent As Collection: Set deletedOffSlideContent = New Collection
    
    Dim content As Shape
    For Each content In ListOffSlideContent
        Call deletedOffSlideContent.Add(DeleteShape(content))
    Next
    
    Set DeleteOffSlideContent = deletedOffSlideContent
End Function

'@Description "_(Description-InspectOffSlideContent)"
Public Function InspectOffSlideContent(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectOffSlideContent.VB_Description = "_(Description-InspectOffSlideContent)"
    Const ProcedureName As String = "InspectOffSlideContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.OffSlideContent)
    Call inspector.Inspect(status, Results)
    
    InspectOffSlideContent = status
End Function

'@Description "_(Description-FixOffSlideContent)"
Public Function FixOffSlideContent(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixOffSlideContent.VB_Description = "_(Description-FixOffSlideContent)"
    Const ProcedureName As String = "FixOffSlideContent"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.OffSlideContent)
    Call inspector.Fix(status, Results)
    
    FixOffSlideContent = status
End Function

'' == Presentation Notes ==
'@Description "_(Description-ContainsPresentationNoteInSlide_)"
Private Function ContainsPresentationNoteInSlide_(ByVal TargetSlide As Slide) As Boolean
Attribute ContainsPresentationNoteInSlide_.VB_Description = "_(Description-ContainsPresentationNoteInSlide_)"
    If TargetSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text <> "" Then
        ContainsPresentationNoteInSlide_ = True
        Exit Function
    End If
    
    ContainsPresentationNoteInSlide_ = False
End Function

'@Description "_(Description-ContainsPresentationNoteInSlide)"
Public Function ContainsPresentationNoteInSlide(ByVal TargetSlide As Slide) As Boolean
Attribute ContainsPresentationNoteInSlide.VB_Description = "_(Description-ContainsPresentationNoteInSlide)"
    Const ProcedureName As String = "SlideContainsPresentationNote"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetSlide, "TargetSlide")
    Call RaiseErrorIfParentIsNotTargetPresentation(ProcedureName, TargetSlide)
    
    ContainsPresentationNoteInSlide = ContainsPresentationNoteInSlide_(TargetSlide)
End Function

'@Description "_(Description-ListSlidesContainPresentationNotes_)"
Private Function ListSlidesContainPresentationNotes_() As Collection
Attribute ListSlidesContainPresentationNotes_.VB_Description = "_(Description-ListSlidesContainPresentationNotes_)"
    Dim listedSlides As Collection: Set listedSlides = New Collection
    
    Dim slide_ As Slide
    For Each slide_ In Target.Slides
        If ContainsPresentationNoteInSlide_(slide_) Then
            Call listedSlides.Add(slide_)
        End If
    Next
    
    Set ListSlidesContainPresentationNotes_ = listedSlides
End Function

'@Description "_(Description-ListSlidesContainPresentationNotes)"
Public Function ListSlidesContainPresentationNotes() As Collection
Attribute ListSlidesContainPresentationNotes.VB_Description = "_(Description-ListSlidesContainPresentationNotes)"
    Const ProcedureName As String = "ListSlidesContainPresentationNotes"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListSlidesContainPresentationNotes = ListSlidesContainPresentationNotes_
End Function

'@Description "_(Description-ClearPresentationNoteInSlide_)"
Private Function ClearPresentationNoteInSlide_(ByVal TargetSlide As Slide) As Boolean
Attribute ClearPresentationNoteInSlide_.VB_Description = "_(Description-ClearPresentationNoteInSlide_)"
    If ContainsPresentationNoteInSlide_(TargetSlide) Then
        TargetSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ""
        ClearPresentationNoteInSlide_ = True
        Exit Function
    End If
End Function

'@Description "_(Description-ClearPresentationNoteInSlide)"
Public Function ClearPresentationNoteInSlide(ByVal TargetSlide As Slide) As Boolean
Attribute ClearPresentationNoteInSlide.VB_Description = "_(Description-ClearPresentationNoteInSlide)"
    Const ProcedureName As String = "ClearPresentationNoteInSlide"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    ClearPresentationNoteInSlide = ClearPresentationNoteInSlide_(TargetSlide)
End Function

'@Description "_(Description-ClearPresentationNotes)"
Public Function ClearPresentationNotes() As Collection
Attribute ClearPresentationNotes.VB_Description = "_(Description-ClearPresentationNotes)"
    Const ProcedureName As String = "ClearPresentationNotes"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim clearedSlidesContainPresentationNotes As Collection: Set clearedSlidesContainPresentationNotes = New Collection
    
    Dim slide_ As Slide
    For Each slide_ In ListSlidesContainPresentationNotes_
        If ClearPresentationNoteInSlide_(slide_) Then
            Call clearedSlidesContainPresentationNotes.Add(slide_)
        End If
    Next
    
    Set ClearPresentationNotes = clearedSlidesContainPresentationNotes
End Function

'@Description "_(Description-InspectPresentationNotes)"
Public Function InspectPresentationNotes(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectPresentationNotes.VB_Description = "_(Description-InspectPresentationNotes)"
    Const ProcedureName As String = "InspectPresentationNotes"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.PresentationNotes)
    Call inspector.Inspect(status, Results)
    
    InspectPresentationNotes = status
End Function

'@Description "_(Description-FixPresentationNotes)"
Public Function FixPresentationNotes(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixPresentationNotes.VB_Description = "_(Description-FixPresentationNotes)"
    Const ProcedureName As String = "FixPresentationNotes"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.PresentationNotes)
    Call inspector.Fix(status, Results)
    
    FixPresentationNotes = status
End Function
