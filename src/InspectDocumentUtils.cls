VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InspectDocumentUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "_(Description-InspectDocumentUtils)"
'Copyright (c) 2022 telitas
'This file is released under the MIT License.
'See the LICENSE.txt file or https://opensource.org/licenses/mit-license.php for details.
'Package version: $PackageVersion
'@Folder "InspectDocumentUtilsProject"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'@ModuleDescription("_(Description-InspectDocumentUtils)")
Option Explicit

Private Const ModuleName As String = "InspectDocumentUtils"

Private Enum InspectorIndex
    CollapsedHeadings = 1
    CustomXMLData = 2
    HeadersAndFooters = 3
    InvisibleContent = 4
    HiddenText = 5
End Enum

'@VariableDescription "_(Description-Target_)"
Private Target_ As Document
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
Public Property Get Target() As Document
Attribute Target.VB_Description = "_(Description-Target)"
    Set Target = Target_
End Property

'@Description "_(Description-Class_Initialize)"
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "_(Description-Class_Initialize)"
    Set Target_ = Nothing
End Sub

'@Description "_(Description-Initialize)"
Public Sub Initialize(ByVal Target As Document)
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

'@Description "_(Description-RaiseErrorIfParentIsNotTargetDocument)"
Private Sub RaiseErrorIfParentIsNotTargetDocument(ByVal ProcedureName As String, ByVal TargetSection As Section)
Attribute RaiseErrorIfParentIsNotTargetDocument.VB_Description = "_(Description-RaiseErrorIfParentIsNotTargetDocument)"
    If Not TargetSection.Parent Is Target Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="This section is not a child of " & Target.Name & ".")
    End If
End Sub

'@Description "_(Description-RaiseErrorIfArgumentIsNothing)"
Private Sub RaiseErrorIfArgumentIsNothing(ByVal ProcedureName As String, ByVal TargetArgument As Object, ByVal ArgumentName As String)
Attribute RaiseErrorIfArgumentIsNothing.VB_Description = "_(Description-RaiseErrorIfArgumentIsNothing)"
    If TargetArgument Is Nothing Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="Argument """ & ArgumentName & """ must be not Nothing.")
    End If
End Sub

'@Description "_(Description-GetParentDocument)"
Public Function GetParentDocument(ByVal TargetObject As Object) As Document
Attribute GetParentDocument.VB_Description = "_(Description-GetParentDocument)"
    If TypeOf TargetObject Is Document Then
        Set GetParentDocument = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    
    Set GetParentDocument = GetParentDocument(TargetObject.Parent)
End Function

'@Description "_(Description-GetDocumentLocation)"
Private Function GetDocumentLocation(ByVal TargetDocument As Document) As String
Attribute GetDocumentLocation.VB_Description = "_(Description-GetDocumentLocation)"
    GetDocumentLocation = "'" & TargetDocument.Path & "[" & TargetDocument.Name & "]'"
End Function

'@Description "_(Description-GetRangeLocation)"
Private Function GetRangeLocation(ByVal TargetRange As Range) As String
Attribute GetRangeLocation.VB_Description = "_(Description-GetRangeLocation)"
    GetRangeLocation = GetDocumentLocation(GetParentDocument(TargetRange)) & "!Range(" & TargetRange.Start & "," & TargetRange.End & ")"
End Function

'@Description "_(Description-GetShapeLocation)"
Private Function GetShapeLocation(ByVal TargetShape As Shape) As String
Attribute GetShapeLocation.VB_Description = "_(Description-GetShapeLocation)"
    GetShapeLocation = GetDocumentLocation(GetParentDocument(TargetShape)) & "!Location(" & TargetShape.Left & "," & TargetShape.Top & ")"
End Function

'' = For all Office documents =
'@Description "_(Description-ListShapes)"
Private Function ListShapes(ByVal ShapeType As MsoShapeType) As Collection
Attribute ListShapes.VB_Description = "_(Description-ListShapes)"
    Dim listedShapes As Collection: Set listedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In Target.Shapes
        If shape_.Type = ShapeType Then
            Call listedShapes.Add(shape_)
        End If
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

'@Description "_(Description-DeleteFormField)"
Private Function DeleteFormField(ByVal TargetFormField As FormField) As Object
Attribute DeleteFormField.VB_Description = "_(Description-DeleteFormField)"
    Dim deletedFormFieldInfo As Object: Set deletedFormFieldInfo = CreateObject("Scripting.Dictionary")
    
    With TargetFormField
        deletedFormFieldInfo("Type") = TypeName(TargetFormField)
        deletedFormFieldInfo("Name") = .Name
        deletedFormFieldInfo("Location") = GetRangeLocation(.Range)
        deletedFormFieldInfo("FormField.Type") = .Type
        Call .Delete
    End With
    
    Set DeleteFormField = deletedFormFieldInfo
End Function

'@Description "_(Description-DeleteRange)"
Private Function DeleteRange(ByVal TargetRange As Range) As Object
Attribute DeleteRange.VB_Description = "_(Description-DeleteRange)"
    Dim deletedRangeInfo As Object: Set deletedRangeInfo = CreateObject("Scripting.Dictionary")
    
    With TargetRange
        deletedRangeInfo("Type") = TypeName(TargetRange)
        deletedRangeInfo("Location") = GetRangeLocation(TargetRange)
        Call .Delete
    End With
    
    Set DeleteRange = deletedRangeInfo
End Function

'@Description "_(Description-ListInlineShapes)"
Private Function ListInlineShapes(ByVal InlineShapeType As WdInlineShapeType) As Collection
Attribute ListInlineShapes.VB_Description = "_(Description-ListInlineShapes)"
    Dim listedInlineShapes As Collection: Set listedInlineShapes = New Collection
    
    Dim inlineShape_ As InlineShape
    For Each inlineShape_ In Target.InlineShapes
        If inlineShape_.Type = InlineShapeType Then
            Call listedInlineShapes.Add(inlineShape_)
        End If
    Next
    
    Set ListInlineShapes = listedInlineShapes
End Function

'@Description "_(Description-DeleteInlineShape)"
Private Function DeleteInlineShape(ByVal TargetInlineShape As InlineShape) As Object
Attribute DeleteInlineShape.VB_Description = "_(Description-DeleteInlineShape)"
    Dim deletedInlineShapeInfo As Object: Set deletedInlineShapeInfo = CreateObject("Scripting.Dictionary")
    
    With TargetInlineShape
        deletedInlineShapeInfo("Type") = TypeName(TargetInlineShape)
        deletedInlineShapeInfo("Location") = GetRangeLocation(.Range)
        deletedInlineShapeInfo("InlineShape.Type") = .Type
        Call .Delete
    End With
    
    Set DeleteInlineShape = deletedInlineShapeInfo
End Function

'@Description "_(Description-DeleteComment)"
Private Function DeleteComment(ByVal TargetComment As Comment) As Object
Attribute DeleteComment.VB_Description = "_(Description-DeleteComment)"
    Dim deletedCommentInfo As Object: Set deletedCommentInfo = CreateObject("Scripting.Dictionary")
    
    With TargetComment
        deletedCommentInfo("Type") = TypeName(TargetComment)
        deletedCommentInfo("Location") = GetRangeLocation(.Range)
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
        deletedPropertyInfo("Location") = GetDocumentLocation(GetParentDocument(TargetDocumentProperty))
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
        deletedMacroInfo("Location") = GetDocumentLocation(Target)
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
        deletedCustomXMLPartInfo("Location") = GetDocumentLocation(GetParentDocument(TargetCustomXMLPart))
        Call .Delete
    End With
    
    Set DeleteCustomXMLPart = deletedCustomXMLPartInfo
End Function

'' == Comments ==
'@Description "_(Description-ListComments_)"
Private Function ListComments_() As Collection
Attribute ListComments_.VB_Description = "_(Description-ListComments_)"
    Dim listedComments As Collection: Set listedComments = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In Target.Comments
        Call listedComments.Add(comment_)
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

'@Description "_(Description-DeleteComments)"
Public Function DeleteComments() As Collection
Attribute DeleteComments.VB_Description = "_(Description-DeleteComments)"
    Dim deletedCommentAddresses As Collection: Set deletedCommentAddresses = New Collection
    
    ' For some reason, this block does not work as expected.
    'Dim comment_ As Comment
    'For Each comment_ In ListComments_
    '    Call deletedCommentAddresses.Add(DeleteComment(comment_))
    'Next
    ' Alternative
    Dim listedComments As Collection
    Set listedComments = ListComments_
    Dim i As Long
    For i = listedComments.Count To 1 Step -1
        Call deletedCommentAddresses.Add(DeleteComment(listedComments(i)))
    Next
    ' End Alternative
    
    Set DeleteComments = deletedCommentAddresses
End Function

'@Description "_(Description-RemoveComments)"
Public Sub RemoveComments()
Attribute RemoveComments.VB_Description = "_(Description-RemoveComments)"
    Const ProcedureName As String = "RemoveComments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(wdRDIComments)
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@Description "_(Description-RemoveDocumentProperties)"
Public Sub RemoveDocumentProperties()
Attribute RemoveDocumentProperties.VB_Description = "_(Description-RemoveDocumentProperties)"
    Const ProcedureName As String = "RemoveDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(wdRDIDocumentProperties)
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
    
    Call Target.RemoveDocumentInformation(wdRDIRemovePersonalInformation)
End Sub

'' == Task Pane Add-ins ==
'@Description "_(Description-RemoveTaskPaneAddIns)"
Public Sub RemoveTaskPaneAddIns()
Attribute RemoveTaskPaneAddIns.VB_Description = "_(Description-RemoveTaskPaneAddIns)"
    Const ProcedureName As String = "RemoveTaskPaneAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(wdRDITaskpaneWebExtensions)
End Sub

'' == Embedded Documents ==
'@Description "_(Description-ListEmbeddedDocuments_)"
Private Function ListEmbeddedDocuments_() As Collection
Attribute ListEmbeddedDocuments_.VB_Description = "_(Description-ListEmbeddedDocuments_)"
    Set ListEmbeddedDocuments_ = ListInlineShapes(wdInlineShapeEmbeddedOLEObject)
End Function

'@Description "_(Description-ListEmbeddedDocuments)"
Public Function ListEmbeddedDocuments() As Collection
Attribute ListEmbeddedDocuments.VB_Description = "_(Description-ListEmbeddedDocuments)"
    Const ProcedureName As String = "ListEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListEmbeddedDocuments = ListEmbeddedDocuments_
End Function

'@Description "_(Description-DeleteEmbeddedDocuments)"
Public Function DeleteEmbeddedDocuments() As Collection
Attribute DeleteEmbeddedDocuments.VB_Description = "_(Description-DeleteEmbeddedDocuments)"
    Const ProcedureName As String = "DeleteEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedEmbeddedDocumentInfos As Collection: Set deletedEmbeddedDocumentInfos = New Collection
    
    ' For some reason, this block does not work as expected.
    'Dim inlineShape_ As InlineShape
    'For Each inlineShape_ In ListEmbeddedDocuments_
    '    Call deletedEmbeddedDocumentInfos.Add(DeleteInlineShape(inlineShape_))
    'Next
    ' Alternative
    Dim listedInlineShapes As Collection
    Set listedInlineShapes = ListEmbeddedDocuments_
    Dim i As Long
    For i = listedInlineShapes.Count To 1 Step -1
        Call deletedEmbeddedDocumentInfos.Add(DeleteInlineShape(listedInlineShapes(i)))
    Next
    ' End Alternative
    
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

'' === Forms ===
'@Description "_(Description-ListForms_)"
Private Function ListForms_() As Collection
Attribute ListForms_.VB_Description = "_(Description-ListForms_)"
    Dim listedForms As Collection: Set listedForms = New Collection
    
    Dim form As FormField
    For Each form In Target.FormFields
        Call listedForms.Add(form)
    Next
    
    Set ListForms_ = listedForms
End Function

'@Description "_(Description-ListForms)"
Public Function ListForms() As Collection
Attribute ListForms.VB_Description = "_(Description-ListForms)"
    Const ProcedureName As String = "ListForms"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListForms = ListForms_
End Function

'@Description "_(Description-DeleteForms)"
Public Function DeleteForms() As Collection
Attribute DeleteForms.VB_Description = "_(Description-DeleteForms)"
    Const ProcedureName As String = "DeleteForms"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedFormInfos As Collection: Set deletedFormInfos = New Collection
    
    ' For some reason, this block does not work as expected.
    'Dim form As FormField
    'For Each form In ListForms_
    '    Call deletedFormInfos.Add(DeleteFormField(form))
    'Next
    ' Alternative
    Dim listedForms As Collection
    Set listedForms = ListForms_
    Dim i As Long
    For i = listedForms.Count To 1 Step -1
        Call deletedFormInfos.Add(DeleteFormField(listedForms(i)))
    Next
    ' End Alternative
    
    Set DeleteForms = deletedFormInfos
End Function

'' === ActiveX Controls ===
'@Description "_(Description-ListActiveXControls)"
Public Function ListActiveXControls() As Collection
Attribute ListActiveXControls.VB_Description = "_(Description-ListActiveXControls)"
    Const ProcedureName As String = "ListActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListActiveXControls = ListInlineShapes(wdInlineShapeOLEControlObject)
End Function

'@Description "_(Description-DeleteActiveXControls)"
Public Function DeleteActiveXControls() As Collection
Attribute DeleteActiveXControls.VB_Description = "_(Description-DeleteActiveXControls)"
    Const ProcedureName As String = "DeleteActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedActiveXControlInfos As Collection: Set deletedActiveXControlInfos = New Collection
    
    ' For some reason, this block does not work as expected.
    'Dim inlineShape_ As InlineShape
    'For Each inlineShape_ In ListInlineShapes(wdInlineShapeOLEControlObject)
    '    Call deletedActiveXControlInfos.Add(DeleteInlineShape(inlineShape_))
    'Next
    ' Alternative
    Dim listedInlineShapes As Collection
    Set listedInlineShapes = ListInlineShapes(wdInlineShapeOLEControlObject)
    Dim i As Long
    For i = listedInlineShapes.Count To 1 Step -1
        Call deletedActiveXControlInfos.Add(DeleteInlineShape(listedInlineShapes(i)))
    Next
    ' End Alternative
    
    Set DeleteActiveXControls = deletedActiveXControlInfos
End Function

'' == Ink ==
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
    
    Call Target.RemoveDocumentInformation(wdRDIInkAnnotations)
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

'' == Headers, Footers, and Watermarks ==
'' === Headers and Footers ===
'@Description "_(Description-InspectHeadersAndFooters)"
Public Function InspectHeadersAndFooters(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectHeadersAndFooters.VB_Description = "_(Description-InspectHeadersAndFooters)"
    Const ProcedureName As String = "InspectHeadersAndFooters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HeadersAndFooters)
    Call inspector.Inspect(status, Results)
    
    InspectHeadersAndFooters = status
End Function

'@Description "_(Description-FixHeadersAndFooters)"
Public Function FixHeadersAndFooters(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixHeadersAndFooters.VB_Description = "_(Description-FixHeadersAndFooters)"
    Const ProcedureName As String = "FixHeadersAndFooters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HeadersAndFooters)
    Call inspector.Fix(status, Results)
    
    FixHeadersAndFooters = status
End Function

'' === Headers ===
'@Description "_(Description-ListHeadersInSection_)"
Private Function ListHeadersInSection_(ByVal TargetSection As Section) As Collection
Attribute ListHeadersInSection_.VB_Description = "_(Description-ListHeadersInSection_)"
    Dim listedHeaders As Collection: Set listedHeaders = New Collection
    
    Dim header_ As HeaderFooter
    Dim Index As Long
    Dim i As Long
    For i = 1 To 3
        Select Case i
            Case 1
                Index = wdHeaderFooterFirstPage
            Case 2
                Index = wdHeaderFooterPrimary
            Case 3
                Index = wdHeaderFooterEvenPages
        End Select
        Set header_ = TargetSection.Headers(Index)
        If header_.Exists And Not header_.IsEmpty Then
            Call listedHeaders.Add(header_)
        End If
    Next
    
    Set ListHeadersInSection_ = listedHeaders
End Function

'@Description "_(Description-ListHeadersInSection)"
Public Function ListHeadersInSection(ByVal TargetSection As Section) As Collection
Attribute ListHeadersInSection.VB_Description = "_(Description-ListHeadersInSection)"
    Const ProcedureName As String = "ListHeadersInSection"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfParentIsNotTargetDocument(ProcedureName, TargetSection)
    
    Set ListHeadersInSection = ListHeadersInSection_(TargetSection)
End Function

'@Description "_(Description-ListHeaders_)"
Private Function ListHeaders_() As Collection
Attribute ListHeaders_.VB_Description = "_(Description-ListHeaders_)"
    Dim listedHeaders As Collection: Set listedHeaders = New Collection
    
    Dim section_ As Section
    Dim header_ As HeaderFooter
    For Each section_ In Target.Sections
        For Each header_ In ListHeadersInSection_(section_)
            If Not header_.LinkToPrevious Then
                Call listedHeaders.Add(header_)
            End If
        Next
    Next
    
    Set ListHeaders_ = listedHeaders
End Function

'@Description "_(Description-ListHeaders)"
Public Function ListHeaders() As Collection
Attribute ListHeaders.VB_Description = "_(Description-ListHeaders)"
    Const ProcedureName As String = "ListHeaders"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHeaders = ListHeaders_()
End Function

'' === Footers ===
'@Description "_(Description-ListFootersInSection_)"
Private Function ListFootersInSection_(ByVal TargetSection As Section) As Collection
Attribute ListFootersInSection_.VB_Description = "_(Description-ListFootersInSection_)"
    Dim listedFooters As Collection: Set listedFooters = New Collection
    
    Dim footer_ As HeaderFooter
    Dim Index As Long
    Dim i As Long
    For i = 1 To 3
        Select Case i
            Case 1
                Index = wdHeaderFooterFirstPage
            Case 2
                Index = wdHeaderFooterPrimary
            Case 3
                Index = wdHeaderFooterEvenPages
        End Select
        Set footer_ = TargetSection.Footers(Index)
        If footer_.Exists And Not footer_.IsEmpty Then
            Call listedFooters.Add(footer_)
        End If
    Next
    
    Set ListFootersInSection_ = listedFooters
End Function

'@Description "_(Description-ListFootersInSection)"
Public Function ListFootersInSection(ByVal TargetSection As Section) As Collection
Attribute ListFootersInSection.VB_Description = "_(Description-ListFootersInSection)"
    Const ProcedureName As String = "ListFootersInSection"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfParentIsNotTargetDocument(ProcedureName, TargetSection)
    
    Set ListFootersInSection = ListFootersInSection_(TargetSection)
End Function

'@Description "_(Description-ListFooters_)"
Private Function ListFooters_() As Collection
Attribute ListFooters_.VB_Description = "_(Description-ListFooters_)"
    Dim listedFooters As Collection: Set listedFooters = New Collection
    
    Dim section_ As Section
    Dim footer_ As HeaderFooter
    For Each section_ In Target.Sections
        For Each footer_ In ListFootersInSection_(section_)
            If Not footer_.LinkToPrevious Then
                Call listedFooters.Add(footer_)
            End If
        Next
    Next
    
    Set ListFooters_ = listedFooters
End Function

'@Description "_(Description-ListFooters)"
Public Function ListFooters() As Collection
Attribute ListFooters.VB_Description = "_(Description-ListFooters)"
    Const ProcedureName As String = "ListFooters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListFooters = ListFooters_()
End Function

'' == Invisible Content ==
'@Description "_(Description-ListInvisibleContent_)"
Private Function ListInvisibleContent_() As Collection
Attribute ListInvisibleContent_.VB_Description = "_(Description-ListInvisibleContent_)"
    Const ProcedureName As String = "ListInvisibleContent_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedInvisibleContent As Collection: Set listedInvisibleContent = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In Target.Shapes
        If shape_.Visible = msoFalse Then
            Call listedInvisibleContent.Add(shape_)
        End If
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

'' = For Word documents =
'' == Collapsed Headings ==
'@Description "_(Description-ListCollapsedHeadings_)"
Private Function ListCollapsedHeadings_() As Collection
Attribute ListCollapsedHeadings_.VB_Description = "_(Description-ListCollapsedHeadings_)"
    Dim listedCollapsedHeadings As Collection: Set listedCollapsedHeadings = New Collection
    
    Dim paragraph_ As Paragraph
    For Each paragraph_ In Target.Content.Paragraphs
        If paragraph_.CollapsedState = True Then
            Call listedCollapsedHeadings.Add(paragraph_)
        End If
    Next
    
    Set ListCollapsedHeadings_ = listedCollapsedHeadings
End Function

'@Description "_(Description-ListCollapsedHeadings)"
Public Function ListCollapsedHeadings() As Collection
Attribute ListCollapsedHeadings.VB_Description = "_(Description-ListCollapsedHeadings)"
    Const ProcedureName As String = "ListCollapsedHeadings"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListCollapsedHeadings = ListCollapsedHeadings_
End Function

'@Description "_(Description-ExpandCollapsedHeadings)"
Public Function ExpandCollapsedHeadings() As Collection
Attribute ExpandCollapsedHeadings.VB_Description = "_(Description-ExpandCollapsedHeadings)"
    Const ProcedureName As String = "ExpandCollapsedHeadings"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim expandedCollapsedHeadings As Collection: Set expandedCollapsedHeadings = New Collection
    
    Dim paragraph_ As Paragraph
    For Each paragraph_ In ListCollapsedHeadings_
        paragraph_.CollapsedState = False
        Call expandedCollapsedHeadings.Add(paragraph_)
    Next
    
    Set ExpandCollapsedHeadings = expandedCollapsedHeadings
End Function

'@Description "_(Description-InspectCollapsedHeadings)"
Public Function InspectCollapsedHeadings(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectCollapsedHeadings.VB_Description = "_(Description-InspectCollapsedHeadings)"
    Const ProcedureName As String = "InspectCollapsedHeadings"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.CollapsedHeadings)
    Call inspector.Inspect(status, Results)
    
    InspectCollapsedHeadings = status
End Function

'@Description "_(Description-FixCollapsedHeadings)"
Public Function FixCollapsedHeadings(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixCollapsedHeadings.VB_Description = "_(Description-FixCollapsedHeadings)"
    Const ProcedureName As String = "FixCollapsedHeadings"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.CollapsedHeadings)
    Call inspector.Fix(status, Results)
    
    FixCollapsedHeadings = status
End Function

'' == Hidden Text ==
'@Description "_(Description-ListHiddenText_)"
Private Function ListHiddenText_() As Collection
Attribute ListHiddenText_.VB_Description = "_(Description-ListHiddenText_)"
    Const ProcedureName As String = "ListHiddenText_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedHiddenText As Collection: Set listedHiddenText = New Collection
    
    Dim findRange As Range
    Set findRange = Target.Content
    Call findRange.Find.ClearFormatting
    With findRange.Find.Font
        .Hidden = True
    End With
    Do While findRange.Find.Execute
        Call listedHiddenText.Add(Target.Range(findRange.Start, findRange.End))
    Loop
    
    Set ListHiddenText_ = listedHiddenText
End Function

'@Description "_(Description-ListHiddenText)"
Public Function ListHiddenText() As Collection
Attribute ListHiddenText.VB_Description = "_(Description-ListHiddenText)"
    Const ProcedureName As String = "ListHiddenText"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenText = ListHiddenText_()
End Function

'@Description "_(Description-VisualizeHiddenText)"
Public Function VisualizeHiddenText() As Collection
Attribute VisualizeHiddenText.VB_Description = "_(Description-VisualizeHiddenText)"
    Const ProcedureName As String = "VisualizeHiddenText"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedHiddenText As Collection: Set visualizedHiddenText = New Collection
    
    Dim hiddenText_ As Range
    For Each hiddenText_ In ListHiddenText_
        hiddenText_.Font.Hidden = False
        Call visualizedHiddenText.Add(hiddenText_)
    Next
    
    Set VisualizeHiddenText = visualizedHiddenText
End Function

'@Description "_(Description-DeleteHiddenText)"
Public Function DeleteHiddenText() As Collection
Attribute DeleteHiddenText.VB_Description = "_(Description-DeleteHiddenText)"
    Const ProcedureName As String = "DeleteHiddenText"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedHiddenTextInfos As Collection: Set deletedHiddenTextInfos = New Collection
    
    Dim hiddenText_ As Range
    For Each hiddenText_ In ListHiddenText_
        Call deletedHiddenTextInfos.Add(DeleteRange(hiddenText_))
    Next
    
    Set DeleteHiddenText = deletedHiddenTextInfos
End Function

'@Description "_(Description-InspectHiddenText)"
Public Function InspectHiddenText(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectHiddenText.VB_Description = "_(Description-InspectHiddenText)"
    Const ProcedureName As String = "InspectHiddenText"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenText)
    Call inspector.Inspect(status, Results)
    
    InspectHiddenText = status
End Function

'@Description "_(Description-FixHiddenText)"
Public Function FixHiddenText(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixHiddenText.VB_Description = "_(Description-FixHiddenText)"
    Const ProcedureName As String = "FixHiddenText"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenText)
    Call inspector.Fix(status, Results)
    
    FixHiddenText = status
End Function
