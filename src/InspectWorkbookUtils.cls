VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InspectWorkbookUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "_(Description-InspectWorkbookUtils)"
'Copyright (c) 2022 telitas
'This file is released under the MIT License.
'See the LICENSE.txt file or https://opensource.org/licenses/mit-license.php for details.
'Package version: $PackageVersion
'@Folder "InspectWorkbookUtilsProject"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'@ModuleDescription("_(Description-InspectWorkbookUtils)")
Option Explicit

Private Const ModuleName As String = "InspectWorkbookUtils"

Private Enum InspectorIndex
    CustomXMLData = 1
    HeadersAndFooters = 2
    HiddenRowsAndColumns = 3
    HiddenWorkSheets = 4
    InvisibleContent = 5
End Enum

'@VariableDescription "_(Description-Target_)"
Private Target_ As Workbook
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
Public Property Get Target() As Workbook
Attribute Target.VB_Description = "_(Description-Target)"
    Set Target = Target_
End Property

'@Description "_(Description-Class_Initialize)"
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "_(Description-Class_Initialize)"
    Set Target_ = Nothing
End Sub

'@Description "_(Description-Initialize)"
Public Sub Initialize(ByVal Target As Workbook)
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

'@Description "_(Description-RaiseErrorIfParentIsNotTargetWorkbook)"
Private Sub RaiseErrorIfParentIsNotTargetWorkbook(ByVal ProcedureName As String, ByVal TargetWorksheet As Worksheet)
Attribute RaiseErrorIfParentIsNotTargetWorkbook.VB_Description = "_(Description-RaiseErrorIfParentIsNotTargetWorkbook)"
    If Not GetParentWorkbook(TargetWorksheet) Is Target Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="This sheet is not a child of " & Target.Name & ".")
    End If
End Sub

'@Description "_(Description-RaiseErrorIfArgumentIsNothing)"
Private Sub RaiseErrorIfArgumentIsNothing(ByVal ProcedureName As String, ByVal TargetArgument As Object, ByVal ArgumentName As String)
Attribute RaiseErrorIfArgumentIsNothing.VB_Description = "_(Description-RaiseErrorIfArgumentIsNothing)"
    If TargetArgument Is Nothing Then
        Call RaiseError(Number:=InvalidArgumentError, ProcedureName:=ProcedureName, Description:="Argument """ & ArgumentName & """ must be not Nothing.")
    End If
End Sub

'@Description "_(Description-GetParentWorkbook)"
Public Function GetParentWorkbook(ByVal TargetObject As Object) As Workbook
Attribute GetParentWorkbook.VB_Description = "_(Description-GetParentWorkbook)"
    If TypeOf TargetObject Is Workbook Then
        Set GetParentWorkbook = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    
    Set GetParentWorkbook = GetParentWorkbook(TargetObject.Parent)
End Function

'@Description "_(Description-GetParentWorksheet)"
Public Function GetParentWorksheet(ByVal TargetObject As Object) As Worksheet
Attribute GetParentWorksheet.VB_Description = "_(Description-GetParentWorksheet)"
    If TypeOf TargetObject Is Worksheet Then
        Set GetParentWorksheet = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Workbook Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    
    Set GetParentWorksheet = GetParentWorksheet(TargetObject.Parent)
End Function

'@Description "_(Description-GetParentRange)"
Public Function GetParentRange(ByVal TargetObject As Object) As Range
Attribute GetParentRange.VB_Description = "_(Description-GetParentRange)"
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

'@Description "_(Description-GetWorkbookLocation)"
Private Function GetWorkbookLocation(ByVal TargetWorkbook As Workbook) As String
Attribute GetWorkbookLocation.VB_Description = "_(Description-GetWorkbookLocation)"
    Dim externalAddressWithoutRange As String
    externalAddressWithoutRange = GetWorksheetLocation(TargetWorkbook.Worksheets(1))
    
    GetWorkbookLocation = Left$(externalAddressWithoutRange, InStrRev(externalAddressWithoutRange, "]"))
End Function

'@Description "_(Description-GetWorksheetLocation)"
Private Function GetWorksheetLocation(ByVal TargetWorksheet As Worksheet) As String
Attribute GetWorksheetLocation.VB_Description = "_(Description-GetWorksheetLocation)"
    Dim externalAddress_ As String
    externalAddress_ = TargetWorksheet.Cells(1, 1).Address(External:=True)
    
    GetWorksheetLocation = Left$(externalAddress_, InStrRev(externalAddress_, "!") - 1)
End Function

'' = For all Office documents =
'@Description "_(Description-ListShapesInWorksheet)"
Private Function ListShapesInWorksheet(ByVal TargetWorksheet As Worksheet, ByVal ShapeType As MsoShapeType) As Collection
Attribute ListShapesInWorksheet.VB_Description = "_(Description-ListShapesInWorksheet)"
    Dim listedShapes As Collection: Set listedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In TargetWorksheet.Shapes
        If shape_.Type = ShapeType Then
            Call listedShapes.Add(shape_)
        End If
    Next
    
    Set ListShapesInWorksheet = listedShapes
End Function

'@Description "_(Description-ListShapes)"
Private Function ListShapes(ByVal ShapeType As MsoShapeType) As Collection
Attribute ListShapes.VB_Description = "_(Description-ListShapes)"
    Dim listedShapes As Collection: Set listedShapes = New Collection
    
    Dim shape_ As Shape
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        For Each shape_ In ListShapesInWorksheet(worksheet_, ShapeType)
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
        deletedShapeInfo("Location") = GetParentWorksheet(TargetShape).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
        deletedShapeInfo("Shape.Type") = .Type
        Call .Delete
    End With
    
    Set DeleteShape = deletedShapeInfo
End Function

'@Description "_(Description-ListOLEObjectsInWorksheet)"
Private Function ListOLEObjectsInWorksheet(ByVal TargetWorksheet As Worksheet, ByVal OLEType As XlOLEType) As Collection
Attribute ListOLEObjectsInWorksheet.VB_Description = "_(Description-ListOLEObjectsInWorksheet)"
    Dim listedOLEObjects As Collection: Set listedOLEObjects = New Collection
    
    Dim oleObject_ As OLEObject
    For Each oleObject_ In TargetWorksheet.OLEObjects
        If oleObject_.OLEType = OLEType Then
            Call listedOLEObjects.Add(oleObject_)
        End If
    Next
    
    Set ListOLEObjectsInWorksheet = listedOLEObjects
End Function

'@Description "_(Description-ListOLEObjects)"
Private Function ListOLEObjects(ByVal OLEType As XlOLEType) As Collection
Attribute ListOLEObjects.VB_Description = "_(Description-ListOLEObjects)"
    Dim listedOLEObjects As Collection: Set listedOLEObjects = New Collection
    
    Dim worksheet_ As Worksheet
    Dim oleObject_ As OLEObject
    For Each worksheet_ In Target.Worksheets
        For Each oleObject_ In ListOLEObjectsInWorksheet(worksheet_, OLEType)
           Call listedOLEObjects.Add(oleObject_)
        Next
    Next
    
    Set ListOLEObjects = listedOLEObjects
End Function

'@Description "_(Description-DeleteOLEObject)"
Private Function DeleteOLEObject(ByVal TargetOLEObject As OLEObject) As Object
Attribute DeleteOLEObject.VB_Description = "_(Description-DeleteOLEObject)"
    Dim deletedOLEObjectInfo As Object: Set deletedOLEObjectInfo = CreateObject("Scripting.Dictionary")
    
    With TargetOLEObject
        deletedOLEObjectInfo("Type") = TypeName(TargetOLEObject)
        deletedOLEObjectInfo("Name") = .Name
        deletedOLEObjectInfo("Location") = GetParentWorksheet(TargetOLEObject).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
        deletedOLEObjectInfo("OLEObject.OLEType") = .OLEType
        Call .Delete
    End With
    
    Set DeleteOLEObject = deletedOLEObjectInfo
End Function

'@Description "_(Description-ConvertShapeToImage)"
Private Function ConvertShapeToImage(ByVal TargetShape As Shape) As Object
Attribute ConvertShapeToImage.VB_Description = "_(Description-ConvertShapeToImage)"
    Dim convertedImage As Object
    
    Dim shapeName As String
    TargetShape.Copy
    Set convertedImage = GetParentWorksheet(TargetShape).Pictures.Paste
    With convertedImage
        .Left = TargetShape.Left
        .Top = TargetShape.Top
        shapeName = TargetShape.Name
        Call TargetShape.Delete
        .Name = shapeName
    End With
    
    Set ConvertShapeToImage = convertedImage
End Function

'@Description "_(Description-ConvertShapesToImagesInWorksheet)"
Private Function ConvertShapesToImagesInWorksheet(ByVal TargetWorksheet As Worksheet, ByVal ShapeType As MsoShapeType) As Collection
Attribute ConvertShapesToImagesInWorksheet.VB_Description = "_(Description-ConvertShapesToImagesInWorksheet)"
    Const ProcedureName As String = "ConvertShapesToImagesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim convertedShapes As Collection: Set convertedShapes = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In ListShapesInWorksheet(TargetWorksheet, ShapeType)
        Call convertedShapes.Add(ConvertShapeToImage(shape_))
    Next
    
    Set ConvertShapesToImagesInWorksheet = convertedShapes
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

'@Description "_(Description-VisualizeComment)"
Private Function VisualizeComment(ByVal TargetComment As Comment) As Comment
Attribute VisualizeComment.VB_Description = "_(Description-VisualizeComment)"
    TargetComment.Visible = True
    
    Set VisualizeComment = TargetComment
End Function

'@Description "_(Description-DeleteComment)"
Private Function DeleteComment(ByVal TargetComment As Comment) As Object
Attribute DeleteComment.VB_Description = "_(Description-DeleteComment)"
    Dim deletedCommentInfo As Object: Set deletedCommentInfo = CreateObject("Scripting.Dictionary")
    
    With TargetComment
        deletedCommentInfo("Type") = TypeName(TargetComment)
        deletedCommentInfo("Location") = GetParentRange(TargetComment).Address(External:=True)
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
        deletedPropertyInfo("Location") = GetWorkbookLocation(GetParentWorkbook(TargetDocumentProperty))
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
        deletedMacroInfo("Location") = GetWorkbookLocation(Target)
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
        deletedCustomXMLPartInfo("Location") = GetWorkbookLocation(Target)
        Call .Delete
    End With
    
    Set DeleteCustomXMLPart = deletedCustomXMLPartInfo
End Function

'' == Comments ==
'@Description "_(Description-ListCommentsInWorksheet_)"
Private Function ListCommentsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCommentsInWorksheet_.VB_Description = "_(Description-ListCommentsInWorksheet_)"
    Dim listedComments As Collection: Set listedComments = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In TargetWorksheet.Comments
        Call listedComments.Add(comment_)
    Next
    
    Set ListCommentsInWorksheet_ = listedComments
End Function

'@Description "_(Description-ListCommentsInWorksheet)"
Public Function ListCommentsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCommentsInWorksheet.VB_Description = "_(Description-ListCommentsInWorksheet)"
    Const ProcedureName As String = "ListCommentsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListCommentsInWorksheet = ListCommentsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListComments_)"
Private Function ListComments_() As Collection
Attribute ListComments_.VB_Description = "_(Description-ListComments_)"
    Dim listedComments As Collection: Set listedComments = New Collection
    
    Dim worksheet_ As Worksheet
    Dim comment_ As Comment
    For Each worksheet_ In Target.Worksheets
        For Each comment_ In ListCommentsInWorksheet_(worksheet_)
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

'@Description "_(Description-VisualizeCommentsInWorksheet)"
Public Function VisualizeCommentsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeCommentsInWorksheet.VB_Description = "_(Description-VisualizeCommentsInWorksheet)"
    Dim visualizedComments As Collection: Set visualizedComments = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In ListCommentsInWorksheet_(TargetWorksheet)
        If Not comment_.Visible Then
            Call visualizedComments.Add(VisualizeComment(comment_))
        End If
    Next
    
    Set VisualizeCommentsInWorksheet = visualizedComments
End Function

'@Description "_(Description-VisualizeComments)"
Public Function VisualizeComments() As Collection
Attribute VisualizeComments.VB_Description = "_(Description-VisualizeComments)"
    Dim visualizedComments As Collection: Set visualizedComments = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In ListComments_()
        If Not comment_.Visible Then
            Call visualizedComments.Add(VisualizeComment(comment_))
        End If
    Next
    
    Set VisualizeComments = visualizedComments
End Function

'@Description "_(Description-DeleteCommentsInWorksheet)"
Public Function DeleteCommentsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteCommentsInWorksheet.VB_Description = "_(Description-DeleteCommentsInWorksheet)"
    Dim deletedCommentAddresses As Collection: Set deletedCommentAddresses = New Collection
    
    Dim comment_ As Comment
    For Each comment_ In ListCommentsInWorksheet_(TargetWorksheet)
        Call deletedCommentAddresses.Add(DeleteComment(comment_))
    Next
    
    Set DeleteCommentsInWorksheet = deletedCommentAddresses
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
    
    Call Target.RemoveDocumentInformation(xlRDIComments)
End Sub

'' == Document Properties and Personal Information ==
'' === Document Properties ===
'@Description "_(Description-RemoveDocumentProperties)"
Public Sub RemoveDocumentProperties()
Attribute RemoveDocumentProperties.VB_Description = "_(Description-RemoveDocumentProperties)"
    Const ProcedureName As String = "RemoveDocumentProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(xlRDIDocumentProperties)
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
    
    Call Target.RemoveDocumentInformation(xlRDIRemovePersonalInformation)
End Sub

'' === Printer Path ===
'@Description "_(Description-RemovePrinterPaths)"
Public Sub RemovePrinterPaths()
Attribute RemovePrinterPaths.VB_Description = "_(Description-RemovePrinterPaths)"
    Const ProcedureName As String = "RemovePrinterPaths"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(xlRDIPrinterPath)
End Sub

'' == Content Add-ins ==
'@Description "_(Description-ListContentAddInsInWorksheet_)"
Private Function ListContentAddInsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListContentAddInsInWorksheet_.VB_Description = "_(Description-ListContentAddInsInWorksheet_)"
    Set ListContentAddInsInWorksheet_ = ListShapesInWorksheet(TargetWorksheet, msoContentApp)
End Function

'@Description "_(Description-ListContentAddInsInWorksheet)"
Public Function ListContentAddInsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListContentAddInsInWorksheet.VB_Description = "_(Description-ListContentAddInsInWorksheet)"
    Const ProcedureName As String = "ListContentAddInsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListContentAddInsInWorksheet = ListContentAddInsInWorksheet_(TargetWorksheet)
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

'@Description "_(Description-ConvertContentAddInsToImagesInWorksheet)"
Public Function ConvertContentAddInsToImagesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ConvertContentAddInsToImagesInWorksheet.VB_Description = "_(Description-ConvertContentAddInsToImagesInWorksheet)"
    Const ProcedureName As String = "ConvertContentAddInsToImagesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ConvertContentAddInsToImagesInWorksheet = ConvertShapesToImagesInWorksheet(TargetWorksheet, msoContentApp)
End Function

'@Description "_(Description-ConvertContentAddInsToImages)"
Public Function ConvertContentAddInsToImages() As Collection
Attribute ConvertContentAddInsToImages.VB_Description = "_(Description-ConvertContentAddInsToImages)"
    Const ProcedureName As String = "ConvertContentAddInsToImages"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ConvertContentAddInsToImages = ConvertShapesToImages(msoContentApp)
End Function

'@Description "_(Description-DeleteContentAddInsInWorksheet)"
Public Function DeleteContentAddInsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteContentAddInsInWorksheet.VB_Description = "_(Description-DeleteContentAddInsInWorksheet)"
    Const ProcedureName As String = "DeleteContentAddInsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedContentAddInInfos As Collection: Set deletedContentAddInInfos = New Collection
    
    Dim contentAddIn As Shape
    For Each contentAddIn In ListContentAddInsInWorksheet_(TargetWorksheet)
        Call deletedContentAddInInfos.Add(DeleteShape(contentAddIn))
    Next
    
    Set DeleteContentAddInsInWorksheet = deletedContentAddInInfos
End Function

'@Description "_(Description-DeleteContentAddIns)"
Public Function DeleteContentAddIns() As Collection
Attribute DeleteContentAddIns.VB_Description = "_(Description-DeleteContentAddIns)"
    Const ProcedureName As String = "DeleteContentAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedContentAddInInfos As Collection: Set deletedContentAddInInfos = New Collection
    
    Dim contentAddIn As Shape
    For Each contentAddIn In ListContentAddIns_
        Call deletedContentAddInInfos.Add(DeleteShape(contentAddIn))
    Next
    
    Set DeleteContentAddIns = deletedContentAddInInfos
End Function

'@Description "_(Description-RemoveContentAddIns)"
Public Sub RemoveContentAddIns()
Attribute RemoveContentAddIns.VB_Description = "_(Description-RemoveContentAddIns)"
    Const ProcedureName As String = "RemoveContentAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(xlRDIInlineWebExtensions)
End Sub

'' == Task Pane Add-ins ==
'@Description "_(Description-RemoveTaskPaneAddIns)"
Public Sub RemoveTaskPaneAddIns()
Attribute RemoveTaskPaneAddIns.VB_Description = "_(Description-RemoveTaskPaneAddIns)"
    Const ProcedureName As String = "RemoveTaskPaneAddIns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(xlRDITaskpaneWebExtensions)
End Sub

'' == Embedded Documents ==
'@Description "_(Description-ListEmbeddedDocumentsInWorksheet_)"
Private Function ListEmbeddedDocumentsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListEmbeddedDocumentsInWorksheet_.VB_Description = "_(Description-ListEmbeddedDocumentsInWorksheet_)"
    Set ListEmbeddedDocumentsInWorksheet_ = ListOLEObjectsInWorksheet(TargetWorksheet, xlOLEEmbed)
End Function

'@Description "_(Description-ListEmbeddedDocumentsInWorksheet)"
Public Function ListEmbeddedDocumentsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListEmbeddedDocumentsInWorksheet.VB_Description = "_(Description-ListEmbeddedDocumentsInWorksheet)"
    Const ProcedureName As String = "ListEmbeddedDocumentsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListEmbeddedDocumentsInWorksheet = ListEmbeddedDocumentsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListEmbeddedDocuments_)"
Private Function ListEmbeddedDocuments_() As Collection
Attribute ListEmbeddedDocuments_.VB_Description = "_(Description-ListEmbeddedDocuments_)"
    Set ListEmbeddedDocuments_ = ListOLEObjects(xlOLEEmbed)
End Function

'@Description "_(Description-ListEmbeddedDocuments)"
Public Function ListEmbeddedDocuments() As Collection
Attribute ListEmbeddedDocuments.VB_Description = "_(Description-ListEmbeddedDocuments)"
    Const ProcedureName As String = "ListEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListEmbeddedDocuments = ListEmbeddedDocuments_
End Function

'@Description "_(Description-DeleteEmbeddedDocumentsInWorksheet)"
Public Function DeleteEmbeddedDocumentsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteEmbeddedDocumentsInWorksheet.VB_Description = "_(Description-DeleteEmbeddedDocumentsInWorksheet)"
    Const ProcedureName As String = "DeleteEmbeddedDocumentsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedEmbeddedDocumentInfos As Collection: Set deletedEmbeddedDocumentInfos = New Collection
    
    Dim oleObject_ As OLEObject
    For Each oleObject_ In ListEmbeddedDocumentsInWorksheet_(TargetWorksheet)
        Call deletedEmbeddedDocumentInfos.Add(DeleteOLEObject(oleObject_))
    Next
    
    Set DeleteEmbeddedDocumentsInWorksheet = deletedEmbeddedDocumentInfos
End Function

'@Description "_(Description-DeleteEmbeddedDocuments)"
Public Function DeleteEmbeddedDocuments() As Collection
Attribute DeleteEmbeddedDocuments.VB_Description = "_(Description-DeleteEmbeddedDocuments)"
    Const ProcedureName As String = "DeleteEmbeddedDocuments"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedEmbeddedDocumentInfos As Collection: Set deletedEmbeddedDocumentInfos = New Collection
    
    Dim oleObject_ As OLEObject
    For Each oleObject_ In ListEmbeddedDocuments_
        Call deletedEmbeddedDocumentInfos.Add(DeleteOLEObject(oleObject_))
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

'' === Forms ===
'@Description "_(Description-ListFormsInWorksheet_)"
Private Function ListFormsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListFormsInWorksheet_.VB_Description = "_(Description-ListFormsInWorksheet_)"
    Set ListFormsInWorksheet_ = ListShapesInWorksheet(TargetWorksheet, msoFormControl)
End Function

'@Description "_(Description-ListFormsInWorksheet)"
Public Function ListFormsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListFormsInWorksheet.VB_Description = "_(Description-ListFormsInWorksheet)"
    Const ProcedureName As String = "ListFormsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListFormsInWorksheet = ListFormsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListForms_)"
Private Function ListForms_() As Collection
Attribute ListForms_.VB_Description = "_(Description-ListForms_)"
    Set ListForms_ = ListShapes(msoFormControl)
End Function

'@Description "_(Description-ListForms)"
Public Function ListForms() As Collection
Attribute ListForms.VB_Description = "_(Description-ListForms)"
    Const ProcedureName As String = "ListForms"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListForms = ListForms_
End Function

'@Description "_(Description-DeleteFormsInWorksheet)"
Public Function DeleteFormsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteFormsInWorksheet.VB_Description = "_(Description-DeleteFormsInWorksheet)"
    Const ProcedureName As String = "DeleteFormsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedFormInfos As Collection: Set deletedFormInfos = New Collection
    
    Dim form As Shape
    For Each form In ListFormsInWorksheet_(TargetWorksheet)
        Call deletedFormInfos.Add(DeleteShape(form))
    Next
    
    Set DeleteFormsInWorksheet = deletedFormInfos
End Function

'@Description "_(Description-DeleteForms)"
Public Function DeleteForms() As Collection
Attribute DeleteForms.VB_Description = "_(Description-DeleteForms)"
    Const ProcedureName As String = "DeleteForms"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedFormInfos As Collection: Set deletedFormInfos = New Collection
    
    Dim form As Shape
    For Each form In ListForms_
        Call deletedFormInfos.Add(DeleteShape(form))
    Next
    
    Set DeleteForms = deletedFormInfos
End Function

'' === ActiveX Controls ===
'@Description "_(Description-ListActiveXControlsInWorksheet_)"
Private Function ListActiveXControlsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListActiveXControlsInWorksheet_.VB_Description = "_(Description-ListActiveXControlsInWorksheet_)"
    Set ListActiveXControlsInWorksheet_ = ListOLEObjectsInWorksheet(TargetWorksheet, xlOLEControl)
End Function

'@Description "_(Description-ListActiveXControlsInWorksheet)"
Public Function ListActiveXControlsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListActiveXControlsInWorksheet.VB_Description = "_(Description-ListActiveXControlsInWorksheet)"
    Const ProcedureName As String = "ListActiveXControlsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListActiveXControlsInWorksheet = ListActiveXControlsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListActiveXControls_)"
Private Function ListActiveXControls_() As Collection
Attribute ListActiveXControls_.VB_Description = "_(Description-ListActiveXControls_)"
    Set ListActiveXControls_ = ListOLEObjects(xlOLEControl)
End Function

'@Description "_(Description-ListActiveXControls)"
Public Function ListActiveXControls() As Collection
Attribute ListActiveXControls.VB_Description = "_(Description-ListActiveXControls)"
    Const ProcedureName As String = "ListActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListActiveXControls = ListActiveXControls_
End Function

'@Description "_(Description-DeleteActiveXControlsInWorksheet)"
Public Function DeleteActiveXControlsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteActiveXControlsInWorksheet.VB_Description = "_(Description-DeleteActiveXControlsInWorksheet)"
    Const ProcedureName As String = "DeleteActiveXControlsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedActiveXControlInfos As Collection: Set deletedActiveXControlInfos = New Collection
    
    Dim oleObject_ As OLEObject
    For Each oleObject_ In ListActiveXControlsInWorksheet_(TargetWorksheet)
        Call deletedActiveXControlInfos.Add(DeleteOLEObject(oleObject_))
    Next
    
    Set DeleteActiveXControlsInWorksheet = deletedActiveXControlInfos
End Function

'@Description "_(Description-DeleteActiveXControls)"
Public Function DeleteActiveXControls() As Collection
Attribute DeleteActiveXControls.VB_Description = "_(Description-DeleteActiveXControls)"
    Const ProcedureName As String = "DeleteActiveXControls"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedActiveXControlInfos As Collection: Set deletedActiveXControlInfos = New Collection
    
    Dim oleObject_ As OLEObject
    For Each oleObject_ In ListActiveXControls_
        Call deletedActiveXControlInfos.Add(DeleteOLEObject(oleObject_))
    Next
    
    Set DeleteActiveXControls = deletedActiveXControlInfos
End Function

'' == Ink ==
'@Description "_(Description-ListInkInWorksheet_)"
Private Function ListInkInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListInkInWorksheet_.VB_Description = "_(Description-ListInkInWorksheet_)"
    Set ListInkInWorksheet_ = ListShapesInWorksheet(TargetWorksheet, msoInkComment)
End Function

'@Description "_(Description-ListInkInWorksheet)"
Public Function ListInkInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListInkInWorksheet.VB_Description = "_(Description-ListInkInWorksheet)"
    Const ProcedureName As String = "ListInkInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListInkInWorksheet = ListInkInWorksheet_(TargetWorksheet)
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

'@Description "_(Description-DeleteInkInWorksheet)"
Public Function DeleteInkInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteInkInWorksheet.VB_Description = "_(Description-DeleteInkInWorksheet)"
    Const ProcedureName As String = "DeleteInkInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedInkInfos As Collection: Set deletedInkInfos = New Collection
    
    Dim ink_ As Shape
    For Each ink_ In ListInkInWorksheet_(TargetWorksheet)
        Call deletedInkInfos.Add(DeleteShape(ink_))
    Next
    
    Set DeleteInkInWorksheet = deletedInkInfos
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
    
    Call Target.RemoveDocumentInformation(xlRDIInkAnnotations)
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

'' == Headers and Footers ==
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
'@Description "_(Description-ContainsHeadersInWorksheet_)"
Private Function ContainsHeadersInWorksheet_(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ContainsHeadersInWorksheet_.VB_Description = "_(Description-ContainsHeadersInWorksheet_)"
    With TargetWorksheet.PageSetup
        If .LeftHeader <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
        If .CenterHeader <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
        If .RightHeader <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
        If GetFilenameWithoutError(.LeftHeaderPicture) <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
        If GetFilenameWithoutError(.CenterHeaderPicture) <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
        If GetFilenameWithoutError(.RightHeaderPicture) <> vbNullString Then
            ContainsHeadersInWorksheet_ = True
            Exit Function
        End If
    End With
End Function

'@Description "_(Description-ContainsHeadersInWorksheet)"
Public Function ContainsHeadersInWorksheet(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ContainsHeadersInWorksheet.VB_Description = "_(Description-ContainsHeadersInWorksheet)"
    Const ProcedureName As String = "ContainsHeadersInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    ContainsHeadersInWorksheet = ContainsHeadersInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListSheetsContainsHeaders_)"
Private Function ListSheetsContainsHeaders_() As Collection
Attribute ListSheetsContainsHeaders_.VB_Description = "_(Description-ListSheetsContainsHeaders_)"
    Dim listedSheetsContainsHeaders As Collection: Set listedSheetsContainsHeaders = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        If ContainsHeadersInWorksheet_(worksheet_) Then
            Call listedSheetsContainsHeaders.Add(worksheet_)
        End If
    Next
    
    Set ListSheetsContainsHeaders_ = listedSheetsContainsHeaders
End Function

'@Description "_(Description-ListSheetsContainsHeaders)"
Public Function ListSheetsContainsHeaders() As Collection
Attribute ListSheetsContainsHeaders.VB_Description = "_(Description-ListSheetsContainsHeaders)"
    Const ProcedureName As String = "ListSheetsContainsHeaders"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListSheetsContainsHeaders = ListSheetsContainsHeaders_
End Function

'@Description "_(Description-ClearHeadersInWorksheet_)"
Private Function ClearHeadersInWorksheet_(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ClearHeadersInWorksheet_.VB_Description = "_(Description-ClearHeadersInWorksheet_)"
    If ContainsHeadersInWorksheet_(TargetWorksheet) Then
        With TargetWorksheet.PageSetup
            .LeftHeader = vbNullString
            .CenterHeader = vbNullString
            .RightHeader = vbNullString
            .LeftHeaderPicture.fileName = vbNullString
            .CenterHeaderPicture.fileName = vbNullString
            .RightHeaderPicture.fileName = vbNullString
        End With
        ClearHeadersInWorksheet_ = True
        Exit Function
    End If
End Function

'@Description "_(Description-ClearHeadersInWorksheet)"
Public Function ClearHeadersInWorksheet(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ClearHeadersInWorksheet.VB_Description = "_(Description-ClearHeadersInWorksheet)"
    Const ProcedureName As String = "ClearHeadersInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    ClearHeadersInWorksheet = ClearHeadersInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ClearHeaders)"
Public Function ClearHeaders() As Collection
Attribute ClearHeaders.VB_Description = "_(Description-ClearHeaders)"
    Const ProcedureName As String = "ClearHeaders"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim clearedSheetsContainsHeaders As Collection: Set clearedSheetsContainsHeaders = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In ListSheetsContainsHeaders_
        Call ClearHeadersInWorksheet_(worksheet_)
        Call clearedSheetsContainsHeaders.Add(worksheet_)
    Next
    
    Set ClearHeaders = clearedSheetsContainsHeaders
End Function

'' === Footers ===
'@Description "_(Description-ContainsFootersInWorksheet_)"
Private Function ContainsFootersInWorksheet_(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ContainsFootersInWorksheet_.VB_Description = "_(Description-ContainsFootersInWorksheet_)"
    With TargetWorksheet.PageSetup
        If .LeftFooter <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
        If .CenterFooter <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
        If .RightFooter <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
        
        If GetFilenameWithoutError(.LeftFooterPicture) <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
        If GetFilenameWithoutError(.CenterFooterPicture) <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
        If GetFilenameWithoutError(.RightFooterPicture) <> vbNullString Then
            ContainsFootersInWorksheet_ = True
            Exit Function
        End If
    End With
End Function

'@Description "_(Description-ContainsFootersInWorksheet)"
Public Function ContainsFootersInWorksheet(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ContainsFootersInWorksheet.VB_Description = "_(Description-ContainsFootersInWorksheet)"
    Const ProcedureName As String = "ContainsFootersInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    ContainsFootersInWorksheet = ContainsFootersInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListSheetsContainsFooters_)"
Private Function ListSheetsContainsFooters_() As Collection
Attribute ListSheetsContainsFooters_.VB_Description = "_(Description-ListSheetsContainsFooters_)"
    Dim listedSheetsContainsFooters As Collection: Set listedSheetsContainsFooters = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        If ContainsFootersInWorksheet_(worksheet_) Then
            Call listedSheetsContainsFooters.Add(worksheet_)
        End If
    Next
    
    Set ListSheetsContainsFooters_ = listedSheetsContainsFooters
End Function

'@Description "_(Description-ListSheetsContainsFooters)"
Public Function ListSheetsContainsFooters() As Collection
Attribute ListSheetsContainsFooters.VB_Description = "_(Description-ListSheetsContainsFooters)"
    Const ProcedureName As String = "ListSheetsContainsFooters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListSheetsContainsFooters = ListSheetsContainsFooters_
End Function

'@Description "_(Description-ClearFootersInWorksheet_)"
Private Function ClearFootersInWorksheet_(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ClearFootersInWorksheet_.VB_Description = "_(Description-ClearFootersInWorksheet_)"
    If ContainsFootersInWorksheet_(TargetWorksheet) Then
        With TargetWorksheet.PageSetup
            .LeftFooter = vbNullString
            .CenterFooter = vbNullString
            .RightFooter = vbNullString
            .LeftFooterPicture.fileName = vbNullString
            .CenterFooterPicture.fileName = vbNullString
            .RightFooterPicture.fileName = vbNullString
        End With
        ClearFootersInWorksheet_ = True
        Exit Function
    End If
End Function

'@Description "_(Description-ClearFootersInWorksheet)"
Public Function ClearFootersInWorksheet(ByVal TargetWorksheet As Worksheet) As Boolean
Attribute ClearFootersInWorksheet.VB_Description = "_(Description-ClearFootersInWorksheet)"
    Const ProcedureName As String = "ClearFootersInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    ClearFootersInWorksheet = ClearFootersInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ClearFooters)"
Public Function ClearFooters() As Collection
Attribute ClearFooters.VB_Description = "_(Description-ClearFooters)"
    Const ProcedureName As String = "ClearFooters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim clearedSheetsContainsHeaders As Collection: Set clearedSheetsContainsHeaders = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In ListSheetsContainsFooters_
        Call ClearFootersInWorksheet_(worksheet_)
        Call clearedSheetsContainsHeaders.Add(worksheet_)
    Next
    
    Set ClearFooters = clearedSheetsContainsHeaders
End Function

'' == Invisible Content ==
'@Description "_(Description-ListInvisibleContentInWorksheet_)"
Private Function ListInvisibleContentInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListInvisibleContentInWorksheet_.VB_Description = "_(Description-ListInvisibleContentInWorksheet_)"
    Dim listedInvisibleContent As Collection: Set listedInvisibleContent = New Collection
    
    Dim shape_ As Shape
    For Each shape_ In TargetWorksheet.Shapes
        If shape_.Visible = msoFalse Then
            Call listedInvisibleContent.Add(shape_)
        End If
    Next
    
    Set ListInvisibleContentInWorksheet_ = listedInvisibleContent
End Function

'@Description "_(Description-ListInvisibleContentInWorksheet)"
Public Function ListInvisibleContentInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListInvisibleContentInWorksheet.VB_Description = "_(Description-ListInvisibleContentInWorksheet)"
    Const ProcedureName As String = "ListInvisibleContentInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListInvisibleContentInWorksheet = ListInvisibleContentInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListInvisibleContent_)"
Private Function ListInvisibleContent_() As Collection
Attribute ListInvisibleContent_.VB_Description = "_(Description-ListInvisibleContent_)"
    Const ProcedureName As String = "ListInvisibleContent_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedInvisibleContent As Collection: Set listedInvisibleContent = New Collection
    
    Dim worksheet_ As Worksheet
    Dim shape_ As Shape
    For Each worksheet_ In Target.Worksheets
        For Each shape_ In ListInvisibleContentInWorksheet_(worksheet_)
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

'@Description "_(Description-VisualizeInvisibleContentInWorksheet)"
Public Function VisualizeInvisibleContentInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeInvisibleContentInWorksheet.VB_Description = "_(Description-VisualizeInvisibleContentInWorksheet)"
    Const ProcedureName As String = "VisualizeInvisibleContentInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim visualizedInvisibleContent As Collection: Set visualizedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContentInWorksheet_(TargetWorksheet)
        Call visualizedInvisibleContent.Add(VisualizeShape(content_))
    Next
    
    Set VisualizeInvisibleContentInWorksheet = visualizedInvisibleContent
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

'@Description "_(Description-DeleteInvisibleContentInWorksheet)"
Public Function DeleteInvisibleContentInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteInvisibleContentInWorksheet.VB_Description = "_(Description-DeleteInvisibleContentInWorksheet)"
    Const ProcedureName As String = "DeleteInvisibleContentInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedInvisibleContent As Collection: Set deletedInvisibleContent = New Collection
    
    Dim content_ As Shape
    For Each content_ In ListInvisibleContentInWorksheet_(TargetWorksheet)
        Call deletedInvisibleContent.Add(DeleteShape(content_))
    Next
    
    Set DeleteInvisibleContentInWorksheet = deletedInvisibleContent
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

'' = For Excel documents =
'@Description "_(Description-FindCellsInWorksheet)"
Private Function FindCellsInWorksheet( _
    ByVal TargetWorksheet As Worksheet, _
    ByVal What As String _
) As Collection
Attribute FindCellsInWorksheet.VB_Description = "_(Description-FindCellsInWorksheet)"
    Dim listedCells As Collection: Set listedCells = New Collection
    
    Dim foundCell As Range
    Dim firstFoundCellAddress As String
    With TargetWorksheet.Cells
        Set foundCell = .Find( _
            What:=What, _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            MatchByte:=False, _
            SearchFormat:=False _
        )
        
        If foundCell Is Nothing Then
            Set FindCellsInWorksheet = listedCells
            Exit Function
        End If
        
        Call listedCells.Add(foundCell)
        firstFoundCellAddress = foundCell.Address
        Do
            Set foundCell = .FindNext(foundCell)
            If foundCell.Address = firstFoundCellAddress Then
                Exit Do
            Else
                Call listedCells.Add(foundCell)
            End If
        Loop
    End With
    
    Set FindCellsInWorksheet = listedCells
End Function

'@Description "_(Description-FindFunctionInWorksheet)"
Private Function FindFunctionInWorksheet(ByVal TargetWorksheet As Worksheet, ByVal FunctionName As String) As Collection
Attribute FindFunctionInWorksheet.VB_Description = "_(Description-FindFunctionInWorksheet)"
    Dim candidateCells As Collection: Set candidateCells = New Collection
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "[+\-*/^=<>&:,\(] *" & FunctionName & " *\("
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim cell_ As Range
    For Each cell_ In FindCellsInWorksheet( _
        TargetWorksheet:=TargetWorksheet, _
        What:=FunctionName _
    )
        If regex.Test(cell_.formula) Then
            Call candidateCells.Add(cell_)
        End If
    Next
    
    Set FindFunctionInWorksheet = candidateCells
End Function

'@Description "_(Description-DeletePivotTable)"
Private Function DeletePivotTable(ByVal TargetPivotTable As PivotTable) As Object
Attribute DeletePivotTable.VB_Description = "_(Description-DeletePivotTable)"
    Dim deletedPivotTableInfo As Object: Set deletedPivotTableInfo = CreateObject("Scripting.Dictionary")
    
    With TargetPivotTable
        deletedPivotTableInfo("Type") = TypeName(TargetPivotTable)
        deletedPivotTableInfo("Name") = .Name
        deletedPivotTableInfo("Location") = .TableRange2.Address(External:=True)
        Call .TableRange2.Clear
    End With
    
    Set DeletePivotTable = deletedPivotTableInfo
End Function

'@Description "_(Description-DeleteSlicerCache)"
Private Function DeleteSlicerCache(ByVal TargetSlicerCache As SlicerCache) As Object
Attribute DeleteSlicerCache.VB_Description = "_(Description-DeleteSlicerCache)"
    Dim deletedSlicerCacheInfo As Object: Set deletedSlicerCacheInfo = CreateObject("Scripting.Dictionary")
    
    With TargetSlicerCache
        deletedSlicerCacheInfo("Type") = TypeName(TargetSlicerCache)
        deletedSlicerCacheInfo("Name") = .Name
        deletedSlicerCacheInfo("Location") = GetWorkbookLocation(GetParentWorkbook(TargetSlicerCache))
        deletedSlicerCacheInfo("SlicerCache.SlicerCacheType") = .SlicerCacheType
        Call .Delete
    End With
    
    Set DeleteSlicerCache = deletedSlicerCacheInfo
End Function

'@Description "_(Description-DeleteChartObject)"
Private Function DeleteChartObject(ByVal TargetChartObject As ChartObject) As Object
Attribute DeleteChartObject.VB_Description = "_(Description-DeleteChartObject)"
    Dim deletedPivotChartInfo As Object: Set deletedPivotChartInfo = CreateObject("Scripting.Dictionary")
    
    With TargetChartObject
        deletedPivotChartInfo("Type") = TypeName(TargetChartObject)
        deletedPivotChartInfo("Name") = .Name
        deletedPivotChartInfo("Location") = GetParentWorksheet(TargetChartObject).Range(.TopLeftCell, .BottomRightCell).Address(External:=True)
        Call .Delete
    End With
    
    Set DeleteChartObject = deletedPivotChartInfo
End Function

'@Description "_(Description-DeleteCustomProperty)"
Private Function DeleteCustomProperty(ByVal TargetCustomProperty As CustomProperty) As Object
Attribute DeleteCustomProperty.VB_Description = "_(Description-DeleteCustomProperty)"
    Dim deletedCustomProperyInfo As Object: Set deletedCustomProperyInfo = CreateObject("Scripting.Dictionary")
    
    With TargetCustomProperty
        deletedCustomProperyInfo("Type") = TypeName(TargetCustomProperty)
        deletedCustomProperyInfo("Name") = .Name
        deletedCustomProperyInfo("Location") = GetWorksheetLocation(GetParentWorksheet(TargetCustomProperty))
        Call .Delete
    End With
    
    Set DeleteCustomProperty = deletedCustomProperyInfo
End Function

'@Description "_(Description-VisualizeWorksheet)"
Private Function VisualizeWorksheet(ByVal TargetWorksheet As Worksheet) As Worksheet
Attribute VisualizeWorksheet.VB_Description = "_(Description-VisualizeWorksheet)"
    TargetWorksheet.Visible = xlSheetVisible
    
    Set VisualizeWorksheet = TargetWorksheet
End Function

'@Description "_(Description-DeleteWorksheet)"
Private Function DeleteWorksheet(ByVal TargetWorksheet As Worksheet) As Object
Attribute DeleteWorksheet.VB_Description = "_(Description-DeleteWorksheet)"
    Dim deletedWorksheetInfo As Object: Set deletedWorksheetInfo = CreateObject("Scripting.Dictionary")
    
    Dim currentDisplayAlerts As Boolean
    currentDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If TargetWorksheet.Visible = xlSheetVeryHidden Then
        TargetWorksheet.Visible = xlSheetHidden
    End If
    
    With TargetWorksheet
        deletedWorksheetInfo("Type") = TypeName(TargetWorksheet)
        deletedWorksheetInfo("Name") = .Name
        deletedWorksheetInfo("Location") = GetWorksheetLocation(TargetWorksheet)
        Call .Delete
    End With
    
    Application.DisplayAlerts = currentDisplayAlerts
    
    Set DeleteWorksheet = deletedWorksheetInfo
End Function

'@Description "_(Description-VisualizeName)"
Private Function VisualizeName(ByVal TargetName As Name) As String
Attribute VisualizeName.VB_Description = "_(Description-VisualizeName)"
    TargetName.Visible = True
    
    VisualizeName = TargetName
End Function

'@Description "_(Description-DeleteName)"
Private Function DeleteName(ByVal TargetName As Name) As Object
Attribute DeleteName.VB_Description = "_(Description-DeleteName)"
    Dim deletedNameInfo As Object: Set deletedNameInfo = CreateObject("Scripting.Dictionary")
    
    With TargetName
        deletedNameInfo("Type") = TypeName(TargetName)
        deletedNameInfo("Name") = .Name
        If TypeOf .Parent Is Workbook Then
            deletedNameInfo("Location") = GetWorkbookLocation(GetParentWorkbook(TargetName))
        Else
            deletedNameInfo("Location") = GetWorksheetLocation(GetParentWorksheet(TargetName))
        End If
        Call .Delete
    End With
    
    Set DeleteName = deletedNameInfo
End Function

'@Description "_(Description-DeleteScenario)"
Private Function DeleteScenario(ByVal TargetScenario As Scenario) As Object
Attribute DeleteScenario.VB_Description = "_(Description-DeleteScenario)"
    Dim deletedScenarioInfo As Object: Set deletedScenarioInfo = CreateObject("Scripting.Dictionary")
    
    With TargetScenario
        deletedScenarioInfo("Type") = TypeName(TargetScenario)
        deletedScenarioInfo("Name") = .Name
        deletedScenarioInfo("Location") = GetWorksheetLocation(GetParentWorksheet(TargetScenario))
        Call .Delete
    End With
    
    Set DeleteScenario = deletedScenarioInfo
End Function

'@Description "_(Description-GetFilenameWithoutError)"
Private Function GetFilenameWithoutError(ByVal Graph As Graphic) As String
Attribute GetFilenameWithoutError.VB_Description = "_(Description-GetFilenameWithoutError)"
    Dim fileName As String
    
    On Error Resume Next
    fileName = Graph.fileName
    On Error GoTo 0
    
    GetFilenameWithoutError = fileName
End Function

'@Description "_(Description-VisualizeRange)"
Private Function VisualizeRange(ByVal targetRange As Range) As Range
Attribute VisualizeRange.VB_Description = "_(Description-VisualizeRange)"
    targetRange.Hidden = False
    
    Set VisualizeRange = targetRange
End Function

'@Description "_(Description-DeleteHiddenRanges)"
Private Function DeleteHiddenRanges(ByVal HiddenRanges As Collection) As Collection
Attribute DeleteHiddenRanges.VB_Description = "_(Description-DeleteHiddenRanges)"
    Dim deletedRangeInfos As Collection: Set deletedRangeInfos = New Collection
    
    Dim hiddenRangesAsArray() As Range
    ReDim hiddenRangesAsArray(1 To HiddenRanges.Count)
    Dim rangeAddresses() As String
    ReDim rangeAddresses(1 To HiddenRanges.Count)
    
    Dim i As Long
    Dim range_ As Range
    For i = 1 To HiddenRanges.Count
        Set range_ = HiddenRanges(i)
        Set hiddenRangesAsArray(i) = range_
        rangeAddresses(i) = range_.Address(External:=True)
    Next
    
    Dim deletedRangeInfo As Object
    For i = 1 To UBound(hiddenRangesAsArray)
        Set deletedRangeInfo = CreateObject("Scripting.Dictionary")
        Set range_ = hiddenRangesAsArray(i)
        deletedRangeInfo("Type") = TypeName(range_)
        deletedRangeInfo("Location") = rangeAddresses(i)
        range_.Hidden = False
        Call range_.Delete
        Call deletedRangeInfos.Add(deletedRangeInfo)
    Next
    
    Set DeleteHiddenRanges = deletedRangeInfos
End Function

'' == Data Model ==
'@Description "_(Description-ListDataModels)"
Public Function ListDataModels() As Collection
Attribute ListDataModels.VB_Description = "_(Description-ListDataModels)"
    Const ProcedureName As String = "ListDataModels"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedDataModels As Collection: Set listedDataModels = New Collection
    
    Dim connection As WorkbookConnection
    Dim table As ModelTable
    For Each connection In Target.Connections
        If connection.Name <> "ThisWorkbookDataModel" And connection.InModel Then
            For Each table In connection.ModelTables
                 If table.RecordCount > 0 Then
                    Call listedDataModels.Add(connection)
                 End If
            Next
        End If
    Next
    
    Set ListDataModels = listedDataModels
End Function

'@Description "_(Description-RemoveDataModels)"
Public Sub RemoveDataModels()
Attribute RemoveDataModels.VB_Description = "_(Description-RemoveDataModels)"
    Const ProcedureName As String = "RemoveDataModels"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Call Target.RemoveDocumentInformation(xlRDIExcelDataModel)
End Sub

'' == PivotTables, PivotCharts, Cube Formulas, Slicers, and Timelines ==
'' === PivotTables ===
'@Description "_(Description-ListPivotTablesInWorksheet_)"
Private Function ListPivotTablesInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListPivotTablesInWorksheet_.VB_Description = "_(Description-ListPivotTablesInWorksheet_)"
    Dim listedPivotTables As Collection: Set listedPivotTables = New Collection
    
    Dim pivotTable_ As PivotTable
    For Each pivotTable_ In TargetWorksheet.PivotTables
        Call listedPivotTables.Add(pivotTable_)
    Next
    
    Set ListPivotTablesInWorksheet_ = listedPivotTables
End Function

'@Description "_(Description-ListPivotTablesInWorksheet)"
Public Function ListPivotTablesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListPivotTablesInWorksheet.VB_Description = "_(Description-ListPivotTablesInWorksheet)"
    Const ProcedureName As String = "ListPivotTablesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListPivotTablesInWorksheet = ListPivotTablesInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListPivotTables_)"
Private Function ListPivotTables_() As Collection
Attribute ListPivotTables_.VB_Description = "_(Description-ListPivotTables_)"
    Const ProcedureName As String = "ListPivotTables_"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedPivotTables As Collection: Set listedPivotTables = New Collection
    
    Dim worksheet_ As Worksheet
    Dim pivotTable_ As PivotTable
    For Each worksheet_ In Target.Worksheets
        For Each pivotTable_ In ListPivotTablesInWorksheet_(worksheet_)
            Call listedPivotTables.Add(pivotTable_)
        Next
    Next
    
    Set ListPivotTables_ = listedPivotTables
End Function

'@Description "_(Description-ListPivotTables)"
Public Function ListPivotTables() As Collection
Attribute ListPivotTables.VB_Description = "_(Description-ListPivotTables)"
    Const ProcedureName As String = "ListPivotTables"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListPivotTables = ListPivotTables_
End Function

'@Description "_(Description-DeletePivotTablesInWorksheet)"
Public Function DeletePivotTablesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeletePivotTablesInWorksheet.VB_Description = "_(Description-DeletePivotTablesInWorksheet)"
    Const ProcedureName As String = "DeletePivotTablesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedPivotTableInfos As Collection: Set deletedPivotTableInfos = New Collection
    
    Dim pivotTable_ As PivotTable
    For Each pivotTable_ In ListPivotTablesInWorksheet_(TargetWorksheet)
        Call deletedPivotTableInfos.Add(DeletePivotTable(pivotTable_))
    Next
    
    Set DeletePivotTablesInWorksheet = deletedPivotTableInfos
End Function

'@Description "_(Description-DeletePivotTables)"
Public Function DeletePivotTables() As Collection
Attribute DeletePivotTables.VB_Description = "_(Description-DeletePivotTables)"
    Const ProcedureName As String = "DeletePivotTables"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedPivotTableInfos As Collection: Set deletedPivotTableInfos = New Collection
    
    Dim pivotTable_ As PivotTable
    For Each pivotTable_ In ListPivotTables_
        Call deletedPivotTableInfos.Add(DeletePivotTable(pivotTable_))
    Next
    
    Set DeletePivotTables = deletedPivotTableInfos
End Function

'' === PivotCharts ===
'@Description "_(Description-ListPivotChartsInWorksheet_)"
Private Function ListPivotChartsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListPivotChartsInWorksheet_.VB_Description = "_(Description-ListPivotChartsInWorksheet_)"
    Dim listedPivotCharts As Collection: Set listedPivotCharts = New Collection
    
    Dim chartObject_ As ChartObject
    For Each chartObject_ In TargetWorksheet.ChartObjects
        If chartObject_.chart.HasPivotFields Then
            Call listedPivotCharts.Add(chartObject_)
        End If
    Next
    
    Set ListPivotChartsInWorksheet_ = listedPivotCharts
End Function

'@Description "_(Description-ListPivotChartsInWorksheet)"
Public Function ListPivotChartsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListPivotChartsInWorksheet.VB_Description = "_(Description-ListPivotChartsInWorksheet)"
    Const ProcedureName As String = "ListPivotChartsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListPivotChartsInWorksheet = ListPivotChartsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListPivotCharts_)"
Private Function ListPivotCharts_() As Collection
Attribute ListPivotCharts_.VB_Description = "_(Description-ListPivotCharts_)"
    Dim listedPivotCharts As Collection: Set listedPivotCharts = New Collection
    
    Dim worksheet_ As Worksheet
    Dim chartObject_ As ChartObject
    For Each worksheet_ In Target.Worksheets
        For Each chartObject_ In ListPivotChartsInWorksheet_(worksheet_)
            Call listedPivotCharts.Add(chartObject_)
        Next
    Next
    
    Set ListPivotCharts_ = listedPivotCharts
End Function

'@Description "_(Description-ListPivotCharts)"
Public Function ListPivotCharts() As Collection
Attribute ListPivotCharts.VB_Description = "_(Description-ListPivotCharts)"
    Const ProcedureName As String = "ListPivotCharts"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListPivotCharts = ListPivotCharts_
End Function

'@Description "_(Description-DeletePivotChartsInWorksheet)"
Public Function DeletePivotChartsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeletePivotChartsInWorksheet.VB_Description = "_(Description-DeletePivotChartsInWorksheet)"
    Const ProcedureName As String = "DeletePivotChartsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedPivotChartInfos As Collection: Set deletedPivotChartInfos = New Collection
    
    Dim chartObject_ As ChartObject
    For Each chartObject_ In ListPivotChartsInWorksheet_(TargetWorksheet)
        If chartObject_.chart.HasPivotFields Then
            Call deletedPivotChartInfos.Add(DeleteChartObject(chartObject_))
        End If
    Next
    
    Set DeletePivotChartsInWorksheet = deletedPivotChartInfos
End Function

'@Description "_(Description-DeletePivotCharts)"
Public Function DeletePivotCharts() As Collection
Attribute DeletePivotCharts.VB_Description = "_(Description-DeletePivotCharts)"
    Const ProcedureName As String = "DeletePivotCharts"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedPivotChartInfos As Collection: Set deletedPivotChartInfos = New Collection
    
    Dim chartObject_ As ChartObject
    For Each chartObject_ In ListPivotCharts_
        If chartObject_.chart.HasPivotFields Then
            Call deletedPivotChartInfos.Add(DeleteChartObject(chartObject_))
        End If
    Next
    
    Set DeletePivotCharts = deletedPivotChartInfos
End Function

'' === Cube Formulas ===
'@Description "_(Description-ListCubeFormulasInWorksheet_)"
Private Function ListCubeFormulasInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCubeFormulasInWorksheet_.VB_Description = "_(Description-ListCubeFormulasInWorksheet_)"
    Dim listedCubeFormulas As Collection: Set listedCubeFormulas = New Collection
    
    Dim FunctionName As Variant
    Dim cell_ As Range
    For Each FunctionName In Array( _
        "CUBEKPIMEMBER", _
        "CUBEMEMBER", _
        "CUBEMEMBERPROPERTY", _
        "CUBERANKEDMEMBER", _
        "CUBESET", _
        "CUBESETCOUNT", _
        "CUBEVALUE" _
    )
        For Each cell_ In FindFunctionInWorksheet(TargetWorksheet, CStr(FunctionName))
            Call listedCubeFormulas.Add(cell_)
        Next
    Next
    
    Set ListCubeFormulasInWorksheet_ = listedCubeFormulas
End Function

'@Description "_(Description-ListCubeFormulasInWorksheet)"
Public Function ListCubeFormulasInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCubeFormulasInWorksheet.VB_Description = "_(Description-ListCubeFormulasInWorksheet)"
    Const ProcedureName As String = "ListCubeFormulasInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListCubeFormulasInWorksheet = ListCubeFormulasInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListCubeFormulas_)"
Private Function ListCubeFormulas_() As Collection
Attribute ListCubeFormulas_.VB_Description = "_(Description-ListCubeFormulas_)"
    Dim listedCubeFormulas As Collection: Set listedCubeFormulas = New Collection
    
    Dim worksheet_ As Worksheet
    Dim cell_ As Range
    For Each worksheet_ In Target.Worksheets
        For Each cell_ In ListCubeFormulasInWorksheet_(worksheet_)
            Call listedCubeFormulas.Add(cell_)
        Next
    Next
    
    Set ListCubeFormulas_ = listedCubeFormulas
End Function

'@Description "_(Description-ListCubeFormulas)"
Public Function ListCubeFormulas() As Collection
Attribute ListCubeFormulas.VB_Description = "_(Description-ListCubeFormulas)"
    Const ProcedureName As String = "ListCubeFormulas"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListCubeFormulas = ListCubeFormulas_
End Function

'' === Slicers ===
'@Description "_(Description-ListSlicers_)"
Private Function ListSlicers_() As Collection
Attribute ListSlicers_.VB_Description = "_(Description-ListSlicers_)"
    Dim listedSlicers As Collection: Set listedSlicers = New Collection
    
    Dim cache As SlicerCache
    For Each cache In Target.SlicerCaches
        If cache.SlicerCacheType = xlSlicer Then
            Call listedSlicers.Add(cache)
        End If
    Next
    
    Set ListSlicers_ = listedSlicers
End Function

'@Description "_(Description-ListSlicers)"
Public Function ListSlicers() As Collection
Attribute ListSlicers.VB_Description = "_(Description-ListSlicers)"
    Const ProcedureName As String = "ListSlicers"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListSlicers = ListSlicers_
End Function

'@Description "_(Description-DeleteSlicers)"
Public Function DeleteSlicers() As Collection
Attribute DeleteSlicers.VB_Description = "_(Description-DeleteSlicers)"
    Const ProcedureName As String = "DeleteSlicers"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedSlicerInfos As Collection: Set deletedSlicerInfos = New Collection
    
    Dim cache As SlicerCache
    For Each cache In ListSlicers_
        If cache.SlicerCacheType = xlSlicer Then
            Call deletedSlicerInfos.Add(DeleteSlicerCache(cache))
        End If
    Next
    
    Set DeleteSlicers = deletedSlicerInfos
End Function

'' === Timelines (cache) ===
'@Description "_(Description-ListTimelines_)"
Private Function ListTimelines_() As Collection
Attribute ListTimelines_.VB_Description = "_(Description-ListTimelines_)"
    Dim listedTimelines As Collection: Set listedTimelines = New Collection
    
    Dim cache As SlicerCache
    For Each cache In Target.SlicerCaches
        If cache.SlicerCacheType = xlTimeline Then
            Call listedTimelines.Add(cache)
        End If
    Next
    
    Set ListTimelines_ = listedTimelines
End Function

'@Description "_(Description-ListTimelines)"
Public Function ListTimelines() As Collection
Attribute ListTimelines.VB_Description = "_(Description-ListTimelines)"
    Const ProcedureName As String = "ListTimelines"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListTimelines = ListTimelines_
End Function

'@Description "_(Description-DeleteTimelines)"
Public Function DeleteTimelines() As Collection
Attribute DeleteTimelines.VB_Description = "_(Description-DeleteTimelines)"
    Const ProcedureName As String = "DeleteTimelines"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedTimelineInfos As Collection: Set deletedTimelineInfos = New Collection
    
    Dim cache As SlicerCache
    For Each cache In ListTimelines_
        Call deletedTimelineInfos.Add(DeleteSlicerCache(cache))
    Next
    
    Set DeleteTimelines = deletedTimelineInfos
End Function

'' == Links to Other Files ==
'@Description "_(Description-ListLinksToOtherFiles_)"
Private Function ListLinksToOtherFiles_() As Collection
Attribute ListLinksToOtherFiles_.VB_Description = "_(Description-ListLinksToOtherFiles_)"
    Dim listedExternalLink As Collection: Set listedExternalLink = New Collection
    
    Dim links As Variant
    Dim link As Variant
    links = Target.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For Each link In links
            Call listedExternalLink.Add(CStr(link))
        Next
    End If
    
    Set ListLinksToOtherFiles_ = listedExternalLink
End Function

'@Description "_(Description-ListLinksToOtherFiles)"
Public Function ListLinksToOtherFiles() As Collection
Attribute ListLinksToOtherFiles.VB_Description = "_(Description-ListLinksToOtherFiles)"
    Const ProcedureName As String = "ListLinksToOtherFiles"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListLinksToOtherFiles = ListLinksToOtherFiles_
End Function

'@Description "_(Description-ListExternalLinkCellsInWorksheet_)"
Private Function ListExternalLinkCellsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListExternalLinkCellsInWorksheet_.VB_Description = "_(Description-ListExternalLinkCellsInWorksheet_)"
    Dim listedExteralLinkCells As Collection: Set listedExteralLinkCells = New Collection
    
    Dim link As Variant
    Dim regex As Object
    Dim cell_ As Range
    For Each link In ListLinksToOtherFiles_
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .pattern = "[" & link & "]"
            .IgnoreCase = True
            .Global = True
        End With
        
        For Each cell_ In FindCellsInWorksheet( _
            TargetWorksheet:=TargetWorksheet, _
            What:=link _
        )
            If regex.Test(cell_.formula) Then
                Call listedExteralLinkCells.Add(cell_)
            End If
        Next
    Next
    
    Set ListExternalLinkCellsInWorksheet_ = listedExteralLinkCells
End Function

'@Description "_(Description-ListExternalLinkCellsInWorksheet)"
Public Function ListExternalLinkCellsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListExternalLinkCellsInWorksheet.VB_Description = "_(Description-ListExternalLinkCellsInWorksheet)"
    Const ProcedureName As String = "ListExternalLinkCellsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListExternalLinkCellsInWorksheet = ListExternalLinkCellsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListExternalLinkCells)"
Public Function ListExternalLinkCells() As Collection
Attribute ListExternalLinkCells.VB_Description = "_(Description-ListExternalLinkCells)"
    Const ProcedureName As String = "ListExternalLinkCells"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim listedExteralLinkCells As Collection: Set listedExteralLinkCells = New Collection
        
    Dim worksheet_ As Worksheet
    Dim cell_ As Range
    For Each worksheet_ In Target.Worksheets
        For Each cell_ In ListExternalLinkCellsInWorksheet_(worksheet_)
            Call listedExteralLinkCells.Add(cell_)
        Next
    Next
    
    Set ListExternalLinkCells = listedExteralLinkCells
End Function

'@Description "_(Description-BreakLinksToOtherFiles)"
Public Function BreakLinksToOtherFiles() As Collection
Attribute BreakLinksToOtherFiles.VB_Description = "_(Description-BreakLinksToOtherFiles)"
    Const ProcedureName As String = "BreakLinksToOtherFiles"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim brokenLinksToOtherFiles As Collection: Set brokenLinksToOtherFiles = New Collection
    
    Dim link As Variant
    For Each link In ListLinksToOtherFiles
        Call Target.BreakLink(link, xlLinkTypeExcelLinks)
        Call brokenLinksToOtherFiles.Add(link)
    Next
    
    Set BreakLinksToOtherFiles = brokenLinksToOtherFiles
End Function

'' == Real Time Data Functions ==
'@Description "_(Description-ListRealTimeDataFunctionsInWorksheet_)"
Private Function ListRealTimeDataFunctionsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListRealTimeDataFunctionsInWorksheet_.VB_Description = "_(Description-ListRealTimeDataFunctionsInWorksheet_)"
    Dim listedRealTimeDataFunctions As Collection: Set listedRealTimeDataFunctions = New Collection
    
    Dim cell_ As Range
    For Each cell_ In FindFunctionInWorksheet(TargetWorksheet, "RTD")
        Call listedRealTimeDataFunctions.Add(cell_)
    Next
    
    Set ListRealTimeDataFunctionsInWorksheet_ = listedRealTimeDataFunctions
End Function

'@Description "_(Description-ListRealTimeDataFunctionsInWorksheet)"
Public Function ListRealTimeDataFunctionsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListRealTimeDataFunctionsInWorksheet.VB_Description = "_(Description-ListRealTimeDataFunctionsInWorksheet)"
    Const ProcedureName As String = "ListRealTimeDataFunctionsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListRealTimeDataFunctionsInWorksheet = ListRealTimeDataFunctionsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListRealTimeDataFunctions_)"
Private Function ListRealTimeDataFunctions_() As Collection
Attribute ListRealTimeDataFunctions_.VB_Description = "_(Description-ListRealTimeDataFunctions_)"
    Dim listedRealTimeDataFunctions As Collection: Set listedRealTimeDataFunctions = New Collection
    
    Dim worksheet_ As Worksheet
    Dim cell_ As Range
    For Each worksheet_ In Target.Worksheets
        For Each cell_ In ListRealTimeDataFunctionsInWorksheet_(worksheet_)
            Call listedRealTimeDataFunctions.Add(cell_)
        Next
    Next
    
    Set ListRealTimeDataFunctions_ = listedRealTimeDataFunctions
End Function

'@Description "_(Description-ListRealTimeDataFunctions)"
Public Function ListRealTimeDataFunctions() As Collection
Attribute ListRealTimeDataFunctions.VB_Description = "_(Description-ListRealTimeDataFunctions)"
    Const ProcedureName As String = "ListRealTimeDataFunctions"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListRealTimeDataFunctions = ListRealTimeDataFunctions_
End Function

'' == Excel Surveys ==

'' == Defined Scenarios ==
'@Description "_(Description-ListDefinedScenariosInWorksheet_)"
Private Function ListDefinedScenariosInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListDefinedScenariosInWorksheet_.VB_Description = "_(Description-ListDefinedScenariosInWorksheet_)"
    Dim listedDefinedScenarios As Collection: Set listedDefinedScenarios = New Collection
    
    Dim scenario_ As Scenario
    For Each scenario_ In TargetWorksheet.Scenarios
        Call listedDefinedScenarios.Add(scenario_)
    Next
    
    Set ListDefinedScenariosInWorksheet_ = listedDefinedScenarios
End Function

'@Description "_(Description-ListDefinedScenariosInWorksheet)"
Public Function ListDefinedScenariosInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListDefinedScenariosInWorksheet.VB_Description = "_(Description-ListDefinedScenariosInWorksheet)"
    Const ProcedureName As String = "ListDefinedScenariosInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListDefinedScenariosInWorksheet = ListDefinedScenariosInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListDefinedScenarios_)"
Private Function ListDefinedScenarios_() As Collection
Attribute ListDefinedScenarios_.VB_Description = "_(Description-ListDefinedScenarios_)"
    Dim listedDefinedScenarios As Collection: Set listedDefinedScenarios = New Collection
    
    Dim worksheet_ As Worksheet
    Dim scenario_ As Scenario
    For Each worksheet_ In Target.Worksheets
        For Each scenario_ In ListDefinedScenariosInWorksheet_(worksheet_)
            Call listedDefinedScenarios.Add(scenario_)
        Next
    Next
    
    Set ListDefinedScenarios_ = listedDefinedScenarios
End Function

'@Description "_(Description-ListDefinedScenarios)"
Public Function ListDefinedScenarios() As Collection
Attribute ListDefinedScenarios.VB_Description = "_(Description-ListDefinedScenarios)"
    Const ProcedureName As String = "ListDefinedScenarios"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListDefinedScenarios = ListDefinedScenarios_
End Function

'@Description "_(Description-DeleteDefinedScenariosInWorksheet)"
Public Function DeleteDefinedScenariosInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteDefinedScenariosInWorksheet.VB_Description = "_(Description-DeleteDefinedScenariosInWorksheet)"
    Const ProcedureName As String = "DeleteDefinedScenariosInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedDefinedScenarioInfos As Collection: Set deletedDefinedScenarioInfos = New Collection
    
    Dim scenario_ As Scenario
    For Each scenario_ In ListDefinedScenariosInWorksheet_(TargetWorksheet)
        Call deletedDefinedScenarioInfos.Add(DeleteScenario(scenario_))
    Next
    
    Set DeleteDefinedScenariosInWorksheet = deletedDefinedScenarioInfos
End Function

'@Description "_(Description-DeleteDefinedScenarios)"
Public Function DeleteDefinedScenarios() As Collection
Attribute DeleteDefinedScenarios.VB_Description = "_(Description-DeleteDefinedScenarios)"
    Const ProcedureName As String = "DeleteDefinedScenarios"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedDefinedScenarioInfos As Collection: Set deletedDefinedScenarioInfos = New Collection
    
    Dim scenario_ As Scenario
    ' For some reason, it fails.
    'For Each scenario_ In ListDefinedScenarios_
    '    Call deletedDefinedScenarioInfos.Add(DeleteScenario(scenario_))
    'Next
    ' Alternative
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        For Each scenario_ In ListDefinedScenariosInWorksheet_(worksheet_)
            Call deletedDefinedScenarioInfos.Add(DeleteScenario(scenario_))
        Next
    Next
    ' End Alternative
    
    Set DeleteDefinedScenarios = deletedDefinedScenarioInfos
End Function

'' == Active Filters ==
'@Description "_(Description-ListActiveFilters_)"
Private Function ListActiveFilters_() As Collection
Attribute ListActiveFilters_.VB_Description = "_(Description-ListActiveFilters_)"
    Dim listedActiveFilters As Collection: Set listedActiveFilters = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        If Not worksheet_.AutoFilter Is Nothing Then
            If worksheet_.AutoFilter.FilterMode Then
                Call listedActiveFilters.Add(worksheet_.AutoFilter)
            End If
        End If
    Next
    
    Set ListActiveFilters_ = listedActiveFilters
End Function

'@Description "_(Description-ListActiveFilters)"
Public Function ListActiveFilters() As Collection
Attribute ListActiveFilters.VB_Description = "_(Description-ListActiveFilters)"
    Const ProcedureName As String = "ListActiveFilters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListActiveFilters = ListActiveFilters_
End Function

'@Description "_(Description-InactivateActiveFilters)"
Public Function InactivateActiveFilters() As Collection
Attribute InactivateActiveFilters.VB_Description = "_(Description-InactivateActiveFilters)"
    Const ProcedureName As String = "InactivateActiveFilters"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim inactivatedActiveFilters As Collection: Set inactivatedActiveFilters = New Collection
    
    Dim activeFilter As AutoFilter
    For Each activeFilter In ListActiveFilters_
        Call activeFilter.ShowAllData
        Call inactivatedActiveFilters.Add(activeFilter)
    Next
    
    Set InactivateActiveFilters = inactivatedActiveFilters
End Function

'' == Custom Worksheet Properties ==
'@Description "_(Description-ListCustomWorksheetPropertiesInWorksheet_)"
Private Function ListCustomWorksheetPropertiesInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCustomWorksheetPropertiesInWorksheet_.VB_Description = "_(Description-ListCustomWorksheetPropertiesInWorksheet_)"
    Dim listedCustomWorksheetProperties As Collection: Set listedCustomWorksheetProperties = New Collection
    
    Dim property As CustomProperty
    For Each property In TargetWorksheet.CustomProperties
        Call listedCustomWorksheetProperties.Add(property)
    Next
    
    Set ListCustomWorksheetPropertiesInWorksheet_ = listedCustomWorksheetProperties
End Function

'@Description "_(Description-ListCustomWorksheetPropertiesInWorksheet)"
Public Function ListCustomWorksheetPropertiesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListCustomWorksheetPropertiesInWorksheet.VB_Description = "_(Description-ListCustomWorksheetPropertiesInWorksheet)"
    Const ProcedureName As String = "ListCustomWorksheetPropertiesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListCustomWorksheetPropertiesInWorksheet = ListCustomWorksheetPropertiesInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListCustomWorksheetProperties_)"
Private Function ListCustomWorksheetProperties_() As Collection
Attribute ListCustomWorksheetProperties_.VB_Description = "_(Description-ListCustomWorksheetProperties_)"
    Dim listedCustomWorksheetProperties As Collection: Set listedCustomWorksheetProperties = New Collection
    
    Dim worksheet_ As Worksheet
    Dim property As CustomProperty
    For Each worksheet_ In Target.Worksheets
        For Each property In ListCustomWorksheetPropertiesInWorksheet_(worksheet_)
            Call listedCustomWorksheetProperties.Add(property)
        Next
    Next
    
    Set ListCustomWorksheetProperties_ = listedCustomWorksheetProperties
End Function

'@Description "_(Description-ListCustomWorksheetProperties)"
Public Function ListCustomWorksheetProperties() As Collection
Attribute ListCustomWorksheetProperties.VB_Description = "_(Description-ListCustomWorksheetProperties)"
    Const ProcedureName As String = "ListCustomWorksheetProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListCustomWorksheetProperties = ListCustomWorksheetProperties_
End Function

'@Description "_(Description-DeleteCustomWorksheetPropertiesInWorksheet)"
Public Function DeleteCustomWorksheetPropertiesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteCustomWorksheetPropertiesInWorksheet.VB_Description = "_(Description-DeleteCustomWorksheetPropertiesInWorksheet)"
    Const ProcedureName As String = "DeleteCustomWorksheetPropertiesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletedCustomWorksheetProperyInfos As Collection: Set deletedCustomWorksheetProperyInfos = New Collection
    
    Dim property As CustomProperty
    For Each property In ListCustomWorksheetPropertiesInWorksheet_(TargetWorksheet)
        Call deletedCustomWorksheetProperyInfos.Add(DeleteCustomProperty(property))
    Next
    
    Set DeleteCustomWorksheetPropertiesInWorksheet = deletedCustomWorksheetProperyInfos
End Function
    
'@Description "_(Description-DeleteCustomWorksheetProperties)"
Public Function DeleteCustomWorksheetProperties() As Collection
Attribute DeleteCustomWorksheetProperties.VB_Description = "_(Description-DeleteCustomWorksheetProperties)"
    Const ProcedureName As String = "DeleteCustomWorksheetProperties"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedCustomWorksheetProperyInfos As Collection: Set deletedCustomWorksheetProperyInfos = New Collection
    
    Dim property As CustomProperty
    For Each property In ListCustomWorksheetProperties_
        Call deletedCustomWorksheetProperyInfos.Add(DeleteCustomProperty(property))
    Next
    
    Set DeleteCustomWorksheetProperties = deletedCustomWorksheetProperyInfos
End Function

'' == Hidden Names ==
'@Description "_(Description-ListHiddenNamesInWorksheet_)"
Private Function ListHiddenNamesInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenNamesInWorksheet_.VB_Description = "_(Description-ListHiddenNamesInWorksheet_)"
    Dim listedHiddenNames As Collection: Set listedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In TargetWorksheet.names
        If Not name_.Visible Then
            Call listedHiddenNames.Add(name_)
        End If
    Next
    
    Set ListHiddenNamesInWorksheet_ = listedHiddenNames
End Function

'@Description "_(Description-ListHiddenNamesInWorksheet)"
Public Function ListHiddenNamesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenNamesInWorksheet.VB_Description = "_(Description-ListHiddenNamesInWorksheet)"
    Const ProcedureName As String = "ListHiddenNamesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListHiddenNamesInWorksheet = ListHiddenNamesInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListHiddenNamesInWorkbook_)"
Private Function ListHiddenNamesInWorkbook_() As Collection
Attribute ListHiddenNamesInWorkbook_.VB_Description = "_(Description-ListHiddenNamesInWorkbook_)"
    Dim listedHiddenNames As Collection: Set listedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In Target.names
        If TypeOf name_.Parent Is Workbook Then
            If Not name_.Visible Then
                Call listedHiddenNames.Add(name_)
            End If
        End If
    Next
    
    Set ListHiddenNamesInWorkbook_ = listedHiddenNames
End Function

'@Description "_(Description-ListHiddenNamesInWorkbook)"
Public Function ListHiddenNamesInWorkbook() As Collection
Attribute ListHiddenNamesInWorkbook.VB_Description = "_(Description-ListHiddenNamesInWorkbook)"
    Const ProcedureName As String = "ListHiddenNamesInWorkbook"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenNamesInWorkbook = ListHiddenNamesInWorkbook_
End Function

'@Description "_(Description-ListHiddenNames_)"
Private Function ListHiddenNames_() As Collection
Attribute ListHiddenNames_.VB_Description = "_(Description-ListHiddenNames_)"
    Dim listedHiddenNames As Collection: Set listedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNamesInWorkbook_
        Call listedHiddenNames.Add(name_)
    Next
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        For Each name_ In ListHiddenNamesInWorksheet_(worksheet_)
            Call listedHiddenNames.Add(name_)
        Next
    Next
    
    Set ListHiddenNames_ = listedHiddenNames
End Function

'@Description "_(Description-ListHiddenNames)"
Public Function ListHiddenNames() As Collection
Attribute ListHiddenNames.VB_Description = "_(Description-ListHiddenNames)"
    Const ProcedureName As String = "ListHiddenNames"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenNames = ListHiddenNames_
End Function

'@Description "_(Description-VisualizeHiddenNamesInWorksheet)"
Public Function VisualizeHiddenNamesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeHiddenNamesInWorksheet.VB_Description = "_(Description-VisualizeHiddenNamesInWorksheet)"
    Const ProcedureName As String = "VisualizeHiddenNamesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim visualizedHiddenNames As Collection: Set visualizedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNamesInWorksheet_(TargetWorksheet)
        Call visualizedHiddenNames.Add(VisualizeName(name_))
    Next
    
    Set VisualizeHiddenNamesInWorksheet = visualizedHiddenNames
End Function

'@Description "_(Description-VisualizeHiddenNamesInWorkbook)"
Public Function VisualizeHiddenNamesInWorkbook() As Collection
Attribute VisualizeHiddenNamesInWorkbook.VB_Description = "_(Description-VisualizeHiddenNamesInWorkbook)"
    Const ProcedureName As String = "VisualizeHiddenNamesInWorkbook"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedHiddenNames As Collection: Set visualizedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNamesInWorkbook_()
        Call visualizedHiddenNames.Add(VisualizeName(name_))
    Next
    
    Set VisualizeHiddenNamesInWorkbook = visualizedHiddenNames
End Function

'@Description "_(Description-VisualizeHiddenNames)"
Public Function VisualizeHiddenNames() As Collection
Attribute VisualizeHiddenNames.VB_Description = "_(Description-VisualizeHiddenNames)"
    Const ProcedureName As String = "VisualizeHiddenNames"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedHiddenNames As Collection: Set visualizedHiddenNames = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNames_()
        Call visualizedHiddenNames.Add(VisualizeName(name_))
    Next
    
    Set VisualizeHiddenNames = visualizedHiddenNames
End Function

'@Description "_(Description-DeleteHiddenNamesInWorksheet)"
Public Function DeleteHiddenNamesInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteHiddenNamesInWorksheet.VB_Description = "_(Description-DeleteHiddenNamesInWorksheet)"
    Const ProcedureName As String = "DeleteHiddenNamesInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim deletededHiddenNameInfos As Collection: Set deletededHiddenNameInfos = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNamesInWorksheet_(TargetWorksheet)
        Call deletededHiddenNameInfos.Add(DeleteName(name_))
    Next
    
    Set DeleteHiddenNamesInWorksheet = deletededHiddenNameInfos
End Function

'@Description "_(Description-DeleteHiddenNamesInWorkbook)"
Public Function DeleteHiddenNamesInWorkbook() As Collection
Attribute DeleteHiddenNamesInWorkbook.VB_Description = "_(Description-DeleteHiddenNamesInWorkbook)"
    Const ProcedureName As String = "DeleteHiddenNamesInWorkbook"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletededHiddenNameInfos As Collection: Set deletededHiddenNameInfos = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNamesInWorkbook_()
        Call deletededHiddenNameInfos.Add(DeleteName(name_))
    Next
    
    Set DeleteHiddenNamesInWorkbook = deletededHiddenNameInfos
End Function

'@Description "_(Description-DeleteHiddenNames)"
Public Function DeleteHiddenNames() As Collection
Attribute DeleteHiddenNames.VB_Description = "_(Description-DeleteHiddenNames)"
    Const ProcedureName As String = "DeleteHiddenNames"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletededNameInfos As Collection: Set deletededNameInfos = New Collection
    
    Dim name_ As Name
    For Each name_ In ListHiddenNames_()
        Call deletededNameInfos.Add(DeleteName(name_))
    Next
    
    Set DeleteHiddenNames = deletededNameInfos
End Function

'' == Long External References ==

'' == Hidden Rows and Columns ==
'' === Hidden Rows and Columns ===
'@Description "_(Description-ListHiddenRowsAndColumnsInWorksheet_)"
Private Function ListHiddenRowsAndColumnsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenRowsAndColumnsInWorksheet_.VB_Description = "_(Description-ListHiddenRowsAndColumnsInWorksheet_)"
    Dim listedHiddenRowsAndColumns As Collection: Set listedHiddenRowsAndColumns = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRowsInWorksheet_(TargetWorksheet)
        Call listedHiddenRowsAndColumns.Add(range_)
    Next
    For Each range_ In ListHiddenColumnsInWorksheet_(TargetWorksheet)
        Call listedHiddenRowsAndColumns.Add(range_)
    Next
    
    Set ListHiddenRowsAndColumnsInWorksheet_ = listedHiddenRowsAndColumns
End Function

'@Description "_(Description-ListHiddenRowsAndColumnsInWorksheet)"
Public Function ListHiddenRowsAndColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenRowsAndColumnsInWorksheet.VB_Description = "_(Description-ListHiddenRowsAndColumnsInWorksheet)"
    Const ProcedureName As String = "ListHiddenRowsAndColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListHiddenRowsAndColumnsInWorksheet = ListHiddenRowsAndColumnsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListHiddenRowsAndColumns_)"
Private Function ListHiddenRowsAndColumns_() As Collection
Attribute ListHiddenRowsAndColumns_.VB_Description = "_(Description-ListHiddenRowsAndColumns_)"
    Dim listedHiddenRowsAndColumns As Collection: Set listedHiddenRowsAndColumns = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRows_
        Call listedHiddenRowsAndColumns.Add(range_)
    Next
    For Each range_ In ListHiddenColumns_
        Call listedHiddenRowsAndColumns.Add(range_)
    Next
    
    Set ListHiddenRowsAndColumns_ = listedHiddenRowsAndColumns
End Function

'@Description "_(Description-ListHiddenRowsAndColumns)"
Public Function ListHiddenRowsAndColumns() As Collection
Attribute ListHiddenRowsAndColumns.VB_Description = "_(Description-ListHiddenRowsAndColumns)"
    Const ProcedureName As String = "ListHiddenRowsAndColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenRowsAndColumns = ListHiddenRowsAndColumns_
End Function

'@Description "_(Description-VisualizeHiddenRowsAndColumnsInWorksheet)"
Public Function VisualizeHiddenRowsAndColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeHiddenRowsAndColumnsInWorksheet.VB_Description = "_(Description-VisualizeHiddenRowsAndColumnsInWorksheet)"
    Const ProcedureName As String = "VisualizeHiddenRowsAndColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim visualizedRows As Collection: Set visualizedRows = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRowsAndColumnsInWorksheet_(TargetWorksheet)
        Call visualizedRows.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenRowsAndColumnsInWorksheet = visualizedRows
End Function

'@Description "_(Description-VisualizeHiddenRowsAndColumns)"
Public Function VisualizeHiddenRowsAndColumns() As Collection
Attribute VisualizeHiddenRowsAndColumns.VB_Description = "_(Description-VisualizeHiddenRowsAndColumns)"
    Const ProcedureName As String = "VisualizeHiddenRowsAndColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedRows As Collection: Set visualizedRows = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRowsAndColumns_()
        Call visualizedRows.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenRowsAndColumns = visualizedRows
End Function

'@Description "_(Description-DeleteHiddenRowsAndColumnsInWorksheet)"
Public Function DeleteHiddenRowsAndColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteHiddenRowsAndColumnsInWorksheet.VB_Description = "_(Description-DeleteHiddenRowsAndColumnsInWorksheet)"
    Const ProcedureName As String = "DeleteHiddenRowsAndColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set DeleteHiddenRowsAndColumnsInWorksheet = DeleteHiddenRanges(ListHiddenRowsAndColumnsInWorksheet_(TargetWorksheet))
End Function

'@Description "_(Description-DeleteHiddenRowsAndColumns)"
Public Function DeleteHiddenRowsAndColumns() As Collection
Attribute DeleteHiddenRowsAndColumns.VB_Description = "_(Description-DeleteHiddenRowsAndColumns)"
    Const ProcedureName As String = "DeleteHiddenRowsAndColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set DeleteHiddenRowsAndColumns = DeleteHiddenRanges(ListHiddenRowsAndColumns)
End Function

'@Description "_(Description-InspectHiddenRowsAndColumns)"
Public Function InspectHiddenRowsAndColumns(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectHiddenRowsAndColumns.VB_Description = "_(Description-InspectHiddenRowsAndColumns)"
    Const ProcedureName As String = "InspectHiddenRowsAndColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenRowsAndColumns)
    Call inspector.Inspect(status, Results)
    
    InspectHiddenRowsAndColumns = status
End Function

'@Description "_(Description-FixHiddenRowsAndColumns)"
Public Function FixHiddenRowsAndColumns(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixHiddenRowsAndColumns.VB_Description = "_(Description-FixHiddenRowsAndColumns)"
    Const ProcedureName As String = "FixHiddenRowsAndColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenRowsAndColumns)
    Call inspector.Fix(status, Results)
    
    FixHiddenRowsAndColumns = status
End Function

'' === Hidden Rows ===
'@Description "_(Description-ListHiddenRowsInWorksheet_)"
Private Function ListHiddenRowsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenRowsInWorksheet_.VB_Description = "_(Description-ListHiddenRowsInWorksheet_)"
    Dim listedHiddenRows As Collection: Set listedHiddenRows = New Collection
    
    Dim hiddenStartRow As Long
    Dim i As Long
    Dim j As Long
    With TargetWorksheet
        For j = 1 To .Rows.Count
            If .Rows(j).Hidden = True Then
                hiddenStartRow = j
                For i = j To .Rows.Count
                    If .Rows(i).Hidden = False Then
                        Call listedHiddenRows.Add(.Range(.Cells(hiddenStartRow, 1), .Cells(i - 1, 1)).EntireRow)
                        hiddenStartRow = 0
                        Exit For
                    End If
                Next
                j = i
            End If
        Next
        If hiddenStartRow > 0 Then
            Call listedHiddenRows.Add(.Range(.Cells(hiddenStartRow, 1), .Cells(.Rows.Count, 1)).EntireRow)
        End If
    End With
    
    Set ListHiddenRowsInWorksheet_ = listedHiddenRows
End Function

'@Description "_(Description-ListHiddenRowsInWorksheet)"
Public Function ListHiddenRowsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenRowsInWorksheet.VB_Description = "_(Description-ListHiddenRowsInWorksheet)"
    Const ProcedureName As String = "ListHiddenRowsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListHiddenRowsInWorksheet = ListHiddenRowsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListHiddenRows_)"
Private Function ListHiddenRows_() As Collection
Attribute ListHiddenRows_.VB_Description = "_(Description-ListHiddenRows_)"
    Dim listedHiddenRows As Collection: Set listedHiddenRows = New Collection
    
    Dim worksheet_ As Worksheet
    Dim range_ As Range
    For Each worksheet_ In Target.Worksheets
        For Each range_ In ListHiddenRowsInWorksheet_(TargetWorksheet:=worksheet_)
            Call listedHiddenRows.Add(range_)
        Next
    Next
    
    Set ListHiddenRows_ = listedHiddenRows
End Function

'@Description "_(Description-ListHiddenRows)"
Public Function ListHiddenRows() As Collection
Attribute ListHiddenRows.VB_Description = "_(Description-ListHiddenRows)"
    Const ProcedureName As String = "ListHiddenRows"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenRows = ListHiddenRows_
End Function

'@Description "_(Description-VisualizeHiddenRowsInWorksheet)"
Public Function VisualizeHiddenRowsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeHiddenRowsInWorksheet.VB_Description = "_(Description-VisualizeHiddenRowsInWorksheet)"
    Const ProcedureName As String = "VisualizeHiddenRowsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim visualizedRows As Collection: Set visualizedRows = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRowsInWorksheet_(TargetWorksheet)
        Call visualizedRows.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenRowsInWorksheet = visualizedRows
End Function

'@Description "_(Description-VisualizeHiddenRows)"
Public Function VisualizeHiddenRows() As Collection
Attribute VisualizeHiddenRows.VB_Description = "_(Description-VisualizeHiddenRows)"
    Const ProcedureName As String = "VisualizeHiddenRows"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedRows As Collection: Set visualizedRows = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenRows_()
        Call visualizedRows.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenRows = visualizedRows
End Function

'@Description "_(Description-DeleteHiddenRowsInWorksheet)"
Public Function DeleteHiddenRowsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteHiddenRowsInWorksheet.VB_Description = "_(Description-DeleteHiddenRowsInWorksheet)"
    Const ProcedureName As String = "DeleteHiddenRowsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set DeleteHiddenRowsInWorksheet = DeleteHiddenRanges(ListHiddenRowsInWorksheet_(TargetWorksheet))
End Function

'@Description "_(Description-DeleteHiddenRows)"
Public Function DeleteHiddenRows() As Collection
Attribute DeleteHiddenRows.VB_Description = "_(Description-DeleteHiddenRows)"
    Const ProcedureName As String = "DeleteHiddenRows"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set DeleteHiddenRows = DeleteHiddenRanges(ListHiddenRows_)
End Function

'' === Hidden Columns ===
'@Description "_(Description-ListHiddenColumnsInWorksheet_)"
Private Function ListHiddenColumnsInWorksheet_(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenColumnsInWorksheet_.VB_Description = "_(Description-ListHiddenColumnsInWorksheet_)"
    Dim listedHiddenColumns As Collection: Set listedHiddenColumns = New Collection
    
    Dim hiddenStartColumn As Long
    Dim i As Long
    Dim j As Long
    With TargetWorksheet
        For j = 1 To .Columns.Count
            If .Columns(j).Hidden = True Then
                hiddenStartColumn = j
                For i = j To .Columns.Count
                    If .Columns(i).Hidden = False Then
                        Call listedHiddenColumns.Add(.Range(.Cells(1, hiddenStartColumn), .Cells(1, i - 1)).EntireColumn)
                        hiddenStartColumn = 0
                        Exit For
                    End If
                Next
                j = i
            End If
        Next
        If hiddenStartColumn > 0 Then
            Call listedHiddenColumns.Add(.Range(.Cells(1, hiddenStartColumn), .Cells(1, .Columns.Count)).EntireColumn)
        End If
    End With
    
    Set ListHiddenColumnsInWorksheet_ = listedHiddenColumns
End Function

'@Description "_(Description-ListHiddenColumnsInWorksheet)"
Public Function ListHiddenColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute ListHiddenColumnsInWorksheet.VB_Description = "_(Description-ListHiddenColumnsInWorksheet)"
    Const ProcedureName As String = "ListHiddenColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set ListHiddenColumnsInWorksheet = ListHiddenColumnsInWorksheet_(TargetWorksheet)
End Function

'@Description "_(Description-ListHiddenColumns_)"
Private Function ListHiddenColumns_() As Collection
Attribute ListHiddenColumns_.VB_Description = "_(Description-ListHiddenColumns_)"
    Dim listedHiddenColumns As Collection: Set listedHiddenColumns = New Collection
    
    Dim worksheet_ As Worksheet
    Dim range_ As Range
    For Each worksheet_ In Target.Worksheets
        For Each range_ In ListHiddenColumnsInWorksheet_(TargetWorksheet:=worksheet_)
            Call listedHiddenColumns.Add(range_)
        Next
    Next
    
    Set ListHiddenColumns_ = listedHiddenColumns
End Function

'@Description "_(Description-ListHiddenColumns)"
Public Function ListHiddenColumns() As Collection
Attribute ListHiddenColumns.VB_Description = "_(Description-ListHiddenColumns)"
    Const ProcedureName As String = "ListHiddenColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenColumns = ListHiddenColumns_
End Function

'@Description "_(Description-VisualizeHiddenColumnsInWorksheet)"
Public Function VisualizeHiddenColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute VisualizeHiddenColumnsInWorksheet.VB_Description = "_(Description-VisualizeHiddenColumnsInWorksheet)"
    Const ProcedureName As String = "VisualizeHiddenColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Dim visualizedColumns As Collection: Set visualizedColumns = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenColumnsInWorksheet_(TargetWorksheet)
        Call visualizedColumns.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenColumnsInWorksheet = visualizedColumns
End Function

'@Description "_(Description-VisualizeHiddenColumns)"
Public Function VisualizeHiddenColumns() As Collection
Attribute VisualizeHiddenColumns.VB_Description = "_(Description-VisualizeHiddenColumns)"
    Const ProcedureName As String = "VisualizeHiddenColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedColumns As Collection: Set visualizedColumns = New Collection
    
    Dim range_ As Range
    For Each range_ In ListHiddenColumns_()
        Call visualizedColumns.Add(VisualizeRange(range_))
    Next
    
    Set VisualizeHiddenColumns = visualizedColumns
End Function

'@Description "_(Description-DeleteHiddenColumnsInWorksheet)"
Public Function DeleteHiddenColumnsInWorksheet(ByVal TargetWorksheet As Worksheet) As Collection
Attribute DeleteHiddenColumnsInWorksheet.VB_Description = "_(Description-DeleteHiddenColumnsInWorksheet)"
    Const ProcedureName As String = "DeleteHiddenColumnsInWorksheet"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    Call RaiseErrorIfArgumentIsNothing(ProcedureName, TargetWorksheet, "TargetWorksheet")
    Call RaiseErrorIfParentIsNotTargetWorkbook(ProcedureName, TargetWorksheet)
    
    Set DeleteHiddenColumnsInWorksheet = DeleteHiddenRanges(ListHiddenColumnsInWorksheet_(TargetWorksheet))
End Function

'@Description "_(Description-DeleteHiddenColumns)"
Public Function DeleteHiddenColumns() As Collection
Attribute DeleteHiddenColumns.VB_Description = "_(Description-DeleteHiddenColumns)"
    Const ProcedureName As String = "DeleteHiddenColumns"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set DeleteHiddenColumns = DeleteHiddenRanges(ListHiddenColumns_)
End Function

'' == Hidden Worksheets ==
'@Description "_(Description-ListHiddenWorksheets_)"
Private Function ListHiddenWorksheets_() As Collection
Attribute ListHiddenWorksheets_.VB_Description = "_(Description-ListHiddenWorksheets_)"
    Dim listedHiddenWorksheets As Collection: Set listedHiddenWorksheets = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In Target.Worksheets
        If worksheet_.Visible <> xlSheetVisible Then
            Call listedHiddenWorksheets.Add(worksheet_)
        End If
    Next
    
    Set ListHiddenWorksheets_ = listedHiddenWorksheets
End Function

'@Description "_(Description-ListHiddenWorksheets)"
Public Function ListHiddenWorksheets() As Collection
Attribute ListHiddenWorksheets.VB_Description = "_(Description-ListHiddenWorksheets)"
    Const ProcedureName As String = "ListHiddenWorksheets"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Set ListHiddenWorksheets = ListHiddenWorksheets_
End Function

'@Description "_(Description-VisualizeHiddenWorksheets)"
Public Function VisualizeHiddenWorksheets() As Collection
Attribute VisualizeHiddenWorksheets.VB_Description = "_(Description-VisualizeHiddenWorksheets)"
    Const ProcedureName As String = "VisualizeHiddenWorksheets"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim visualizedWorksheets As Collection: Set visualizedWorksheets = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In ListHiddenWorksheets_
        Call visualizedWorksheets.Add(VisualizeWorksheet(worksheet_))
    Next
    
    Set VisualizeHiddenWorksheets = visualizedWorksheets
End Function

'@Description "_(Description-DeleteHiddenWorksheets)"
Public Function DeleteHiddenWorksheets() As Collection
Attribute DeleteHiddenWorksheets.VB_Description = "_(Description-DeleteHiddenWorksheets)"
    Const ProcedureName As String = "DeleteHiddenWorksheets"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim deletedWorksheetInfos As Collection: Set deletedWorksheetInfos = New Collection
    
    Dim worksheet_ As Worksheet
    For Each worksheet_ In ListHiddenWorksheets_
        Call deletedWorksheetInfos.Add(DeleteWorksheet(worksheet_))
    Next
    
    Set DeleteHiddenWorksheets = deletedWorksheetInfos
End Function

'@Description "_(Description-InspectHiddenWorksheets)"
Public Function InspectHiddenWorksheets(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute InspectHiddenWorksheets.VB_Description = "_(Description-InspectHiddenWorksheets)"
    Const ProcedureName As String = "InspectHiddenWorksheets"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenWorkSheets)
    Call inspector.Inspect(status, Results)
    
    InspectHiddenWorksheets = status
End Function

'@Description "_(Description-FixHiddenWorksheets)"
Public Function FixHiddenWorksheets(Optional ByRef Results As String = vbNullString) As MsoDocInspectorStatus
Attribute FixHiddenWorksheets.VB_Description = "_(Description-FixHiddenWorksheets)"
    Const ProcedureName As String = "FixHiddenWorksheets"
    Call RaiseErrorIfNotInitialized(ProcedureName)
    
    Dim status As MsoDocInspectorStatus: status = msoDocInspectorStatusDocOk
    
    Dim inspector As DocumentInspector
    Set inspector = Target.DocumentInspectors(InspectorIndex.HiddenWorkSheets)
    Call inspector.Fix(status, Results)
    
    FixHiddenWorksheets = status
End Function
