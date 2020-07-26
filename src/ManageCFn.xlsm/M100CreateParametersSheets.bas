Attribute VB_Name = "M100CreateParametersSheets"
''
' Create Parameters Sheets
'
' Copyright (c) 2020, Masakazu Kayano
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const CFnJsonPath_Row = 6
Private Const CFnJsonPath_Column = 2

Private Const ResouceFile_Row = 12
Private Const ResouceFile_Column = 2

Private Parameter_Row As Integer

Dim ResouceParameters As Object


Public Sub Create_ParameterSheets()

    Dim Row As Integer
    Dim FilePath As String
    Dim BoxReturn As String
    
    Row = ResouceFile_Row
    
    BoxReturn = MsgBox("リソースの設定が消えてしまます。`" & vbCrLf & "パラメータシートの作成を行いますか?", vbOKCancel)
    If BoxReturn <> vbOK Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    DeleteAllSheetsWithoutSettingSheets

    Do While Worksheets("Resources").Cells(Row, ResouceFile_Column) <> ""
    

        FilePath = Worksheets("Resources").Cells(CFnJsonPath_Row, ResouceFile_Column) & "\" & Worksheets("Resources").Cells(Row, ResouceFile_Column)
        ReadJsonFile FilePath
        
        Set_ResourceSpecificationVersion
        
        CreateSheet Left(GetResouceName(Get_Information(ResouceParameters, "ResourceType")), 31)
        
        Create_ParameterSheet
        
        Row = Row + 1
        
        M200Shaping.Shaping Left(GetResouceName(Get_Information(ResouceParameters, "ResourceType")), 31)
    Loop

    Application.ScreenUpdating = True

End Sub


Private Sub Set_ResourceSpecificationVersion()

    Worksheets("Resources").Cells(ResourceSpecificationVersion_Row, ResourceSpecificationVersion_Column).Value = Get_Information(ResouceParameters, "ResourceSpecificationVersion")

End Sub


Private Sub Create_ParameterSheet()

    Dim MySheet As String
    Dim ResourceType As String

    MySheet = Left(GetResouceName(GetResouceName(Get_Information(ResouceParameters, "ResourceType"))), 31)
    Parameter_Row = Parameter_Start_Row
 
    Set_Parameter MySheet, 1, "Name:", "String", "Resource Name"
    Parameter_Row = Parameter_Row + 1
    
    ResourceType = Get_Information(ResouceParameters, "ResourceType")
    Set_Parameter MySheet, 2, "Type: " & ResourceType, "'", Get_Information(ResouceParameters("ResourceType")(ResourceType), "Documentation")
    Parameter_Row = Parameter_Row + 1
    Set_Parameter MySheet, 2, "Properties: ", "'", "'"
    Parameter_Row = Parameter_Row + 1
    
    Analysis MySheet, ResouceParameters("ResourceType")(ResourceType)("Properties"), 3

End Sub


Private Sub Analysis(ByVal MySheet As String, ByVal TargetObject As Object, Indent As Integer)

    Dim Member
    Dim ResourceType As String
    Dim PropertyTypes As String
    
    ResourceType = Get_Information(ResouceParameters, "ResourceType")
    
    For Each Member In TargetObject
    
        If Get_Information(TargetObject(Member), "PrimitiveType") <> "" Then
        
            Set_Parameter MySheet, Indent, Member & ": ", Get_Information(TargetObject(Member), "PrimitiveType"), Get_Information(TargetObject(Member), "Documentation")
            Parameter_Row = Parameter_Row + 1
            
        Else
        
            If Get_Information(TargetObject(Member), "Type") = "List" Then

            
                If Get_Information(TargetObject(Member), "PrimitiveItemType") <> "" Then
            
                    Set_Parameter MySheet, Indent, Member & ": ", "'", Get_Information(TargetObject(Member), "Documentation")
                    Parameter_Row = Parameter_Row + 1
                    Set_Parameter_list MySheet, Indent + 1, "'", Get_Information(TargetObject(Member), "PrimitiveItemType"), ""
                    Parameter_Row = Parameter_Row + 1
                    Set_Parameter_list MySheet, Indent + 1, "'", Get_Information(TargetObject(Member), "PrimitiveItemType"), ""
                    Parameter_Row = Parameter_Row + 1
            
                Else
                          
                    If Get_Information(TargetObject(Member), "ItemType") <> "" Then
            
                        Set_Parameter MySheet, Indent, Member & ": ", "'", Get_Information(TargetObject(Member), "Documentation")
                        Parameter_Row = Parameter_Row + 1
                        
                        If Get_Information(TargetObject(Member), "ItemType") = "Tag" Then
                        
                            PropertyTypes = Get_Information(TargetObject(Member), "ItemType")
                        
                        Else
                        
                            PropertyTypes = ResourceType & "." & Get_Information(TargetObject(Member), "ItemType")
                            
                        End If
                        
                        Set_list MySheet, Indent + 1
                        Analysis MySheet, ResouceParameters("PropertyTypes")(PropertyTypes)("Properties"), Indent + 2

                        Set_list MySheet, Indent + 1
                        Analysis MySheet, ResouceParameters("PropertyTypes")(PropertyTypes)("Properties"), Indent + 2

                    End If
                    
                End If
                
            Else
                
                Set_Parameter MySheet, Indent, Member & ": ", "'", Get_Information(TargetObject(Member), "Documentation")
                Parameter_Row = Parameter_Row + 1
                    
 
                PropertyTypes = ResourceType & "." & Get_Information(TargetObject(Member), "Type")
                
                If Get_Information(TargetObject(Member), "Type") <> "Map" Then
                
                    Analysis MySheet, ResouceParameters("PropertyTypes")(PropertyTypes)("Properties"), Indent + 1
                    
                Else
                
                    Set_Parameter MySheet, Indent + 1, "Map(Key : Value)", Get_Information(TargetObject(Member), "PrimitiveItemType"), ""
                    Parameter_Row = Parameter_Row + 1
                    
                End If
      
            End If
        
        End If
           
    Next

End Sub


Private Function Get_Information(ByVal TargetObject As Object, ByVal SearchString As String) As String

    Dim Member
    Dim Find As Boolean

    Find = False
    Get_Information = ""
    
    For Each Member In TargetObject
        
        If SearchString = Member Then
            Find = True
        End If
        
    Next
    
    If Find Then
        
        If VBA.VarType(TargetObject(SearchString)) = vbObject Then
            
        
            For Each Member In TargetObject(SearchString)
                
                Get_Information = Member
                
            Next
            
        Else
        
            Get_Information = TargetObject(SearchString)
        
        End If
        
    End If

End Function


Private Sub Set_Parameter(ByVal SheetName As String, ByVal Indent As Integer, ByVal Name As String, ByVal ValueType As String, ByVal LinkUrl As String)

    Worksheets(SheetName).Cells(Parameter_Row, Nunber_Column) = "=row()-" & Nunber_Column
    Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Indent) = Name
    Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Max_Indent) = ValueType

    If Left(LinkUrl, 4) <> "http" Then
        Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Max_Indent + 1) = LinkUrl
    Else
        ActiveSheet.Hyperlinks.Add Anchor:=Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Max_Indent + 1), ScreenTip:=LinkUrl, Address:=LinkUrl, TextToDisplay:="Link"
    End If

End Sub


Private Sub Set_Parameter_list(ByVal SheetName As String, ByVal Indent As Integer, ByVal Name As String, ByVal ValueType As String, ByVal LinkUrl As String)

    Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Indent) = "'- "
    Set_Parameter SheetName, Indent + 1, Name, ValueType, LinkUrl

End Sub


Private Sub Set_list(ByVal SheetName As String, ByVal Indent As Integer)

    Worksheets(SheetName).Cells(Parameter_Row, Parameter_Start_Column + Indent) = "'- "

End Sub


Private Function GetResouceName(ByVal ResourceType As String) As String

    GetResouceName = Mid(ResourceType, InStrRev(ResourceType, ":") + 1)

End Function


'--- Convert Json Files to object --------------------------------------------------

Private Sub ReadJsonFile(ByVal Json_File_Name As String)

    Set ResouceParameters = M999JsonConverter.ParseJson(ReadSJISText(Json_File_Name))

End Sub

Private Function ReadSJISText(ByVal Json_File_Name As String) As String

    Dim Buf  As String

    With CreateObject("ADODB.Stream")
        .Charset = "Shift-JIS"
        .Type = 2           'adTypeText
        .LineSeparator = -1 'adCrLf
        .Open
        .LoadFromFile Json_File_Name
        Buf = .ReadText(-1) 'adReadAll
        .Close
    End With

    ReadSJISText = Buf

End Function
'--- Convert Json Files to object end ----------------------------------------------
