Attribute VB_Name = "M300CreateCfn"
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

Private Const CFnDir = "01_CFn"
Private Const CreatedResourceDir = "11_CreatedResource"
Private Const BackupDir = "99_Backup"

Private Const IndentString = "  "

Private Const ToolVersion = "Create Cloud Formation 2020/07/25"
Private Const CreateInformationOnTags = True

Private CFnFileName As String

Private CFnFullPath As String
Private CreatedResourcFullPath As String
Private BackupFullPath As String

Private CFn_Resources As String
Private CFn_Outputs As String

Private CFnTpye As String
Private ResouceRow As Integer
Private ExternalValueRow As Integer

Public Sub CreateCFn()

    Dim SheetNo As Integer

    Create_Dir
    Set_FileName
    set_CFnVersion
    
    ResouceRow = ResouceType_Start_Row
    Worksheets("CreatingResource").Range("B3:C2002").ClearContents

    
    GetExternalValue
    
    For SheetNo = 1 To Worksheets.Count
        
        If Worksheets(SheetNo).Cells(1, 1) = "CFn" Then
            Set_ResoucesValues SheetNo
        End If
    
    Next
    
    CreateCFnYamlFile
    CreateExternalValue
    BackupExcelFile

    
End Sub

Private Sub BackupExcelFile()

    Workbooks(ThisWorkbook.Name).Save
    DeleteTool

End Sub

Private Sub DeleteTool()

    Set_FileName

    Application.DisplayAlerts = False
    
        Worksheets("Setting").Buttons.Delete
        Worksheets("Resources").Delete
        ThisWorkbook.SaveAs FileName:=BackupFullPath, FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = True
    
    Workbooks(ThisWorkbook.Name).Close

End Sub

Private Sub CreateExternalValue()

    Dim CheckRow As Integer
    Dim ExternalValue As String
    Dim FileNo As Integer
    
    CheckRow = ResouceType_Start_Row
    ExternalValue = ""
    
    Do While Worksheets("CreatingResource").Cells(CheckRow, ResouceType_Start_Column).Value <> ""
    
        ExternalValue = ExternalValue & Worksheets("CreatingResource").Cells(CheckRow, ResouceType_Start_Column).Value & vbTab & Worksheets("CreatingResource").Cells(CheckRow, ResouceType_Start_Column + 1).Value & vbCrLf
        CheckRow = CheckRow + 1
    Loop
    
    FileNo = FreeFile
     
    Open CreatedResourcFullPath For Append As #FileNo
    Print #FileNo, ExternalValue
    Close #FileNo

End Sub

Public Sub GetExternalValue()

    Dim FileName As String
    
    ExternalValueRow = ResouceType_Start_Row
    Worksheets("ExternalValue").Range("B3:C2002").ClearContents
    
    FileName = Dir(ActiveWorkbook.Path & "\" & CreatedResourceDir & "\*.txt")
    
    Do While FileName <> ""
    
        GetExternalValueFromFile ActiveWorkbook.Path & "\" & CreatedResourceDir & "\" & FileName

        FileName = Dir

    Loop


End Sub


Private Sub GetExternalValueFromFile(ByVal FileFilePath As String)

    Dim ExternalValue As String
    Dim FileNo As Integer

    FileNo = FreeFile
    
    Open FileFilePath For Input As #FileNo
    
        Line Input #FileNo, ExternalValue

        Do Until EOF(1)
        
            SetExternalValue ExternalValue
            Line Input #FileNo, ExternalValue

        Loop

    Close #FileNo

End Sub


Private Sub SetExternalValue(ByVal Value As String)

    Dim ResouceType As String
    Dim ResouceName As String
    Dim CheckRow As Integer

    CheckRow = ResouceType_Start_Row
    
    ResouceType = Left(Value, InStr(Value, vbTab) - 1)
    ResouceName = Right(Value, Len(Value) - Len(ResouceType) - 1)

    
    Do While Worksheets("ExternalValue").Cells(CheckRow, ResouceType_Start_Column + 1).Value <> ""
    
        If Worksheets("ExternalValue").Cells(CheckRow, ResouceType_Start_Column + 1).Value = ResouceName Then
        
            Exit Sub
        
        End If
        
        CheckRow = CheckRow + 1
    
    Loop
    
    Worksheets("ExternalValue").Cells(ExternalValueRow, ResouceType_Start_Column).Value = ResouceType
    Worksheets("ExternalValue").Cells(ExternalValueRow, ResouceType_Start_Column + 1).Value = ResouceName
    
    ExternalValueRow = ExternalValueRow + 1

End Sub


Private Sub CreateCFnYamlFile()

    Dim Config As String
    Dim FileNo As Integer

    If CFn_Outputs = "Outputs:" & vbCrLf Then
    
        Config = CFn_Resources
    
    Else
    
        Config = CFn_Resources & vbCrLf & CFn_Outputs
    
    End If
    
    FileNo = FreeFile
     
    Open CFnFullPath For Append As #FileNo
    Print #FileNo, Config
    Close #FileNo

End Sub


Private Sub Set_ResoucesValues(ByVal SheetNo As Integer)

    Dim ValuesColumn As Integer
    Dim CheckRow As Integer
    Dim j As Integer
    Dim IsList As Boolean
    
    ValuesColumn = Parameter_Start_Column + Max_Indent + Max_AddSetting
    CheckRow = Parameter_Start_Row
    
    Do While Worksheets(SheetNo).Cells(CheckRow, ValuesColumn).Value <> ""
        
        CFn_Resources = CFn_Resources & IndentString & RemoveSymbols(Worksheets(SheetNo).Cells(CheckRow, ValuesColumn).Value) & ": " & vbCrLf
        CheckRow = CheckRow + 1
        
        CFnTpye = Right(IndentString & Worksheets(SheetNo).Cells(CheckRow, Parameter_Start_Column + 2).Value, Len(IndentString & Worksheets(SheetNo).Cells(CheckRow, Parameter_Start_Column + 2).Value) - Len("Type:  "))
        CFn_Resources = CFn_Resources & IndentString & IndentString & Worksheets(SheetNo).Cells(CheckRow, Parameter_Start_Column + 2).Value & vbCrLf
        CheckRow = CheckRow + 1
        
        CFn_Resources = CFn_Resources & IndentString & IndentString & "DeletionPolicy: " & Worksheets("Setting").Cells(DeletionPolicy_Row, DeletionPolicy_Column).Value & vbCrLf

        CFn_Resources = CFn_Resources & IndentString & IndentString & Worksheets(SheetNo).Cells(CheckRow, Parameter_Start_Column + 2).Value & vbCrLf
        CheckRow = CheckRow + 1

        IsList = False
        
        Do While Worksheets(SheetNo).Cells(CheckRow, Nunber_Column).Value <> ""

            If Worksheets(SheetNo).Cells(CheckRow, ValuesColumn).Value = "" Then
            
                
                If SatisfyingValue(SheetNo, CheckRow, ValuesColumn) Then

                    If Worksheets(SheetNo).Cells(CheckRow, Parameter_Start_Column + Max_Indent).Value = "" Then
                    
                        Set_Value2Yaml SheetNo, CheckRow, ValuesColumn
                    
                    Else
                        If Is_List(SheetNo, CheckRow) Then
                            IsList = True
                        End If
                    End If
                End If
            
            Else
            
                If IsList Then
                
                    Set_ValueList2Yaml SheetNo, CheckRow, ValuesColumn
                    IsList = False
                
                Else
                    Set_Value2Yaml SheetNo, CheckRow, ValuesColumn
                End If
            
            End If
            
            If Is_Tags(SheetNo, CheckRow) Then
                ' ここで書いちゃうとその後の指定で、文字列にしないといけないのに、!Refとなってしまう。考えること。
                TagsTreatment SheetNo, CheckRow, ValuesColumn
            End If
        
        CheckRow = CheckRow + 1
        
        Loop

        ValuesColumn = ValuesColumn + 1
        CheckRow = Parameter_Start_Row
        CFn_Resources = CFn_Resources & vbCrLf

    Loop


End Sub


Private Function SatisfyingValue(ByVal SheetNo, ByVal Row As Integer, ByVal ValuesColumn) As Boolean

    Dim CheckRow As Integer
    Dim TragetColumn As Integer
    Dim i As Integer
    
    CheckRow = Row
    SatisfyingValue = False
    
    For TragetColumn = Parameter_Start_Column To Parameter_Start_Column + Max_Indent
        If Worksheets(SheetNo).Cells(CheckRow, TragetColumn) <> "" Then
                Exit For
        End If
    Next
        
    If Worksheets(SheetNo).Cells(CheckRow, TragetColumn) = "Tags: " Then
    
        SatisfyingValue = True
        Exit Function
        
    End If
        
    CheckRow = CheckRow + 1
        
    Do While Worksheets(SheetNo).Cells(CheckRow, Nunber_Column).Value <> ""

        For i = Parameter_Start_Column To TragetColumn
            If Worksheets(SheetNo).Cells(CheckRow, i) <> "" Then
                Exit Function
            End If
        Next
        
        If Worksheets(SheetNo).Cells(CheckRow, ValuesColumn) <> "" Then
            SatisfyingValue = True
            Exit Function
        End If

        CheckRow = CheckRow + 1
    Loop

End Function


Private Function Is_Tags(ByVal SheetNo, ByVal Row As Integer) As Boolean

    Dim TragetColumn As Integer
    
    Is_Tags = False
    
    For TragetColumn = Parameter_Start_Column To Parameter_Start_Column + Max_Indent
        
        If Worksheets(SheetNo).Cells(Row, TragetColumn) = "Tags: " Then
            Is_Tags = True
            Exit For
        End If
    Next
        

End Function


Private Function Is_List(ByVal SheetNo, ByVal Row As Integer) As Boolean

    Dim TragetColumn As Integer
    
    Is_List = False
    
    For TragetColumn = Parameter_Start_Column To Parameter_Start_Column + Max_Indent
        
        If Worksheets(SheetNo).Cells(Row, TragetColumn) = "- " Then
            If Worksheets(SheetNo).Cells(Row, TragetColumn + 1) <> "" Then
                Is_List = True
                Exit For
            End If
        End If
    Next
        

End Function


Private Sub Set_Value2Yaml(ByVal SheetNo As Integer, ByVal Row As Integer, ByVal ValuesColumn As Integer)

    Dim i As Integer
    
    For i = 0 To Max_Indent

        If Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value = "" Then
        
           CFn_Resources = CFn_Resources & IndentString
    
        Else
            
            
            If Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value = "- " Then
            
                If Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i + 1).Value = "" Then
                    CFn_Resources = CFn_Resources & Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value & GetValue(Worksheets(SheetNo).Cells(Row, ValuesColumn).Value) & vbCrLf
                    Exit Sub
                Else
                    CFn_Resources = CFn_Resources & Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value
                    CFn_Resources = CFn_Resources & Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i + 1).Value & GetValue(Worksheets(SheetNo).Cells(Row, ValuesColumn).Value) & vbCrLf
                    Exit Sub
                End If
            Else
            
                CFn_Resources = CFn_Resources & Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value & GetValue(Worksheets(SheetNo).Cells(Row, ValuesColumn).Value) & vbCrLf
                Exit Sub
            
            End If
                
        End If
    
    Next

End Sub


Private Sub Set_ValueList2Yaml(ByVal SheetNo As Integer, ByVal Row As Integer, ByVal ValuesColumn As Integer)

    Dim i As Integer
    Dim Buf As String
    
    Buf = ""
    
    For i = 0 To Max_Indent

        If Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value = "" Then
        
            Buf = Buf & IndentString
    
        Else

            Buf = Left(Buf, Len(Buf) - 2) & "- "
            CFn_Resources = CFn_Resources & Buf & Worksheets(SheetNo).Cells(Row, Parameter_Start_Column + i).Value & GetValue(Worksheets(SheetNo).Cells(Row, ValuesColumn).Value) & vbCrLf
            Exit Sub
                
        End If
    
    Next
End Sub


Private Sub TagsTreatment(ByVal SheetNo, ByVal Row As Integer, ByVal ValuesColumn)

    Dim TragetColumn As Integer
    Dim IndentNumber  As Integer
    
    IndentNumber = 0
    
    For TragetColumn = Parameter_Start_Column To Parameter_Start_Column + Max_Indent
        
        If Worksheets(SheetNo).Cells(Row, TragetColumn) = "Tags: " Then
            Exit For
        End If
        
        IndentNumber = IndentNumber + 1
        
    Next
    
    CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber) & "  - Key: Name" & vbCrLf
    CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber + 2) & "Value: " & Worksheets(SheetNo).Cells(Parameter_Start_Row, ValuesColumn).Value & vbCrLf

    If CreateInformationOnTags Then
        
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber) & "  - Key: Tool Version" & vbCrLf
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber + 2) & "Value: " & ToolVersion & vbCrLf
    
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber) & "  - Key: CloudFormation File" & vbCrLf
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber + 2) & "Value: " & CFnFileName & vbCrLf
        
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber) & "  - Key: ResourceSpecificationVersion" & vbCrLf
        CFn_Resources = CFn_Resources & WorksheetFunction.Rept(IndentString, IndentNumber + 2) & "Value: " & Worksheets("Resources").Cells(ResourceSpecificationVersion_Row, ResourceSpecificationVersion_Column).Value & vbCrLf

    End If
    
    Worksheets("CreatingResource").Cells(ResouceRow, ResouceType_Start_Column).Value = CFnTpye
    Worksheets("CreatingResource").Cells(ResouceRow, ResouceType_Start_Column + 1).Value = Worksheets(SheetNo).Cells(Parameter_Start_Row, ValuesColumn).Value
    ResouceRow = ResouceRow + 1
    
    CFn_Outputs = CFn_Outputs & WorksheetFunction.Rept(IndentString, 1) & "Export" & RemoveSymbols(Worksheets(SheetNo).Cells(Parameter_Start_Row, ValuesColumn).Value) & ": " & vbCrLf
    CFn_Outputs = CFn_Outputs & WorksheetFunction.Rept(IndentString, 2) & "Value: !Ref " & RemoveSymbols(Worksheets(SheetNo).Cells(Parameter_Start_Row, ValuesColumn).Value) & vbCrLf
    CFn_Outputs = CFn_Outputs & WorksheetFunction.Rept(IndentString, 2) & "Export: " & vbCrLf
    CFn_Outputs = CFn_Outputs & WorksheetFunction.Rept(IndentString, 3) & "Name: " & Worksheets(SheetNo).Cells(Parameter_Start_Row, ValuesColumn).Value & vbCrLf
    CFn_Outputs = CFn_Outputs & WorksheetFunction.Rept(IndentString, 2) & vbCrLf

End Sub


Private Function GetValue(ByVal Value As String) As String

    Dim CheckRow As String

    GetValue = Value
    CheckRow = ResouceType_Start_Row
    
    Do While Worksheets("ExternalValue").Cells(CheckRow, ResouceType_Start_Column + 1).Value <> ""
    
        If Value = Worksheets("ExternalValue").Cells(CheckRow, ResouceType_Start_Column + 1).Value Then
        
            GetValue = "!ImportValue " & Value
            
        End If
    
        CheckRow = CheckRow + 1
    
    Loop
    
    CheckRow = ResouceType_Start_Row
    
    Do While Worksheets("CreatingResource").Cells(CheckRow, ResouceType_Start_Column + 1).Value <> ""
    
        If Value = Worksheets("CreatingResource").Cells(CheckRow, ResouceType_Start_Column + 1).Value Then
        
            GetValue = "!Ref " & RemoveSymbols(Value)
            
        End If
    
        CheckRow = CheckRow + 1
    
    Loop
        
    
End Function


Private Sub set_CFnVersion()

    CFn_Resources = "AWSTemplateFormatVersion: ""2010-09-09""" & vbCrLf & vbCrLf
    CFn_Resources = CFn_Resources + "Resources:" & vbCrLf
    CFn_Outputs = "Outputs:" & vbCrLf
    
End Sub


Private Sub Set_FileName()

    Dim NowTime As String
    
    NowTime = Format(Now, "yymmdd_hhnnss")
    
    CFnFileName = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & "_" & NowTime & ".yaml"
    
    CFnFullPath = ActiveWorkbook.Path & "\" & CFnDir & "\" & Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & "_" & NowTime & ".yaml"
    CreatedResourcFullPath = ActiveWorkbook.Path & "\" & CreatedResourceDir & "\" & Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & "_" & NowTime & ".txt"
    BackupFullPath = ActiveWorkbook.Path & "\" & BackupDir & "\" & Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & "_" & NowTime & ".xlsx"

End Sub


Private Sub Create_Dir()

    Dim FilePath

    FilePath = ActiveWorkbook.Path & "\" & BackupDir

    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
    End If

    FilePath = ActiveWorkbook.Path & "\" & CreatedResourceDir

    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
    End If

    FilePath = ActiveWorkbook.Path & "\" & CFnDir

    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
    End If

End Sub


Private Function RemoveSymbols(ByVal Origen As String) As String

    RemoveSymbols = Origen
    RemoveSymbols = Replace(RemoveSymbols, "(", "")
    RemoveSymbols = Replace(RemoveSymbols, "_", "")
    RemoveSymbols = Replace(RemoveSymbols, ".", "")
    RemoveSymbols = Replace(RemoveSymbols, ":", "")
    RemoveSymbols = Replace(RemoveSymbols, "/", "")
    RemoveSymbols = Replace(RemoveSymbols, "=", "")
    RemoveSymbols = Replace(RemoveSymbols, "+", "")
    RemoveSymbols = Replace(RemoveSymbols, "-", "")
    RemoveSymbols = Replace(RemoveSymbols, "@", "")
    RemoveSymbols = Replace(RemoveSymbols, ")", "")
    
End Function
