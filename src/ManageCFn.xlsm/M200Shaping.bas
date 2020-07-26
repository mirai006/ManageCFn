Attribute VB_Name = "M200Shaping"
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


Public Sub Shaping(ByVal SheetName As String)

    Set_Titles SheetName
    Fit_ColumnWidth SheetName
    Set_Lines SheetName
    Set_Color SheetName

       
End Sub


Private Sub Set_Color(ByVal SheetName As String)

    Dim i As Integer
    Dim j As Integer
    Dim IsPaint As Boolean
    Dim NowRowColorNo As Integer
    Dim IsColor(Max_Indent) As Boolean
    Dim ColoredColumn As Integer


    i = Parameter_Start_Row

    
    ColoredColumn = 0

    Do While Worksheets(SheetName).Cells(i, Nunber_Column) <> ""
    
        IsPaint = False
        IsColor(0) = True
        IsColor(1) = True
    
        For j = 0 To Max_Indent - 1
        

        
            If Worksheets(SheetName).Cells(i, Parameter_Start_Column + j) <> "" Then
                

                    ColoredColumn = j - 1
                    IsColor(j) = False

                
                If Worksheets(SheetName).Cells(i, Parameter_Start_Column + Max_Indent) = "" Then
                
                    IsPaint = True
                    IsColor(j) = True
                    NowRowColorNo = j
                    
                End If
            
            End If
            
            
            If Worksheets(SheetName).Cells(i, Parameter_Start_Column + j) = "- " Then
            
                 Worksheets(SheetName).Cells(i, Parameter_Start_Column + j).Interior.Color = RGB(255, 255, 153)
            
            End If
            
            If IsPaint Then
            
                Worksheets(SheetName).Cells(i, Parameter_Start_Column + j).Interior.Color = Get_SetColor(NowRowColorNo)
                Worksheets(SheetName).Cells(i, Parameter_Start_Column + j + Max_AddSetting + 1).Interior.Color = Get_SetColor(NowRowColorNo)
            
            End If
            
        
        Next
        
        For j = 0 To ColoredColumn
        
            If IsColor(j) Then
                
                Worksheets(SheetName).Cells(i, Parameter_Start_Column + j).Interior.Color = Get_SetColor(j)
                
                
                If Worksheets(SheetName).Cells(i, Parameter_Start_Column + j) = "- " Then
            
                 Worksheets(SheetName).Cells(i, Parameter_Start_Column + j).Interior.Color = RGB(255, 255, 153)
            
            End If
            
            End If
        
        Next

        i = i + 1
        
    Loop

End Sub


Private Sub Set_Lines(ByVal SheetName As String)

    Dim i As Integer
    Dim j As Integer
    
    i = Parameter_Start_Row
    
    Do While Worksheets(SheetName).Cells(i, Nunber_Column) <> ""
    
        Worksheets(SheetName).Cells(i, Nunber_Column).BorderAround LineStyle:=xlContinuous
        Worksheets(SheetName).Range(Cells(i, Nunber_Column + 1), Cells(i, Nunber_Column + Max_Indent)).BorderAround LineStyle:=xlContinuous
        
        For j = 1 To Max_AddSetting + 1
        
            Worksheets(SheetName).Cells(i, Nunber_Column + Max_Indent + j).BorderAround LineStyle:=xlContinuous
        
        Next

        i = i + 1
    Loop

End Sub


Private Sub Fit_ColumnWidth(ByVal SheetName As String)

    Dim i As Integer

    Worksheets(SheetName).Columns(Nunber_Column).ColumnWidth = 4

    For i = Nunber_Column + 1 To Nunber_Column + Max_Indent - 1
        Worksheets(SheetName).Columns(i).ColumnWidth = 2.25
    Next

    Worksheets(SheetName).Columns(Nunber_Column + Max_Indent).ColumnWidth = 20
    Worksheets(SheetName).Columns(Nunber_Column + Max_Indent + 1).ColumnWidth = 8
    Worksheets(SheetName).Columns(Nunber_Column + Max_Indent + 2).ColumnWidth = 15
    Worksheets(SheetName).Columns(Nunber_Column + Max_Indent + 3).ColumnWidth = 4
    Worksheets(SheetName).Columns(Nunber_Column + Max_Indent + Max_AddSetting + 1).ColumnWidth = 20

End Sub


Private Function Get_SetColor(i As Integer) As Long

    Select Case i Mod 7
    
    Case 0
        Get_SetColor = RGB(255, 239, 239)
    Case 1
        Get_SetColor = RGB(255, 247, 239)
    Case 2
        Get_SetColor = RGB(255, 255, 239)
    Case 3
        Get_SetColor = RGB(247, 255, 239)
    Case 4
        Get_SetColor = RGB(239, 255, 247)
    Case 5
        Get_SetColor = RGB(239, 247, 255)
    Case 6
        Get_SetColor = RGB(247, 239, 255)
    End Select


End Function


Private Sub Set_Titles(ByVal SheetName As String)

    Dim j As Integer

    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column) = "No"
    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column + 1) = "Resources"
    
    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column + 1 + Max_Indent) = "Type"
    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column + 1 + Max_Indent + 1) = "Remarks"

    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column + 1 + Max_Indent + 1 + 1) = "CFi"
    
    Worksheets(SheetName).Cells(Titles_Row, Nunber_Column + 1 + Max_Indent + Max_AddSetting) = "=""Value"" & COLUMN() - " & Str(Nunber_Column + Max_Indent + Max_AddSetting)
    
    For j = Nunber_Column To Nunber_Column + Max_Indent + Max_AddSetting + 1
    
        Worksheets(SheetName).Cells(Titles_Row, j).Font.Bold = True
        Worksheets(SheetName).Cells(Titles_Row, j).Font.Color = RGB(255, 255, 255)
        Worksheets(SheetName).Cells(Titles_Row, j).Interior.Color = RGB(0, 100, 100)
    
    Next

End Sub
