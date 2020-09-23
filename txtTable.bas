Attribute VB_Name = "txtTable"
Type tblAttrib
tblThick As Integer
tblTop As Boolean
tblBottom As Boolean
tblHeight As Integer
tblWidth As Integer
tblPadding As String
tblWall As String
tblData As String
End Type

Function DrawTable(Attributes As tblAttrib) As String
Dim iCount As Integer
Dim sTable As String
iCount = Attributes.tblHeight
    For x = 1 To Attributes.tblHeight
        For y = 1 To Attributes.tblWidth
            If x = Attributes.tblHeight And Attributes.tblBottom = True Then
                sTable = sTable & Attributes.tblWall
            ElseIf x = 1 And Attributes.tblTop = True Then
                sTable = sTable & Attributes.tblWall
            Else
                If (y > Attributes.tblThick) And (y <= Attributes.tblWidth - Attributes.tblThick) Then
                    sTable = sTable & Attributes.tblPadding
                Else
                    sTable = sTable & Attributes.tblWall
                End If
            End If
        Next
    sTable = sTable & Chr$(13) & Chr$(10)
    Next
    
Dim stab As Variant
stab = Split(sTable, vbCrLf)
    For x = 0 To UBound(stab)
        If IsRow(Attributes.tblData, CInt(x)) = True Then
            buff = ((Attributes.tblWidth - (Attributes.tblThick * 2)) - Len(GetData(Attributes.tblData, CInt(x))))
                buf1 = 0
                buf2 = 0
                If 2 Mod buff > 0 Then
                    buf1 = buff / 2
                    buf2 = buff / 2 - 1
                Else
                    buf1 = buff / 2
                    buf2 = buff / 2
                End If
                    kk = ""
                    For k = 1 To Attributes.tblThick
                        kk = kk & Attributes.tblWall
                    Next
                    stab(x - 1) = kk & Space(buf1) & GetData(Attributes.tblData, CInt(x)) & Space(buf2) & kk
stov:
                        If Len(stab(x - 1)) < Attributes.tblWidth Then
                            stab(x - 1) = Mid$(stab(x - 1), 1, Len(stab(x - 1)) - Attributes.tblThick) & " "
                                 ll = 1
                                 For ll = 1 To Attributes.tblThick
                                    stab(x - 1) = stab(x - 1) & Attributes.tblWall
                                 Next
                            GoTo stov
                        End If
        End If
    Next

For z = 0 To UBound(stab)
t = t & stab(z) & vbCrLf
Next
t = Left(t, Len(t) - 1)

DrawTable = t
End Function

Private Function IsRow(sDat As String, CurrRow As Integer) As Boolean
Dim s As Variant
s = Split(sDat, ",")

For d = 0 To UBound(s) Step 2
    If s(d) = CurrRow Then IsRow = True: Exit Function
Next
IsRow = False: Exit Function
End Function

Private Function GetData(sDat As String, iRow As Integer) As String
Dim s As Variant
s = Split(sDat, ",")

For a = 0 To UBound(s) Step 2
    If s(a) = iRow Then GetData = s(a + 1): Exit Function
Next

GetData = "ERROR"
Exit Function
End Function

Sub Main()
Dim snew(1) As tblAttrib
    With snew(0)
        .tblBottom = True
        .tblTop = True
        .tblHeight = 3
        .tblWidth = 20
        .tblThick = 1
        .tblWall = "*"
        .tblPadding = " "
        .tblData = "2,Test"
    End With
    
    With snew(1)
        .tblBottom = True
        .tblTop = False
        .tblHeight = 6
        .tblWidth = 20
        .tblThick = 1
        .tblWall = "*"
        .tblPadding = " "
        .tblData = "2,Hello,3,World!,4,!"
    End With
Form1.Text1.Text = DrawTable(snew(0)) & DrawTable(snew(1))
Form1.Show
End Sub
