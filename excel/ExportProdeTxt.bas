Option Explicit

Private Const SHEET_GENERAL As String = "TABLA_GENERAL"
Private Const SHEET_OFFICIAL As String = "RESULTADOS_OFICIALES"

Public Sub Auto_Open()
    ExportarTXT_Prode
End Sub

Public Sub ExportarTXT_Prode()
    Dim basePath As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then Exit Sub

    EnsureFolder basePath & "\data"
    EnsureFolder basePath & "\data\users"

    ExportTablaGeneral basePath & "\data\tabla_general.txt"
    ExportResultadosOficiales basePath & "\data\resultados_oficiales.txt"
    ExportUsuarios basePath & "\data\users"
End Sub

Private Sub ExportTablaGeneral(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim json As String
    Dim isFirst As Boolean
    Dim playedCount As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_GENERAL)
    playedCount = GetOfficialCount()
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    json = "{" & DQ() & "updatedAt" & DQ() & ":" & DQ() & IsoNow() & DQ() & "," & DQ() & "rows" & DQ() & ":["
    isFirst = True

    For rowIndex = 2 To lastRow
        If Len(Trim$(ws.Cells(rowIndex, 2).Text)) > 0 And IsNumeric(ws.Cells(rowIndex, 4).Value2) Then
            If Not isFirst Then json = json & ","
            json = json & "{" & _
                DQ() & "position" & DQ() & ":" & CLng(Val(ws.Cells(rowIndex, 1).Value2)) & "," & _
                DQ() & "usuario_tab" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 2).Text) & DQ() & "," & _
                DQ() & "nombre_fantasia" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 3).Text) & DQ() & "," & _
                DQ() & "points" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 4).Value2) & "," & _
                DQ() & "exactos" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 5).Value2) & "," & _
                DQ() & "acertados" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 6).Value2) & "," & _
                DQ() & "played" & DQ() & ":" & CStr(playedCount) & _
            "}"
            isFirst = False
        End If
    Next rowIndex

    json = json & "]}"
    WriteUtf8Text outputPath, json
End Sub

Private Sub ExportResultadosOficiales(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim json As String
    Dim isFirst As Boolean

    Set ws = ThisWorkbook.Worksheets(SHEET_OFFICIAL)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    json = "{" & DQ() & "updatedAt" & DQ() & ":" & DQ() & IsoNow() & DQ() & "," & DQ() & "matches" & DQ() & ":["
    isFirst = True

    For rowIndex = 1 To lastRow
        If IsNumeric(ws.Cells(rowIndex, 1).Value2) Then
            If Not isFirst Then json = json & ","
            json = json & "{" & _
                DQ() & "number" & DQ() & ":" & CLng(Val(ws.Cells(rowIndex, 1).Value2)) & "," & _
                DQ() & "fecha" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 2).Text) & DQ() & "," & _
                DQ() & "grupo" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 3).Text) & DQ() & "," & _
                DQ() & "local" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 4).Text) & DQ() & "," & _
                DQ() & "goles_local" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 5).Value2) & "," & _
                DQ() & "goles_visitante" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 7).Value2) & "," & _
                DQ() & "visitante" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 8).Text) & DQ() & "," & _
                DQ() & "score" & DQ() & ":" & DQ() & JsonScore(ws.Cells(rowIndex, 5).Value2, ws.Cells(rowIndex, 7).Value2) & DQ() & "," & _
                DQ() & "status" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 10).Text) & DQ() & "," & _
                DQ() & "stadium" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 11).Text) & DQ() & _
            "}"
            isFirst = False
        End If
    Next rowIndex

    json = json & "]}"
    WriteUtf8Text outputPath, json
End Sub

Private Sub ExportUsuarios(ByVal folderPath As String)
    Dim ws As Worksheet
    Dim sheetIndex As Long
    Dim userSlug As String
    Dim fantasyName As String
    Dim outputPath As String
    Dim officialMap As Object
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim json As String
    Dim isFirst As Boolean
    Dim officialRow As Variant
    Dim points As Long
    Dim matchText As String
    Dim predictionText As String
    Dim officialText As String

    Set officialMap = BuildOfficialMap()

    For sheetIndex = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(sheetIndex)
        If ws.Name Like "MIPRODE_*" Then
            userSlug = Mid$(ws.Name, 9)
            fantasyName = Trim$(ws.Range("E2").Text)
            outputPath = folderPath & "\" & LCase$(userSlug) & ".txt"

            json = "{" & _
                DQ() & "usuario_tab" & DQ() & ":" & DQ() & JsonEscape(userSlug) & DQ() & "," & _
                DQ() & "nombre_fantasia" & DQ() & ":" & DQ() & JsonEscape(fantasyName) & DQ() & "," & _
                DQ() & "updatedAt" & DQ() & ":" & DQ() & IsoNow() & DQ() & "," & _
                DQ() & "matches" & DQ() & ":["

            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            isFirst = True

            For rowIndex = 1 To lastRow
                If IsNumeric(ws.Cells(rowIndex, 1).Value2) Then
                    officialRow = Empty
                    If officialMap.Exists(CLng(Val(ws.Cells(rowIndex, 1).Value2))) Then
                        officialRow = officialMap(CLng(Val(ws.Cells(rowIndex, 1).Value2)))
                    End If

                    matchText = Trim$(ws.Cells(rowIndex, 5).Text) & " vs " & Trim$(ws.Cells(rowIndex, 9).Text)
                    predictionText = JsonScore(ws.Cells(rowIndex, 6).Value2, ws.Cells(rowIndex, 8).Value2)

                    If IsArray(officialRow) Then
                        officialText = JsonScore(officialRow(0), officialRow(1))
                        points = ScoreMatch(ws.Cells(rowIndex, 6).Value2, ws.Cells(rowIndex, 8).Value2, officialRow(0), officialRow(1))
                    Else
                        officialText = ""
                        points = 0
                    End If

                    If Not isFirst Then json = json & ","
                    json = json & "{" & _
                        DQ() & "number" & DQ() & ":" & CLng(Val(ws.Cells(rowIndex, 1).Value2)) & "," & _
                        DQ() & "fecha" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 2).Text) & DQ() & "," & _
                        DQ() & "horario" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 3).Text) & DQ() & "," & _
                        DQ() & "grupo" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 4).Text) & DQ() & "," & _
                        DQ() & "local" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 5).Text) & DQ() & "," & _
                        DQ() & "prediction_local" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 6).Value2) & "," & _
                        DQ() & "prediction_visitante" & DQ() & ":" & JsonNumber(ws.Cells(rowIndex, 8).Value2) & "," & _
                        DQ() & "visitante" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 9).Text) & DQ() & "," & _
                        DQ() & "stadium" & DQ() & ":" & DQ() & JsonEscape(ws.Cells(rowIndex, 10).Text) & DQ() & "," & _
                        DQ() & "match" & DQ() & ":" & DQ() & JsonEscape(matchText) & DQ() & "," & _
                        DQ() & "prediction" & DQ() & ":" & DQ() & JsonEscape(predictionText) & DQ() & "," & _
                        DQ() & "official" & DQ() & ":" & DQ() & JsonEscape(officialText) & DQ() & "," & _
                        DQ() & "points" & DQ() & ":" & CStr(points) & _
                    "}"
                    isFirst = False
                End If
            Next rowIndex

            json = json & "]}"
            WriteUtf8Text outputPath, json
        End If
    Next sheetIndex
End Sub

Private Function BuildOfficialMap() As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim map As Object
    Dim key As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SHEET_OFFICIAL)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 1 To lastRow
        If IsNumeric(ws.Cells(rowIndex, 1).Value2) Then
            key = CLng(Val(ws.Cells(rowIndex, 1).Value2))
            map(key) = Array(CLng(Val(ws.Cells(rowIndex, 5).Value2)), CLng(Val(ws.Cells(rowIndex, 7).Value2)))
        End If
    Next rowIndex

    Set BuildOfficialMap = map
End Function

Private Function GetOfficialCount() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim count As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_OFFICIAL)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 1 To lastRow
        If IsNumeric(ws.Cells(rowIndex, 1).Value2) Then
            count = count + 1
        End If
    Next rowIndex

    GetOfficialCount = count
End Function

Private Function ScoreMatch(ByVal predictedHome As Variant, ByVal predictedAway As Variant, ByVal actualHome As Variant, ByVal actualAway As Variant) As Long
    Dim predictedResult As Long
    Dim actualResult As Long

    predictedResult = Sgn(CLng(Val(predictedHome)) - CLng(Val(predictedAway)))
    actualResult = Sgn(CLng(Val(actualHome)) - CLng(Val(actualAway)))

    If predictedResult = actualResult Then
        ScoreMatch = 2
        If CLng(Val(predictedHome)) = CLng(Val(actualHome)) And CLng(Val(predictedAway)) = CLng(Val(actualAway)) Then
            ScoreMatch = ScoreMatch + 1
        End If
    End If
End Function

Private Function JsonNumber(ByVal value As Variant) As String
    If IsNumeric(value) Then
        JsonNumber = CStr(CLng(Val(value)))
    Else
        JsonNumber = "0"
    End If
End Function

Private Function JsonScore(ByVal homeGoals As Variant, ByVal awayGoals As Variant) As String
    JsonScore = CStr(CLng(Val(homeGoals))) & " - " & CStr(CLng(Val(awayGoals)))
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace$(value, Chr$(92), Chr$(92) & Chr$(92))
    value = Replace$(value, Chr$(34), Chr$(92) & Chr$(34))
    value = Replace$(value, vbCrLf, "\n")
    value = Replace$(value, vbCr, "\n")
    value = Replace$(value, vbLf, "\n")
    JsonEscape = value
End Function

Private Function IsoNow() As String
    IsoNow = Format$(Now, "yyyy-mm-dd\Thh:nn:ss") & "-03:00"
End Function

Private Function DQ() As String
    DQ = Chr$(34)
End Function

Private Sub WriteUtf8Text(ByVal filePath As String, ByVal content As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim parentPath As String

    If Len(folderPath) = 0 Then Exit Sub
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    parentPath = Left$(folderPath, InStrRev(folderPath, "\") - 1)
    If Len(parentPath) > 0 And Dir(parentPath, vbDirectory) = "" Then
        EnsureFolder parentPath
    End If

    MkDir folderPath
End Sub