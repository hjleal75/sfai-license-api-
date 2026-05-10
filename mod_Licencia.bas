Attribute VB_Name = "mod_Licencia"
' ============================================================
' mod_Licencia — SFAI OneXcel License Manager
' Backend: Vercel + Upstash (en lugar de Cloudflare Worker)
' Reemplaza el mod_Licencia anterior — misma API, diferente URL
' ============================================================
Option Explicit

' ── Endpoint de tu proyecto en Vercel ────────────────────────
' Cambiar "sfai-license" por el nombre real de tu proyecto
Private Const API_URL     As String = "https://sfai-license.vercel.app"
Private Const EXCEL_TOKEN As String = "sfai_xl_2026_prod_TU_TOKEN_AQUI"

' ── Orquestador principal ────────────────────────────────────
Function ValidarLicencia() As Boolean
    ValidarLicencia = False
    If Not ValidarUsuario() Then Exit Function
    If Not ValidarFecha()   Then Exit Function
    If Not ValidarKey()     Then Exit Function
    ValidarLicencia = True
End Function

' ── Check 1: Usuario del PC ──────────────────────────────────
Function ValidarUsuario() As Boolean
    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("_LIC")
    Dim userPC As String
    userPC = LCase(Environ("USERNAME"))
    Dim emailReg As String
    emailReg = LCase(Trim(wsLic.Range("B2").Value))
    Dim userReg As String
    If InStr(emailReg, "@") > 0 Then
        userReg = Left(emailReg, InStr(emailReg, "@") - 1)
    Else
        userReg = emailReg
    End If
    ValidarUsuario = (InStr(userReg, userPC) > 0 Or InStr(userPC, userReg) > 0)
End Function

' ── Check 2: Fecha local (fallback offline) ──────────────────
Function ValidarFecha() As Boolean
    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("_LIC")
    On Error Resume Next
    Dim fechaExpira As Date
    fechaExpira = CDate(wsLic.Range("B3").Value)
    On Error GoTo 0
    If fechaExpira = 0 Then
        ValidarFecha = False : Exit Function
    End If
    ValidarFecha = (Date <= fechaExpira)
End Function

' ── Check 3: Validar Key via Vercel ─────────────────────────
Function ValidarKey() As Boolean
    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("_LIC")

    Dim email As String, key As String
    email = Trim(LCase(wsLic.Range("B2").Value))
    key   = Trim(wsLic.Range("B4").Value)

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    On Error GoTo FallbackOffline

    ' ── Llamada al endpoint de Vercel ────────────────────────
    http.Open "POST", API_URL & "/api/validate", False
    http.setRequestHeader "Content-Type",   "application/json"
    http.setRequestHeader "X-Excel-Token",  EXCEL_TOKEN
    http.setRequestHeader "User-Agent",     "SFAI-Excel/1.0"

    Dim body As String
    body = "{""email"":""" & email & """,""key"":""" & key & """}"
    http.Send body

    If http.Status = 200 Then
        Dim resp As String
        resp = http.responseText

        If InStr(resp, """valid"":true") > 0 Then
            ValidarKey = True
            ' Tier viene del servidor
            If InStr(resp, "AGENCIA") > 0 Then wsLic.Range("B5").Value = "AGENCIA"
            If InStr(resp, "PRO")     > 0 Then wsLic.Range("B5").Value = "PRO"
            If InStr(resp, "FREE")    > 0 Then wsLic.Range("B5").Value = "FREE"
        Else
            ValidarKey = False
            Dim reason As String
            If InStr(resp, "expired")     > 0 Then reason = "Licencia expirada"
            If InStr(resp, "revoked")     > 0 Then reason = "Licencia revocada"
            If InStr(resp, "not_found")   > 0 Then reason = "No encontrada en servidor"
            If InStr(resp, "invalid_key") > 0 Then reason = "Key invalida"
            If reason = "" Then reason = "Error: " & Left(resp, 100)
            wsLic.Range("B6").Value = reason
        End If
    Else
        GoTo FallbackOffline
    End If
    Exit Function

FallbackOffline:
    ' Sin internet: validar hash local como respaldo
    ValidarKey = ValidarKeyOffline()
End Function

Function ValidarKeyOffline() As Boolean
    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("_LIC")
    Dim keyAlmacenada As String
    keyAlmacenada = wsLic.Range("B4").Value
    Dim keyEsperada As String
    keyEsperada = GenerarKey(wsLic.Range("B2").Value, wsLic.Range("B3").Value)
    ValidarKeyOffline = (keyAlmacenada = keyEsperada)
End Function

Function GenerarKey(email As String, fecha As String) As String
    Dim i As Long, suma As Long
    suma = 0
    Dim semilla As String
    semilla = LCase(Trim(email)) & "|" & Trim(fecha)
    For i = 1 To Len(semilla)
        suma = suma + (Asc(Mid(semilla, i, 1)) * i)
    Next i
    GenerarKey = "SFAI-" & Format(suma Mod 99999, "00000") & "-" & Format(Len(email), "00")
End Function

' ── Control de hojas por tier ────────────────────────────────
Sub DesbloquearPorTier()
    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("_LIC")
    Dim tier As String
    tier = UCase(Trim(wsLic.Range("B5").Value))
    Call BloquearTodo
    Select Case tier
        Case "FREE"
            ShowSheet "Dashboard"
        Case "PRO"
            ShowSheet "Dashboard"
            ShowSheet "Reportes"
            ShowSheet "Exportar"
        Case "AGENCIA"
            Dim ws As Worksheet
            For Each ws In ThisWorkbook.Sheets
                If ws.Name <> "_LIC" Then ws.Visible = xlSheetVisible
            Next ws
        Case Else
            Exit Sub
    End Select
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub

Sub BloquearTodo()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVeryHidden
    Next ws
End Sub

Sub ShowSheet(nombre As String)
    On Error Resume Next
    ThisWorkbook.Sheets(nombre).Visible = xlSheetVisible
    On Error GoTo 0
End Sub
