Attribute VB_Name = "Module1"
' No Módulo (por exemplo, Module1.bas)
Option Explicit

Public Tipo As String
Public networkPath As String
Public Ver As String
Public NivelLog As Integer

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_TRAYMESSAGE As Long = &H800 + 1
Private lpPrevWndProc As Long
Private TrayCallbackMessage As Long
Private intervaloTimer As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
    hBalloonIcon As Long
End Type

Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_ICON As Long = &H2
Private Const NIF_TIP As Long = &H4

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDataType As String
End Type

Private trayIcon As NOTIFYICONDATA
Private TimerID As Long
Private isProcessing As Boolean

Private Sub TrayIconCallback(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If Msg = WM_TRAYMESSAGE Then
        If lParam = WM_LBUTTONDBLCLK Then
            ' Chamada para encerrar o programa
            Unload Form1
            End
        End If
    End If
End Sub

Public Sub StartTimer(ByVal interval As Long)
    WriteToLog "StartTimer", 1
    TimerID = SetTimer(0&, 0&, interval, AddressOf TimerProc)
    If TimerID = 0 Then
        WriteToLog "Não foi possível criar o timer! TimerID = 0", 1
        MsgBox "Não foi possível criar o timer!", vbCritical
    End If
End Sub

Public Sub StopTimer()
    WriteToLog "StopTimer", 2
    KillTimer 0&, TimerID
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
    Dim sourceFile As String
    Dim fileFound As String
    Dim mostRecentFile As String
    Dim mostRecentDate As Date
    Dim currentFileDate As Date
    Static lastProcessingTime As Double

    If isProcessing Then
        If Timer - lastProcessingTime > 60 Then ' 60 segundos = 1 minuto
            WriteToLog "isProcessing passado para TRUE", 1
            isProcessing = False
        Else
            WriteToLog "Saiu de TimerProc porque isProcessing = true", 2
            Exit Sub
        End If
    Else
        lastProcessingTime = Timer ' Atualiza apenas quando isProcessing passa de True para False
    End If
    isProcessing = True
'    If isProcessing Then
'        WriteToLog "Saiu de TimerProc porque isProcessing = true", 2
'        Exit Sub
'    End If
'    isProcessing = True

    WriteToLog "TimerProc Tipo = " & Tipo, 2
    
    If Tipo = "" Then
        Tipo = "0"
        WriteToLog "Variavel Tipo esta vazia", 2
        Ver = "C:\OrCarro"
        WriteToLog "Variavel Tipo ajustada para " & Tipo, 2
    End If

    If Tipo = "0" Then
        'COPIA PARA A REDE
        sourceFile = Ver & "\Impres.prn"
        WriteToLog "Procura por " & sourceFile, 2
        If Dir(sourceFile) <> "" Then
            WriteToLog "Procurou por " & sourceFile & " e ACHOU", 1
            MoveECopia sourceFile
        Else
            WriteToLog "Procurou por " & sourceFile & " e NÃO ACHOU", 1
        End If
    Else
        'IMPRIME
        If Ver = "" Then
            WriteToLog "Variavel Ver esta vazia", 2
            Ver = "C:\OrCarro"
            WriteToLog "Variavel Ver ajustada para " & Ver, 2
        End If
        WriteToLog "Procura por tudo que estiver em " & Ver, 2
        fileFound = Dir(Ver & "\*.*")
        If fileFound > "" Then
            sourceFile = Ver & "\" & fileFound
            WriteToLog "Procurou em " & Ver & " e ACHOU", 1
            Imprime sourceFile
        Else
            WriteToLog "Procurou em " & Ver & " e NÃO ACHOU", 1
        End If
    End If
    isProcessing = False
End Sub

Private Sub Imprime(sourceFile As String)
    Dim destinationFile As String
    Dim printerName As String

    destinationFile = App.Path & "\Bak\" & Mid(sourceFile, InStrRev(sourceFile, "\") + 1)
    If Dir(App.Path & "\Bak", vbDirectory) = "" Then
        MkDir App.Path & "\Bak"
    End If
    FileCopy sourceFile, destinationFile
    printerName = GetPrinterName()
    PrintFile printerName, destinationFile
    DelayedKill sourceFile
    WriteToLog "Arquivo enviado para impressão: " & sourceFile, 1
End Sub

Function ReadIniValue(section As String, key As String, iniFileName As String) As String
    Dim returnValue As String
    Dim length As Long
    
    returnValue = Space$(255)
    length = GetPrivateProfileString(section, key, "", returnValue, Len(returnValue), iniFileName)
    ReadIniValue = Left$(returnValue, length)
End Function

Private Sub MoveECopia(sourceFile As String)
    Dim destinationFile As String
    Dim dateTimeStamp As String
    Dim NmArquivo As String
    
    dateTimeStamp = Format(Now, "yyyymmdd_HHMMSS")
    NmArquivo = "impress_" & dateTimeStamp & ".txt"
    destinationFile = App.Path & "\Bak\" & NmArquivo
    Sleep intervaloTimer
    FileCopy sourceFile, destinationFile
    FileCopy sourceFile, networkPath & "\" & NmArquivo
    DelayedKill sourceFile
    WriteToLog "Arquivo copiado: " & sourceFile, 1
End Sub

Sub PrintFile(ByVal printerName As String, ByVal fileName As String)
    Dim hPrinter As Long
    Dim di As DOCINFO
    Dim fileContent() As Byte
    Dim bytesWritten As Long
    
    Open fileName For Binary Access Read As #1
    ReDim fileContent(LOF(1) - 1)
    Get #1, , fileContent
    Close #1
    If OpenPrinter(printerName, hPrinter, 0&) Then
        di.pDocName = "Meu Documento"
        di.pOutputFile = vbNullString
        di.pDataType = "RAW"
        If StartDocPrinter(hPrinter, 1, di) Then
            If StartPagePrinter(hPrinter) Then
                WritePrinter hPrinter, fileContent(0), UBound(fileContent) + 1, bytesWritten
                EndPagePrinter hPrinter
            End If
            EndDocPrinter hPrinter
        End If
        ClosePrinter hPrinter
    Else
        MsgBox "Não foi possível abrir a impressora."
        End
    End If
End Sub

Public Sub AddTrayIcon()
    Shell_NotifyIcon NIM_ADD, trayIcon
End Sub

Sub SetupTrayIcon()
    With trayIcon
        .cbSize = Len(trayIcon)
        .hWnd = Form1.hWnd ' Referencia o hWnd do Form1
        .uID = 1 ' Um identificador único para o ícone
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .hIcon = Form1.Icon ' Usa o ícone do Form1
        .szTip = "Monitor do OrCarro" & Chr$(0)
        .uCallbackMessage = WM_TRAYMESSAGE
    End With
    TrayCallbackMessage = RegisterWindowMessage("TrayIconCallbackMessage")
    lpPrevWndProc = SetWindowLong(Form1.hWnd, GWL_WNDPROC, AddressOf TrayIconCallback)
End Sub

Sub RestoreWindowProc()
    Call SetWindowLong(Form1.hWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Sub Main()
    Dim iniFileName As String
    Dim lastRunDate As String
    Dim currentDate As String
    Dim AusTimer As String

    Load Form1
    Form1.Visible = False
    SetupTrayIcon
    iniFileName = App.Path & "\monitor.ini"

    iniFileName = App.Path & "\monitor.ini"
    lastRunDate = ReadIniValue("Config", "LastRunDate", iniFileName)
    currentDate = Format(Date, "yyyymmdd")
    If lastRunDate <> currentDate Then
        PerformMaintenance
        WriteIniValue "Config", "LastRunDate", currentDate, iniFileName
    End If
    NivelLog = Val(ReadIniValue("Config", "NivelLog", iniFileName))
    networkPath = ReadIniValue("Config", "NetworkPath", iniFileName)
    Tipo = ReadIniValue("Config", "Tipo", iniFileName)
    
    AusTimer = ReadIniValue("Config", "TimerInterval", iniFileName)
    If AusTimer = "" Then AusTimer = "1000"
    intervaloTimer = CLng(AusTimer)
    Ver = ReadIniValue("Config", "Ver", iniFileName)
    
    StartTimer intervaloTimer
    AddTrayIcon
End Sub

Private Sub PerformMaintenance()
    ' Verificar e criar a pasta Log se não existir
    Dim logFolderPath As String
    logFolderPath = App.Path & "\Log"
    If Dir(logFolderPath, vbDirectory) = "" Then
        MkDir logFolderPath
    End If

    ' Mover o log
    Dim logFileName As String
    Dim newLogFileName As String
    logFileName = App.Path & "\monitor.log"
    newLogFileName = logFolderPath & "\monitor_" & Format(Date, "yyyymmdd") & ".log"
    If Dir(logFileName) <> "" Then
        FileCopy logFileName, newLogFileName
        Kill logFileName
    End If

    ' Apagar arquivos em Bak que têm mais de 48 horas
    Dim fileName As String
    Dim filePath As String
    Dim fileCreationDate As Date

    fileName = Dir(App.Path & "\Bak\*.*")
    Do While fileName <> ""
        filePath = App.Path & "\Bak\" & fileName
        fileCreationDate = FileDateTime(filePath)

        ' Verifica se a diferença de tempo é maior que 48 horas
        If DateDiff("h", fileCreationDate, Now) > 48 Then
            Kill filePath
        End If

        fileName = Dir() ' Próximo arquivo
    Loop
End Sub

Function GetPrinterName() As String
    Dim printerName As String
    Dim bufferSize As Long

    bufferSize = 254
    printerName = Space$(bufferSize)
    GetDefaultPrinter printerName, bufferSize
    GetPrinterName = Left$(printerName, InStr(printerName, Chr$(0)) - 1)
End Function

Public Sub WriteToLog(message As String, Nivel As Integer)
    Dim logFilePath As String
    Dim fileNumber As Integer
    
    If Nivel <= NivelLog Then
        logFilePath = App.Path & "\monitor.txt"
        fileNumber = FreeFile ' Obter um número de arquivo livre
    
        Open logFilePath For Append As #fileNumber ' Abre o arquivo para adição de conteúdo
        Print #fileNumber, Format(Now, "yyyy-mm-dd HH:MM:SS") & " - " & message ' Escreve data/hora e mensagem
        Close #fileNumber ' Fecha o arquivo
    End If
End Sub

Public Sub DelayedKill(filePath As String)
    Sleep intervaloTimer
    On Error Resume Next ' Para evitar erro se o arquivo já foi excluído ou está inacessível
    Kill filePath
    On Error GoTo 0 ' Desativa o gerenciamento de erro
End Sub

Public Sub WriteIniValue(section As String, key As String, value As String, iniFileName As String)
WritePrivateProfileString section, key, value, iniFileName
End Sub

