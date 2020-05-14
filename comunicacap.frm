VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmFTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização 5.30"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "comunicacap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command16 
      Caption         =   "resolvedor cont arquivos"
      Height          =   375
      Left            =   12600
      TabIndex        =   44
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "contar arquivos"
      Height          =   615
      Left            =   12360
      TabIndex        =   43
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "caixa do windos"
      Height          =   495
      Left            =   13680
      TabIndex        =   42
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "matar processo"
      Height          =   735
      Left            =   15720
      TabIndex        =   40
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "mover"
      Height          =   735
      Left            =   15360
      TabIndex        =   39
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "apagar order"
      Height          =   495
      Left            =   14040
      TabIndex        =   38
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "apagar diretorio"
      Height          =   735
      Left            =   12720
      TabIndex        =   37
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "criar um arquivo"
      Height          =   735
      Left            =   14040
      TabIndex        =   36
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cria a pasta"
      Height          =   495
      Left            =   12600
      TabIndex        =   35
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "apagar a pasta"
      Height          =   495
      Left            =   14760
      TabIndex        =   34
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   7800
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   375
         Left            =   5280
         TabIndex        =   41
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Instalar"
         Height          =   495
         Left            =   5280
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   1680
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transmitindo.."
         Height          =   975
         Left            =   1800
         TabIndex        =   28
         Top             =   120
         Width           =   3255
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Atualização do sistema Order Taker"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A ultima atualização ainda esta válida no sistema!"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Iniciando...."
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   120
         Picture         =   "comunicacap.frx":0A02
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7800
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   11520
      TabIndex        =   24
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   11760
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton cmdRecebe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "<--"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton cmdEnvia 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-->"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.Frame fraLocal 
      Caption         =   "Local Host"
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
      Begin VB.FileListBox filList 
         Height          =   3600
         Left            =   2040
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.DirListBox dirList 
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Host Remoto"
      Height          =   4215
      Left            =   5040
      TabIndex        =   11
      Top             =   3600
      Width           =   8535
      Begin VB.CommandButton Command5 
         Caption         =   "acesso suport"
         Height          =   375
         Left            =   3960
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   3600
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdMkDir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&MkDir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ListBox lstRemote 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2595
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblRemoteDirectory 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   6735
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Pronto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame fraFTP 
      Caption         =   "FTP :"
      Height          =   1575
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1620
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1620
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   5
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblUsername 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   4
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Endereço :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   3
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdConecta 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Conectar com o Host Remoto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   5775
   End
   Begin VB.TextBox txtLog 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   8415
   End
   Begin InetCtlsObjects.Inet itcFTP 
      Left            =   7680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
 

Dim lret As Long
Dim fileop As SHFILEOPSTRUCT
'http://www.macoratti.net/vb_ftp1.htm
Private Sub cmdConecta_Click()
    ' Verifica a propriedade Caption para desconectar ou conectar
    If Left(cmdConecta.Caption, 4) = "&Con" Then
        conectaHost
    Else
        desconectaHost
    End If
End Sub

Private Sub cmdDelete_Click()
    
    Dim operacao As String
    Dim nomeArquivo As String
    Dim response As Integer, contador As Integer
    
    response = MsgBox("Confirma a operação ?", vbQuestion + vbYesNo, "Delete")
    If response = vbYes Then
        For contador = 0 To lstRemote.ListCount - 1
            If lstRemote.Selected(contador) = True Then
                nomeArquivo = lstRemote.List(contador)
                ' Verifica se é um diretorio ou arquivo
                If Right(nomeArquivo, 1) = "/" Then
                    operacao = "rmdir " & Left(nomeArquivo, Len(nomeArquivo) - 1)
                Else
                    operacao = "delete " & nomeArquivo
                End If
                executaComando operacao, False
            End If
        Next contador
        listaDir
    End If
End Sub

Private Sub cmdMkDir_Click()
    Dim dir As String, operacao As String
    
    dir = InputBox("Informe o nome da pasta", "Cria Diretório")
    If dir <> "" Then
        operacao = "mkdir " & dir
        executaComando operacao, True
    End If
End Sub

Private Sub cmdRecebe_Click()
    Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    For contador = 0 To lstRemote.ListCount - 1
        If lstRemote.Selected(contador) = True Then
            nomeArquivo = lstRemote.List(contador)
            If Len(dirList.Path) > 3 Then
                arquivoSaida = dirList.Path & "\" & nomeArquivo
            Else
                arquivoSaida = dirList.Path & nomeArquivo
            End If
            operacao = "recv " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
            lstRemote.Selected(contador) = False
        End If
    Next contador
   filList.Refresh
End Sub

Private Sub cmdEnvia_Click()
    Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    For contador = 0 To filList.ListCount - 1
        If filList.Selected(contador) = True Then
            nomeArquivo = filList.List(contador)
            arquivoSaida = lblRemoteDirectory.Caption & "/" & nomeArquivo
            operacao = "send " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
        End If
    Next contador
    listaDir
    filList.Refresh
End Sub




Private Sub Command1_Click()
   Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    
    selecioneTodos
    For contador = 0 To lstRemote.ListCount - 1
        If lstRemote.Selected(contador) = True Then
            nomeArquivo = lstRemote.List(contador)
            If Len(dirList.Path) > 3 Then
            'lugar onde vai ser transferidos
                arquivoSaida = dirList.Path & "\" & nomeArquivo
                Debug.Print arquivoSaida
               
               '  arquivoSaida = "C:\myordertaker\../"
            Else
                arquivoSaida = dirList.Path & nomeArquivo
                 'arquivoSaida = Text1.Text
            End If
            
            operacao = "recv " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
            lstRemote.Selected(contador) = False
        End If
    Next contador
   filList.Refresh
End Sub



Private Sub Command10_Click()

'remover pasta e subpasta
Dim command As String
'command = "c:\windows\notepad.exe"
'command = "del /q \MYorderTaker\*.*"

command = "del/q \Order_taker\*.*"""
Shell "cmd.exe /c " & command

command = "del /q \MYorderTaker\*.*"
Shell "cmd.exe /c " & command

'criarPastaParaAlocacaoDaAtualizacao


End Sub

Private Sub Command11_Click()
'mover
Dim command As String
'command = "c:\windows\notepad.exe"
command = "G:\visualBasic\FtpInet\setup.exe"
Shell "cmd.exe /c " & command
End Sub


Private Sub Command12_Click()
Dim appName As String

Dim Comando As String
appName = "Calc.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
End Sub

Private Sub Command13_Click()
Dim resp As Integer
        
   resp = MsgBox("A atualização falhou devido a configuração de portas deste pc , precisa de ajuda para resolver?", vbYesNo, "Ajuda para resolver o problema")
   If resp = 6 Then
   Call cmdConecta_Click
                                resp = MsgBox("você permite a liberação e transição deste arquivo de atualização", vbYesNo, "Libere a porta #200")
                                   If resp = 6 Then
   Timer1.Interval = 500
                                 MsgBox "verificando se o problema foi resolvido ! ", vbYes, "Robofild"
   
   
   End If
   
   End If
    
End Sub

Private Sub Command14_Click()
'With fileop
 ' .hwnd = 0
  
  '.wFunc = FO_COPY
  
  '.pFrom = txtorigem & vbNullChar & vbNullChar
  
  '.pTo = Txtdestino.Text & vbNullChar & vbNullChar
  
 ' .lpszProgressTitle = "Aguarde, realizando copia..."
  
 ' .fFlags = FOF_SIMPLEPROGRESS Or FOF_RENAMEONCOLLISION

'End With

'lret = SHFileOP(fileop)

'If result <> 0 Then 'a operaçao falhou
 '  MsgBox Err.LastDllError 'exibe o erro retornado pela API
'Else
 ' If fileop.fAnyOperationsAborted <> 0 Then
  '   MsgBox "Operação falhou !!!"
 ' End If
'End If
End Sub

Private Sub Command15_Click()
Dim nomeDoArquivosLidos As String


        Dim itens As Integer
    For itens = 0 To filList.ListCount - 1
    
    filList.Selected(itens) = True
    nomeDoArquivosLidos = filList.FileName
    
    Next itens
    MsgBox itens

End Sub

Private Sub Command2_Click()

   Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    
    selecioneTodos
               
                 arquivoSaida = "C:\myordertaker\../"
            
            operacao = "recv ../" & " " & arquivoSaida
            executaComando operacao, False
            lstRemote.Selected(contador) = False
        
   
   filList.Refresh

End Sub

Private Sub Command3_Click()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "G:\visualBasic\FtpInet\setup.exe"
Shell "cmd.exe /c " & command
End Sub

Private Sub Command4_Click()
liparSuborder_taker
InstalarautoBat
matarprocessoAtualizador
Unload Me
Exit Sub

End Sub

Private Sub Command5_Click()
OperadorDelistasSuporte
End Sub

Private Sub Command6_Click()
Apaga
End Sub



Private Sub Command7_Click()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "MkDir c:\MYordertaker"
Shell "cmd.exe /c " & command
command = "MkDir c:\Order_Taker"
Shell "cmd.exe /c " & command
End Sub

Private Sub Command8_Click()


  'cria a Pasta pelo bat no caminho abaixo
Dim command As String
'command = "c:\windows\notepad.exe"
command = "del /q  \C:\Program Files (x86)\test\*.*"
'command = "del /q \MYorderTaker\*.*"
Shell "cmd.exe /c " & command

End Sub

Private Sub Command9_Click()

Dim command As String
'command = "c:\windows\notepad.exe"
command = "MkDir c:\MYordertaker\"
Shell "cmd.exe /c " & command




End Sub

Private Sub dirList_Change()
    filList.Path = dirList.Path
End Sub

Private Sub drvList_Change()
    On Error GoTo driveError
    
    dirList.Path = drvList.Drive
    Exit Sub
driveError:
    MsgBox Err.Description, vbExclamation, "Drive Error"
End Sub

Private Sub Form_Load()
inicialProcedimentos
 

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
matarprocessoAtualizador
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo trata_erro
    If cmdEnvia.Enabled = True Then
        itcFTP.Execute , "Quit"
    End If
    Exit Sub
trata_erro:
        MsgBox "Erro na aplicação FTP", vbCritical
End Sub

Private Sub itcFTP_StateChanged(ByVal State As Integer)
    Select Case State
        Case icResolvingHost
            lblStatus.Caption = "Resolvendo Host"
            Form1.ProgressBar1.Value = 10
        Case icHostResolved
            lblStatus.Caption = "Host Resolvido"
            Form1.ProgressBar1.Value = 25
        Case icConnecting
            lblStatus.Caption = "Conectando ..."
            Form1.ProgressBar1.Value = 80
        Case icConnected
            lblStatus.Caption = "Conectado"
            Form1.ProgressBar1.Value = 100
        Case icRequesting
            lblStatus.Caption = "Requesitando ..."
            Form1.ProgressBar1.Value = 0
        Case icRequestSent
            lblStatus.Caption = "Requesição enviada"
            Form1.ProgressBar1.Value = 50
        Case icReceivingResponse
            lblStatus.Caption = "Recebendo ..."
            Form1.ProgressBar1.Value = 25
        Case icResponseReceived
            lblStatus.Caption = "Resposta recebida"
            Form1.ProgressBar1.Value = 100
        Case icDisconnecting
            lblStatus.Caption = "Desconectando ..."
            Form1.ProgressBar1.Value = 30
        Case icDisconnected
            lblStatus.Caption = "Desconectado"
            Form1.ProgressBar1.Value = 0
            
        Case icError
          'caso o arquivo já exista ou outros
           'Label3.Visible = True
           'Command4.Visible = True
           
           'Command4.Caption = "Atualizado"
            lblStatus.Caption = itcFTP.ResponseInfo
            'MsgBox lblStatus.Caption
            'desconectaHost
            'Timer2.Interval = 1000
           ' Command4.Caption "Aguarde"
            'Command4.Enabled = False
            
            txtLog.Text = txtLog.Text & itcFTP.ResponseInfo & vbCrLf
            Form1.ProgressBar1.Value = 0
        Case icResponseCompleted
            lblStatus.Caption = "Resolvendo erros"
            txtLog.Text = txtLog.Text & "Resolvendo erros" & vbCrLf
            Form1.ProgressBar1.Value = 100
             'Command4.Caption "Instalar"
            'Command4.Enabled = False
    End Select
    If lblStatus.Caption = "Pronto" Then
    Form1.ProgressBar1.Value = 100
      Form1.ProgressBar1.Value = 100
             'Command4.Caption "Instalar"
            'Command4.Enabled = False
    Call Command3_Click
    End If
    
    
    
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub lblRemoteDirectory_Change()
Text1.Text = lblRemoteDirectory.Caption

End Sub


Private Sub lstRemote_DblClick()
    Dim operacao As String, dir As String
    
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
       ' dir = lstRemote.List(lstRemote.ListIndex)
        dir = lstRemote.List(lstRemote.ListIndex)
        'dir = lstRemote.List(21)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        executaComando operacao, True
    End If
End Sub

Private Sub Text1_Change()
Label2.Caption = Replace(Text1.Text, "public_ftp", "Atualizando...")
End Sub

Private Sub Timer1_Timer()
Timer2.Interval = 0
ProgressBar1.Value = 20
MoverodiretroriodeRecebimento
ProgressBar1.Value = 26
'Form1.Show
conectaHost
ProgressBar1.Value = 50
OperadorDelista
ProgressBar1.Value = 75
Call Command1_Click
ProgressBar1.Value = 100
'InstalarautoBat
' Animation1.Close
Timer1.Interval = 0
'Animation1.Visible = False

resovedorParaArquivosFaltantes2
'desconectaHost
End Sub

Private Sub Timer2_Timer()
conectaHost
mataProcessos
LimparpastaeSubPasta
ProgressBar1.Value = 10
CriarpastaEsubPasta
ProgressBar1.Value = 15


Timer2.Interval = 0
Timer1.Interval = 500

End Sub

Private Sub txtLog_GotFocus()
    txtAddress.SetFocus
End Sub

Private Sub conectaHost()
    Dim operacao  As String
    
     On Error GoTo connectError
    
    If txtAddress.Text = "" Then
      itcFTP.URL = "robofild.info"
      itcFTP.UserName = "robofi61"
      itcFTP.Password = "jCu70q4e6Q"
      listaDir
      ProgressBar1.Value = 28

      cmdEnvia.Enabled = True
      cmdRecebe.Enabled = True
      cmdMkDir.Enabled = True
      cmdDelete.Enabled = True
      lstRemote.Enabled = True
        ProgressBar1.Value = 29
      cmdConecta.Caption = "D&esconectar do Host Remoto"
    Else
      MsgBox "Informe o nome do servidor FTP.", vbCritical
      txtAddress.SetFocus
    End If
      ProgressBar1.Value = 30
    Exit Sub
    
connectError:
    MsgBox Err.Description, vbExclamation, "FTP"
      ProgressBar1.Value = 20
End Sub

Private Sub desconectaHost()
  

   On Error GoTo trata_erro

    Dim operacao  As String
    
    itcFTP.Execute , "quit"
    operacao = "quit"
    executaComando operacao, False
    cmdEnvia.Enabled = False
    cmdRecebe.Enabled = False
    cmdMkDir.Enabled = False
    cmdDelete.Enabled = False
    lstRemote.Enabled = False
    cmdConecta.Caption = "&Conectar ao Host Remoto"
    ProgressBar1.Value = 0
    Exit Sub
trata_erro:
    MsgBox "Erro ao efetuar a operacao com : " & txtAddress.Text & vbCrLf & " erro : " & Err.Number
    
End Sub

Private Sub executaComando(ByVal op As String, ByVal ld As Boolean)
 
    On Error GoTo trata_erro

    If itcFTP.StillExecuting Then
        itcFTP.Cancel
    End If
    txtLog.Text = txtLog.Text & "Comando: " & op & vbCrLf
    itcFTP.Execute , op
    terminaComando
    If ld = True Then
        listaDir
        terminaComando
    End If
    Exit Sub
    
trata_erro:
    MsgBox "Não foi possivel efetuar operacao com : " & txtAddress.Text & vbCrLf & " erro : " & Err.Number
End Sub

Private Sub terminaComando()
    Do While itcFTP.StillExecuting
        DoEvents
    Loop
End Sub

Private Sub listaDir()
    Dim operacao As String
    Dim data As Variant, contador As Integer
    Dim inicio As Integer, length As Integer
    
    inicio = 1
    lstRemote.Clear
    operacao = "dir"
    executaComando operacao, False
    Do
        data = itcFTP.GetChunk(1024, icString)
        Text1.Text = data
        'data = itcFTP.GetChunk(1024, "0")
          
        DoEvents
        For contador = 1 To Len(data)
            If Mid(data, contador, 1) = Chr(13) Then
                If length > 0 And Mid(data, inicio, length) <> "./" Then
                    lstRemote.AddItem Mid(data, inicio, length)
                End If
                inicio = contador + 2
                length = -1
            Else
                length = length + 1
            End If
        Next contador
    Loop While LenB(data) > 0
    operacao = "pwd"
    executaComando operacao, False
    lblRemoteDirectory.Caption = itcFTP.GetChunk(1024, icString)
End Sub


Public Sub selecioneTodos()
Dim itens As Integer
    For itens = 0 To lstRemote.ListCount - 1
     Sleep 5000
    lstRemote.Selected(itens) = True
    Next itens
End Sub

Public Sub OperadorDelista()
Dim contador As Integer
    For contador = 0 To 4
     Sleep 3000
     ProgressBar1.Value = (50 + contador)
    Dim operacao As String, dir As String
    
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
       ' dir = lstRemote.List(lstRemote.ListIndex)
        'dir = lstRemote.List(lstRemote.ListIndex)
         Select Case contador
                Case 0
                dir = lstRemote.List(30)
                ProgressBar1.Value = 52
                    
                Case 1
                dir = lstRemote.List(2)
                ProgressBar1.Value = 55
                Case 2
                dir = lstRemote.List(1)
                ProgressBar1.Value = 62
                Case 3
                dir = lstRemote.List(1)
                ProgressBar1.Value = 75
                Case 4
        End Select
        'dir = lstRemote.List(21)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        executaComando operacao, True
    End If
    
     Next contador
     selecioneTodos
     Form1.ProgressBar1.Value = 50
   Sleep 3000
'causes program to pause for 3 seconds
     DownloadAutomatico
     
End Sub
Public Sub OperadorDelistasSuporte()
Dim contador As Integer
    For contador = 0 To 4
     ProgressBar1.Value = (50 + contador)
    Dim operacao As String, dir As String
    
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
       ' dir = lstRemote.List(lstRemote.ListIndex)
        'dir = lstRemote.List(lstRemote.ListIndex)
         Select Case contador
                Case 0
                dir = lstRemote.List(23)
                ProgressBar1.Value = 52
                    
                Case 1
                dir = lstRemote.List(1)
                ProgressBar1.Value = 55
                Case 2
                dir = lstRemote.List(1)
                ProgressBar1.Value = 62
                Case 3
                dir = lstRemote.List(1)
                ProgressBar1.Value = 75
                Case 4
                 dir = lstRemote.List(3)
                ProgressBar1.Value = 80
        End Select
        'dir = lstRemote.List(21)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        executaComando operacao, True
    End If
    
     Next contador
     selecioneTodos
     Form1.ProgressBar1.Value = 50
     DownloadAutomatico
     
End Sub


Public Sub DownloadAutomatico()
  Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    selecioneTodos
    ProgressBar1.Value = 80
    For contador = 0 To lstRemote.ListCount - 1
       Sleep 3000
'causes program to pause for 3 seconds
        If lstRemote.Selected(contador) = True Then
            nomeArquivo = lstRemote.List(contador)
            If Len(dirList.Path) > 3 Then
            
               Form1.ProgressBar1.Value = (50 + contador)
            'lugar onde vai ser transferidos
                arquivoSaida = dirList.Path & "\" & nomeArquivo
                 ProgressBar1.Value = 85
                 
            Else
                arquivoSaida = dirList.Path & nomeArquivo
                 'arquivoSaida = Text1.Text
            End If
            operacao = "recv " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
            lstRemote.Selected(contador) = False
        End If
        Form1.ProgressBar1.Value = (100)
    Next contador
   filList.Refresh
   Form1.ProgressBar1.Value = (0)
End Sub

Public Sub InstalarautoBat()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "C:\MYordertaker\setup.exe"
Shell "cmd.exe /c " & command

End Sub

Public Sub criarPastaParaAlocacaoDaAtualizacao()
'cria a Pasta pelo bat no caminho abaixo
Dim command As String
'command = "c:\windows\notepad.exe"
command = " C:\Users\Developers\Desktop\GerasPstOrder.bat"
Shell "cmd.exe /c " & command
End Sub

Public Sub MoverodiretroriodeRecebimento()
'caminho da pasta para descarregar a atualizacao
'Dir1.Path = Left $ (Drive1.Drive, 1) e ": \"
  dirList.Path = "C:\MYordertaker"
  ProgressBar1.Value = 23
  filList.Path = dirList.Path
   ProgressBar1.Value = 25
'filList.Path = "C:\myordertaker"

End Sub



Public Sub Apaga()

'cria a Pasta pelo bat no caminho abaixo
Dim command As String
'command = "c:\windows\notepad.exe"
command = "del /q \MYorderTaker\*.*"

Shell "cmd.exe /c " & command
criarPastaParaAlocacaoDaAtualizacao
End Sub


Public Sub liparSuborder_taker()



'cria a Pasta pelo bat no caminho abaixo
Dim command As String
'command = "c:\windows\notepad.exe"
command = "C:\Users\Developers\Desktop\delMyOrderTaker.bat"

Shell "cmd.exe /c " & command
'criarPastaParaAlocacaoDaAtualizacao
'
End Sub

Public Sub CriarpastaEsubPasta()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "MkDir c:\MYordertaker"
Shell "cmd.exe /c " & command
command = "MkDir c:\Order_Taker"
Shell "cmd.exe /c " & command
End Sub

Public Sub LimparpastaeSubPasta()
'
'remover pasta e subpasta
Dim command As String
'command = "c:\windows\notepad.exe"
'command = "del /q \MYorderTaker\*.*"

command = "del/q \Order_taker\*.*"""
Shell "cmd.exe /c " & command

command = "del /q \MYorderTaker\*.*"
Shell "cmd.exe /c " & command

'criarPastaParaAlocacaoDaAtualizacao
End Sub

Public Sub mataProcessos()
Dim appName As String

Dim Comando As String
appName = "Amarelinho.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
Shell Comando
Shell Comando
End Sub

Public Sub inicialProcedimentos()
mataProcessos
LimparpastaeSubPasta
ProgressBar1.Value = 10
CriarpastaEsubPasta
ProgressBar1.Value = 15
End Sub

Public Sub resovedorParaArquivosFaltantes()
Dim nomeDoArquivosLidos As String

    Dim amarelinho, setup, SetupLst As Boolean
    
    Dim itens As Integer
    For itens = 0 To filList.ListCount - 1
    
    filList.Selected(itens) = True
    nomeDoArquivosLidos = filList.FileName
    
    Next itens
  
        If itens >= 3 Then
        desconectaHost
        Else
        'verificar os principais aquivos para encontar o que esta falatando
            For itens = 0 To filList.ListCount - 1
    
            filList.Selected(itens) = True
            nomeDoArquivosLidos = filList.FileName
            
               Select Case nomeDoArquivosLidos
            Case "Amarelinho.CAB"
               amarelinho = True
            Case "setup.exe"
                 setup = True
            Case "SETUP.LST"
                 SetupLst = True
          
                End Select
                
                
            
            Next itens
        
        
        'verificar qual nao foi true
         If amarelinho <> True Then
         'refaça a rota buscando este arquivo
         End If
           If setup <> True Then
           desconectaHost
           conectaHost
           busquepelosetup ("setup.exe")
            'refaça a rota buscando este arquivo
         End If
           If SetupLst <> True Then
            'refaça a rota buscando este arquivo
         End If
  End If
  
End Sub

Public Sub busquepelosetup(ArquivoFaltante)
conectaHost
'operdor desta busca
Dim contador As Integer
    For contador = 0 To 4
     ProgressBar1.Value = (50 + contador)
    Dim operacao As String, dir As String
    
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
       ' dir = lstRemote.List(lstRemote.ListIndex)
        'dir = lstRemote.List(lstRemote.ListIndex)
         Select Case contador
                Case 0
                dir = lstRemote.List(23)
                ProgressBar1.Value = 52
                    
                Case 1
                dir = lstRemote.List(1)
                ProgressBar1.Value = 55
                Case 2
                dir = lstRemote.List(1)
                ProgressBar1.Value = 62
                Case 3
                dir = lstRemote.List(1)
                ProgressBar1.Value = 75
                Case 4
                 dir = lstRemote.List(3)
                ProgressBar1.Value = 80
        End Select
        'dir = lstRemote.List(21)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        'dir = lstRemote.List(1)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        executaComando operacao, True
    End If
    
     Next contador
   selecioneTodossetup (ArquivoFaltante)
     Form1.ProgressBar1.Value = 50
     DownloadAutomatico
     desconectaHost
End Sub
Public Sub selecioneTodossetup(ArquivoFaltante As String)
Dim itens As Integer
    For itens = 0 To lstRemote.ListCount - 1
    If filList.FileName = ArquivoFaltante Then
    lstRemote.Selected(itens) = True
    End If
    Next itens
End Sub

Public Sub resovedorParaArquivosFaltantes2()
   Dim resp As Integer
    Dim itens As Integer
    For itens = 0 To filList.ListCount - 1
    
    filList.Selected(itens) = True

    
    Next itens
  
        If itens >= 3 Then
         Timer1.Interval = 0
           Label3.Caption = "Aguarde já esta quase teminando"""
          desconectaHost
            Label3.Visible = False
           resp = MsgBox("Os erros foram corrigidos, Deseja seguir para instalação da atualização ", vbYesNo, "Instalar atualizações")
            If resp = 6 Then
             Command4.Visible = True
             Call Command4_Click
             
            Else
          Command4.Visible = True
            End If
       Else
       Label3.Visible = True
       
       Label3.Caption = "Resolvendo problemas em seu pc para otimizar a instalação"
       desconectaHost
       Timer1.Interval = 500
       
       End If


End Sub

Public Sub matarprocessoAtualizador()
Dim appName As String

Dim Comando As String
appName = "FTPInternetControl.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
appName = "FTPInternetControl.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
End Sub

