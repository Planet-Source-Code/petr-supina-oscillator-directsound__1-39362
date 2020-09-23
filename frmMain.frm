VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscillator 1.0"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMod 
      Caption         =   "Amplitude Modulation"
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.HScrollBar hscMod 
         Height          =   200
         LargeChange     =   10
         Left            =   120
         Max             =   100
         TabIndex        =   4
         Top             =   280
         Width           =   3015
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Details"
      ForeColor       =   &H00800080&
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   5415
      Begin VB.Label lblDetails 
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame fraOsc 
      Caption         =   "Oscilloscope"
      ForeColor       =   &H00800080&
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   5415
      Begin VB.PictureBox picOsc 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004000&
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   5175
      End
      Begin VB.Line lin2 
         BorderColor     =   &H000000FF&
         X1              =   5280
         X2              =   5280
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line lin1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   120
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line linGraph 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5280
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblGraph 
         Alignment       =   2  'Center
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   5175
      End
   End
   Begin VB.Frame fraOpt 
      Caption         =   "Oscillator"
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optOsc 
         Appearance      =   0  'Flat
         Caption         =   "Square"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optOsc 
         Appearance      =   0  'Flat
         Caption         =   "Sine"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraVolume 
      Caption         =   "Volume"
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   5415
      Begin VB.HScrollBar hscVol 
         Height          =   200
         LargeChange     =   50
         Left            =   120
         Max             =   1000
         TabIndex        =   9
         Top             =   280
         Value           =   100
         Width           =   4335
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame fraFreq 
      Caption         =   "Frequency"
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5415
      Begin VB.HScrollBar hscFreq 
         Height          =   200
         LargeChange     =   50
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   6
         Top             =   280
         Value           =   185
         Width           =   4335
      End
      Begin VB.Label lblFreq 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Timer tmrMod 
      Left            =   2160
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const nSamples = 44100
Const nBasicBufferSize = 4096
Const pi = 3.14159265358979

Dim DX7 As New DirectX7, DS As DirectSound, DSB As DirectSoundBuffer
Dim PCM As WAVEFORMATEX, DSBD As DSBUFFERDESC
Dim nFreq&, nMod!, nModDir%

Private Sub SinBuffer(ByVal nFrequency&, ByVal nVolume!, Optional ByVal bSquare As Boolean)
Dim lpBuffer() As Byte, I&, C!, nBuffer&
lblFreq = FormatNumber(nFreq, 0) & " Hz"
lblVol = FormatPercent(nVolume, 0)
C = nSamples / nFrequency
nBuffer = (nBasicBufferSize \ C) * C
If nBuffer = 0 Then nBuffer = C
ReDim lpBuffer(nBuffer - 1)
For I = 0 To nBuffer - 1
C = Sin(I * 2 * pi / nSamples * nFrequency)
If bSquare Then
C = Sgn(C)
If C = 0 Then C = 1
End If
lpBuffer(I) = (C * nMod * nVolume + 1) * 127.5!
Next
If DSBD.lBufferBytes <> nBuffer Then
DSBD.lBufferBytes = nBuffer
Set DSB = DS.CreateSoundBuffer(DSBD, PCM)
End If
DSB.WriteBuffer 0, 0, lpBuffer(0), DSBLOCK_ENTIREBUFFER
DSB.Play DSBPLAY_LOOPING
lblDetails = "Channels: 1 (Mono)" & vbCrLf _
& "Bits per sample: 8" & vbCrLf _
& "Samples per second: " & FormatNumber(44100, 0) & vbCrLf _
& "DirectSound buffer size: " & FormatNumber(nBuffer, 0) & " bytes" & vbCrLf _
& "Period: " & FormatNumber(1000 / nFrequency, 3) & " ms"
C = 1000
Do While nFrequency * 20 > picOsc.ScaleWidth
nFrequency = nFrequency \ 2
C = C / 2
Loop
lblGraph = FormatNumber(C, 1) & " ms"
picOsc.Cls
picOsc.Line (0, picOsc.ScaleHeight \ 2)-(picOsc.ScaleWidth, picOsc.ScaleHeight \ 2), &H8000&
picOsc.Line (0, (picOsc.ScaleHeight \ 2) * (1 - nVolume))-(picOsc.ScaleWidth, (picOsc.ScaleHeight \ 2) * (1 - nVolume)), &H6000&
picOsc.Line (0, (picOsc.ScaleHeight \ 2) * (1 + nVolume))-(picOsc.ScaleWidth, (picOsc.ScaleHeight \ 2) * (1 + nVolume)), &H6000&
If optOsc(1).Value Then
For I = 0 To picOsc.ScaleWidth
C = Sgn(Sin(I / picOsc.ScaleWidth * pi * 2 * nFrequency))
If C = 0 Then C = 1
picOsc.PSet (I, ((picOsc.ScaleHeight - 1) \ 2) * (1 - C * nMod * nVolume)), vbGreen
Next
Else
picOsc.Line (0, picOsc.ScaleHeight \ 2)-(0, picOsc.ScaleHeight \ 2)
For I = 0 To picOsc.ScaleWidth
picOsc.Line -(I, (picOsc.ScaleHeight \ 2) * (1 - Sin(I / picOsc.ScaleWidth * pi * 2 * nFrequency) * nMod * nVolume)), vbGreen
Next
End If
Refresh
End Sub

Private Sub Form_Load()
nMod = 1
Set DS = DX7.DirectSoundCreate(vbNullString)
DS.SetCooperativeLevel hWnd, DSSCL_NORMAL
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = nSamples
PCM.nBitsPerSample = 8
PCM.nBlockAlign = 1
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
DSBD.lFlags = DSBCAPS_STATIC
hscFreq_Scroll
End Sub

Private Sub hscFreq_Change()
hscFreq_Scroll
End Sub

Private Sub hscFreq_Scroll()
nFreq = 1 + hscFreq.Value * 22.049! * Log(1 + hscFreq.Value / 1000) / Log(2)
SinBuffer nFreq, hscVol.Value / 1000, optOsc(1).Value
End Sub

Private Sub hscMod_Change()
hscMod_Scroll
End Sub

Private Sub hscMod_Scroll()
If hscMod.Value = 0 Then
tmrMod.Interval = 0
nMod = 1
Else
tmrMod.Interval = 1
End If
End Sub

Private Sub hscVol_Change()
hscVol_Scroll
End Sub

Private Sub hscVol_Scroll()
SinBuffer nFreq, hscVol.Value / 1000, optOsc(1).Value
End Sub

Private Sub optOsc_Click(Index As Integer)
SinBuffer nFreq, hscVol.Value / 1000, optOsc(1).Value
End Sub

Private Sub tmrMod_Timer()
If nModDir >= 0 Then
nMod = nMod + 0.2! / (101 - hscMod.Value)
If nMod > 1 Then nMod = 1: nModDir = -1
Else
nMod = nMod - 0.2! / (101 - hscMod.Value)
If nMod < -1 Then nMod = -1: nModDir = 1
End If
SinBuffer nFreq, hscVol.Value / 1000, optOsc(1).Value
End Sub
