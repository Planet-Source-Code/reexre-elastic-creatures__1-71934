VERSION 5.00
Begin VB.Form frmMAIN
Caption         =   "Form1"
ClientHeight    =   9885
ClientLeft      =   60
ClientTop       =   450
ClientWidth     =   13935
Icon            =   "frmMAIN.frx":0000
LinkTopic       =   "Form1"
ScaleHeight     =   9885
ScaleWidth      =   13935
StartUpPosition =   1 'CenterOwner
Begin VB.TextBox cNAME
Height          =   375
Left            =   1320
TabIndex        =   45
Text            =   "Creature"
Top             =   8400
Visible         =   0 'False
Width           =   2655
End
Begin BOUNCY_CLS.MINI MINI
Height          =   1215
Left            =   5880
TabIndex        =   50
Top             =   8520
Width           =   4815
_ExtentX        =   8493
_ExtentY        =   1296
End
Begin VB.PictureBox picPREV
Appearance      =   0 'Flat
AutoRedraw      =   -1 'True
BackColor       =   &H00000000&
FillStyle       =   0 'Solid
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   13.5
Charset         =   0
Weight          =   400
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
ForeColor       =   &H0000FFFF&
Height          =   975
Left            =   0
ScaleHeight     =   63
ScaleMode       =   3 'Pixel
ScaleWidth      =   63
TabIndex        =   49
Top             =   0
Visible         =   0 'False
Width           =   975
End
Begin VB.CommandButton cmdSaveProject
Caption         =   "Save Project"
Height          =   735
Left            =   1680
TabIndex        =   48
Top             =   8760
Width           =   1335
End
Begin VB.CommandButton cmdEDITnewCreature
Caption         =   "Start Edit NEW Creature"
Height          =   735
Left            =   3120
TabIndex        =   47
Top             =   8760
Width           =   855
End
Begin VB.CommandButton cmdLOADCreature
Caption         =   "Load Creature"
Height          =   375
Left            =   4200
TabIndex        =   46
Top             =   8520
Width           =   1455
End
Begin VB.CommandButton cmdSaveCreature
Caption         =   "Save Creature"
Height          =   375
Left            =   4200
TabIndex        =   44
Top             =   9120
Width           =   1455
End
Begin VB.FileListBox File1
Height          =   4185
Left            =   1320
TabIndex        =   43
Top             =   4200
Visible         =   0 'False
Width           =   2655
End
Begin VB.CommandButton cmdLoadProject
Caption         =   "Load Project"
Height          =   735
Left            =   240
TabIndex        =   42
Top             =   8760
Width           =   1335
End
Begin VB.Frame fPoint
Caption         =   "POINT Parameters"
Height          =   2535
Left            =   11160
TabIndex        =   27
Top             =   6720
Visible         =   0 'False
Width           =   2295
Begin VB.CheckBox ChGlue
Caption         =   "Glue"
Height          =   375
Left            =   120
TabIndex        =   35
Top             =   1320
Width           =   855
End
Begin VB.CheckBox chDELA
Caption         =   "Delaunay"
Height          =   375
Left            =   120
TabIndex        =   34
Top             =   960
Width           =   1095
End
Begin VB.TextBox txtFaseLen
Height          =   285
Left            =   1680
TabIndex        =   31
Text            =   "0"
ToolTipText     =   "How Long Stay FIX (0-1)"
Top             =   2160
Width           =   495
End
Begin VB.ComboBox cmbDFpoint
Height          =   315
Left            =   720
TabIndex        =   30
Text            =   "Combo1"
ToolTipText     =   "Phase (0-1) *  360°"
Top             =   2160
Width           =   825
End
Begin VB.TextBox txtDTpoint
Height          =   285
Left            =   120
TabIndex        =   29
Text            =   "0"
ToolTipText     =   "Dynamic Speed"
Top             =   2160
Width           =   495
End
Begin VB.CheckBox chAUTOpoint
Caption         =   "UpDate on MouseOver"
Height          =   375
Left            =   120
Style           =   1 'Graphical
TabIndex        =   28
Top             =   240
Width           =   1935
End
Begin VB.CheckBox chPFIX
Caption         =   "Fixed"
Height          =   375
Left            =   120
TabIndex        =   33
Top             =   600
Width           =   855
End
Begin VB.Label Label15
Caption         =   "FIX"
Height          =   255
Left            =   1680
TabIndex        =   41
Top             =   1920
Width           =   495
End
Begin VB.Label Label14
Caption         =   "Phase"
Height          =   255
Left            =   720
TabIndex        =   40
Top             =   1920
Width           =   495
End
Begin VB.Label Label13
Caption         =   "Speed"
Height          =   255
Left            =   120
TabIndex        =   39
Top             =   1920
Width           =   495
End
Begin VB.Label Label10
Caption         =   "Cicle Dynamics"
Height          =   255
Left            =   120
TabIndex        =   32
Top             =   1680
Width           =   1575
End
End
Begin VB.PictureBox picGravity
Appearance      =   0 'Flat
BackColor       =   &H0000C000&
ForeColor       =   &H80000008&
Height          =   615
Left            =   13200
ScaleHeight     =   39
ScaleMode       =   3 'Pixel
ScaleWidth      =   39
TabIndex        =   18
ToolTipText     =   "WorldGravity"
Top             =   2040
Width           =   615
End
Begin VB.TextBox Text1
BeginProperty Font
Name            =   "Courier New"
Size            =   9
Charset         =   0
Weight          =   400
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   2895
Left            =   11040
MultiLine       =   -1 'True
TabIndex        =   17
Text            =   "frmMAIN.frx":030A
Top             =   6360
Width           =   2775
End
Begin VB.ComboBox cmbSTIFF
Height          =   315
Left            =   11040
TabIndex        =   16
Text            =   "Combo1"
Top             =   6000
Width           =   975
End
Begin VB.Timer Timer1
Enabled         =   0 'False
Interval        =   25
Left            =   11760
Top             =   240
End
Begin VB.CommandButton cmdCLEAR
Caption         =   "Clear ALL"
Height          =   735
Left            =   12240
TabIndex        =   6
ToolTipText     =   "Clear Project"
Top             =   120
Width           =   735
End
Begin VB.CommandButton cmdLOAD
Caption         =   "Load"
Enabled         =   0 'False
Height          =   735
Left            =   13320
TabIndex        =   5
ToolTipText     =   "Load"
Top             =   840
Visible         =   0 'False
Width           =   735
End
Begin VB.CheckBox chSimulate
Caption         =   "Simulate"
Height          =   615
Left            =   10920
Style           =   1 'Graphical
TabIndex        =   4
ToolTipText     =   "Start / Stop Simulation"
Top             =   960
Width           =   2055
End
Begin VB.CommandButton cmdSAVE
Caption         =   "Save"
Enabled         =   0 'False
Height          =   735
Left            =   13320
TabIndex        =   3
ToolTipText     =   "Save Current State"
Top             =   120
Visible         =   0 'False
Width           =   735
End
Begin VB.VScrollBar PLdraw
Height          =   975
Left            =   10920
Max             =   1
TabIndex        =   1
Top             =   1680
Width           =   375
End
Begin VB.PictureBox PIC
AutoRedraw      =   -1 'True
BackColor       =   &H80000007&
ForeColor       =   &H80000003&
Height          =   8175
Left            =   120
ScaleHeight     =   541
ScaleMode       =   3 'Pixel
ScaleWidth      =   709
TabIndex        =   0
Top             =   120
Width           =   10695
End
Begin VB.Frame fLINK
Caption         =   "LINK Parameters"
Height          =   2775
Left            =   10920
TabIndex        =   7
Top             =   2760
Visible         =   0 'False
Width           =   2295
Begin VB.CheckBox chAuto
Caption         =   "UpDate on MouseOver"
Height          =   375
Left            =   120
Style           =   1 'Graphical
TabIndex        =   25
Top             =   240
Width           =   1935
End
Begin VB.ComboBox cmbDF
Height          =   315
Left            =   1440
TabIndex        =   24
Text            =   "Combo1"
ToolTipText     =   "Phase (0-1) *  360°"
Top             =   2400
Width           =   825
End
Begin VB.TextBox txtDT
Height          =   285
Left            =   840
TabIndex        =   23
Text            =   "0"
ToolTipText     =   "Dynamic Speed"
Top             =   2400
Width           =   495
End
Begin VB.TextBox txtDL
Height          =   285
Left            =   120
TabIndex        =   22
Text            =   "0"
ToolTipText     =   "Dynamic Length (Multiple)"
Top             =   2400
Width           =   615
End
Begin VB.CommandButton cmdSetToAll
Caption         =   "Set these Values to ALL Links"
Height          =   975
Left            =   1440
TabIndex        =   14
Top             =   960
Width           =   735
End
Begin VB.ComboBox cmbBREAK
Height          =   315
Left            =   120
TabIndex        =   11
Text            =   "Combo1"
Top             =   960
Width           =   975
End
Begin VB.ComboBox cmbSPRING
Height          =   315
Left            =   120
TabIndex        =   10
Text            =   "Combo1"
Top             =   1560
Width           =   975
End
Begin VB.Label Label12
Caption         =   "Phase"
Height          =   255
Left            =   1440
TabIndex        =   38
Top             =   2160
Width           =   495
End
Begin VB.Label Label11
Caption         =   "Speed"
Height          =   255
Left            =   840
TabIndex        =   37
Top             =   2160
Width           =   495
End
Begin VB.Label Label9
Caption         =   "Len"
Height          =   255
Left            =   120
TabIndex        =   36
Top             =   2160
Width           =   495
End
Begin VB.Line Line2
X1              =   1080
X2              =   1440
Y1              =   1920
Y2              =   1440
End
Begin VB.Line Line1
X1              =   1080
X2              =   1440
Y1              =   960
Y2              =   1440
End
Begin VB.Label Label7
Caption         =   "Cicle Dynamics"
Height          =   255
Left            =   120
TabIndex        =   21
Top             =   1920
Width           =   1695
End
Begin VB.Label Label2
Caption         =   "Spring Strength"
Height          =   255
Left            =   120
TabIndex        =   9
Top             =   1320
Width           =   1815
End
Begin VB.Label Label1
Caption         =   "BreakLength  (Multiple)"
Height          =   255
Left            =   120
TabIndex        =   8
Top             =   720
Width           =   1815
End
End
Begin VB.CheckBox chLFix
Caption         =   "Fixed"
Height          =   375
Left            =   11400
TabIndex        =   12
Top             =   2400
Visible         =   0 'False
Width           =   855
End
Begin VB.Label Label8
Caption         =   "Gravity"
Height          =   255
Left            =   13200
TabIndex        =   26
Top             =   2640
Width           =   615
End
Begin VB.Label Label16
Caption         =   "Creature Selector"
Height          =   255
Left            =   5880
TabIndex        =   51
Top             =   8280
Width           =   1815
End
Begin VB.Label Label6
Caption         =   "vy"
Height          =   255
Left            =   13200
TabIndex        =   20
Top             =   1800
Width           =   615
End
Begin VB.Label Label5
Caption         =   "vx"
Height          =   255
Left            =   13200
TabIndex        =   19
Top             =   1560
Width           =   615
End
Begin VB.Label Label4
Caption         =   "Body Stiffness"
Height          =   255
Left            =   11040
TabIndex        =   15
Top             =   5760
Width           =   1095
End
Begin VB.Label Label3
Alignment       =   2 'Center
Appearance      =   0 'Flat
BackColor       =   &H80000005&
BorderStyle     =   1 'Fixed Single
Caption         =   "LEFT Click to ADD. RIGHT Click to DELETE"
BeginProperty Font
Name            =   "Arial"
Size            =   6.75
Charset         =   0
Weight          =   400
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
ForeColor       =   &H80000008&
Height          =   735
Left            =   10920
TabIndex        =   13
Top             =   120
Width           =   855
End
Begin VB.Label PLlabel
Alignment       =   2 'Center
BackColor       =   &H80000013&
Caption         =   "Point"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   12
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   375
Left            =   11400
TabIndex        =   2
Top             =   1680
Width           =   855
End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CRE As clsBODY

Dim L As tLINK
Dim P As tPOINT
Dim i As Integer
Dim NW As Integer
Dim NH As Integer

Dim II As Integer
Dim FirstPoint As Boolean
Dim nP1 As Integer
Dim nP2 As Integer

Dim PervLINK As Integer
Dim PervPOINT As Integer

Dim Command As String

Dim Mous As Vector2
Dim copyP(100) As tPOINT



Private Sub chDELA_Click()
If chDELA.Value = Checked Then
    ChGlue.Value = Unchecked
    ChGlue.Visible = False
    B(ICre).FirstDELALink = B(ICre).NLink + 1
    fLINK.Visible = True
End If

If chDELA.Value = Unchecked Then
    ChGlue.Visible = True
    clearDELAresult
    B(ICre).UnCheckDELApoints
    fLINK.Visible = False
End If


End Sub

Private Sub chPFIX_Click()
If chPFIX.Value = Checked Then
    txtFaseLen = 1
Else
    txtFaseLen = 0
End If


End Sub

Private Sub chSimulate_Click()
If chSimulate.Value = vbChecked Then
    
    For BB = 1 To NCre
        B(BB).FindCROSSLinks_InitialState
    Next
    
    Timer1.Enabled = True
    With picGravity
        picGravity.Circle (.ScaleWidth \ 2, .ScaleHeight \ 2), 3, 0
        picGravity.Line (.ScaleWidth \ 2, .ScaleHeight \ 2)- _
                (.ScaleWidth \ 2 + GravityX * .ScaleWidth * 4, _
                .ScaleHeight \ 2 + GravityY * .ScaleHeight * 4), 0
        Label5 = "vx" & Format(GravityX * 10, "0.0")
        Label6 = "vy" & Format(GravityY * 10, "0.0")
    End With
    
Else
    Timer1.Enabled = False
End If

chDELA.Value = Unchecked
chDELA_Click

End Sub

Private Sub cmbBREAK_LostFocus()
Dim Found As Boolean
Dim s As String
s = cmbBREAK
For i = 0 To cmbBREAK.ListCount - 1
    cmbBREAK.ListIndex = i
    If cmbBREAK = s Then Found = True: Exit For
Next
If Not (Found) Then cmbBREAK.AddItem s: cmbBREAK.ListIndex = cmbBREAK.ListCount - 1
End Sub

Private Sub cmbSPRING_LostFocus()
Dim Found As Boolean
Dim s As String
s = cmbSPRING
For i = 0 To cmbSPRING.ListCount - 1
    cmbSPRING.ListIndex = i
    If cmbSPRING = s Then Found = True: Exit For
Next
If Not (Found) Then cmbSPRING.AddItem s: cmbSPRING.ListIndex = cmbSPRING.ListCount - 1
End Sub

Private Sub cmbSTIFF_LostFocus()
Dim Found As Boolean
Dim s As String


GlobalStiffness = Val(Replace(cmbSTIFF, ",", "."))

s = cmbSTIFF
For i = 0 To cmbSTIFF.ListCount - 1
    cmbSTIFF.ListIndex = i
    If cmbSTIFF = s Then Found = True: Exit For
Next
If Not (Found) Then cmbSTIFF.AddItem s: cmbSTIFF.ListIndex = cmbSTIFF.ListCount - 1


End Sub

Private Sub cmdCLEAR_Click()
PIC.Cls
PLdraw = 0


clearDELAresult
For BB = 1 To NCre
    B(BB).UnCheckDELApoints
Next

chDELA.Value = Unchecked
chDELA_Click
'Stop

For BB = 1 To NCre
    B(BB).Clear
Next

PervLINK = 0
PervPOINT = 0 ''
NCre = 0
ICre = 0

MINI.GeneraMiniaturas App.Path & "\empty.bmp"


End Sub

Private Sub cmdEDITnewCreature_Click()
'Stop

NCre = NCre + 1
ReDim Preserve B(NCre + 1)

picPREV.Visible = True

picPREV.Line (0, 0)-(500, 500), 0, BF
picPREV.CurrentX = 1
picPREV.CurrentY = 1

picPREV.Print "NEW" & NCre
picPREV.Refresh

SavePicture picPREV.Image, App.Path & "\tmp" & NCre & ".bmp"
picPREV = LoadPicture(App.Path & "\tmp" & NCre & ".bmp")
picPREV.Refresh

MINI.AddPicture picPREV
ICre = NCre
B(NCre).Clear
slSELECTOR = ICre
PervPOINT = 1
PervLINK = 1
picPREV.Visible = False

End Sub



Private Sub cmdLOAD_Click()
B(ICre).LoadMe "Struct.txt"
PervLINK = 0
PervPOINT = 0
''B.FindCROSSLinks_InitialState

PIC.Cls
DrawALL

'cmbSTIFF.Clear: cmbSTIFF.AddItem GlobalStiffness
'cmbSTIFF.ListIndex = cmbSTIFF.ListCount - 1


chDELA.Value = Unchecked
chDELA_Click

PLdraw = 0


End Sub

Private Sub cmdLOADCreature_Click()

picPREV.Visible = True

File1.Left = cmdLOADCreature.Left



Command = "LC"

File1.Pattern = "*CRE.txt"
File1.Path = App.Path
File1.Refresh

File1.Visible = True
File1.Enabled = True

FirstPoint = True



End Sub

Private Sub cmdLoadProject_Click()
Command = "LP"

File1.Left = cmdLoadProject.Left


File1.Enabled = True

File1.Pattern = "*PRJ.txt"
File1.Path = App.Path
File1.Refresh

File1.Visible = True
End Sub

Private Sub cmdSAVE_Click()
'B.SaveMe "Struct.txt"

End Sub

Private Sub cmdSaveCreature_Click()


File1.Left = cmdSaveCreature.Left
cNAME.Left = cmdSaveCreature.Left


'cmdLOADCreature.Enabled = False


Command = "SC"
'MsgBox "Select TopLeft and BottomRight points of Creature Bounding Box", , "2 Points..."
'FirstPoint = True

File1.Pattern = "*CRE.txt"
File1.Path = App.Path
File1.Visible = True
File1.Enabled = False
cNAME.Visible = True
cNAME.SetFocus

End Sub

Private Sub cmdSaveProject_Click()

File1.Left = cmdSaveProject.Left
cNAME.Left = cmdSaveProject.Left


'cmdLOADCreature.Enabled = False



Command = "SP"

File1.Pattern = "*prj.txt"
File1.Path = App.Path
File1.Refresh
File1.Visible = True
File1.Enabled = False
cNAME.Visible = True
cNAME.SetFocus


End Sub

Private Sub cmdSetToAll_Click()
Dim L As tLINK


For i = 1 To B(ICre).NLink
    
    L = B(ICre).GetLink(i)
    L.BreakDist = Val(Replace(cmbBREAK, ",", ".")) * L.MainLenght
    L.SpringStrenght = Val(Replace(cmbSPRING, ",", "."))
    B(ICre).SetLink(i) = L
    
    'L = B.GetLink(I)
    'MsgBox L.BreakDist
    
Next

End Sub




Private Sub cNAME_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FileName = App.Path & "\" & cNAME
    File1.Visible = False
    cNAME.Visible = False
    
    
    Select Case Command
        Case "SC"
            
            
            'Set CRE = B(ICre).CreateFromAABB
            'CRE.SaveMe FileName & "_CRE.txt"
            'CRE.CreatePreview picPREV, FileName & "_CRE.txt"
            
            B(ICre).SaveMe FileName & "_CRE.txt"
            B(ICre).CreatePreview picPREV, FileName & "_CRE.txt"
            
            
            
            MsgBox FileName & " Saved!", vbInformation, "OK"
            Command = ""
            'cmdLOADCreature.Enabled = True
            
            
            
        Case "SP"
            
            
            
            
            Open FileName & "_PRJ.txt" For Output As 1
            Dim tmpP As tPOINT
            Dim tmpL As tLINK
            
            '    Stop
            
            Print #1, "N-CREATURES"
            Print #1, NCre
            
            ReDim Preserve B(NCre + 1)
            For BB = 1 To NCre
                
                
                Print #1, s
                Print #1, B(BB).Npoint
                For i = 1 To B(BB).Npoint
                    tmpP = B(BB).GetPoint(i)
                    B(BB).PrintPoint tmpP
                Next
                Print #1, s
                Print #1, B(BB).NLink
                For i = 1 To B(BB).NLink
                    tmpL = B(BB).GetLink(i)
                    B(BB).PrintLink tmpL
                Next
                
                
            Next BB
            Close 1
            ' Stop
            
            Command = ""
            
            
            
            
            
            
End Select

End If


End Sub





Private Sub File1_Click()


If Command = "LC" Then
    Dim F As String
    Dim s As String
    Dim II As Integer
    '   Stop
    F = App.Path & "\" & File1.FileName
    II = InStrRev(F, ".")
    s = Left$(F, II - 1) & ".bmp"
    picPREV = LoadPicture(s)
    picPREV.Refresh
    
End If

End Sub

Private Sub File1_DblClick()

Dim s As String

FileName = App.Path & "\" & File1.FileName
File1.Visible = False

Select Case Command
        
        
        
    Case "LC"
        
        
        NCre = NCre + 1
        ReDim Preserve B(NCre + 1)
        
        
        B(NCre).LoadMe FileName
        
        B(NCre).FindAABB
        
        Command = "LC2"
        '    slSELECTOR.Max = IIf(NCre > 2, NCre, 2)
        
        F = App.Path & "\" & File1.FileName
        II = InStrRev(F, ".")
        s = Left$(F, II - 1) & ".bmp"
        '    Stop
        
        
        
        picPREV = LoadPicture(s)
        picPREV.Refresh
        
        MINI.AddPicture picPREV
        '    Stop
        
        picPREV.Visible = False
        
    Case "LP"
        
        cmdCLEAR_Click
        
        
        
        
        Kill App.Path & "\tmp*.bmp"
        
        Open FileName For Input As 1
        Input #1, s
        Input #1, NCre
        
        ReDim Preserve B(NCre + 1)
        For BB = 1 To NCre
            
            
            Input #1, s
            Input #1, Npoint
            B(BB).SetNpoint Npoint
            For i = 1 To Npoint
                B(BB).SetPoint(i) = B(BB).InputPoint
            Next
            Input #1, s
            Input #1, NLink
            B(BB).SetNlink NLink
            For i = 1 To B(BB).NLink
                B(BB).SetLink(i) = B(BB).InputLink
            Next
            
            B(BB).CreatePreview picPREV, App.Path & "\tmp" & BB & ".txt"
            
        Next BB
        Close 1
        
        
        Command = ""
        
        PIC.Cls
        DrawALL
        
        
        MINI.GeneraMiniaturas App.Path & "\tmp*.bmp"
        
        
End Select





End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command = "LC" Then
    picPREV.Left = X + File1.Left + Screen.TwipsPerPixelX * 10
    picPREV.Top = Y + File1.Top + Screen.TwipsPerPixelY * 10
    
    
End If

End Sub

Private Sub Form_Initialize()
SavePicture picPREV.Image, App.Path & "\tmp0.bmp"
'MINI.GeneraMiniaturas App.Path & "\tmp*.bmp"
End Sub

Private Sub Form_Load()

Set CRE = New clsBODY
ReDim Preserve B(2) As New clsBODY
Set B(1) = New clsBODY





NCre = 1
'ReDim B(NCre)
ICre = 1


'Gravity = 0.1
GravityY = 0 '0.1
GravityX = 0

fPoint.Visible = True
fPoint.Top = fLINK.Top
fPoint.Left = fLINK.Left

fLINK.Visible = False

For bu = 0 To 1 Step 0.125
    cmbDF.AddItem CStr(bu)
Next
cmbDF.ListIndex = 0
For bu = 0 To 1 Step 0.125
    cmbDFpoint.AddItem CStr(bu)
Next
cmbDFpoint.ListIndex = 0

PervLINK = 0
PervPOINT = 0



PI = Atn(1) * 4

'''GlobalTime = 0
'''TimeSTEP = 0.005  '0.003
FirstPoint = True
cmbBREAK.Clear
cmbSPRING.Clear
cmbSTIFF.Clear


MaxH = PIC.ScaleHeight - 3
MaxW = PIC.ScaleWidth - 3



GlobalStiffness = 0.5


NW = 12 '17
NH = 5 '3
Ww = 25 '30.1 '50
'B(ICre).GlobalStiffness = 0.5 '.35 '0.5 '0.5  '0.1 ' 0.75 ' 0.5
GlobalBreak = 2 '3 '1.7 '1.5
GlobalSPRING = 20 '15


cmbBREAK.AddItem GlobalBreak
cmbBREAK.ListIndex = 0

cmbSPRING.AddItem GlobalSPRING
cmbSPRING.ListIndex = 0

cmbSTIFF.AddItem GlobalStiffness
cmbSTIFF.ListIndex = 0



GoTo AddNothing
For h = 1 To NH
    For W = 1 To NW
        'Stop
        B(ICre).ADDPoint 20 + W * Ww, 80 + h * Ww, False
    Next
Next


For i = 0 To NW * NH - NW - 2
    If (i + 1) Mod NW <> 0 Then
        II = i + 1
        B(ICre).ADDLink II, II + 1, GlobalBreak, , GlobalSPRING '-
        B(ICre).ADDLink II, II + NW, GlobalBreak, , GlobalSPRING '        '|
        B(ICre).ADDLink II, II + NW + 1, GlobalBreak, , GlobalSPRING '    '\
        B(ICre).ADDLink II + 1, II + NW, GlobalBreak, , GlobalSPRING '   '/
    End If
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''quelli sotto
For i = NW * NH - NW To NW * NH - 2
    II = i + 1
    B(ICre).ADDLink II, II + 1, GlobalBreak, , GlobalSPRING '
Next
'' quelli lato destro
For i = NW - 1 To NW * NH - NW - 1 Step NW
    II = i + 1
    B(ICre).ADDLink II, II + NW, GlobalBreak, , GlobalSPRING 'e
Next

'    Stop

'    B.ADDPoint 100, 100, False
'    B.ADDPoint 150, 100, False
'    B.ADDPoint 100, 150, False
'    B.ADDPoint 150, 155, False


'    B.ADDLink 1, 2, 2
'    B.ADDLink 2, 4, 2
'   B.ADDLink 4, 3, 2
'   B.ADDLink 3, 1, 2
'    B.ADDLink 1, 4, 2
'    B.ADDLink 2, 3, 2
'Stop

'p = B.GetPoint(NW * NH - NW + 1)
'p.IsFIX = True
'B.SetPoint(NW * NH - NW + 1) = p
'
'p = B.GetPoint(NW * NH)
'p.IsFIX = True
'B.SetPoint(NW * NH) = p

AddNothing:
NCre = 0
ICre = 0
MINI.GeneraMiniaturas App.Path & "\empty.bmp"

B(ICre).FindCROSSLinks_InitialState

B(ICre).FindEXTERNALlinks

B(ICre).DRAW PIC

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If NCre > 0 Then
    Text1 = "Creature N " & ICre & " of " & NCre & vbCrLf & _
            "POINTS: " & B(ICre).Npoint & vbCrLf & "LINKS: " & B(ICre).NLink
    
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
Kill App.Path & "\tmp*.bmp"
End Sub

Private Sub MINI_Click(N As Integer)
PervPOINT = 1
PervLINK = 1

ICre = N
PIC.Cls
DrawALL

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dmin
Dim II As Integer
Dim tmpP As tPOINT

Dim Pt As tPOINT
Dim PP1 As tPOINT
Dim PP2 As tPOINT



'Stop
'Stop

If Button = 1 Then
    
    If Command <> "" Then
        Select Case Command
            Case "SC"
                'If FirstPoint Then
                '
                '    PIC.Line (X, Y)-(X + 50, Y), vbWhite
                '    PIC.Line (X, Y)-(X, Y + 50), vbWhite
                '    FirstPoint = False
                '    AABB1.X = X
                '    AABB1.Y = Y
                '
                'Else
                '    PIC.Line (X, Y)-(X - 50, Y), vbWhite
                '    PIC.Line (X, Y)-(X, Y - 50), vbWhite
                '    FirstPoint = True
                '    AABB2.X = X
                '    AABB2.Y = Y
                '
                '    File1.Pattern = "*.txt"
                '    File1.Path = App.Path
                '    File1.Visible = True
                '    File1.Enabled = False
                '    cNAME.Visible = True
                '    cNAME.SetFocus
                
                ' End If
            Case "LC2"
                '        Stop
                
                For i = 1 To B(NCre).Npoint
                    tmpP = B(NCre).GetPoint(i)
                    tmpP.X = tmpP.X + X - AABB1.X
                    tmpP.Y = tmpP.Y + Y - AABB1.Y
                    B(NCre).SetPoint(i) = tmpP
                Next
                '        Stop
                
                'NCre = NCre + 1
                'ReDim Preserve B(NCre) As New clsBODY
                'Set B(NCre) = CRE
                
                Command = ""
                '        Stop
                
                Mous.X = 0
                Mous.Y = 0
                
                
    End Select
    
    
Else '''' not command
    
    
    Select Case PLdraw
        Case 0
            '            Stop
            
            B(ICre).ADDPoint CDbl(X), CDbl(Y), _
                    IIf(chPFIX.Value = Checked, True, False), _
                    IIf(chDELA.Value = Checked, True, False), _
                    Val(Replace(txtDTpoint, ",", ".")), _
                    Val(Replace(cmbDFpoint, ",", ".")), _
                    Val(Replace(txtFaseLen, ",", "."))
            
            
            If ChGlue.Value = Checked Then B(ICre).GLUE
            
            If chDELA.Value = Checked Then B(ICre).Delaunay
            
        Case 1
            If FirstPoint Then
                Dmin = 999999999
                Pt.X = X
                Pt.Y = Y
                For i = 1 To B(ICre).Npoint
                    D = Distance(Pt, B(ICre).GetPoint(i))
                    If D < Dmin Then Dmin = D: nP1 = i
                Next
                
                Me.Caption = B(ICre).Npoint & " " & nP1
                FirstPoint = False
            Else
                Dmin = 999999999
                Pt.X = X
                Pt.Y = Y
                For i = 1 To B(ICre).Npoint
                    D = Distance(Pt, B(ICre).GetPoint(i))
                    If D < Dmin Then Dmin = D: nP2 = i
                Next
                
                If nP1 <> nP2 Then
                    'B.ADDLink nP1, nP2, GlobalBreak, IIf(chLFix.Value = Checked, True, False), GlobalSPRING
                    
                    'B.ADDLink nP1, nP2, Val(Replace(cmbBREAK, ",", ".")), _
                    IIf(chLFix.Value = Checked, True, False), _
                            Val(Replace(cmbSPRING, ",", "."))
                    
                    B(ICre).ADDLink nP1, nP2, Val(Replace(cmbBREAK, ",", ".")), _
                            IIf(chLFix.Value = Checked, True, False), _
                            Val(Replace(cmbSPRING, ",", ".")), _
                            Val(Replace(txtDL, ",", ".")), _
                            Val(Replace(txtDT, ",", ".")), _
                            Val(Replace(cmbDF, ",", "."))
                End If
                
                Me.Caption = B(ICre).Npoint & " " & nP2
                FirstPoint = True
            End If
End Select

End If

Else 'button=2
If Command = "LC2" Then
    Mous.X = X
    Mous.Y = Y
    B(NCre).FindAABB
    For i = 1 To B(NCre).Npoint
        copyP(i) = B(NCre).GetPoint(i)
        'copyP(i).X = copyP(i).X - B(NCre).pAABB1x
        ''copyP(i).Y = copyP(i).Y - B(NCre).pAABB1y
        
    Next
    
    
Else
    
    Select Case PLdraw
        Case 0
            Dmin = 999999999
            Pt.X = X
            Pt.Y = Y
            For i = 1 To B(ICre).Npoint
                D = Distance(Pt, B(ICre).GetPoint(i))
                If D < Dmin Then Dmin = D: nP1 = i
            Next
            B(ICre).REMOVEPoint nP1
            PervLINK = 0
            PervPOINT = 0
            PIC.Cls
            
        Case 1
            Dmin = 999999999
            Pt.X = X
            Pt.Y = Y
            For i = 1 To B(ICre).NLink
                PP1 = B(ICre).GetPoint(B(ICre).GetLink(i).P1)
                PP2 = B(ICre).GetPoint(B(ICre).GetLink(i).P2)
                PP1.X = (PP1.X + PP2.X) / 2
                PP1.Y = (PP1.Y + PP2.Y) / 2
                
                'D = Distance(Pt, B.GetPoint(B.GetLink(I).P1))
                'D = D + Distance(Pt, B.GetPoint(B.GetLink(I).p2))
                D = Distance(Pt, PP1)
                D = D + Distance(Pt, PP1)
                
                If D < Dmin Then Dmin = D: nP1 = i
            Next
            
            B(ICre).BREAKLink nP1
            PervLINK = 0
            PervPOINT = 0
            PIC.Cls
End Select

End If
End If

DrawALL

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = X & "   " & Y
Dim tmpP As tPOINT
Dim tmpP2 As tPOINT
Dim tmpL As tLINK
Dim TmpL2 As tLINK

Dim P As Integer
Dim NearP As Integer
Dim NearL As Integer
Dim Med As tPOINT
Dim L As Integer
Dim P1 As tPOINT
Dim P2 As tPOINT
Dim PoI As tPOINT
Dim PoJ As tPOINT

Dim ZeroV As Vector2
Dim Delta As Vector2


Dim mANG As Double
Dim PAng As Double


If NCre = 0 Then Exit Sub

If Command = "LC2" Then
    Select Case Button
            
        Case 0
            PIC.Cls
            '    Stop
            AABB1.X = B(NCre).pAABB1x
            AABB1.Y = B(NCre).pAABB1y
            
            '    Stop
            
            For i = 1 To B(NCre).Npoint
                tmpP = B(NCre).GetPoint(i)
                tmpP.X = tmpP.X + X - AABB1.X
                tmpP.Y = tmpP.Y + Y - AABB1.Y
                B(NCre).SetPoint(i) = tmpP
            Next
            'B(NCre).DRAW PICù
            DrawALL
            For i = 1 To B(NCre).Npoint
                tmpP = B(NCre).GetPoint(i)
                tmpP.X = tmpP.X - X + AABB1.X
                tmpP.Y = tmpP.Y - Y + AABB1.Y
                B(NCre).SetPoint(i) = tmpP
            Next
        Case 2
            
            
            
            ZeroV.X = 0
            ZeroV.Y = 0
            
            Delta.X = X - Mous.X
            Delta.Y = Y - Mous.Y
            
            mANG = GetAngle(Delta, ZeroV)
            Debug.Print mANG
            Me.Caption = a
            For i = 1 To B(NCre).Npoint
                tmpP = copyP(i) '
                Delta.X = tmpP.X - copyP(1).X
                Delta.Y = tmpP.Y - copyP(1).Y
                ZeroV.X = 0
                ZeroV.Y = 0
                PAng = GetAngle(Delta, ZeroV)
                DIst = Distance(tmpP, copyP(1))
                tmpP.X = DIst * Cos(PAng + mANG) + Mous.X '- B(NCre).pAABB1x
                tmpP.Y = DIst * Sin(PAng + mANG) + Mous.Y '- B(NCre).pAABB1y
                B(NCre).SetPoint(i) = tmpP
                
            Next
            PIC.Cls
            
            DrawALL
            
            
End Select


Else ' not LC


If PLdraw = 1 And FirstPoint = False Then
    PIC.Cls
    DrawALL
    PIC.Line (B(ICre).GetPoint(nP1).X, B(ICre).GetPoint(nP1).Y)-(X, Y), vbRed
Else
    
    If PLdraw = 0 Then '''point
        If B(ICre).Npoint > 0 Then
            tmpP.X = X
            tmpP.Y = Y
            Dmin = 999999999
            For P = 1 To B(ICre).Npoint
                D = Distance(B(ICre).GetPoint(P), tmpP)
                If D < Dmin Then Dmin = D: NearP = P
            Next
            tmpP = B(ICre).GetPoint(NearP)
            Text1 = ""
            Text1 = Text1 & "POINT:" & NearP & vbCrLf
            Text1 = Text1 & "X:" & tmpP.X & " Y:" & tmpP.Y & vbCrLf
            Text1 = Text1 & "Is" & IIf(tmpP.isFix, "", " Not") & " Fixed" & vbCrLf
            Text1 = Text1 & "Linked " & tmpP.HowManyLinks & " times" & vbCrLf
            Text1 = Text1 & "VX:" & tmpP.vX & " VY:" & tmpP.vY & vbCrLf
            
            Text1 = Text1 & "Dyn Speed =" & tmpP.DynamicSpeed & vbCrLf
            Text1 = Text1 & "Dyn  Fase =" & tmpP.DynamicFase & vbCrLf
            Text1 = Text1 & "DynHowLong=" & tmpP.DynamicFaseLEN & vbCrLf
            ''''''''''''''' yellow point
            
            tmpP2 = B(ICre).GetPoint(PervPOINT)
            If tmpP2.isFix Then
                PIC.Line (tmpP2.X - 2, tmpP2.Y - 3)-(tmpP2.X + 4, tmpP2.Y + 3), vbRed
                PIC.Line (tmpP2.X + 3, tmpP2.Y - 3)-(tmpP2.X - 3, tmpP2.Y + 3), vbRed
            Else
                PIC.Circle (tmpP2.X, tmpP2.Y), 2, vbRed
            End If
            If tmpP.isFix Then
                PIC.Line (tmpP.X - 2, tmpP.Y - 3)-(tmpP.X + 4, tmpP.Y + 3), vbYellow
                PIC.Line (tmpP.X + 3, tmpP.Y - 3)-(tmpP.X - 3, tmpP.Y + 3), vbYellow
            Else
                PIC.Circle (tmpP.X, tmpP.Y), 2, vbYellow
            End If
            PervPOINT = NearP
            
            '''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If chAUTOpoint.Value = Checked Then
                
                
                'ì cmbBREAK = Round(TmpL.BreakDist / TmpL.MainLenght * 100) / 100
                'ì cmbSPRING = Round(TmpL.SpringStrenght * 10) / 10
                'ì   txtDL = Round(TmpL.DynamicLenght / TmpL.MainLenght * 100) / 100
                txtDTpoint = Round(tmpP.DynamicSpeed * 1000) / 1000
                cmbDFpoint = Round(tmpP.DynamicFase * 1000) / 1000
                txtFaseLen = Round(tmpP.DynamicFaseLEN * 1000) / 1000
            End If
            
            
            
        End If
    Else ''''LINK
        If B(ICre).NLink > 0 Then
            tmpP.X = X
            tmpP.Y = Y
            Dmin = 99999999999#
            For L = 1 To B(ICre).NLink
                tmpL = B(ICre).GetLink(L)
                P1 = B(ICre).GetPoint(tmpL.P1)
                P2 = B(ICre).GetPoint(tmpL.P2)
                Med.X = (P1.X + P2.X) / 2
                Med.Y = (P1.Y + P2.Y) / 2
                D = Distance(Med, tmpP)
                If D < Dmin Then Dmin = D: NearL = L
            Next
            tmpL = B(ICre).GetLink(NearL)
            P1 = B(ICre).GetPoint(tmpL.P1)
            P2 = B(ICre).GetPoint(tmpL.P2)
            Text1 = ""
            Text1 = Text1 & "LINK:" & NearL & vbCrLf
            Text1 = Text1 & "P1 " & tmpL.P1 & "-P2 " & tmpL.P2 & vbCrLf
            Text1 = Text1 & "Main Len=" & tmpL.MainLenght & vbCrLf
            Text1 = Text1 & "Curr Len=" & Distance(P1, P2) & vbCrLf
            Text1 = Text1 & "BreakLen=" & tmpL.BreakDist & vbCrLf
            Text1 = Text1 & "SpringStrength=" & tmpL.SpringStrenght & vbCrLf
            Text1 = Text1 & "Dyn  Len=" & tmpL.DynamicLenght & vbCrLf
            Text1 = Text1 & "DynSpeed=" & tmpL.DynamicSpeed & vbCrLf
            Text1 = Text1 & "Dyn Fase=" & tmpL.DynamicFase & vbCrLf
            
            '''''' Evident Link Mouse Over YELLOW
            TmpL2 = B(ICre).GetLink(PervLINK)
            PoI = B(ICre).GetPoint(TmpL2.P1)
            PoJ = B(ICre).GetPoint(TmpL2.P2)
            PIC.Line (PoI.X, PoI.Y)-(PoJ.X, PoJ.Y), IIf(TmpL2.DynamicLenght <> 0, vbGreen, vbRed)
            PoI = B(ICre).GetPoint(tmpL.P1)
            PoJ = B(ICre).GetPoint(tmpL.P2)
            PIC.Line (PoI.X, PoI.Y)-(PoJ.X, PoJ.Y), vbYellow
            PervLINK = NearL
            '''''''''''
            
            
            If chAuto.Value = Checked Then
                
                
                cmbBREAK = Round(tmpL.BreakDist / tmpL.MainLenght * 1000) / 1000
                cmbSPRING = Round(tmpL.SpringStrenght * 10) / 10
                txtDL = Round(tmpL.DynamicLenght / tmpL.MainLenght * 1000) / 1000
                txtDT = Round(tmpL.DynamicSpeed * 1000) / 1000
                cmbDF = Round(tmpL.DynamicFase * 1000) / 1000
                
            End If
        End If
    End If
    
End If

End If

End Sub

Private Sub picGravity_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picGravity.Cls

With picGravity
    picGravity.Circle (.ScaleWidth \ 2, .ScaleHeight \ 2), 3, 0
    
    
    picGravity.Line (.ScaleWidth \ 2, .ScaleHeight \ 2)-(X, Y), 0
    
    GravityY = (Y - .ScaleHeight \ 2) / (.ScaleHeight * 4)
    GravityX = (X - .ScaleWidth \ 2) / (.ScaleWidth * 4)
    Label5 = "vx" & Format(GravityX * 10, "0.0")
    Label6 = "vy" & Format(GravityY * 10, "0.0")
End With


End Sub

Private Sub picGravity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    
    picGravity.Cls
    
    With picGravity
        picGravity.Circle (.ScaleWidth \ 2, .ScaleHeight \ 2), 3, 0
        
        
        picGravity.Line (.ScaleWidth \ 2, .ScaleHeight \ 2)-(X, Y), 0
        
        GravityY = (Y - .ScaleHeight \ 2) / (.ScaleHeight * 4)
        GravityX = (X - .ScaleWidth \ 2) / (.ScaleWidth * 4)
        Label5 = "vx" & Format(GravityX * 10, "0.0")
        Label6 = "vy" & Format(GravityY * 10, "0.0")
    End With
    
End If

End Sub

Private Sub PLdraw_Change()
If B(ICre).Npoint = 0 Then PLdraw = 0


Select Case PLdraw
        
    Case 0
        PLlabel = "Point"
        FirstPoint = False
        chPFIX.Visible = True
        chLFix.Visible = False
        
        PIC.MousePointer = 0
        
        chDELA.Visible = True
        If chDELA.Value = Checked Then fLINK.Visible = False
        fLINK.Visible = False
        fPoint.Visible = True
        
        ChGlue.Visible = True
        
        txtDTpoint.Visible = True
        cmbDFpoint.Visible = True
        txtFaseLen.Visible = True
        
    Case 1
        PLlabel = "Link"
        FirstPoint = True
        chLFix.Visible = True
        chPFIX.Visible = False
        ChGlue.Visible = False
        PIC.MousePointer = 2
        chDELA.Visible = False
        fLINK.Visible = True
        fPoint.Visible = False
        
        chLFix.Visible = False 'True
        
        
        txtDTpoint.Visible = False
        cmbDFpoint.Visible = False
        txtFaseLen.Visible = False
End Select

End Sub



Private Sub slGRAVITY_Change()
Gravity = slGRAVITY.Value / 100
End Sub

Private Sub slGRAVITY_Click()
Gravity = slGRAVITY.Value / 100
End Sub

Private Sub slGRAVITY_Scroll()
Gravity = slGRAVITY.Value / 100
End Sub

Private Sub slSELECTOR_Change()
PervPOINT = 1
PervLINK = 1

ICre = slSELECTOR
PIC.Cls
DrawALL

End Sub

Private Sub Timer1_Timer()
PIC.Cls

'''GlobalTime = GlobalTime + TimeSTEP


For BB = 1 To NCre
    B(BB).DRAW PIC
    B(BB).UpDateForces
Next

ReactToCreature
'B.FindCROSSLinks_andUPDATE


End Sub

Private Sub txtDL_Change()
'If Val(txtDL) = 0 Then txtDT = "0": cmbDF = "0"


End Sub

Sub DrawALL()

For BB = 1 To NCre
    'Stop
    
    
    B(BB).DRAW PIC
Next
End Sub
