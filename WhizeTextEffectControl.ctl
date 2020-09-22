VERSION 5.00
Begin VB.UserControl WhizeTextEffectControl 
   BackColor       =   &H80000007&
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   PropertyPages   =   "WhizeTextEffectControl.ctx":0000
   ScaleHeight     =   1170
   ScaleWidth      =   2370
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2595
      TabIndex        =   1
      Text            =   "Whize Text Effect Control 1.0.0"
      Top             =   900
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox picEffect 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   0
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1185
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "WhizeTextEffectControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'***********************************************************************************'
'** Project     : Whize Text Effect Control ver 1.0.0
'** Author      : Zubair Ahmed M.
'** Description : A simple ActiveX control which applies a shadow or a 3D effect
'**               to the text
'***********************************************************************************'
'ThanQ for downloading this project. The code is pretty simple thats why I have not
'commented it, but still if u find anything difficult to understand, remember I'm
'just a mail away. And if u like this project, plz dont forget to vote.
'***********************************************************************************'

Enum Effects
[Shadow] = 0
[3D Effect] = 1
End Enum

'Default Property Values:
Const m_def_Effect = [Shadow]
Const m_def_ForeColor = &H80000008

'Property Variables:
Dim m_Effect As Effects
Dim m_ForeColor As Long
Dim m_EffectColor As Long
Dim m_Distance As Integer
Dim m_AutoSize As Boolean
Dim m_Font As Font
Private Sub Timer1_Timer()
Label1.Caption = Text1.Text
picEffect.Width = Label1.Width
picEffect.Cls
picEffect.Print Text1.Text
ApplyEffect (m_Effect)
Timer1.Enabled = False
End Sub
Private Sub UserControl_InitProperties()
    m_Effect = m_def_Effect
    m_ForeColor = m_def_ForeColor
    Set m_Font = picEffect.Font
End Sub
Private Sub UserControl_Resize()
With UserControl
    picEffect.Width = .Width
    picEffect.Height = .Height
End With
End Sub
Public Property Get AutoSize() As Boolean
AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
m_AutoSize = New_AutoSize
PropertyChanged "AutoSize"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = picEffect.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picEffect.BackColor() = New_BackColor
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get Distance() As Integer
Attribute Distance.VB_Description = "Returns/sets the value for specifying the distance the desired effect will be cast."
Attribute Distance.VB_ProcData.VB_Invoke_Property = "Distance"
Distance = m_Distance
End Property
Public Property Let Distance(ByVal New_Distance As Integer)
If (New_Distance < 0 Or New_Distance > 100) Then
MsgBox "An integer between 0 and 100 is required. Closest value inserted.", vbCritical + vbOKOnly
If New_Distance < 0 Then m_Distance = 0
If New_Distance > 100 Then m_Distance = 100
Else
m_Distance = New_Distance
End If
PropertyChanged "Distance"
End Property
Public Property Get Effect() As Effects
Attribute Effect.VB_Description = "Returns/sets the desired effect."
Effect = m_Effect
End Property
Public Property Let Effect(ByVal New_Effect As Effects)
m_Effect = New_Effect
PropertyChanged "Effect"
End Property
Public Property Get Font() As Font
    Set Font = picEffect.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set picEffect.Font = New_Font
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Reutrns/sets the foreground color used to display text and graphics in the control."
    ForeColor = picEffect.ForeColor
    End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picEffect.ForeColor() = New_ForeColor
    m_ForeColor = New_ForeColor
    ApplyEffect (m_Effect)
    PropertyChanged "ForeColor"
End Property
Public Property Get EffectColor() As OLE_COLOR
Attribute EffectColor.VB_Description = "Returns/sets the desired effect's color."
    EffectColor = m_EffectColor
    End Property
Public Property Let EffectColor(ByVal New_EffectColor As OLE_COLOR)
    m_EffectColor = New_EffectColor
    PropertyChanged "EffectColor"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = -518
    Text = Text1.Text
    Label1.Caption = Text1.Text
    picEffect.Width = Label1.Width
    picEffect.Height = Label1.Height
    picEffect.Cls
    picEffect.Print Text1.Text
    ApplyEffect (m_Effect)
End Property
Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_Effect = .ReadProperty("Effect", m_def_Effect)
    m_ForeColor = .ReadProperty("ForeColor", &H80000007)
    m_EffectColor = .ReadProperty("EffectColor", &H8000000D)
    m_Distance = .ReadProperty("Distance", 50)
    m_AutoSize = .ReadProperty("AutoSize", True)
    picEffect.BackColor = .ReadProperty("BackColor", &H8000000F)
    Text1.Text = .ReadProperty("Text", "Whize Text Effect Control")
    UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
    Set picEffect.Font = .ReadProperty("Font", picEffect.Font)
    Set Label1.Font = .ReadProperty("Font", picEffect.Font)
End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("BackColor", picEffect.BackColor, &H8000000F)
    Call .WriteProperty("Text", Text1.Text, "Whize Text Effect Control")
    Call .WriteProperty("Effect", m_Effect, m_def_Effect)
    Call .WriteProperty("ForeColor", m_ForeColor, &H80000007)
    Call .WriteProperty("EffectColor", m_EffectColor, &H8000000D)
    Call .WriteProperty("Font", picEffect.Font)
    Call .WriteProperty("Distance", m_Distance, 50)
    Call .WriteProperty("AutoSize", m_AutoSize, False)
End With
End Sub
Private Sub ApplyEffect(Effect As Effects)
'**********************************************************************************'
'This procedure here applies the desired effect based on the effect selected
'**********************************************************************************'

Dim loopValue As Single
Select Case Effect
Case 0
picEffect.Width = picEffect.Width + m_Distance
picEffect.Cls
picEffect.ForeColor = m_EffectColor
picEffect.Print Text1.Text
picEffect.CurrentX = m_Distance
picEffect.CurrentY = m_Distance
picEffect.ForeColor = m_ForeColor
picEffect.Print Text1.Text
Case 1
picEffect.Width = picEffect.Width + m_Distance
picEffect.Cls
picEffect.ForeColor = m_EffectColor
For loopValue = 1 To m_Distance
picEffect.CurrentX = loopValue
picEffect.CurrentY = loopValue
picEffect.Print Text1.Text
Next loopValue
picEffect.ForeColor = m_ForeColor
picEffect.CurrentX = loopValue
picEffect.CurrentY = loopValue
picEffect.Print Text1.Text
End Select
If m_AutoSize = True Then UserControl.Width = picEffect.Width
End Sub
