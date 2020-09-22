VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OpenGL Particle Explosion Demo..."
   ClientHeight    =   5565
   ClientLeft      =   360
   ClientTop       =   2040
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Axes"
      Height          =   615
      Left            =   4800
      TabIndex        =   27
      Top             =   4800
      Width           =   4500
      Begin VB.CheckBox chkAxes 
         Caption         =   "Show XYZ Axes"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Particle Size..."
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   4530
      Begin MSComctlLib.Slider sldParticleSize 
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   7
         Value           =   7
      End
      Begin VB.Label Label8 
         Caption         =   "10.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "20.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4040
         TabIndex        =   25
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   520
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Maximum Particle Size:"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Particle Color..."
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   4530
      Begin VB.CheckBox chkRandomColors 
         Caption         =   "Use Random Colors"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gravity..."
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4530
      Begin VB.CheckBox chkEnableGravity 
         Caption         =   "Enable Gravity"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin MSComctlLib.Slider sldGravity 
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   2
         Min             =   -10
      End
      Begin VB.Label lblG1 
         Caption         =   "Gravity Level:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblG2 
         Caption         =   "-1.0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblG4 
         Caption         =   "+1.0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4000
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblG3 
         Caption         =   "0.0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2325
         TabIndex        =   10
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Particle Quantity..."
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4530
      Begin MSComctlLib.Slider sldParticleCount 
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   20
         Min             =   10
         Max             =   3000
         SelStart        =   1000
         TickFrequency   =   100
         Value           =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "1500"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2320
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "3000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4050
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Particles:"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Timer timAnimate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9000
      Top             =   0
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00000000&
      Height          =   4500
      Left            =   4800
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   0
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'* OpenGL Particle Explosion Demonstration in Visual Basic       *
'*      By: Daniel S. Soper                                      *
'*****************************************************************
'*                  If you like it, vote for it!                 *
'*****************************************************************

Option Explicit 'explicitly declare all variables


Private Type ParticleDescriptor 'describes each particle's properties
    X As Single 'X coordinate
    Y As Single 'Y coordinate
    Z As Single 'Z coordinate
    TranslateX As Single 'Amount to translate the X coordinate by
    TranslateY As Single 'Amount to translate the Y coordinate by
    TranslateZ As Single 'Amount to translate the Z coordinate by
    Red As Single 'Red color value
    Green As Single 'Green color value
    Blue As Single 'Blue color value
    Size As Single 'Radius of the particle
End Type

Dim hGLRC As Long 'handle to the gl rendering context
Dim Particle(3000) As ParticleDescriptor 'the particles
Dim ParticleCount As Integer 'Number of particles to display
Dim FrameCount As Integer 'Current rendering frame
Dim GravityLevel As Integer 'Current level of gravity

Private Function Initialize() As Boolean
    Dim PFD As PIXELFORMATDESCRIPTOR 'Describes the pixel format of a drawing surface
    Dim RetVal As Long 'Return value
    
    PFD.nSize = Len(PFD)
    PFD.nVersion = 1
    PFD.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    PFD.iPixelType = PFD_TYPE_RGBA 'set the pixel type
    PFD.cColorBits = 24 'set the color bits
    PFD.cDepthBits = 16 'set the depth bits
    PFD.iLayerType = PFD_MAIN_PLANE 'set the layer type
    RetVal = ChoosePixelFormat(picOutput.hDC, PFD) 'Retreive a compatible pixel format for picOutput
    If RetVal = 0 Then 'If a compatible pixel format could not be found
        MsgBox "Could not find a compatible pixel format.", vbCritical + vbOKOnly, "Initialization Error" 'Display error message
        Initialize = False 'Initialization of the OpenGL drawing surface was not successful
        End 'Exit program
    End If
    RetVal = SetPixelFormat(picOutput.hDC, RetVal, PFD) 'Set the pixel format of picOutput
 
    hGLRC = wglCreateContext(picOutput.hDC) 'Return the handle to an OpenGL rendering context that is compatible with picOutput
    wglMakeCurrent picOutput.hDC, hGLRC 'Set the current OpenGL rendering context to picOutput
    glClearColor 0, 0, 0, 1 'Set the clear color
    glEnable glcColorMaterial 'Allow material parameters track the current color
    glColorMaterial faceFront, cmmAmbientAndDiffuse 'Enable color tracking
    glClearDepth 1 'Set the clear value for the depth buffer
    glShadeModel smFlat 'Set the shading mode to flat
    
    glEnable glcDepthTest   'Enable depth comparisons
    glEnable glcLighting    'Enable lighting operations
    glEnable glcLight0      'Enable this light
    SetupLight 'Setup this light's color and position properties
    FrameCount = 0 'Set the rendering frame to 0
    GravityLevel = 0 'Set the gravity level to 0
    timAnimate.Enabled = True 'enable the timer that controls the explosion
    Initialize = True 'initialization succeeded
End Function

Private Sub DrawParticles()
    Dim Obj As Long 'holds the handle to the current particle
    Dim X As Integer 'loop variable
    
    If chkAxes.Value = 1 Then 'draw XYZ axes
        glBegin bmLines
            glColor3f 1, 0, 0 'X axis
            glVertex3f -450, 0, 0
            glVertex3f 450, 0, 0
            glColor3f 0, 1, 0 'Y axis
            glVertex3f 0, -300, 0
            glVertex3f 0, 300, 0
            glColor3f 0, 0, 1 'Z axis
            glVertex3f 0, 0, -1000
            glVertex3f 0, 0, 500
        glEnd
    End If
    
    For X = 0 To ParticleCount 'for each particle
        Particle(X).X = Particle(X).X + Particle(X).TranslateX 'translate this particle's X location
        If GravityLevel = 0 Then 'if no gravity
            Particle(X).Y = Particle(X).Y + Particle(X).TranslateY  'translate this particle's Y location
        Else 'if gravity
            Particle(X).Y = (Particle(X).Y + Particle(X).TranslateY) - ((FrameCount / 10) * GravityLevel) 'translate this particle's Y location
        End If
        Particle(X).Z = Particle(X).Z + Particle(X).TranslateZ 'translate this particle's Z location
        Obj = gluNewQuadric 'create new quadric object
        glPushMatrix 'push the matrix stack down by one
            glRGBA Particle(X).Red, Particle(X).Green, Particle(X).Blue, 1 'set this particle's color
            glTranslatef Particle(X).X, Particle(X).Y, Particle(X).Z 'translate this particle's location in 3D space
            gluSphere Obj, Particle(X).Size, 4, 4 'draw the sphere that represents the particle
        glPopMatrix 'pop the matrix stack up by one
        gluDeleteQuadric Obj 'delete the quadric from memory
    Next
End Sub

Private Sub Render()
    Static Busy As Boolean 'holds the current rendering process status

    If Busy = True Then 'if currently rendering
        Exit Sub
    Else
        Busy = True
    End If
    
    wglMakeCurrent picOutput.hDC, hGLRC 'sets this thread's current rendering context
    
    glClear clrColorBufferBit Or clrDepthBufferBit 'clear the color and depth buffers

    DrawParticles 'call the subroutine that draws the particles

    SwapBuffers picOutput.hDC 'exchange the front and back buffers
    wglMakeCurrent 0, 0
    FrameCount = FrameCount + 1 'increment the frame count
    Busy = False 'rendering process is complete
End Sub

Private Sub SetupParticles()
    Dim X As Integer 'loop variable
    
    For X = 0 To ParticleCount
        Particle(X).X = 0 'initialize this particle's X value to 0
        Particle(X).Y = 0 'initialize this particle's Y value to 0
        Particle(X).Z = 0 'initialize this particle's Z value to 0
        Randomize
        Particle(X).TranslateX = CInt(Rnd * 20) 'create a random X translation amount between 0 and 20
        Randomize
        Particle(X).TranslateY = CInt(Rnd * 20) 'create a random Y translation amount between 0 and 20
        Randomize
        Particle(X).TranslateZ = CInt(Rnd * 20) 'create a random Z translation amount between 0 and 20
        If chkRandomColors.Value = 1 Then 'if use random colors
            Randomize
            Particle(X).Red = CInt(Rnd * 255) 'Create a random red color value
            Randomize
            Particle(X).Green = CInt(Rnd * 255) 'Create a random green color value
            Randomize
            Particle(X).Blue = CInt(Rnd * 255) 'Create a random blue color value
        Else 'if use a specific color
            If picColor(0).BorderStyle = 1 Then 'white-ish
                Randomize
                Particle(X).Red = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Green = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Blue = CInt(Rnd * 25) + 230
            ElseIf picColor(1).BorderStyle = 1 Then 'yellow-ish
                Randomize
                Particle(X).Red = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Green = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Blue = CInt(Rnd * 100)
            ElseIf picColor(2).BorderStyle = 1 Then 'red-ish
                Randomize
                Particle(X).Red = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Green = CInt(Rnd * 100)
                Randomize
                Particle(X).Blue = CInt(Rnd * 100)
            ElseIf picColor(3).BorderStyle = 1 Then 'green-ish
                Randomize
                Particle(X).Red = CInt(Rnd * 100)
                Randomize
                Particle(X).Green = CInt(Rnd * 25) + 230
                Randomize
                Particle(X).Blue = CInt(Rnd * 100)
            ElseIf picColor(4).BorderStyle = 1 Then 'blue-ish
                Randomize
                Particle(X).Red = CInt(Rnd * 100)
                Randomize
                Particle(X).Green = CInt(Rnd * 100)
                Randomize
                Particle(X).Blue = CInt(Rnd * 25) + 230
            End If
        End If
        Randomize
        Particle(X).Size = CInt(Rnd * sldParticleSize.Value) 'set the particle radius
    Next
    For X = 0 To ParticleCount 'generate random negative vectors for the particles
        Randomize
        If Rnd >= 0.5 Then
            Particle(X).TranslateX = Particle(X).TranslateX * -1
        End If
        Randomize
        If Rnd >= 0.5 Then
            Particle(X).TranslateY = Particle(X).TranslateY * -1
        End If
        Randomize
        If Rnd >= 0.5 Then
            Particle(X).TranslateZ = Particle(X).TranslateZ * -1
        End If
    Next
End Sub

Private Sub SetupLight()
    Dim LightAmbient(0 To 3) As Single 'holds the light's ambient color values
    Dim LightDiffuse(0 To 3) As Single 'holds the light's diffuse color values
    Dim LightSpecular(0 To 3) As Single 'holds the light's specular color values
    Dim LightPosition(0 To 3) As Single 'holds the light's position values
    Dim X As Integer 'loop variable
    
    For X = 0 To 3
        LightAmbient(X) = 0.8 'set the light's ambient color to white
        LightDiffuse(X) = 1 'set the light's diffuse color to white
        LightSpecular(X) = 0.2 'set the light's specular color to white
    Next

    LightPosition(0) = 0    'set the light's X position
    LightPosition(1) = 0    'set the light's Y position
    LightPosition(2) = -100 'set the light's Z position

    glDisable glcLighting 'temporarily disable lighting
        glLightfv ltLight0, lpmAmbient, LightAmbient(0) 'set the light's ambient color values
        glLightfv ltLight0, lpmDiffuse, LightDiffuse(0) 'set the light's diffuse color values
        glLightfv ltLight0, lpmSpecular, LightSpecular(0) 'set the light's specular color values
        glLightfv ltLight0, lpmPosition, LightPosition(0) 'set the position of the light
    glEnable glcLighting 're-enable lighting
End Sub

Private Sub chkEnableGravity_Click()
    If chkEnableGravity.Value = 1 Then
        sldGravity.Enabled = True
        lblG1.Enabled = True
        lblG2.Enabled = True
        lblG3.Enabled = True
        lblG4.Enabled = True
        GravityLevel = sldGravity.Value * 2 'set the gravity level
    Else
        sldGravity.Enabled = False
        lblG1.Enabled = False
        lblG2.Enabled = False
        lblG3.Enabled = False
        lblG4.Enabled = False
        GravityLevel = 0 'set the gravity level to 0
    End If
End Sub

Private Sub chkRandomColors_Click()
    Dim X As Integer
    
    For X = 0 To 4
        picColor(X).BorderStyle = 0 'set the border style to none
    Next
    If chkRandomColors.Value = 0 Then
        picColor(0).BorderStyle = 1 'set the border style to solid
    End If
End Sub

Private Sub Form_Load()
    Initialize 'Execute the initialization subroutine
End Sub

Private Sub Form_Paint()
    Render 'Execute the rendering subroutine
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If hGLRC <> 0 Then 'if the OpenGL rendering context still exists
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC 'remove the OpenGL rendering context from memory
    End If
End Sub

Private Sub Form_Resize()
    Dim viewportW As Long 'gl viewport width
    Dim viewportH As Long 'gl viewport height

    viewportW = picOutput.Width  'set gl viewport width
    viewportH = picOutput.Height  'set gl viewport height
    
    wglMakeCurrent picOutput.hDC, hGLRC 'sets this thread's current rendering context
        glViewport 0, 0, viewportW, viewportH 'set the viewport
        glMatrixMode mmProjection 'set the current matrix mode
        glLoadIdentity 'load identity matrix
        glOrtho -500, 500, -500, 500, -500, 500 'create a parallel projection
        gluLookAt 60, 20, 100, 0, 0, 0, 0, 1, 0
        glMatrixMode mmModelView 'set the current matrix mode
        glLoadIdentity 'load the identity matrix
    wglMakeCurrent 0, 0
    Render 'Execute the rendering subroutine
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hGLRC <> 0 Then 'if the OpenGL rendering context still exists
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC 'remove the OpenGL rendering context from memory
    End If
End Sub

Private Sub glRGBA(RedIn, GreenIn, BlueIn, AlphaIn) 'Converts standard RGB values into the gl model (+ Alpha level)
    glColor4f (RedIn / 255), (GreenIn / 255), (BlueIn / 255), AlphaIn
End Sub

Private Sub picColor_Click(Index As Integer) 'controls the appearance of the color selector
    Dim X As Integer
    
    If chkRandomColors.Value = 1 Then
        chkRandomColors.Value = 0
    End If
    picColor(Index).BorderStyle = 1
    For X = 0 To 4
        If X <> Index Then
            picColor(X).BorderStyle = 0
        End If
    Next
End Sub

Private Sub sldGravity_Change()
    GravityLevel = sldGravity.Value * 2 'set the gravity level
End Sub

Private Sub timAnimate_Timer()
    If FrameCount < 150 Then
        Render 'Execute the rendering subroutine
    Else
        timAnimate.Enabled = False
        ParticleCount = sldParticleCount.Value 'update the number of particles for the next iteration
        SetupParticles 'setup the particles
        FrameCount = 0 'reset the frame count
        timAnimate.Enabled = True
    End If
End Sub
