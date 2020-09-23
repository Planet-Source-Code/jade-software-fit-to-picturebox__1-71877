VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pix2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   5040
      ScaleHeight     =   3345
      ScaleWidth      =   3945
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fit To PictureBox"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Pix1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   720
      ScaleHeight     =   3345
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   6600
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2280
      Top             =   3720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-//===========================================================
'-// Thanks to pscode.
'-// If you like this code vote for it.
'-// Website : http://www.rjsoftware.co.cc
'-// Email   : keencoder@gmail.com
'-//===========================================================
Function ResizePicture_To_fit(Img As Image, PicS As PictureBox)

Dim InputHeight, InputWidth
Dim OutputHeight, OutputWidth
Dim NewHeight, NewWidth
Dim Width_Ratio, Height_Ratio, smallest_ratio
Dim remaining_excess, spacer

Img.Visible = False
 
    '-//Actual size of a source image
     InputHeight = Img.Height
    InputWidth = Img.Width
    
    '-//Size of Picture where we put the New calculated Picture
    OutputHeight = PicS.Height
    OutputWidth = PicS.Width
    
    Width_Ratio = OutputWidth / InputWidth
    Height_Ratio = OutputHeight / InputHeight
    
    '-//get the smallest ratio
    If Width_Ratio < Height_Ratio Then
        smallest_ratio = Width_Ratio
    End If
    
    If Height_Ratio < Width_Ratio Then
        smallest_ratio = Height_Ratio
    End If
    
    '-//New Calculated Height/Width
    NewWidth = smallest_ratio * InputWidth
    NewHeight = smallest_ratio * InputHeight
    
    '-//FOR LANDSCAPE
    If NewWidth > NewHeight Then
        remaining_excess = OutputHeight - NewHeight
        spacer = remaining_excess / 2
        PicS.PaintPicture Img.Picture, 0, spacer, NewWidth, NewHeight
    End If
    
    '-//FOR PORTRAIT
    If NewHeight > NewWidth Then
        remaining_excess = OutputWidth - NewWidth
        spacer = remaining_excess / 2
        PicS.PaintPicture Img.Picture, spacer, 0, NewWidth, NewHeight
    End If
    
End Function

Private Sub Command1_Click()
    Image1.Picture = LoadPicture(App.Path & "\night.jpg")
    Pix1.Cls
    Call ResizePicture_To_fit(Image1, Pix1)
    
    Image2.Picture = LoadPicture(App.Path & "\portrait.jpg")
    Pix2.Cls
    Call ResizePicture_To_fit(Image2, Pix2)
    
End Sub
