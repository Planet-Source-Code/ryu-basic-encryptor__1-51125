VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryptor"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear all"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Your Decrypted text here"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Input your Encrypted text here"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Your text when encrypted"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Input your text here"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "="
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "="
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
'                                                           '
'   a tutorial in basics of encrypting strings              '
'   the fundamentals of encrypting strings starts           '
'   with a small conversion of strings into ascii codes     '
'   into a vast and complex conversion of strings to codes  '
'   this tutorial will focus on how to encrypt and decrypt  '
'   strings into ascii codes and vice versa.                '
'   this code is also essential for password making         '
'   and entering of data in the database                    '
'   have fun examining the source!                          '
'                                                           '
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
'                                                           '
'                                                           '
'   written by: paolo parungao / cherry anne gascon         '
'   9 December 2003                                         '
'   release 1.0                                             '
'                                                           '
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'

Option Explicit     ' this is supposed to be the first line of command
                    ' you must supply for the compiler to check the
                    ' variables used

Private Sub Command1_Click()
Encrypt Text1.Text
Text3.Text = ShowEncryptedText
End Sub

Private Sub Command2_Click()
Decrypt Text2.Text
Text4.Text = ShowDecryptedText
End Sub

Private Sub Command3_Click()        'this will reset all values
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
ShowDecryptedText = vbNullString
ShowEncryptedText = vbNullString
End Sub

'have fun, enjoy, and learn from this simple tutorial.
'have a nice day :)

'paolo and cherry

