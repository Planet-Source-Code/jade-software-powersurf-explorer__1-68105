VERSION 5.00
Begin VB.Form frmverifyremove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verification ::."
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "|"
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmverifyremove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
            
    If Text1 = "" Then
            
                MsgBox "Invalid Entry. Please try again.", vbCritical
                Text1 = ""
                Exit Sub
            
            Else
            
            Set RsBLockPass = New ADODB.Recordset
                RsBLockPass.Open "SELECT * FROM TBL_LOCK WHERE PASS='" & Text1 & "'", CN, adOpenStatic, adLockOptimistic
                
                If Not RsBLockPass.EOF Then
                    
                        Set RSDelBlock = New ADODB.Recordset
                        RSDelBlock.Open "SELECT * FROM TBL_BLOCKED_WEBSITES WHERE URLAdd='" & frmoptions.LstURLDisplay.List(frmoptions.LstURLDisplay.ListIndex) & "'", CN, adOpenStatic, adLockOptimistic
                        
                        With RSDelBlock
                        
                           If .RecordCount > 0 Then
                                .Delete
                                .Requery '-// REFRESH
                            Else
                                
                                Unload Me
                                frmoptions.Text1 = ""
                                MsgBox "Record Not Found!", vbCritical, "Not Found"
                                
                                Exit Sub
                           End If
                           
                        End With
                        
                            frmoptions.LstURLDisplay.RemoveItem (frmoptions.LstURLDisplay.ListIndex)
                            Unload Me
                            MsgBox "Record successfully removed.", vbInformation
                            frmoptions.Text1.Text = ""
                        Else
                    
                            MsgBox "Invalid Password. Please try again!", vbCritical
                            Text1 = ""
                            Exit Sub
                  End If
                        
                        Set rsBlocksite = Nothing
End If


End Sub
