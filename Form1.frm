VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnroll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdEnroll 
      Caption         =   "Enroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdIDSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtSubject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Total Units Enrolled:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Subject ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Studen 
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsStudent As Recordset
Dim rsEnroll As Recordset

Private Sub cmdEnroll_Click()
    
    SQL = "select * from Enroll where StudID = '" & txtID.Text & "' and SubNum = '" & txtSubject.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
        If rsFields("seats") = 0 Then
            MsgBox "Seats already full"
        Else
            SQL = "insert into Enroll values('" & txtID.Text & "','" & txtSubject.Text & "')"
            rsFields ("Seats")
            db.Execute (SQL)
    End If
    txtName.Text = ""
    txtID.Text = ""
    txtAge.Text = ""
    txtCourse.Text = ""
    txtYear.Text = ""
    txtID.SetFocus
    
End Sub

Private Sub cmdIDSearch_Click()
    SQL = "select * from Student where StudID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    
    If rsStudent.BOF = True Then
        MsgBox "ID No. doesn't exist"
    Else
        txtName.Text = rsStudent.Fields("Name")
        cmdSubSearch.Enabled = True

    End If
End Sub

Private Sub cmdSubSearch_Click()
    SQL = "select * from Schedule where SubNum = '" & txtSubject.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    
    If rsStudent.BOF = True Then
        MsgBox "ID No. doesn't exist"
    Else
        txtTitle.Text = rsStudent.Fields("Title")
        txtUnits.Text = rsStudent.Fields("Units")
        cmdIDSearch.Enabled = False
        cmdEnroll.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Student.mdb")
    Set rsStudent = db.OpenRecordset("Student")
    cmdSubSearch.Enabled = False
    cmdIDSearch.Enabled = True
    cmdEnroll.Enabled = False
End Sub


