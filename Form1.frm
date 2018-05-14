VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "barcodes.jpg"
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode only Code128, EAN13 and UPCA"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Create a new instance of Barcode Decoder object
    Dim dec As Object
    Set dec = CreateObject("BarcodeReader.BarcodeDecoder.1")
    dec.BarcodeTypes = &H1 Or &H8 Or &H80 'recognize only Code128, EAN13 and UPCA
    dec.ShowImage = False
    dec.LinearFindBarcodes = 7

    'decode file
    dec.DecodeFile (Text1.Text)

    'show results
    For i = 0 To dec.Barcodes.length - 1
        Dim bc As Barcode
        Set bc = dec.Barcodes.Item(i)
        txt = ""

        If bc.BarcodeType = Codabar Then txt = txt & "Codabar"
        If bc.BarcodeType = Code11 Then txt = txt & "Code11"
        If bc.BarcodeType = Code128 Then txt = txt & "Code128"
        If bc.BarcodeType = Code39 Then txt = txt & "Code39"
        If bc.BarcodeType = Code93 Then txt = txt & "Code93"
        If bc.BarcodeType = EAN13 Then txt = txt & "EAN13"
        If bc.BarcodeType = EAN8 Then txt = txt & "EAN8"
        If bc.BarcodeType = Interl25 Then txt = txt & "Interl25"
        If bc.BarcodeType = Industr25 Then txt = txt & "Industr25"
        If bc.BarcodeType = UPCA Then txt = txt & "UPCA"
        If bc.BarcodeType = UPCE Then txt = txt & "UPCE"
        If bc.BarcodeType = PDF417 Then txt = txt & "PDF417"
        If bc.BarcodeType = LinearUnrecognized Then txt = txt & "Linear Unrecognized"
        If bc.BarcodeType = PDF417Unrecognized Then txt = txt & "PDF417 Unrecognized"

        txt = txt & ": " & bc.Text
        txt = txt & " (" & bc.X1 & "," & bc.Y1 & ")," & "(" & bc.X2 & "," & _
            bc.Y2 & ")," & "(" & bc.x3 & "," & bc.y3 & ")," & "(" & bc.x4 & _
            "," & bc.y4 & ")"
        MsgBox txt
    Next i

    Set dec = Nothing
End Sub

Private Sub Command2_Click()
    'Create a new instance of Barcode Decoder object
    Dim dec As Object
    Set dec = CreateObject("BarcodeReader.BarcodeDecoder.1")

    dec.AboutBox
End Sub
