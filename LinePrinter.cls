VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
        ' -----------------------------------------------------------------
'
' Clase para poder imprimir en forma directa
'
' -----------------------------------------------------------------
Private mDebug As Boolean
Public mPicture As PictureBox
Private mFino As Integer
Private mGrueso
Private LastX As Single
Private LastY As Single

Public Property Get Thin() As Integer
  Thin = mFino
End Property

Public Property Let Thin(ByVal vNewValue As Integer)
  mFino = vNewValue
End Property

Public Property Get Fat() As Integer
  Fat = mGrueso
End Property

Public Property Let Fat(ByVal vNewValue As Integer)
  mGrueso = vNewValue
End Property

Public Sub Box(x As Single, y As Single, w As Single, H As Single, Optional Fat, Optional color)
Dim aux As Long
Dim aux1 As Long

  If mDebug Then
  
    aux = mPicture.DrawWidth
    mPicture.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    aux1 = mPicture.FillColor
    If Not IsMissing(color) Then
      mPicture.FillStyle = 0 'SOLIDO PARA QUE TOME EL COLOR
      mPicture.FillColor = vbYellow
      mPicture.Line (x, y)-(x + w, y + H), , B
      mPicture.FillStyle = 1 ' TRANSPARENTE
      mPicture.FillColor = aux1
    Else
      mPicture.Line (x, y)-(x + w, y + H), , B
    End If
    mPicture.DrawWidth = aux
      
  Else
    Printer.FontTransparent = True
    aux = Printer.DrawWidth
    Printer.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    aux1 = Printer.FillColor
    If Not IsMissing(color) Then
      Printer.FillStyle = 0 'SOLIDO PARA QUE TOME EL COLOR
      Printer.FillColor = vbYellow
      Printer.Line (x, y)-(x + w, y + H), , B
      Printer.FillStyle = 1 ' TRANSPARENTE
      Printer.FillColor = aux1
    Else
      Printer.Line (x, y)-(x + w, y + H), , B
    End If
    Printer.DrawWidth = aux
  End If
End Sub

Public Sub LineH(x As Single, y As Single, l As Single, Optional Fat)
Dim aux As Long
  If mDebug Then
    aux = mPicture.DrawWidth
    mPicture.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    mPicture.Line (x, y)-(x + l, y)
    mPicture.DrawWidth = aux

  Else
    aux = Printer.DrawWidth
    Printer.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    Printer.Line (x, y)-(x + l, y)
    Printer.DrawWidth = aux
  End If
End Sub

Public Sub LineV(x As Single, y As Single, l As Single, Optional Fat)
Dim aux As Long
  If mDebug Then
    aux = mPicture.DrawWidth
    mPicture.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    mPicture.Line (x, y)-(x, y + l)
    mPicture.DrawWidth = aux

  Else
    aux = Printer.DrawWidth
    Printer.DrawWidth = IIf(IsMissing(Fat), mFino, mGrueso)
    Printer.Line (x, y)-(x, y + l)
    Printer.DrawWidth = aux
  End If
End Sub

Public Function XCentrar(Texto As String) As Single
        If mDebug Then
            XCentrar = (mPicture.width - mPicture.TextWidth(Texto)) / 2
        Else
            XCentrar = (Printer.width - Printer.TextWidth(Texto)) / 2
        End If
End Function

Public Sub Text(Optional ByVal x, Optional ByVal y, Optional ByVal Texto, Optional size, Optional bold, Optional font, Optional Italic)
  If mDebug Then
    mPicture.font = IIf(IsMissing(font), "Arial", font)
    mPicture.FontSize = IIf(IsMissing(size), 10, size)
    mPicture.FontBold = IIf(IsMissing(bold), False, bold)
    mPicture.FontItalic = IIf(IsMissing(Italic), False, Italic)
    mPicture.CurrentX = IIf(IsMissing(x), LastX, x)
    mPicture.CurrentY = IIf(IsMissing(y), LastY, y)
    LastX = mPicture.CurrentX
    LastY = mPicture.CurrentY
    mPicture.Print IIf(IsMissing(Texto), "", Texto)
  Else
    Printer.font = IIf(IsMissing(font), "Arial", font)
    Printer.FontSize = IIf(IsMissing(size), 10, size)
    Printer.FontBold = IIf(IsMissing(bold), False, bold)
    Printer.FontItalic = IIf(IsMissing(Italic), False, Italic)
    Printer.CurrentX = IIf(IsMissing(x), LastX, x)
    Printer.CurrentY = IIf(IsMissing(y), LastY, y)
    LastX = Printer.CurrentX
    LastY = Printer.CurrentY
    Printer.Print IIf(IsMissing(Texto), "", Texto)
End If
End Sub


Public Sub Bmp(x As Single, y As Single, BMPFileName As String, MergeBMP As Boolean, Optional width1, Optional Height1)
  On Error Resume Next
  If mDebug Then
   If MergeBMP Then 'Mezclamos con el fondo ?
      If IsMissing(width1) Or IsMissing(Height1) Then
          mPicture.PaintPicture LoadPicture(BMPFileName), x, y, , , , , , , vbSrcAnd
      Else
          mPicture.PaintPicture LoadPicture(BMPFileName), x, y, width1, Height1, , , , , vbSrcAnd
      End If
    Else
      If IsMissing(width1) Or IsMissing(Height1) Then
          mPicture.PaintPicture LoadPicture(BMPFileName), x, y
      Else
          mPicture.PaintPicture LoadPicture(BMPFileName), x, y, width1, Height1
      End If
    End If
  Else
    If MergeBMP Then
      If IsMissing(width1) Or IsMissing(Height1) Then
        Printer.PaintPicture LoadPicture(BMPFileName), x, y, , , , , , , vbSrcAnd
      Else
        Printer.PaintPicture LoadPicture(BMPFileName), x, y, width1, Height1, , , , , vbSrcAnd
      End If
    Else
      If IsMissing(width1) Or IsMissing(Height1) Then
        Printer.PaintPicture LoadPicture(BMPFileName), x, y
      Else
        Printer.PaintPicture LoadPicture(BMPFileName), x, y, width1, Height1
      End If
    End If
  End If
  On Error GoTo 0
End Sub


Public Sub Init(top As Single, left As Single, Optional ByRef PreViewPicture, Optional PaperSize)
' Si no enviamos picture entiende que lo envia a la impresora
  
  mDebug = Not IsMissing(PreViewPicture)
  
  If mDebug Then
    Set mPicture = PreViewPicture
  
    mPicture.ScaleMode = 6 ' milimetros
    mPicture.CurrentX = 0
    mPicture.CurrentY = 0
    mPicture.ScaleLeft = -left
    mPicture.ScaleTop = -top
    mFino = mFino / 4
    mGrueso = mGrueso / 4
  Else
    
    'Printer.PaperSize = IIf(IsMissing(PaperSize), vbPRPSLetter, PaperSize)
    
    'Printer.PaperSize = 1
    Printer.ScaleMode = 6 ' milimetros
    Printer.ScaleLeft = -left
    Printer.ScaleTop = -top
    
    '--- Anti Bug para que funcione FontTransparent = True
    '--- cuando se especifica antes PaperSize :
    Printer.FontTransparent = False
    Printer.Print
    Printer.FontTransparent = True
    '-------------------------------------------------------
        
    Printer.CurrentX = 0
    Printer.CurrentY = 0
  End If
  LastX = 0
End Sub

Public Sub SendToPrinter()
 If Not mDebug Then Printer.EndDoc
End Sub

Public Sub NewPage()
 If Not mDebug Then Printer.NewPage
End Sub

Public Sub Kill()
  If Not mDebug Then Printer.KillDoc
End Sub

Private Sub Class_Initialize()
    If mFino = 0 Then mFino = 1
    If mFat = 0 Then mFat = 2
End Sub
