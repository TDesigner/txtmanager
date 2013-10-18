' Classe che permette la formattazione del testo della RichTextBox 
' Creata da Terze parti 
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class AdvRichTextBox
    Inherits RichTextBox

    ''' <summary>
    ''' Specifies how text in a <see cref="AdvRichTextBox"/> is
    ''' horizontally aligned.
    ''' </summary>
    Public Enum TextAlign As Short
        ''' <summary>
        ''' The text is aligned to the left.
        ''' </summary>
        Left = 1

        ''' <summary>
        ''' The text is aligned to the right.
        ''' </summary>
        Right = 2

        ''' <summary>
        ''' The text is aligned in the center.
        ''' </summary>
        Center = 3

        ''' <summary>
        ''' The text is justified.
        ''' </summary>
        Justify = 4
    End Enum

    ' Constants from the Platform SDK.
    Private Const EM_GETPARAFORMAT As Integer = 1085
    Private Const EM_SETPARAFORMAT As Integer = 1095
    Private Const EM_SETTYPOGRAPHYOPTIONS As Integer = 1226
    Private Const TO_ADVANCEDTYPOGRAPHY As Integer = 1
    Private Const PFM_ALIGNMENT As Integer = 8
    Private Const SCF_SELECTION As Integer = 1

    ' It makes no difference if we use PARAFORMAT or
    ' PARAFORMAT2 here, so I have opted for PARAFORMAT2.
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure PARAFORMAT
        Public cbSize As Integer
        Public dwMask As UInteger
        Public wNumbering As Short
        Public wReserved As Short
        Public dxStartIndent As Integer
        Public dxRightIndent As Integer
        Public dxOffset As Integer
        Public wAlignment As Short
        Public cTabCount As Short
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> _
        Public rgxTabs As Integer()

        ' PARAFORMAT2 from here onwards.
        Public dySpaceBefore As Integer
        Public dySpaceAfter As Integer
        Public dyLineSpacing As Integer
        Public sStyle As Short
        Public bLineSpacingRule As Byte
        Public bOutlineLevel As Byte
        Public wShadingWeight As Short
        Public wShadingStyle As Short
        Public wNumberingStart As Short
        Public wNumberingStyle As Short
        Public wNumberingTab As Short
        Public wBorderSpace As Short
        Public wBorderWidth As Short
        Public wBorders As Short
    End Structure

    <DllImport("user32", CharSet:=CharSet.Auto)> _
    Private Shared Function SendMessage(hWnd As HandleRef, msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    <DllImport("user32", CharSet:=CharSet.Auto)> _
    Private Shared Function SendMessage(hWnd As HandleRef, msg As Integer, wParam As Integer, ByRef lp As PARAFORMAT) As Integer
    End Function

    Public Shadows Property SelectionAlignment() As TextAlign
        Get
            Dim fmt As New PARAFORMAT()
            fmt.cbSize = Marshal.SizeOf(fmt)

            ' Get the alignment.
            SendMessage(New HandleRef(Me, Handle), EM_GETPARAFORMAT, SCF_SELECTION, fmt)

            ' Default to Left align.
            If (fmt.dwMask And PFM_ALIGNMENT) = 0 Then
                Return TextAlign.Left
            End If

            Return CType(fmt.wAlignment, TextAlign)
        End Get

        Set(value As TextAlign)
            Dim fmt As New PARAFORMAT()
            fmt.cbSize = Marshal.SizeOf(fmt)
            fmt.dwMask = PFM_ALIGNMENT
            fmt.wAlignment = CShort(value)

            ' Set the alignment.
            SendMessage(New HandleRef(Me, Handle), EM_SETPARAFORMAT, SCF_SELECTION, fmt)
        End Set
    End Property

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)

        ' Enable support for justification.
        SendMessage(New HandleRef(Me, Handle), EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY)
    End Sub

End Class
