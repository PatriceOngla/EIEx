Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Partial Module ExcelCommander
    'Code pour la récupération des coordonnnées Windows d'une cellule Excel.
    'Pas testé.

#Region "APIs"

    <DllImport("gdi32.dll")>
    Private Function GetDeviceCaps(hdc As IntPtr, nIndex As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Function GetDC(hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Boolean
    End Function
    Private Const LOGPIXELSX As Integer = 88
    Private Const LOGPIXELSY As Integer = 90

    Public Structure RECT
        Dim Left As Long
        Dim Top As Long
        Dim Right As Long
        Dim Bottom As Long
    End Structure

#End Region

#Region "1er implémentation"
    'From https://www.add-in-express.com/forum/read.php?FID=5&TID=10884

    Private Function GetCellPosition_1(range As Range) As System.Drawing.Point
        Dim ws As Worksheet = range.Worksheet
        Dim hdc As IntPtr = GetDC(CType(0, IntPtr))
        Dim px As Long = GetDeviceCaps(hdc, LOGPIXELSX)
        Dim py As Long = GetDeviceCaps(hdc, LOGPIXELSY)
        ReleaseDC(CType(0, IntPtr), hdc)
        Dim zoom As Double = XL.ActiveWindow.Zoom

        Dim pointsPerInch = XL.Application.InchesToPoints(1)
        ' usually 72   
        Dim zoomRatio = zoom / 100
        Dim x = XL.ActiveWindow.PointsToScreenPixelsX(0)

        ' Coordinates of current column   
        x = Convert.ToInt32(x + range.Left * zoomRatio * px / pointsPerInch)

        ' Coordinates of next column   
        'x = Convert.ToInt32(x + (((Range)(ws.Columns)[range.Column]).Width + range.Left) * zoomRatio * px / pointsPerInch);   
        Dim y = XL.ActiveWindow.PointsToScreenPixelsY(0)
        y = Convert.ToInt32(y + range.Top * zoomRatio * py / pointsPerInch)

        Marshal.ReleaseComObject(ws)
        Marshal.ReleaseComObject(range)

        Return New System.Drawing.Point(x, y)
    End Function

#End Region

#Region "2ème implémentation"
    'From http://www.mrexcel.com/forum/excel-questions/765416-how-get-x-y-screen-coordinates-excel-cell-range.html

    Private Function ScreenDPI(bVert As Boolean) As Long
        'in most cases this simply returns 96
        Static lDPI&(1), lDC&
        If lDPI(0) = 0 Then
            lDC = GetDC(0)
            lDPI(0) = GetDeviceCaps(lDC, 88&)    'horz
            lDPI(1) = GetDeviceCaps(lDC, 90&)    'vert
            lDC = ReleaseDC(0, lDC)
        End If
        'ScreenDPI = lDPI(Abs(bVert))
        ScreenDPI = lDPI(If(bVert, 1, 0))
    End Function

    Private Function PTtoPX(Points As Single, bVert As Boolean) As Long
        PTtoPX = Points * ScreenDPI(bVert) / 72
    End Function

    Sub GetRangeRect(ByVal rng As Range, ByRef rc As RECT)
        Dim wnd As Window

        'Il faut du code en plus pour vérfier le scroll et la visibilité du range.

        wnd = rng.Parent.Parent.Windows(1)
        With rng
            rc.Left = PTtoPX(.Left * wnd.Zoom / 100, 0) _
              + wnd.PointsToScreenPixelsX(0)
            rc.Top = PTtoPX(.Top * wnd.Zoom / 100, 1) _
             + wnd.PointsToScreenPixelsY(0)
            rc.Right = PTtoPX(.Width * wnd.Zoom / 100, 0) _
               + rc.Left
            rc.Bottom = PTtoPX(.Height * wnd.Zoom / 100, 1) _
                + rc.Top
        End With
    End Sub


#End Region

End Module
