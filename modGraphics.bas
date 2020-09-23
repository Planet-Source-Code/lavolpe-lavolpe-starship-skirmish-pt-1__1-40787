Attribute VB_Name = "modGraphics"
Option Explicit

'----------------------------------------------------------------------
'Public API Declares...
'----------------------------------------------------------------------
Private Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Const DSna = &H220326
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long

Public Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetMapMode Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function SetMapMode Lib "GDI32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long

Public Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetBkColor Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function GetTextColor Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "GDI32" (lpBitmap As Any) As Long

Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long


'----------------------------------------------------------------------
'Public Type Defs...
'----------------------------------------------------------------------
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'-----------------------------------------------------------------
' Originally from MSDN's sample screensaver application and intermixed
' with a post from VB Thunder and a lot of creativity!

' Modified to include stretching/resizing the source image and to use the top left corner
' pixel as a default transparency color, and flipping images

Public Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, _
                                lBMPsource As Long, bmpRect As RECT, _
                                Optional lMaskColor As Long = -1, _
                                Optional lNewBmpCx As Long, _
                                Optional lNewBmpCy As Long, _
                                Optional lBkgHDC As Long, _
                                Optional bkgX As Long, _
                                Optional bkgY As Long, _
                                Optional FlipHorz As Boolean = False, _
                                Optional FlipVert As Boolean = False)
' Above parameters are described...
' lHDCdest is the DC where the drawing will take place
' destRect is a RECT type indicating the left, top, right & bottom coords where drawing will be done
'   ... by default, this is the size of a background image/mask also
' lBMPsource is the handle to the bitmap to be made transparent and be re-drawn on lHDCdest
' bmpRect is a Rect type indicating the bitmaps coords to use for drawing
'   -- Note: If not provided, the entire bitmap is used.
'   -- If the bitmap is a sheet of frames, the rectangle is the frame coords
' lMaskColor is the bitmap color to be made transparent. The value of -1 picks the top left corner pixel
' lNewBmpCx is the eventual destination width of the source bitmap
'  -- Note: If not provided, the bitmap width is drawn with a 1:1 ratio
' lNewBmpCy is the destination height of the source bitmap
' -- Note: If not provided, the bitmap height is drawn with a 1:1 ratio
' -- however, if overall size is greater than supplied dRect.Right or dRect.Bottom, it will be shrunk
' ************ Following parameters are used if a separate HDC is used as a background or mask
'                 to be used for drawing. This option is used primarily as a background for animation
' lBkgHDC is the DC of the background image container
' bkgX, bkgYare the upper left/top coords to use on the background/mask DC for drawing on the
'   the destination DC. The width and height are determined by destRect's overall width/height
' FlipHorz if passed will flip the image 180 degrees horizontally
' FlipVert if passed will flip the image 180 degrees vertically
'-----------------------------------------------------------------
    Dim udtBitMap As BITMAP
    Dim lMask2Use As Long
    ' Bitmaps to use
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    ' HDCs to use
    Dim lHDCMem As Long, lHDCscreen As Long, lHDCsrc As Long
    Dim lHDCMask As Long, lHDCcolor As Long
    ' Used if image is to be flipped vertically or horizontally
    Dim OrientX As Long, OrientY As Long
    ' X,Y is the drawing region width,height
    ' srcX,srcY is the unresized image size
    Dim X As Long, Y As Long, srcX As Long, srcY As Long
    Dim hPalSrc As Long, hPalMem As Long
    ' used to shrink the image if needed
    Dim lRatio(0 To 1) As Single
'-----------------------------------------------------------------
    lHDCscreen = GetDC(0&)                       ' use desktop as template
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    
    SelectObject lHDCsrc, lBMPsource             'Select the bitmap
    GetObject lBMPsource, Len(udtBitMap), udtBitMap
    ' get the background/transparent color if not provided
    lMask2Use = lMaskColor
    If lMask2Use < 0 Then lMask2Use = GetPixel(lHDCsrc, 0, 0)
    
    ' Bmp size needed for original source
        srcX = udtBitMap.bmWidth                  'Get width of bitmap
        srcY = udtBitMap.bmHeight                 'Get height of bitmap
        If lNewBmpCx = 0 Then
            ' if the image height isn't being resized, then set the width here
            If bmpRect.Right > 0 Then lNewBmpCx = bmpRect.Right - bmpRect.Left Else lNewBmpCx = srcX
        End If
        If lNewBmpCy = 0 Then
            ' if the image height isn't being resized, then set the height here
            If bmpRect.Bottom > 0 Then lNewBmpCy = bmpRect.Bottom - bmpRect.Top Else lNewBmpCy = srcY
        End If
        ' now calculate the dimensions of the passed image if needed
        If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
        If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
        ' Calculate size needed for drawing
        If (destRect.Right) = 0 Then X = lNewBmpCx Else X = (destRect.Right - destRect.Left)
        If (destRect.Bottom) = 0 Then Y = lNewBmpCy Else Y = (destRect.Bottom - destRect.Top)
'=========================================================================
' This routine will fail to draw properly if you try to draw a  larger dimension (lNewBmpCX or lNewBmpCy
' that is larger than the destination dimensions. Therefore, if the source dimensions are larger, then
' the routine will attempt to automatically scale the source image as needed.
'=========================================================================
        If lNewBmpCx > X Or lNewBmpCy > Y Then
            lRatio(0) = (X / (bmpRect.Right - bmpRect.Left))
            lRatio(1) = (Y / (bmpRect.Bottom - bmpRect.Top))
            If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)
            lNewBmpCx = lRatio(0) * (bmpRect.Right - bmpRect.Left)
            lNewBmpCy = lRatio(0) * (bmpRect.Bottom - bmpRect.Top)
            Erase lRatio
        End If
            
    
    'Create some DCs to hold temporary data
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    'Create a bitmap for each DC.  DCs are required for a number of GDI functions
    'Compatible DC's
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, X, Y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    
    'Each DC must select a bitmap object to store pixel data.
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ReleaseDC 0&, lHDCscreen
    
' ====================== Start working here ======================
    
    'Set proper mapping mode.
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    'Copy the background of the destination DC
    If (lBkgHDC <> 0) Then
        BitBlt lHDCMem, 0, 0, X, Y, lBkgHDC, bkgX, bkgY, vbSrcCopy
    Else
        BitBlt lHDCMem, 0&, 0&, X, Y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
    End If
    
    ' set the back/forecolor of the working DC to match the original image
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    ' Get working copy of source bitmap
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
    If FlipHorz Then StretchBlt lHDCcolor, srcX - 1, 0, -srcX, srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
    If FlipVert Then StretchBlt lHDCcolor, 0, srcY - 1, srcX, -srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
    
    ' set working color back/fore colors. These colors will help create the mask
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    
    'Create the object mask for the bitmap by performaing a BitBlt
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    ' This will create a mask of the source color
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna

    'Mask out the places where the bitmap will be placed while resizing as needed
    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    
    'XOR the bitmap with the background on the destination DC while resizing as needed
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    
    'Copy to the destination
    BitBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0&, 0&, vbSrcCopy
    
    
    'Delete memory bitmaps
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    
    'Delete memory DC's
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc
    
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

Public Function ExtractBmpRects(hBmpHandle As Long, destHDC As Long, _
                        bmpRect() As RECT, Optional lBackColor As Long, _
                        Optional HardX As Long, Optional HardY As Long) As Boolean

' Homegrown function.  If an animated bitmap is created using the Code function. Then each frame in
' the bitmap has its left, top, right & bottom coords encoded on the last line(s) of the bitmap.  The
' number of frames in the bitmap is also encoded

' hBmpHandle is the handle to the coded, animated bitmap. Usually loaded via the LoadPicture function
' destHDC is the hDC where animation will occur
' bmpRect are an empty array of rectangles to store the frame coordinates in
' lBackColor will be given the backcolor of the animated bitmap to be used as a transparency color

If hBmpHandle = 0 Then Exit Function

Dim lHDCTemp As Long, udtBitMap As BITMAP, LastX As Long, LastY As Long
Dim yOffset As Long, X As Long, Looper As Long, NextY As Long, RectID As Long

' create a temporary DC to load the bitmap & get pixel colors
lHDCTemp = CreateCompatibleDC(destHDC)
SelectObject lHDCTemp, hBmpHandle
GetObject hBmpHandle, Len(udtBitMap), udtBitMap

' now we need to find out how many lines of coding were needed, this is always the 1st
' pixel of the last line of the bitmap
yOffset = GetPixel(lHDCTemp, 0, udtBitMap.bmHeight - 1)

' quick test to help determine if this is a coded bitmap or not...
If yOffset < 1 Or yOffset > 2000 Then GoTo CloseFunction

' let's try to continue.
' Now we want to determine where the 1st line of code actually starts. Most bitmaps only require
' one line of code, but if you have a small frame (say 64bytes) and have several of them, let's say
' over 20, then the number of pixels needed for coding exceed one line and 2 or 3 may be needed
NextY = udtBitMap.bmHeight - yOffset

' re-index the Rectangle array to that needed. The 2nd pixel always indicates number of frames
ReDim bmpRect(0 To GetPixel(lHDCTemp, 1, NextY))

' in coded animated bitmaps, the top/left corner is always the transparency color
lBackColor = GetPixel(lHDCTemp, 0, 0)

' we count pixels to know when to go to the next line of coding, if needed
X = 2
Do Until RectID > UBound(bmpRect)
    For Looper = 2 To 4         '   (left, top, right, bottom)
        If X = udtBitMap.bmWidth Then   ' got to edge of bitmap, let's go to next line
            NextY = NextY + 1  ' increment line number
            X = 2                      ' reset pixel start point
        End If
        Select Case Looper
            ' Left coordinate is always zero
        Case 2: ' Top coordinate
            bmpRect(RectID).Top = GetPixel(lHDCTemp, X, NextY)
        Case 3: ' Right coordinate
            bmpRect(RectID).Right = GetPixel(lHDCTemp, X, NextY)
        Case 4: ' Bottom coordinate
            bmpRect(RectID).Bottom = GetPixel(lHDCTemp, X, NextY)
        End Select
        X = X + 1           ' increment pixel count
    Next
    If RectID Then
        ' here we calculate the maximum height and width needed to
        ' display all frames. Some animated images contain frames
        ' of different heights and widths
        If bmpRect(RectID).Right <> LastX Then
            If bmpRect(RectID).Right > LastX Then
                HardX = bmpRect(RectID).Right
                LastX = HardX
            Else
                HardX = LastX
            End If
        End If
        If bmpRect(RectID).Bottom - bmpRect(RectID).Top <> LastY Then
            If bmpRect(RectID).Bottom - bmpRect(RectID).Top > LastY Then
                HardY = bmpRect(RectID).Bottom - bmpRect(RectID).Top
                LastY = HardY
            Else
                HardY = LastY
            End If
        End If
    Else
        ' set a starting point for max height and width
        LastX = bmpRect(RectID).Right
        LastY = bmpRect(RectID).Bottom - bmpRect(RectID).Top
    End If
    RectID = RectID + 1     ' increment Rectangle count
Loop
ExtractBmpRects = True
CloseFunction:
DeleteDC lHDCTemp       ' remove HDC from memory
DoEvents
End Function

'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Function ResizeBMP(dispHdc As Long, DestDC As Long, hBmp As Long, RatioX As Single, RatioY As Single) As Long
'-----------------------------------------------------------------
    Dim hBmpOut As Long                             ' output bitmap handle
    Dim bm1 As BITMAP, bm2 As BITMAP                ' temporary bitmap structs
    Dim hdcMem1 As Long, hdcMem2 As Long            ' temporary memory bitmap handles...
'-----------------------------------------------------------------
    hdcMem1 = CreateCompatibleDC(dispHdc)           ' create mem DC compatible to the display DC
    hdcMem2 = CreateCompatibleDC(dispHdc)           ' create mem DC compatible to the display DC
  
    GetObject hBmp, LenB(bm1), bm1                  ' select bitmap object
  
    LSet bm2 = bm1                                  ' copy bitmap object
  
    bm2.bmWidth = CLng(bm2.bmWidth * RatioX)        ' scale output bitmap width
    bm2.bmHeight = CLng(bm2.bmHeight * RatioY)      ' scale output bitmap height
    bm2.bmWidthBytes = ((((bm2.bmWidth * bm2.bmBitsPixel) + 15) \ 16) * 2) ' calculate bitmap width bytes

    hBmpOut = CreateBitmapIndirect(bm2)             ' create handle to output bitmap indirectly from new bm2
    
    SelectObject hdcMem1, hBmp                      ' select original bitmap into mem dc
    ' stretch old bitmap into new bitmap
    If DestDC Then
        'SelectObject DestDC, hBmpOut
        StretchBlt DestDC, 0, 0, bm2.bmWidth, bm2.bmHeight, _
               hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy
    Else
        SelectObject hdcMem2, hBmpOut                   ' select new bitmap into mem dc
        StretchBlt hdcMem2, 0, 0, bm2.bmWidth, bm2.bmHeight, _
               hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy
    End If

    
    DeleteDC hdcMem1                                ' delete memory dc
    DeleteDC hdcMem2                                ' delete memory dc
    If DestDC = 0 Then
        ResizeBMP = hBmpOut                             ' return handle to new bitmap
    Else
       DeleteObject hBmpOut
    End If
        
'-----------------------------------------------------------------
End Function
