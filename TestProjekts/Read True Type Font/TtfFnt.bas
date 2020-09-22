Attribute VB_Name = "Module1"
Option Explicit
Dim Symbolfont As Boolean
Dim Temp As String
Const MAXPONT = 1000  '9000
Const maxhead = 32
Const maxendp = 2000
Const maxN = 200
Const csomoK = 3

Type polypoints
    x As Integer
    y As Integer
End Type

Type headertipus
    name1 As String * 4
    checksum As Long
    offset As Long
    size As Long
End Type

Type headtipus
    version As Long
    revision As Long
    checksum As Long
    magicnum As Long
    flags As String * 2
    unitsPerEm As String * 2
    date1 As String * 8
    date2 As String * 8
    xmin As Integer
    ymin As Integer
    xmax As Integer
    ymax As Integer
    macStyle As String * 2
    minPixels As String * 2
    direction As Integer
    locatype As Integer
    glyftype As Integer
End Type

Dim pontok(MAXPONT) As polypoints
Dim ttfheader(maxhead) As headertipus
Dim maxgly As Long
Dim ttfhead As headtipus
Dim x0 As Long
Dim y0 As Long
Dim SplinePrecision As Integer
Dim headerofheader(5) As Long
Dim ttfsize As Long
Dim ttfoffset As Long
Dim NumberOffContours As Integer
Dim flagek(MAXPONT) As Integer
Dim pontdb As Long
Dim xpontdb As Long
Dim EndpointsOffContours(maxendp) As Long
Dim scale1 As Integer
Dim gly As Integer
Dim BXX(maxN) As Integer
Dim BYY(maxN) As Integer
Dim csomoN As Integer
Dim csomo(maxN * 2) As Integer
Dim hh2 As Long
Dim bottomy As Integer
Dim topy As Integer
Dim printdata As Integer
Dim glyftable() As Long
Dim hmtx() As Long
Dim cmap1() As Integer
Dim a9a As Long
Dim numSubtables As Long
Dim subtableFound As Long
Dim thisPlatformID As Long
Dim thisSpecificID As Long
Dim thisSubtableOffset As Long
Dim subtableOffset As Long
Dim subtableLength As Long
Dim cmapFormat As Long
Dim Fontfile As String
Dim Filenumber As Long

Sub Main(Filename As String, Text1 As String, Optional Text2 As String)
Dim Text As String
Dim fo As Long
Dim sgly As Long
Dim z As String
Dim i1 As Integer
Dim b As Integer
Dim Character As String
Dim minx As Integer
Dim miny As Integer
Dim maxx As Integer
Dim maxy As Integer
Dim I As Integer

ClearVariables
Form1.Picture1.Cls

OpenFontFile Filename
printdata = 0

If printdata = 0 Then 'And Symbolfont = False Then
scale1 = 30
x0 = -180
y0 = 300
printstring Text1
x0 = -100
y0 = 200
fo = 0
scale1 = 50
printstring Text2
Exit Sub
End If
fo = 1
sgly = 1
x0 = -600
y0 = 400
For gly = 1 To maxgly
    If printdata = 1 Then
        scale1 = 15
        x0 = 0
        y0 = 0
    End If
     printgly
    If printdata = 0 Then
    If gly < maxgly Then
    If x0 + (hmtx(0, gly + 1) / scale1) > 600 Then
        x0 = -600
        y0 = y0 + (bottomy / scale1 - topy / scale1) - 8
        If y0 + (bottomy / scale1 - topy / scale1) - 8 < -400 Then
            'Debug.Print "Glyphs:"; sgly; "-"; gly
            sgly = gly + 1
            Form1.Picture1.Cls
            fo = 1
            x0 = -600
            y0 = 400
        End If
    End If
    End If
    Else
        If NumberOffContours > 0 Then
        Form1.Picture1.Cls
        End If
    End If
Next gly

'Debug.Print "Glyphs:"; sgly; "-"; gly - 1
Close Filenumber



End Sub
Private Sub AnError(k As Integer)
Dim b As Integer

If printdata = 1 Then Debug.Print "Error!!!! Code="; k
If printdata = 1 Then Debug.Print "NumberOffContours="; NumberOffContours
If printdata = 1 Then Debug.Print "pontdb="; pontdb
If printdata = 1 Then Debug.Print "b="; b
End
End Sub

Private Function bread() As Integer
Dim a11 As Integer

a11 = 0
Temp$ = " "
Get Filenumber, , Temp$
a11 = Asc(Temp$)
bread = a11
End Function

Private Sub bspline(n1 As Integer)
Dim I As Integer
Dim ex9 As Integer
Dim ey9 As Integer
Dim XX9 As Integer
Dim YY9 As Integer
Dim b9 As Double
Dim x9 As Double
Dim y9 As Double
Dim u As Double
Dim ce As Long
Dim uu As Long


b9 = 0
x9 = 0
y9 = 0
u = 0

n1 = n1 - 1
If n1 <= 0 Then Exit Sub
csomoN = n1

For I = 1 To n1
Form1.Picture1.ForeColor = vbGreen
' xline BXX(I), BYY(I), BXX(I + 1), BYY(I + 1) 'without rendering
'Form1.Picture1.Circle (x0 + BXX(I + 1) / scale1, y0 + BYY(I + 1) / scale1), 4, 12
'Form1.Picture1.Circle (x0 + BXX(I) / scale1, y0 + BYY(I) / scale1), 6, 12

'Form1.Picture1.Circle (x0 + (hmtx(0, gly) / 2) / scale1, y0), 1, 13
'Form1.Picture1.Line (x0 + (hmtx(0, gly) / 2) / scale1, y0)-(x0 + BXX(I) / scale1, y0 + BYY(I) / scale1), 8
Next I

Form1.Picture1.ForeColor = vbBlack

If n1 = 1 Then
    xline BXX(1), BYY(1), BXX(2), BYY(2)
    Exit Sub
End If

For I = 0 To n1 * 2
    csomo(I) = CsomoF(I)
Next I

ce = (csomoN - csomoK + 2) * SplinePrecision - 1
        Form1.Picture1.ForeColor = vbBlack

For uu = 0 To ce
    u = uu / SplinePrecision
    x9 = 0
    y9 = 0

    For I = 0 To csomoN
        b9 = NSuly#(I, csomoK, u)
        x9 = x9 + BXX(1 + I) * b9
        y9 = y9 + BYY(1 + I) * b9
    Next I
    XX9 = CInt(x9)
    YY9 = CInt(y9)

    If uu <> 0 Then xline ex9, ey9, XX9, YY9
    ex9 = XX9
    ey9 = YY9
Next uu
Form1.Picture1.ForeColor = vbBlack

xline XX9, YY9, BXX(n1 + 1), BYY(n1 + 1)

End Sub

Private Function CsomoF(I9 As Integer) As Integer
Dim v9 As Long

v9 = 0
If I9 < csomoK Then
    v9 = 0
Else
    If I9 > csomoN Then
        v9 = csomoN - csomoK + 2
    Else
        v9 = I9 - csomoK + 1
    End If
End If
CsomoF = v9
End Function

Private Function GetGlyphIndex(C9 As Integer) As Integer
Dim segCount As Long
Dim endcount As Long
Dim startCount As Long
Dim idDelta As Long
Dim idRangeOffset As Long
Dim glyphIdArray As Long
Dim end1 As Long
Dim start As Long
Dim range As Long
Dim delta As Long
Dim seg1 As Long

segCount = 0
endcount = 0
startCount = 0
idDelta = 0
idRangeOffset = 0
glyphIdArray = 0

segCount = 0
end1 = 0
start = 0
range = 0
delta = 0
seg1 = 0

If subtableLength = 0 Then
    GetGlyphIndex = 0
    Exit Function
End If

Select Case cmapFormat

    Case 0

        glyphIdArray = 6
        C9 = C9 + 20
            If (C9 < 256) Then
                GetGlyphIndex = cmap1(glyphIdArray + C9)
                
                Exit Function
            Else
                GetGlyphIndex = 0
                Exit Function
            End If

    Case 4

        segCount = (256 * cmap1(6) + cmap1(7)) / 2
        endcount = 14
        startCount = 16 + 2 * segCount
        idDelta = 16 + 4 * segCount
        idRangeOffset = 16 + 6 * segCount
        glyphIdArray = 16 + 8 * segCount

        seg1 = 0
        end1 = 256 * cmap1(endcount) + cmap1(endcount + 1)
        While (end1 < C9)
            seg1 = seg1 + 1
            end1 = 256 * cmap1(endcount + seg1 * 2) + cmap1(endcount + seg1 * 2 + 1)
        Wend

        start = 256 * cmap1(startCount + seg1 * 2) + cmap1(startCount + seg1 * 2 + 1)
        delta = 256& * cmap1(idDelta + seg1 * 2) + cmap1(idDelta + seg1 * 2 + 1)
        range = 256 * cmap1(idRangeOffset + seg1 * 2) + cmap1(idRangeOffset + seg1 * 2 + 1)

        If (start > C9) Then
            GetGlyphIndex = 0
            Exit Function
        End If

        If (range = 0) Then
            segCount = C9 + delta
        Else
            segCount = range + (C9 - start) * 2 + ((16 + 6 * segCount) + seg1 * 2)
            segCount = 256 * cmap1(segCount) + cmap1(segCount + 1)
            If (segCount = 0) Then segCount = segCount + delta
        End If

        If segCount > 65535 Then segCount = segCount - 65536
        GetGlyphIndex = segCount

    Case Else
        GetGlyphIndex = segCount

    End Select


End Function

Private Function iread() As Integer
Dim a10 As Integer
Dim a2 As Long
Dim C As Integer
Dim d9 As Integer

a10 = 0
Get Filenumber, , a10
a2 = a10
If a2 < 0 Then a2 = 65536 + a2
C = Int(a2 / 256)
d9 = a2 - 256& * C
a2 = C + 256& * d9
If a2 > 32767 Then a2 = -(32768 - (a2 - 32768))
iread = a2
End Function

Private Function iswap(a4 As Integer) As Integer
Dim a2 As Long
Dim C As Integer
Dim d9 As Integer

a2 = a4
If a2 < 0 Then a2 = 65536 + a2
C = Int(a2 / 256)
d9 = a2 - 256& * C
a2 = C + 256& * d9
If a2 > 32767 Then a2 = -(32768 - (a2 - 32768))
iswap = a2
End Function

Private Function liread() As Long
liread& = iread
End Function

Function lswap(a1 As Long) As Long
Dim a3 As Long
Dim a2 As Double
Dim h As Double
Dim A As Integer
Dim b1 As Double
Dim b9 As Integer
Dim c1 As Long
Dim C As Integer
Dim d9 As Integer

a3 = a1&
a2 = a1&
If a2 < 0 Then a2 = 4294967296# + a2
h = 256 * 256& * 256
A = Int(a2 / h)
b1 = a2 - (h * A)
b9 = Int(b1 / (256 * 256&))
c1 = b1 - ((256 * 256&) * b9)
C = Int(c1 / 256)
d9 = c1 - 256& * C
a2 = h * d9 + (256 * 256&) * C + 256& * b9 + A
If a2 > 2147483647# Then a2 = -(2147483648# - (a2 - 2147483648#))
lswap& = a2
End Function

Private Function NSuly(I10 As Integer, k1 As Integer, u9 As Double) As Double
Dim v As Double
Dim t As Long

v = 0
If k1 = 1 Then
    If (csomo(I10) <= u9#) And (u9# < csomo(I10 + 1)) Then v = 1
Else
    t = csomo(I10 + k1 - 1) - csomo(I10)
    If t <> 0 Then v = (u9# - csomo(I10)) * NSuly#(I10, k1 - 1, u9#) / t
    t = csomo(I10 + k1) - csomo(I10 + 1)
    If t <> 0 Then v = v + (csomo(I10 + k1) - u9#) * NSuly#(I10 + 1, k1 - 1, u9#) / t
End If
NSuly# = v
End Function

Private Sub OpenFont(openfile As String)
Dim A As Integer
Dim I As Integer
Dim a9 As String
Dim numberOfHMetrics As Integer
Dim lastAdvanceWidth As Long
  A = 0
  Open openfile For Binary As Filenumber
     
      ttfseek "head"
      ttfseek "hhea"
      ttfseek "maxp"
      ttfseek "loca"
      ttfseek "hmtx"
      ttfseek "hdmx"
      ttfseek "glyf"
      ttfseek "cmap"
     
      ttfseek "head"
      Get Filenumber, , ttfhead
     
      ttfseek "hdmx"
      'Get Filenumber, , ttfhdmx
    
     ttfseek "loca"
        Select Case Fix((ttfhead.locatype) / 256)
          Case 0
              'If printdata Then Debug.Print "case 0"
              For I = 0 To maxgly
                  glyftable(I) = liread * 2
              Next I
          Case 1
              'If printdata Then Debug.Print "case 1"
              For I = 0 To maxgly
              Get Filenumber, , glyftable(I)
              Next I
              For I = 0 To maxgly
                  glyftable(I) = lswap(glyftable(I))
              Next I
          Case Else
              AnError 16
      End Select
     
      ttfseek "hhea"
        a9 = Space$(36)
        Get Filenumber, , a9
        numberOfHMetrics = 256 * Asc(Mid$(a9, Len(a9) - 1)) + Asc(Right$(a9, 1))
      
      ttfseek "hmtx"
      a9 = "  "
      For I = 0 To numberOfHMetrics
          Get Filenumber, , a9
          hmtx(0, I) = 256& * Asc(Mid$(a9, Len(a9) - 1)) + Asc(Right$(a9, 1))
          Get Filenumber, , a9
          a9a& = 256& * Asc(Mid$(a9, Len(a9) - 1)) + Asc(Right$(a9, 1))
          If a9a& < 32768 Then
              hmtx(1, I) = a9a&
          Else
              hmtx(1, I) = -(32768 - (a9a& - 32768))
          End If
          hmtx(2, I) = hmtx(0, I) - hmtx(1, I)
      Next I
      If numberOfHMetrics < maxgly Then
          lastAdvanceWidth = hmtx(0, I - 1)
          For I = 1 To maxgly - numberOfHMetrics
              hmtx(0, I + numberOfHMetrics) = lastAdvanceWidth
              Get Filenumber, , a9
              a9a& = 256& * Asc(Mid$(a9, Len(a9) - 1)) + Asc(Right$(a9, 1))
             
              If a9a& < 32768 Then
                  hmtx(1, I + numberOfHMetrics) = a9a&
              Else
                  hmtx(1, I + numberOfHMetrics) = -(32768 - (a9a& - 32768))
              End If
              hmtx(2, I + numberOfHMetrics) = hmtx(0, I + numberOfHMetrics) - hmtx(1, I + numberOfHMetrics)
          Next I
      End If
  
    ttfseek "glyf"
      
End Sub

Private Sub ttfseek(tag As String)
Dim I As Integer

For I = 1 To hh2
    If ttfheader(I).name1 = tag$ Then
        ttfsize = lswap(ttfheader(I).size)
        ttfoffset = lswap(ttfheader(I).offset)
        Temp$ = " "
        Get Filenumber, ttfoffset, Temp$
        'If printdata Then Debug.Print "Tag="; tag$; "   Offs="; ttfoffset; "   Size="; ttfsize
        Exit Sub
    End If
Next I
Debug.Print "tag$ Not found: "; tag$

'maxp not found
If tag$ = "maxp" Then
    maxgly = 100
    Exit Sub
End If
'head NOT found
If tag$ = "head" Then Exit Sub
If tag$ = "hhea" Then Exit Sub
If tag$ = "hdmx" Then Exit Sub
AnError 14

End Sub

Private Sub xline(x19 As Integer, y19 As Integer, x29 As Integer, y29 As Integer)
Dim Randx As Long
Dim Randy As Long
Randx = 200
Randy = 100
       Form1.Picture1.Line (x0 + x19 \ scale1 + Randx, 500 - (y0 + y19 \ scale1 + Randy))-(x0 + x29 \ scale1 + Randx, 500 - (y0 + y29 \ scale1 + Randy))
       Form1.Picture1.ForeColor = vbBlack
       Form1.Picture1.ForeColor = vbBlack
End Sub


Private Sub OpenFontFile(Filename As String)
Dim A As Integer
Dim I As Integer
Dim a9 As String
Dim skod As Integer
Dim minx As Integer
Dim miny As Integer
Dim maxx As Integer
Dim maxy As Integer

Filenumber = FreeFile
x0 = -640                      'origin offset
y0 = 100
SplinePrecision = 5
printdata = 1
  
Fontfile$ = Filename ' Enter filename
If Fontfile$ = "" Then Stop

scale1 = 36

A = 0
Open Fontfile$ For Binary As Filenumber
For I = 0 To 5
    Get Filenumber, , A
    headerofheader(I) = A
Next I
hh2 = Fix(headerofheader(2) / 256)
'If printdata = 1 Then Debug.Print "Headers="; hh2
If hh2 > maxhead Then AnError 10
   
For I = 1 To hh2
    Get Filenumber, , ttfheader(I)
Next I
   
ttfseek "maxp"
a9 = Space$(6)
Get Filenumber, , a9
maxgly = 256 * Asc(Mid$(a9, Len(a9) - 1)) + Asc(Right$(a9, 1))
If printdata = 1 Then
    'Debug.Print
    'Debug.Print "maxgly= "; maxgly
    'Debug.Print
End If

ReDim glyftable(maxgly) As Long
ReDim hmtx(2, maxgly) As Long

ttfseek "cmap"

a9 = "  "
Dim te As Long
te = Seek(1)

Get Filenumber, , a9
Get Filenumber, , a9
a9a = 256 * Asc(Left$(a9, 1)) + Asc(Right$(a9, 1))
numSubtables = a9a
'If printdata Then Debug.Print "numSubtables"; numSubtables
subtableFound = 0
I = 0
While ((subtableFound = 0) And (I < numSubtables))
    a9 = "  "
    Get Filenumber, , a9
    a9a = 256 * Asc(Left$(a9, 1)) + Asc(Right$(a9, 1))
    thisPlatformID = a9a
    'If printdata Then Debug.Print "thisPlatformID"; thisPlatformID
   
    Get Filenumber, , a9
    a9a = 256 * Asc(Left$(a9, 1)) + Asc(Right$(a9, 1))
    thisSpecificID = a9a
    'If printdata Then Debug.Print "thisSpecificID"; thisSpecificID
   
    a9 = "    "
    Get Filenumber, , a9
    a9a& = 16777216 * Asc(Left$(a9, 1)) + 65536 * Asc(Mid$(a9, 2, 1)) + 256 * Asc(Mid$(a9, 3, 1)) + Asc(Right$(a9, 1))
    thisSubtableOffset = a9a&
    'If printdata Then Debug.Print "thisSubtableOffset"; thisSubtableOffset

    If (thisPlatformID = 3) And (thisSpecificID = 1) Then '1,0 for MAC
        subtableOffset = thisSubtableOffset
        subtableFound = 1
    End If

    I = I + 1
Wend

If subtableFound = 0 Then
Symbolfont = True 'Symbol Font
Else
Symbolfont = False 'Unicode Font
End If
                                                                                              
a9 = "  "
Get Filenumber, ttfoffset + subtableOffset + 1, a9
a9a& = 256& * Asc(Left$(a9, 1)) + Asc(Right$(a9, 1))
cmapFormat = a9a&
If printdata Then
   ' Debug.Print "cmapFormat "; cmapFormat
End If

Get Filenumber, , a9
a9a = 256 * Asc(Left$(a9, 1)) + Asc(Right$(a9, 1))
subtableLength = a9a
'If printdata Then Debug.Print "subtableLength"; subtableLength

If ((cmapFormat <> 0) And (cmapFormat <> 4)) Then
    Beep
    Stop
End If
If cmapFormat = 0 Then
subtableLength = 262
End If
ReDim cmap1(subtableLength) As Integer
For I = 0 To subtableLength
    a9 = " "
    Get Filenumber, ttfoffset + subtableOffset + I + 1, a9
    cmap1(I) = Asc(a9)
    'Debug.Print "subtableLength"; subtableLength; "cmap1("; I; ")"; cmap1(I)
Next I
Close Filenumber

'If printdata Then GoSub waitkey
OpenFont Fontfile$

skod = 0

gly = 0
Temp$ = " "

For gly = 0 To 85 'rem maxgly
    Get Filenumber, ttfoffset + glyftable(gly), Temp$
    NumberOffContours = iread
    minx = iread
    miny = iread
    maxx = iread
    maxy = iread
    If miny < bottomy Then bottomy = miny
    If maxy > topy Then topy = maxy
Next gly

'Form1.Picture1.Line (0, y0 + topy / scale1)-(640, y0 + topy / scale1), 8        'rem topline
'Form1.Picture1.Line (0, y0)-(640, y0), 8                      'rem baseline
'Form1.Picture1.Line (0, y0 + bottomy / scale1)-(640, y0 + bottomy / scale1), 8  'rem bottomline
'Form1.Picture1.PSet (x0, y0), 1                               'set screen drawpoint

End Sub

Private Sub printgly()
Dim b As Integer
Dim minx As Integer
Dim miny As Integer
Dim maxx As Integer
Dim maxy As Integer
Dim I As Integer
Dim InstructionLength As Integer
Dim eflagrep As Integer
Dim eflag As Integer
Dim XX As Integer
Dim YY As Integer
Dim endpmut As Long
Dim eee As Integer
Dim A As Integer
Dim b8 As Integer

If printdata = 1 Then
    'Debug.Print "maxgly "; maxgly
End If
 
Get Filenumber, ttfoffset + glyftable(gly), Temp$

NumberOffContours = iread
'If printdata = 1 Then Debug.Print "NumberOffContours"; NumberOffContours


If NumberOffContours <= 0 Then
        x0 = x0 + (hmtx(0, gly) / scale1)
        'If printdata Then Form1.Picture1.Circle (x0, y0), 5, 5
         DoNextCharacter
         Exit Sub
End If

If NumberOffContours > maxendp Then AnError 11
minx = iread
miny = iread
maxx = iread
maxy = iread

hmtx(2, gly) = hmtx(2, gly) - maxx + minx

'If printdata Then Form1.Picture1.Line (x0 + minx / scale1, y0 - (miny / scale1))-(x0 + maxx / scale1, y0 - maxy / scale1), 7, B

' If printdata = 1 Then Debug.Print "+- = scale1 "; scale1
'If printdata = 1 Then Debug.Print "/* = SplinePrecision "; SplinePrecision
' If printdata = 1 Then Debug.Print USING; "minx #####.###  "; minx / scale1;
'If printdata = 1 Then Debug.Print USING; "miny #####.###"; miny / scale1
    
' If printdata = 1 Then Debug.Print USING; "maxx #####.###  "; maxx / scale1;
'If printdata = 1 Then Debug.Print USING; "maxy #####.###"; maxy / scale1

If printdata Then
    'Form1.Picture1.Circle (x0 + hmtx(0, gly) / scale1, y0), 2, 14
    'Form1.Picture1.Line (x0, y0 + topy / scale1)-(x0, y0 + bottomy / scale1), 8
    'Form1.Picture1.Line (x0 + hmtx(0, gly) / scale1, y0 + topy / scale1)-(x0 + hmtx(0, gly) / scale1, y0 + bottomy / scale1), 8
End If
      
hmtx(2, gly) = hmtx(0, gly) - maxx + minx
'If printdata Then Debug.Print USING; "advancew #####.###  "; hmtx(0, gly) / scale1;
'If printdata Then Debug.Print "scale1"; scale1
'If printdata = 1 Then Debug.Print USING; "lsb #####.###  "; hmtx(1, gly) / scale1;
'If printdata = 1 Then Debug.Print USING; "rsb #####.###  "; hmtx(2, gly) / scale1

If printdata Then
   ' Form1.Picture1.Circle (x0 + (hmtx(0, gly) / 2) / scale1, y0), 5, 14
   ' Form1.Picture1.Circle (x0 + (hmtx(0, gly) / 2) / scale1, y0), 3, 3
End If
Temp$ = "  "
For I = 1 To NumberOffContours
    Get Filenumber, , Temp$
    EndpointsOffContours(I) = Asc(Left$(Temp$, 1)) + 256& * Asc(Right$(Temp$, 1))
Next I

If NumberOffContours >= 1 Then
    For I = 1 To NumberOffContours
        If EndpointsOffContours(I) < 32768 Then
            a9a = EndpointsOffContours(I)
        Else
            a9a = -(32768 - (EndpointsOffContours(I) - 32768))
        End If
        EndpointsOffContours(I) = iswap(CInt(a9a))
    Next I
End If

InstructionLength = iread
Temp$ = Space$(InstructionLength)
Get Filenumber, , Temp$
Temp$ = " "

pontdb = EndpointsOffContours(NumberOffContours) + 1
If pontdb > MAXPONT Then AnError 12

eflagrep = 0
For I = 0 To pontdb - 1
    If eflagrep <> 0 Then
        eflagrep = eflagrep - 1
    Else
        Get Filenumber, , Temp$
        eflag = Asc(Temp$)
        eflagrep = 0
        If (eflag And 8) <> 0 Then
            Get Filenumber, , Temp$
            eflagrep = Asc(Temp$)
        End If
    End If
    flagek(I) = eflag
Next I

XX = 0
YY = 0

For I = 0 To pontdb - 1
    eflag = flagek(I)
    '{ XX }
    If (eflag And 2) = 0 Then
        '{ +integer / same }
        If (eflag And 16) = 0 Then XX = XX + iread
    Else
        '{ -byte / +byte }
        If (eflag And 16) = 0 Then XX = XX - bread Else XX = XX + bread
    End If
    pontok(I).x = XX
Next I

For I = 0 To pontdb - 1
    eflag = flagek(I)
    '{ YY }
    If (eflag And 4) = 0 Then
        '{ +integer / same }
        If (eflag And 32) = 0 Then YY = YY + iread
    Else
        '{ -byte / +byte }
        If (eflag And 32) = 0 Then YY = YY - bread Else YY = YY + bread
    End If
    pontok(I).y = YY
Next I

If pontdb > 0 And (flagek(0) And 1) = 0 Then
    'If printdata = 1 Then Debug.Print "Hib s! Az els“ pont nincs rajta a g”rb‚n!!" '+#7
End If

xpontdb = 0
For endpmut = 1 To NumberOffContours
    eee = EndpointsOffContours(endpmut) + 1
    I = xpontdb
    b = 1
 
    If printdata Then
        Form1.Picture1.ForeColor = QBColor(12)
        'Form1.Picture1.Line -(x0 + pontok(I).x \ scale1, y0 + pontok(I).y \ scale1)
        'Form1.Picture1.Circle (x0 + pontok(I).x \ scale1, y0 - pontok(I).y \ scale1), 4
    End If
   
    While I <= eee
        A = I
        If I = eee Then A = xpontdb
        BXX(b) = pontok(A).x
        BYY(b) = pontok(A).y
     
        If (flagek(A) And 1) <> 0 Then
            b8 = b
            Form1.Picture1.ForeColor = vbRed
            bspline b
            b = b8
            BXX(1) = BXX(b)
            BYY(1) = BYY(b)
            b = 1
        End If
        I = I + 1
        b = b + 1
        If b >= maxN Then AnError 13
    Wend
    'Color 10
    bspline b - 1

    xpontdb = eee

Next endpmut

End Sub

Private Sub printstring(Text As String)
Dim i1 As Integer
Dim Character As String
For i1 = 1 To Len(Text)
    Character = Mid$(Text, i1, 1)
    gly = GetGlyphIndex(Asc(Character))
    If gly <> 3 Then  'Space
     printgly
     End If
     DoNextCharacter
Next i1
End Sub

Public Sub DoNextCharacter()
x0 = x0 + (hmtx(0, gly) / scale1)
End Sub

Private Sub ClearVariables()
maxgly = 0
x0 = 0
y0 = 0
SplinePrecision = 0
ttfsize = 0
ttfoffset = 0
NumberOffContours = 0
pontdb = 0
xpontdb = 0
scale1 = 0
gly = 0
csomoN = 0
hh2 = 0
bottomy = 0
topy = 0
printdata = 0
a9a = 0
numSubtables = 0
subtableFound = 0
thisPlatformID = 0
thisSpecificID = 0
thisSubtableOffset = 0
subtableOffset = 0
subtableLength = 0
cmapFormat = 0
Fontfile = 0
Filenumber = 0

End Sub
