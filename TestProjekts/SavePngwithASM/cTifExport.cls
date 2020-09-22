VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTifExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Classe permettant d'exporter une image au format TIFF
'


Public Function CreateTIFFImage(cPath As String, ByRef cPic As PictureBox)

    'Crée une image TIFF avec le contenu du PictureBox
    
    Dim Offset  As Long
    Dim cBts()  As String
    Dim I       As Long
    Dim J       As Long
    
    'Taille de l'image + header + offset
    Offset = cPic.Width * cPic.Height * 3 + 8
    
    Open cPath For Binary As #1
        
        'Header TIFF (Little Endian)
        WriteHexString "4d4d002a", 1
        
        'Taille de l'image
        DecomposeLong Offset, cBts
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        'Données de l'image
        For J = 0 To cPic.Height - 1
        For I = 0 To cPic.Width - 1
            DecomposeLong cPic.Point(I, J), cBts
            'Couleurs R, G, B
            Put #1, , cBts(0): Put #1, , cBts(1): Put #1, , cBts(2)
        Next
        Next
        
        'Le "Footer" : nombre d'entrées (ici 14)
        WriteHexString "000e", 1
        
        'La largeur de l'image
        WriteHexString "0100000300000001", 1
        DecomposeLong cPic.Width, cBts
        Put #1, , cBts(1): Put #1, , cBts(0)
        WriteHexString "0000", 1
        
        'La hauteur de l'image
        WriteHexString "0101000300000001", 1
        DecomposeLong cPic.Height, cBts
        Put #1, , cBts(1): Put #1, , cBts(0)
        WriteHexString "0000", 1
        
        'Bits par echantillon
        WriteHexString "0102000300000003", 1
        Offset = cPic.Width * cPic.Height * 3 + 182
        DecomposeLong Offset, cBts
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        WriteHexString "010300030000000100010000", 1    'Compression : aucune
        WriteHexString "010600030000000100020000", 1    'Iterpolation photo : normale
        WriteHexString "011100040000000100000008", 1    'Strip Offset : aucun
        WriteHexString "011200030000000100010000", 1    'Orientation : normale
        WriteHexString "011500030000000100030000", 1    'Echantillon par pixel
        
        'Colonnes par échantillon
        WriteHexString "0116000300000001", 1
        DecomposeLong cPic.Height, cBts
        Put #1, , cBts(1): Put #1, , cBts(0)
        WriteHexString "0000", 1
        
        'Taille de l'image
        WriteHexString "0117000400000001", 1
        Offset = cPic.Width * cPic.Height * 3
        DecomposeLong Offset, cBts
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        'Position de l'échantillon
        WriteHexString "0118000300000003", 1
        Offset = cPic.Width * cPic.Height * 3 + 188
        DecomposeLong Offset, cBts
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        'Tag d'échantillon
        WriteHexString "0119000300000003", 1
        Offset = cPic.Width * cPic.Height * 3 + 194
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        WriteHexString "011c00030000000100010000", 1    'Configuration du plan : normale
        
        'Format d'échantillon
        WriteHexString "0153000300000003", 1
        Offset = cPic.Width * cPic.Height * 3 + 200
        Put #1, , cBts(3): Put #1, , cBts(2)
        Put #1, , cBts(1): Put #1, , cBts(0)
        
        WriteHexString "00000000", 1                    'Fin de l'entrée
        WriteHexString "000800080008", 1                'Mode de couleurs : RGB
        WriteHexString "000000000000", 1                'Valeur minimum : 0
        WriteHexString "00ff00ff00ff", 1                'Valeur maximum : 255
        WriteHexString "000100010001", 1                'Echantillons par canal
        
    Close #1

End Function

Public Function DecomposeLong(ByVal cL As Long, cB() As String)

    'Décompose un Long en octets

    ReDim cB(3)
    
    cB(0) = Chr$(cL Mod 256)
    cB(1) = Chr$((cL \ 256) Mod 256)
    cB(2) = Chr$((cL \ 65536) Mod 256)
    cB(3) = Chr$((cL \ 16777216) Mod 256)
End Function

Public Function WriteHexString(cHS As String, ByVal cFN As Long)
    
    'Place une chaine hexadécimale dans le fichier
    
    Dim I   As Long
    Dim C   As String
    
    For I = 1 To Len(cHS) Step 2
        C = Mid$(cHS, I, 2)
        Put #cFN, , Chr$(Val("&H" & C))
    Next
End Function
