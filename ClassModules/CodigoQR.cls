VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoQR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Autor:  Leandro Ascierto
'Web:    http://leandroascierto.com
'Date:   09/09/2011
'Help:
'     Encode:  http://qrserver.com/api/documentation/create-qr-code/
'     Decode:  http://zxing.org/w/decode.jspx
'----------------------------------

'wininet.dll
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternet As Long) As Boolean
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'gdiplus.dll
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, ByRef Bitmap As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal mstream As IUnknown, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mHbmReturn As Long, ByVal mBackground As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
'KERNEL32.dll
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'ole32.dll
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As GUID) As Long
'olepro32.dll
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
'msvbvm60.dll
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

'Structure
Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type
 
Private Type GUID
    data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type
 
Private Type EncoderParameter
    GUID            As GUID
    NumberOfValues  As Long
    Type            As Long
    Value           As Long
End Type
 
Private Type EncoderParameters
    Count           As Long
    Parameter(15)   As EncoderParameter
End Type

Private Type PICTDESC
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

'Constans
Private Const INTERNET_DEFAULT_HTTP_PORT        As Long = 80
Private Const INTERNET_OPEN_TYPE_PRECONFIG      As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT         As Long = 1
Private Const INTERNET_SERVICE_HTTP             As Long = 3
Private Const INTERNET_FLAG_NO_AUTO_REDIRECT    As Long = &H200000
Private Const INTERNET_FLAG_RELOAD              As Long = &H80000000

Private Const HTTP_QUERY_STATUS_CODE            As Long = 19
Private Const HTTP_STATUS_OK                    As Long = 200
Private Const CP_UTF8                           As Long = 65001
Private Const GdiPlusVersion                    As Long = 1
Private Const ImageCodecPNG                     As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"


Dim GdipToken As Long
'---------------------------------------------------------------------
'Encoding: "ISO-8859-1" or "UTF-8"
'
'ErrCorrectionLevel:
'L (low, ~7% destroyed data may be corrected)
'M (middle, ~15% destroyed data may be corrected)
'Q (quality, ~25% destroyed data may be corrected)
'H (high, ~30% destroyed data may be corrected)
'
'Read More http://qrserver.com/api/documentation/create-qr-code/
'----------------------------------------------------------------------
Public Function GetPictureQrCode(ByVal sText As String, _
                                 ByVal Width As Long, _
                                 ByVal Height As Long, _
                                 Optional ByVal Encoding As String = "UTF-8", _
                                 Optional ByVal ErrCorrectionLevel As String = "L", _
                                 Optional ByVal ForeColor As OLE_COLOR = vbWhite, _
                                 Optional ByVal BackColor As OLE_COLOR = vbBlack, _
                                 Optional ByVal Margin As Long _
                                 ) As Picture
                                 
                                 
    Dim IIStream As IUnknown
    Dim StrOut As String
    Dim hImage As Long
    Dim hBmp As Long
    Dim sAPI As String

    'optional api by Goolge
    'sAPI = "http://chart.apis.google.com/chart?cht=qr&chs=" & Width & "x" & Height & "&chl=" & GetSafeURL(Unicode2UTF8(sText)) & "&choe=" & Encoding & "&chld=" & ErrCorrectionLevel

    '----------- By qrserver
    sAPI = "http://api.qrserver.com/v1/create-qr-code/?data=" & GetSafeURL(Unicode2UTF8(sText)) & "&size=" & Width & "x" & Height
    If Margin > 0 Then sAPI = sAPI & "&qzone=" & Margin
    If UCase(ErrCorrectionLevel) <> "L" Then sAPI = sAPI & "&ecc=" & ErrCorrectionLevel
    If ForeColor <> vbWhite Then sAPI = sAPI & "&color=" & LongToHtml(ForeColor)
    If BackColor <> vbBlack Then sAPI = sAPI & "&bgcolor=" & LongToHtml(BackColor)
    If UCase(Encoding) <> "UTF-8" Then sAPI = sAPI & "&charset-target=" & UCase(Encoding)
    '-----------

   
    If GetCode(sAPI, StrOut) Then

        Set IIStream = CreateStream(StrConv(StrOut, vbFromUnicode))
        
        If Not IIStream Is Nothing Then
            InitGDI
            If GdipLoadImageFromStream(IIStream, hImage) = 0 Then
                If GdipCreateHBITMAPFromBitmap(hImage, hBmp, 0&) = 0 Then
                    Set GetPictureQrCode = CreateBitmapPicture(hBmp, 0)
                End If
                GdipDisposeImage hImage
            End If
            TerminateGDI
        End If
        
    End If

End Function

Private Function GetSafeURL(ByVal sText As String) As String
    Dim aASC As String
    Dim bChr As Byte
    Dim sHex As String
    Dim i As Long
    
    For i = 1 To Len(sText)
        aASC = Mid$(sText, i, 1)
        bChr = Asc(aASC)
        If (bChr > 47 And bChr < 58) Or (bChr > 64 And bChr < 91) Or (bChr > 96 And bChr < 123) Then
            GetSafeURL = GetSafeURL & aASC
        Else
            sHex = Hex(bChr)
            If Len(sHex) = 1 Then
                GetSafeURL = GetSafeURL & "%0" & sHex
            Else
                GetSafeURL = GetSafeURL & "%" & sHex
            End If
        End If
    Next

End Function


Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long, Pic As PICTDESC, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function


'http://zxing.org/w/decode.jspx
Private Function UpLoadImage(InBytes() As Byte, OutStrDecode As String) As Boolean
    Dim hInternetOpen               As Long
    Dim hInternetConnect            As Long
    Dim hHttpOpenRequest            As Long
    Dim sHeader                     As String
    Dim sBuffer                     As String
    Dim sReadBuffer                 As String * 2048
    Dim lNumberOfBytesRead          As Long

    Const BOUNDARY = "---------------------------30862264243566"

    hInternetOpen = InternetOpen("IE", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    If hInternetOpen Then
        hInternetConnect = InternetConnect(hInternetOpen, "zxing.org", INTERNET_DEFAULT_HTTP_PORT, vbNullString, "HTTP/1.1", INTERNET_SERVICE_HTTP, 0, 0)
        
        If hInternetConnect Then
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", "/w/decode?full=false", "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_NO_AUTO_REDIRECT, 0)
        
            If hHttpOpenRequest Then
            
                sHeader = "Content-Type: multipart/form-data; boundary=" & BOUNDARY & vbCrLf

                sBuffer = "--" & BOUNDARY & vbCrLf
                sBuffer = sBuffer & "Content-Disposition: form-data; name=" & Chr$(34) & "f" & Chr$(34) & "; filename=" & Chr$(34) & "FileName" & Chr$(34) & vbCrLf
                sBuffer = sBuffer & "Content-Type: image/png" & vbCrLf & vbCrLf
                sBuffer = sBuffer & StrConv(InBytes, vbUnicode) & vbCrLf
            
                sBuffer = sBuffer & "--" & BOUNDARY & "--" & vbCrLf

                If HttpSendRequest(hHttpOpenRequest, sHeader, Len(sHeader), sBuffer, Len(sBuffer)) Then
            
                    sBuffer = String(200, 0)
            
                    If HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, Len(sBuffer), 0) Then
                        If (Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1) = HTTP_STATUS_OK) Then
                            OutStrDecode = vbNullString
                            Do
                                sBuffer = String(2048, 0)
                                Call InternetReadFile(hHttpOpenRequest, sBuffer, Len(sBuffer), lNumberOfBytesRead)
                                
                                OutStrDecode = OutStrDecode & Left$(sBuffer, lNumberOfBytesRead)
                                If lNumberOfBytesRead = 0 Then Exit Do
                            Loop
                            
                            OutStrDecode = UTF82Unicode(OutStrDecode)
                            UpLoadImage = True
                            
                        End If
                        
                    End If
            
                End If

                Call InternetCloseHandle(hHttpOpenRequest)
            End If
            
            Call InternetCloseHandle(hInternetConnect)
        End If
    
        Call InternetCloseHandle(hInternetOpen)
    End If
 
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdiPlusVersion
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub
 
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Public Function DecodeFromStream(InBytes() As Byte, OutStrDecode As String) As Boolean
    DecodeFromStream = UpLoadImage(InBytes, OutStrDecode)
End Function

Public Function DecodeFromPicture(ByVal oPicture As StdPicture, OutStrDecode As String) As Boolean
    Dim hImage As Long
    Dim ArrBytes() As Byte
    If Not oPicture Is Nothing Then
        InitGDI
        If GdipCreateBitmapFromHBITMAP(oPicture.Handle, 0, hImage) = 0 Then
            If SaveImageToStream(hImage, ArrBytes) Then
                DecodeFromPicture = UpLoadImage(ArrBytes, OutStrDecode)
            End If
            GdipDisposeImage hImage
        End If
        TerminateGDI
    End If
End Function

Public Function DecodeFromFile(ByVal sFile As String, OutStrDecode As String) As Boolean
    Dim hImage As Long
    Dim ArrBytes() As Byte
    
    InitGDI
    If GdipLoadImageFromFile(StrPtr(sFile), hImage) = 0 Then
        If SaveImageToStream(hImage, ArrBytes) Then
            DecodeFromFile = UpLoadImage(ArrBytes, OutStrDecode)
            GdipDisposeImage hImage
        End If
    End If
    TerminateGDI
    
End Function

Public Function DecodeFromUrl(ByVal sUrl As String, OutStrDecode As String) As Boolean
    DecodeFromUrl = GetCode("http://zxing.org/w/decode?full=false&u=" & GetSafeURL(sUrl), OutStrDecode)
End Function

Private Function GetCode(ByVal sUrl As String, ByRef StrOut As String) As Boolean

    Dim hOpen As Long, hFile As Long, sBuffer As String * 1024
    Dim lRet As Long
    
    StrOut = vbNullString
    hOpen = InternetOpen("IE", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
    If hOpen Then
    
        hFile = InternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_AUTO_REDIRECT, ByVal 0&)
    
        If hFile Then
            If HttpQueryInfo(hFile, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, 5&, 0) Then
                If (Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1) = HTTP_STATUS_OK) Then
                    Do
                        InternetReadFile hFile, sBuffer, 1024, lRet
                        StrOut = StrOut & Left$(sBuffer, lRet)
                        If lRet = 0 Then Exit Do
                    Loop
                    GetCode = True
                End If
            End If
            InternetCloseHandle hFile
        End If
    
        InternetCloseHandle hOpen
    End If
    
End Function

Private Function SaveImageToStream(ByVal hImage As Long, ByRef outStream() As Byte) As Boolean
    Dim IIStream    As IUnknown
    Dim tEncoder    As GUID
    Dim tParams     As EncoderParameters
    
    Erase outStream
       
    Set IIStream = CreateStream(outStream)

    If Not IIStream Is Nothing Then
        
        CLSIDFromString StrPtr(ImageCodecPNG), tEncoder

        tParams.Count = 1


        If GdipSaveImageToStream(hImage, IIStream, tEncoder, tParams) = 0& Then
            SaveImageToStream = ArrayFromStream(IIStream, outStream())
        End If
        
    End If
End Function


Private Function CreateStream(byteContent() As Byte, Optional byteOffset As Long = 0&) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_lngLowerBound As Long
    Dim o_lngByteCount  As Long
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If iparseIsArrayEmpty(VarPtrArray(byteContent)) = 0& Then ' create a growing stream as needed
         Call CreateStreamOnHGlobal(0, 1, CreateStream)
    Else                                        ' create a fixed stream
         o_lngByteCount = UBound(byteContent) - byteOffset + 1
         o_hMem = GlobalAlloc(&H2&, o_lngByteCount)
         If o_hMem <> 0 Then
             o_lpMem = GlobalLock(o_hMem)
             If o_lpMem <> 0 Then
                 CopyMemory ByVal o_lpMem, byteContent(byteOffset), o_lngByteCount
                 Call GlobalUnlock(o_hMem)
                 Call CreateStreamOnHGlobal(o_hMem, 1, CreateStream)
             End If
         End If
     End If
    
HandleError:
End Function

Private Function ArrayFromStream(Stream As IUnknown, arrayBytes() As Byte) As Boolean

    ' Purpose: Return the array contained in an IUnknown interface
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If Not Stream Is Nothing Then
    
        If GetHGlobalFromStream(ByVal ObjPtr(Stream), o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    ArrayFromStream = True
                End If
            End If
        End If

    End If
    
End Function


Private Function iparseIsArrayEmpty(FarPointer As Long) As Long
  ' test to see if an array has been initialized
  CopyMemory iparseIsArrayEmpty, ByVal FarPointer, 4&
End Function


Private Function LongToHtml(ByVal lColor As OLE_COLOR) As String
    Dim cl(3)         As Byte
    OleTranslateColor lColor, 0, VarPtr(cl(0))
    LongToHtml = Format(Hex(cl(0)), "00") & Format(Hex(cl(1)), "00") & Format(Hex(cl(2)), "00")
End Function

Private Function UTF82Unicode(ByVal sUTF8 As String) As String

    Dim UTF8Size As Long
    Dim BufferSize As Long
    Dim BufferUNI As String
    Dim LenUNI As Long
    Dim bUTF8() As Byte
    
    If LenB(sUTF8) = 0 Then Exit Function
    
    bUTF8 = StrConv(sUTF8, vbFromUnicode)
    UTF8Size = UBound(bUTF8) + 1
    
    BufferSize = UTF8Size * 2
    BufferUNI = String$(BufferSize, vbNullChar)
    
    LenUNI = MultiByteToWideChar(CP_UTF8, 0, bUTF8(0), UTF8Size, StrPtr(BufferUNI), BufferSize)
    
    If LenUNI Then
        UTF82Unicode = Left$(BufferUNI, LenUNI)
    End If

End Function


Private Function Unicode2UTF8(ByVal strUnicode As String) As String

    Dim LenUNI As Long
    Dim BufferSize As Long
    Dim LenUTF8 As Long
    Dim bUTF8() As Byte
    
    LenUNI = Len(strUnicode)
    
    If LenUNI = 0 Then Exit Function
    
    BufferSize = LenUNI * 3 + 1
    ReDim bUTF8(BufferSize - 1)
    
    LenUTF8 = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), LenUNI, bUTF8(0), BufferSize, vbNullString, 0)
    
    If LenUTF8 Then
        ReDim Preserve bUTF8(LenUTF8 - 1)
        Unicode2UTF8 = StrConv(bUTF8, vbUnicode)
    End If

End Function


Public Function CodigoQRRequest(rCodigo As Long, rWidth As Long, rHeight As Long) As Picture
    Dim url As String
    Dim CadenaRequest As String, z As Double
    Dim tempFactura
    Dim rVersion As String, rFecha As String, rCuit As String, rPtoVta As String, rTipoCmp As String, rNroCmp As String, rImporte As String, rMoneda As String, rCtz As String, rTipoDocRec As String, rNroDocRec As String, rTipoCodAut As String, rCodAut As String
    tempFactura = obtenerDeSQL("select '1' as Version,fecha,'' as cuitempresa,puntoventa,tipodoc,nrofactura,total,moneda,cotizacion,'' as tipodocrec,cuit as nrodocrec,'E' as tipocae,cae from facturaventa where codigo=" & rCodigo)
    rVersion = tempFactura(0)
    rFecha = Format(tempFactura(1), "yyyy-mm-dd")
    rCuit = Replace(sSinNull(Trim(obtenerDeSQL("select CUITEMPRESA from datosempresa where idempresa=" & gEMPR_idEmpresa))), "-", "")
    rPtoVta = Format(tempFactura(3), "00000")
    rTipoCmp = sSinNull(obtenerDeSQL("SELECT CODFACTURA FROM DOCUMENTOSCAE WHERE TIPO=" & ssTexto(tempFactura(4)) & " AND PUNTOVENTA=" & ssTexto(tempFactura(3))))
    rTipoCmp = Format(rTipoCmp, "000") '3 digitos
    rNroCmp = Format(sSinNull(tempFactura(5)), "00000000")
    rImporte = Format(Replace(s2n(tempFactura(6), 2, True), ",", ""), "000000000000000")
    rMoneda = Trim(ObtenerDescripcion("Monedas", nSinNull(tempFactura(7))))
    rMoneda = UCase(CORTO(rMoneda, 0, Len(rMoneda) - 3))
    z = s2n(tempFactura(8), 4, True)
    If z = 0 Then z = 1
    rCtz = Format(Replace(s2n(z, 6, True), ",", ""), "0000000000000000000")
    rNroDocRec = Replace(sSinNull(tempFactura(10)), "-", "")
    If Len(s2n(rNroDocRec)) <= 8 Then
        rTipoDocRec = "96"
    ElseIf Len(s2n(rNroDocRec)) = 11 Then
        rTipoDocRec = "80"
    Else
        rTipoDocRec = "94"
    End If
    rTipoCodAut = tempFactura(11)
    rCodAut = tempFactura(12)
    
    'testing
    'url = "https://titotest.urquizamotos.com.ar/api/orden/" & nroOrden & "?key=RXC7Q31DDS"
    'produccion
    CadenaRequest = rVersion & rFecha & rCuit & rPtoVta & rTipoCmp & rNroCmp & rImporte & rMoneda & rCtz & rTipoDocRec & rNroDocRec & rTipoCodAut & rCodAut
    CadenaRequest = "{" & Chr(34) & "ver" & Chr(34) & ":" & rVersion & "," & Chr(34) & "fecha" & Chr(34) & ":" & Chr(34) & "" & rFecha & "" & Chr(34) & "," & Chr(34) & "cuit" & Chr(34) & ":" & rCuit & "," & Chr(34) & "ptoVta" & Chr(34) & ": " & rPtoVta & "," & Chr(34) & "tipoCmp" & Chr(34) & ": " & rTipoCmp & "," & Chr(34) & "nroCmp" & Chr(34) & ":" & rNroCmp & "," & Chr(34) & "importe" & Chr(34) & ":" & rImporte & "," & Chr(34) & "moneda" & Chr(34) & ":" & Chr(34) & rMoneda & Chr(34) & "," & Chr(34) & "ctz" & Chr(34) & ":" & rCtz & "," & Chr(34) & "tipoDocRec" & Chr(34) & ":" & rTipoDocRec & "," & Chr(34) & "nroDocRec" & Chr(34) & ":" & rNroDocRec & "," & Chr(34) & "tipoCodAut" & Chr(34) & ":" & Chr(34) & rTipoCodAut & Chr(34) & "," & Chr(34) & "codAut" & Chr(34) & ":" & rCodAut & "}"
    'ejemplo de json sin codificar
    '{"ver":1,"fecha":"2020-10-13","cuit":30000000007,"ptoVta":10,"tipoCmp":1,"nroCmp":94,"importe":12100,"moneda":"DOL","ctz":65,"tipoDocRec":80,"nroDocRec":20000000001,"tipoCodAut":"E","codAut":70417054367476}
    
    'CODIFICAR
    'CadenaRequest = EncryptString(64, CadenaRequest, 2)
    CadenaRequest = Encode64(CadenaRequest)
    'CadenaRequest = "eyJ2ZXIiOjEsImZlY2hhIjoiMjAyMC0wOC0wMyIsImN1aXQiOjIwMjAyNTEyOTE4LCJwdG9WdGEiOiAwMDAwMywidGlwb0NtcCI6IDAwMSwibnJvQ21wIjowMDAwMDg3NywiaW1wb3J0ZSI6MDAwMDAwMDAwMjU2NTAxLCJtb25lZGEiOiJQRVMiLCJjdHoiOjAwMDAwMDAwMDAwMDEwMDAwMDAsInRpcG9Eb2NSZWMiOjgwLCJucm9Eb2NSZWMiOjMwNTU3OTkwOTgwLCJ0aXBvQ29kQXV0IjoiRSIsImNvZEF1dCI6NjkxNjg3Mzk3NjMxODZ9"
    
    
    'url = "https://www.afip.gob.ar/fe/qr/" & CadenaRequest & "p=eyJ2ZXIiOjEsImZlY2hhIjoiMjAyMC0xMC0xMyIsImN1aXQiOjMwMDAwMDAwMDA3LCJwdG9WdGEiOjEwLCJ0aXBvQ21wIjoxLCJucm9DbXAiOjk0LCJpbXBvcnRlIjoxMjEwMCwibW9uZWRhIjoiRE9MIiwiY3R6Ijo2NSwidGlwb0RvY1JlYyI6ODAsIm5yb0RvY1JlYyI6MjAwMDAwMDAwMDEsInRpcG9Db2RBdXQiOiJFIiwiY29kQXV0Ijo3MDQxNzA1NDM2NzQ3Nn0="
    url = "https://www.afip.gob.ar/fe/qr/?p=" & CadenaRequest

    Set CodigoQRRequest = GetPictureQrCode(url, rWidth, rHeight)
    
 
End Function


Public Function EncryptString( _
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKeyX As String
    Dim temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
      
    '//Get UserKey characters
    n = Len(UserKey)
    ReDim UserKeyASCIIS(1 To n)
    For i = 1 To n
        UserKeyASCIIS(i) = Asc(Mid$(UserKey, i, 1))
    Next
          
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(Mid$(Text, i, 1))
    Next
      
    '//Encryption/Decryption
    If Action = 1 Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           temp = TextASCIIS(i) + UserKeyASCIIS(j)
           If temp > 255 Then
              temp = temp - 255
           End If
           rtn = rtn + Chr$(temp)
       Next
    ElseIf Action = 2 Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           temp = TextASCIIS(i) - UserKeyASCIIS(j)
           If temp < 0 Then
              temp = temp + 255
           End If
           rtn = rtn + Chr$(temp)
       Next
    End If
      
    '//Return
    EncryptString = rtn
End Function

Public Function Encode64(Texto As String) As String
Dim Data As New ChilkatBinData
Dim nCorrecto As Long, sCodificado As String
nCorrecto = Data.AppendString(Texto, "utf-8")
'codifica
sCodificado = Data.GetEncoded("base64")
Encode64 = sCodificado
'decodifica
'nCorrecto = data.AppendEncoded(sCodificado, "base64")
Dim sDecodificado As String
'sDecodificado = data.GetString("utf-8")


End Function