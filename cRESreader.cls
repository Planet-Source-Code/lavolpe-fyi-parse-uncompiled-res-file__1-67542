VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRESreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' http://www.filespecs.com/localview.jsp?link=15
' ^^ a site I re-visit from time to time

' TO ALL WHO PLAY WITH THIS.....

' I would certainly appreciate feedback whether this is compatible with VB5 or not.

' Now that being said. I have not found anything out there that can read & extract
' a VB resource file that isn't compiled into an application.  I'm sure now that
' I have said that someone will direct me to a link where I can download the
' code :)

' Because I couldn't find anything, I did this the hard way; basically reverse
' engineering the Res file and have only deciphered a few resource sections:
' Bitmap, Icon, Cursor, & "Custom" resource sections.  Custom is anything that
' you can add to VB's Res file GUI.

' Again, I did this the hard way, so I am not 100% sure it will work for all
' icons, cursors, and bitmaps loaded into an existing Res file, but it has
' worked, so far, on the several Res files I have tested it with.


Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long

Private Enum eResTypes    ' Values < 256 are reserved for system use.
    RT_CURSOR = 1
    RT_BITMAP = 2
    RT_ICON = 3
    RT_MENU = 4
    RT_DIALOG = 5
    RT_STRING = 6
    RT_FONTDIR = 7
    RT_FONT = 8
    RT_ACCELERATOR = 9
    RT_RCDATA = 10
    RT_MESSAGETABLE = 11
    RT_GROUP_CURSOR = 12
    RT_GROUP_ICON = 14
    RT_VERSION = 16
    RT_DLGINCLUDE = 17
    RT_PLUGPLAY = 19
    RT_VXD = 20
    RT_ANICURSOR = 21
    RT_ANIICON = 22
    RT_HTML = 23
    RT_MANIFEST = 24
End Enum

Private Type uResGroupInfo
    ID As String            ' ID or Name (i.e., 101, "MYGIF")
    DataOffset As Long      ' Where in file the data begins for this item
    DataLength As Long      ' The length of the data
End Type
Private Type uResGroup
    Name As String          ' ID or Name (i.e., RT_ICON, "CUSTOM")
    Data() As uResGroupInfo ' collection of child resources
    Attr As Byte            ' resource type
End Type

' PROTOTYPE RESOURCEHEADER
'    DataSize    AS DWORD ' bytes of resource data following header, not counting any padding used for alignment
'    HeaderSize  AS DWORD
'    ResType(0)  AS WORD  ' variable length: numeric id or UnicodeZ string data
'    ResName(0)  AS WORD  ' variable length: numeric id or UnicodeZ string data
'    Padding     AS ????  ' 0-1 WORDs of padding to DWORD-align next item
'    DataVersion AS DWORD
'    MemoryFlags AS WORD
'    LanguageId  AS WORD
'    Version     AS DWORD
'    Characteristics AS DWORD
' END PROTOTYPE

Private m_ResHeader As String * 32
Private m_Resources() As uResGroup
Private m_FileName As String
Private m_ManagedGroups As Long ' number of groups we will expose

Private Function AddResGroup(wID As Variant) As Long

    ' Function creates a group UDT to store only the following
    ' resource types where the Attr < 255
    ' The rest of the types are left in for reference only
    
    ' BTW: Where Attr=255, these aren't available in a standard VB Res file
    ' However, they could be in other RES files, and you are welcome to
    ' remove the 255 Attr flag so they can be enumerated. However, you
    ' should also provide a way of exposing/extracting the resource in
    ' the ExtractResource routine.

    Dim resName As String, Index As Long, Attr As Byte
    If VarType(wID) = vbString Then
        ' non-standard resources (i.e., "CUSTOM")
        resName = wID
    Else
        Attr = CByte(wID)
        Select Case wID
        Case RT_CURSOR: resName = "Application Cursor"
            ' these are the actual cursor image bytes, but are formatted
            ' slightly different than a cursor file
        Case RT_BITMAP: resName = "Bitmap"
            ' these are the actual bitmap bytes, less the BitmapFileHeader
        Case RT_ICON: resName = "Application Icon"
            ' these are the actual icon bytes when combined with the
            ' RT_GROUP_ICON form a complete icon
        Case RT_GROUP_CURSOR: resName = "Cursor"
            ' Basically the CUSORDIR & CURSORDIRENTRY structures, but different
        Case RT_GROUP_ICON: resName = "Icon"
            ' This is an ICONDIR & ICONDIRENTRY structures
            
        Case RT_MENU: resName = "Menu": Attr = 255
        Case RT_DIALOG: resName = "Dialog Box": Attr = 255
        Case RT_STRING: resName = "String Table": Attr = 255
            ' VB has string tables; haven't had time to reverse engineer
            ' how they are stored in VB's Res file
        Case RT_FONTDIR: resName = "Font Directory": Attr = 255
        Case RT_FONT: resName = "Font": Attr = 255
        Case RT_ACCELERATOR: resName = "Accelerator": Attr = 255
        Case RT_RCDATA: resName = "Raw Data": Attr = 255
        Case RT_MESSAGETABLE: resName = "Message Table": Attr = 255
        Case RT_VERSION: resName = "Version information": Attr = 255
        Case RT_DLGINCLUDE: resName = ".RC header filename": Attr = 255
        Case RT_PLUGPLAY: resName = "Plug and Play resource": Attr = 255
        Case RT_VXD: resName = "VXD device driver": Attr = 255
        Case RT_ANICURSOR: resName = "Cursor - Animated": Attr = 255
        Case RT_ANIICON: resName = "Icon - Animated": Attr = 255
        Case RT_HTML: resName = "HTML": Attr = 255
        Case RT_MANIFEST: resName = "XML Manifest for Windows XP": Attr = 255
        End Select
    End If
    
    If Attr = 255 Then
        AddResGroup = -1    ' we won't cache this data
    Else
        If IsArrayEmpty(Not m_Resources) Then
            ReDim m_Resources(0 To 0)
        Else
            For Index = 0 To UBound(m_Resources)
                If m_Resources(Index).Name = resName Then Exit For
            Next
            If Index > UBound(m_Resources) Then
                ReDim Preserve m_Resources(0 To Index)
            Else
                AddResGroup = Index
                ReDim Preserve m_Resources(Index).Data(0 To UBound(m_Resources(Index).Data) + 1)
                Exit Function
            End If
        End If
        m_Resources(Index).Name = resName
        m_Resources(Index).Attr = Attr
        ReDim m_Resources(Index).Data(0 To 0)
        AddResGroup = Index
    End If
    
End Function

Private Function AddResItem(gID As Long, wID As Variant) As Long
    
    ' These are the child resources of the resource groups
    
    If gID = -1 Then
        AddResItem = -1 ' we don't want data for groups we aren't caching
    Else
        Dim Index As Long
        Index = UBound(m_Resources(gID).Data)
        m_Resources(gID).Data(Index).ID = wID
        AddResItem = Index
    End If
End Function

Private Function RetrieveFileData(ByVal nFile As Integer) As Boolean

    Dim dcbDataSize   As Long   ' size of resource data
    Dim dcbHeaderSize As Long   ' size of header record
    Dim FileOffset   As Long    ' offset of header record within file
    ' Long, Integer & String variables for reading the file
    Dim resDataL As Long, resDataI As Integer
    Dim resDataS As String, resChar As Integer
    
    Dim resGrpID As Long        ' Group collection item
    Dim resItemID As Long       ' Child Resource collection item

    FileOffset = Seek(nFile)    ' Begin at this location

    Get #nFile, , dcbDataSize   ' get the data size of the resource (byte size)
    If Err Then Exit Function

    Get #nFile, , dcbHeaderSize ' the header of the Prototype structure
    ' Note: the header and must be word aligned at a specific point which
    ' means we may have to move the file pointer at times. Additionally,
    ' the dcbData must also be word aligned, which means the file pointer]
    ' may be moved yet again when we are done reading this resource
    If Err Then Exit Function
    
    ' minimum header size is 32, but can exceed 32 when string names are used for resources
    If dcbHeaderSize < Len(m_ResHeader) Then Exit Function
    
    Get #nFile, , resDataI      ' get resource ID flag
    If resDataI = &HFFFF Then   ' if -1 then numerical ID
        Get #nFile, , resDataI  ' and next 2bytes is the ID
        resGrpID = AddResGroup(resDataI)
    Else
        ' we have a unicode, null-terminated string...
        Do Until resDataI = 0       ' count characters, unicode=2bytes per char
            resChar = resChar + 1
            Get #nFile, , resDataI
        Loop
        ' now go back to the beginning to cache it; guess we could have concatenated string
        Seek nFile, Seek(nFile) - resChar * 2 - 2
        ' build string of appropriate length
        resDataS = String$(resChar * 2, Chr$(0))
        Get #nFile, , resDataS  ' read the string
        Get #nFile, , resDataI  ' and dispose of the null terminator too
        resChar = 0             ' reset for next section if needed
        ' add this resource group to our collection
        resGrpID = AddResGroup(StrConv(resDataS, vbFromUnicode))
    End If

    Get #nFile, , resDataI      ' get ID type of child resource
    If resDataI = &HFFFF Then   ' numerical ID
        Get #nFile, , resDataI
        resItemID = AddResItem(resGrpID, resDataI)
    Else
        ' same exact comments as above
        Do Until resDataI = 0
            resChar = resChar + 1
            Get #nFile, , resDataI
        Loop
        Seek nFile, Seek(nFile) - resChar * 2 - 2
        resDataS = String$(resChar * 2, Chr$(0))
        Get #nFile, , resDataS
        resItemID = AddResItem(resGrpID, StrConv(resDataS, vbFromUnicode))
    End If
    
    If Not resItemID = -1 Then  ' else a resource we aren't caching
        ' but the following is important when we are caching the resource
        ' As mentioned above, the resource data must start on a DWORD
        ' aligned boundary; here we are ensuring that...
        resDataL = (FileOffset + dcbHeaderSize + 2) And Not 3
        m_Resources(resGrpID).Data(resItemID).DataOffset = resDataL + 1
        m_Resources(resGrpID).Data(resItemID).DataLength = dcbDataSize
    End If
    
    ' There is more information that follows; specifically 16 bytes of info
    ' But for our purposes, we won't be needing it.
    
    ' Now, the next record must start on a DWORD boundary to. Ensure it...
    FileOffset = ((FileOffset + dcbHeaderSize + dcbDataSize) + 2) And Not 3
    
    ' are we at the end?
    If Not FileOffset >= LOF(nFile) Then
        Seek nFile, FileOffset + 1  ' set position for next resource
        RetrieveFileData = True
    End If

End Function

Public Function ScanResources(RESfilename As String) As Boolean

    Dim sFile   As String       ' file name
    Dim nFile   As Integer      ' file number
    Dim sBuffer As String * 32  ' string input

    Dim i As Integer, v As Byte
    
    ' 0,0,0,0,&H20,0,0,0,&HFF,&HFF,0,0,&HFF,&HFF,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
    For i = 1 To 32 ' build the resource file header (1st 32 bytes of a RES file)
        Select Case i
        Case 5:
            v = &H20
        Case 9, 10, 13, 14
            v = &HFF
        Case Else
            v = &H0
        End Select
        Mid$(m_ResHeader, i, 1) = Chr$(v)
    Next

    ' reset & clear if class is being reused
    m_FileName = vbNullString
    Erase m_Resources
    m_ManagedGroups = 0
    
    sFile = Trim$(RESfilename)
    ' Ensure if file exists; no extension checking being made

    On Error Resume Next    ' attempt to open file
    If Len(Dir$(sFile, vbHidden Or vbReadOnly Or vbReadOnly Or vbSystem)) = 0 Then
        Exit Function
    End If
    
    m_FileName = sFile        ' cache file name for later if needed
    nFile = FreeFile
    Open sFile For Binary Access Read As #nFile
    If Err Then Exit Function

    If LOF(nFile) < Len(m_ResHeader) Then
        Close nFile
        Exit Function
    End If

    Get #nFile, 1, sBuffer
    If sBuffer <> m_ResHeader Then
        ' first 32bytes better match our magic string above
        Close nFile
        Exit Function
    End If

    ' now we just read til there is nothing more to read
    Do Until RetrieveFileData(nFile) = False
    
    Loop

    Close nFile
    
    ' here we are going to place our resources we won't expose on the end
    ' of the collection; just to keep them organized and out of the way...
    
    If IsArrayEmpty(Not m_Resources) Then
        m_FileName = vbNullString
        
    Else
    
        Dim tmpResGroup As uResGroup
        m_ManagedGroups = UBound(m_Resources)
        For i = m_ManagedGroups To 0 Step -1
            Select Case m_Resources(i).Attr
            Case RT_ICON, RT_CURSOR, RT_FONT
                If Not i = m_ManagedGroups Then
                    tmpResGroup = m_Resources(i)
                    m_Resources(i) = m_Resources(m_ManagedGroups)
                    m_Resources(m_ManagedGroups) = tmpResGroup
                End If
                m_ManagedGroups = m_ManagedGroups - 1
            Case Else
            End Select
        Next
        m_ManagedGroups = m_ManagedGroups + 1
        If m_ManagedGroups = 0 Then
            m_FileName = vbNullString
        Else
            ScanResources = True
        End If
    End If
            
End Function

Public Function ResourceSection(ByVal Index As Long) As String
    
    ' Function returns the group name of the index being passed (i.e., "CUSTOM")
    ' Index is the zero-bound index of the resource groups, continue calling
    ' routine until ResourceSection returns vbNullString
    
    If Index < m_ManagedGroups Then ResourceSection = m_Resources(Index).Name

End Function

Public Function ResourceID(ResSection As Variant, ByVal Index As Long) As String
        
    ' Function returns the child resource id or name (i.e., 101, "MYGIF")
    ' ResSection can be vbResBitmap , vbResIcon, vbResCursor, Custom
    ' can also be just Bitmap, Icon, Cursor
    ' Index is the zero-bound index of the resource, continue calling
    ' routine until ResourceID returns vbNullString
    
    Dim grpID As Long
    ResolveResNames ResSection, 0, grpID, 0
    If grpID < m_ManagedGroups Then
        If Index <= UBound(m_Resources(grpID).Data) Then
            ResourceID = m_Resources(grpID).Data(Index).ID
        End If
    End If

End Function

Public Function ExtractResPicture(ResSection As Variant, ResID As Variant) As StdPicture

    ' ResSection can be vbResBitmap , vbResIcon, vbResCursor, Custom
    ' can also be just Bitmap, Icon, Cursor. But use a little caution,
    ' VB will allow you to have a Custom resource group named "BITMAP"
    ' but this is different than the standard resource Bitmap. Therefore
    ' use VB constants when appropriate
    
    Dim grpID As Long, Index As Long, aDummy() As Byte
    ResolveResNames ResSection, ResID, grpID, Index
    If grpID < m_ManagedGroups Then
        If Index <= UBound(m_Resources(grpID).Data) Then
            Select Case m_Resources(grpID).Attr
                Case RT_GROUP_ICON, RT_GROUP_CURSOR, RT_BITMAP
                    Call ExtractResource(grpID, Index, ExtractResPicture, aDummy(), True)
                Case Else
            End Select
        End If
    End If
    
End Function

Public Function ExtractResStream(ResSection As Variant, ResID As Variant, outStream() As Byte) As Boolean

    ' ResSection can be vbResBitmap , vbResIcon, vbResCursor, Custom
    ' can also be just Bitmap, Icon, Cursor. But use a little caution,
    ' VB will allow you to have a Custom resource group named "BITMAP"
    ' but this is different than the standard resource Bitmap. Therefore
    ' use VB constants when appropriate
    
    Dim grpID As Long, Index As Long
    ResolveResNames ResSection, ResID, grpID, Index
    If grpID < m_ManagedGroups Then
        If Index <= UBound(m_Resources(grpID).Data) Then
            ExtractResStream = ExtractResource(grpID, Index, Nothing, outStream(), False)
        End If
    End If

End Function

Private Sub ResolveResNames(grpCommon As Variant, itmCommon As Variant, grpID As Long, ItemID As Long)

    
    If VarType(grpCommon) = vbString Then
        ' compare string IDs/Names
        For grpID = 0 To m_ManagedGroups - 1
            If StrComp(m_Resources(grpID).Name, grpCommon, vbTextCompare) = 0 Then Exit For
        Next
    Else
        ' compare to VB constants: vbResBitmap , vbResIcon, vbResCursor
        Dim bCriteria As Byte
        Select Case grpCommon
        Case vbResBitmap: bCriteria = RT_BITMAP
        Case vbResIcon: bCriteria = RT_GROUP_ICON
        Case vbResCursor: bCriteria = RT_GROUP_CURSOR
        Case Else
        End Select
        If bCriteria = 0 Then
            grpID = m_ManagedGroups
        Else
            For grpID = 0 To m_ManagedGroups - 1
                If m_Resources(grpID).Attr = bCriteria Then Exit For
            Next
        End If
    End If
    
    If grpID < m_ManagedGroups Then
        ' look for the specific item
        For ItemID = 0 To UBound(m_Resources(grpID).Data)
            If StrComp(m_Resources(grpID).Data(ItemID).ID, itmCommon, vbTextCompare) = 0 Then Exit For
        Next
    End If
        
End Sub

Public Function ExtractResource(GroupID As Long, ItemIndex As Long, stdPic As StdPicture, outStream() As Byte, bAsStdPic As Boolean) As Boolean
    
    Dim fNR As Integer
    Dim iBitCount As Integer
    Dim lSize As Long, lSize2 As Long
    
    fNR = FreeFile()
    On Error Resume Next
    Open m_FileName For Binary Access Read As #fNR
    If Err Then
        Close #fNR
        Err.Clear
        Exit Function
    End If
    
    With m_Resources(GroupID).Data(ItemIndex)
        
        ReDim outStream(0 To .DataLength - 1)
        Get #fNR, .DataOffset, outStream()
        Select Case m_Resources(GroupID).Attr
        
        Case RT_BITMAP: ' bitmap
            ' The vbResBitmap section contains everything one would need to
            ' create a DIB bitmap. I am just supplyng the BITMAPFILEHEADER
            ' so that the stream can be read by APIs or saved to file
            
            CopyMemory lSize, outStream(20), 4&    ' size Image
            CopyMemory lSize2, outStream(0), 4&    ' BMPinfoHeader Size
            ReDim Preserve outStream(0 To .DataLength + 13)
            CopyMemory outStream(14), outStream(0), .DataLength
            CopyMemory outStream(0), &H4D42, 2& ' bmp magic number
            CopyMemory outStream(2), CLng(.DataLength + 14), 4& ' overall size of image
            ' ^^ 54 = 14 byte bmp header + 40 for the tBMPI structure
            CopyMemory outStream(6), 0&, 4&
            CopyMemory iBitCount, outStream(28), 2& ' bit count
            CopyMemory lSize, outStream(18), 4&
            CopyMemory lSize2, outStream(22), 4&
            lSize = .DataLength - (ByteAlignOnWord(iBitCount, lSize) * lSize2) + 14
            CopyMemory outStream(10), lSize, 4& ' image offset from beginning of file
            
            If bAsStdPic Then Set stdPic = CreateStdPicFromArray(outStream(), 0, .DataLength + 14)
        
        Case RT_GROUP_ICON: ' icon; send to routine to gather multi-icons if needed
            GatherGroupItems RT_ICON, fNR, .DataOffset, outStream()
            If bAsStdPic Then Set stdPic = CreateStdPicFromArray(outStream(), 0, UBound(outStream) + 1)
            
        Case RT_GROUP_CURSOR: ' icon; send to routine to gather multi-cursors if needed
            GatherGroupItems RT_CURSOR, fNR, .DataOffset, outStream()
            If bAsStdPic Then Set stdPic = CreateStdPicFromArray(outStream(), 0, UBound(outStream) + 1)
            
        Case Else
            ' add code for any other resources you care to extract
            
        End Select
    End With
    Close #fNR
    ExtractResource = True

End Function

Private Sub GatherGroupItems(whichGroup As Long, FileNum As Integer, fOffset As Long, rtnArray() As Byte)

    Dim itemData() As Long
    Dim fileDat() As Byte
    Dim arraySize As Long
    Dim nrItems As Long
    Dim iValue As Integer, lValue As Long
    Dim rtIcon As Long
    
    ' for each icon and/or cursor, there exists a RT_ICON / RT_CURSOR
    ' section that contains bascially the ICONDIR/CUSORDIR information.
    ' Find that directory
    For rtIcon = 0 To UBound(m_Resources)
        If m_Resources(rtIcon).Attr = whichGroup Then Exit For
    Next
    
    arraySize = 6   ' icons/cursors both begin with a 6 byte header
    Get #FileNum, fOffset + 4, iValue ' retrieve nr of icons/cursors
    arraySize = arraySize + iValue * 16 ' add 16byte header for each one
    
    ' calculate the size of the array we will need
    ReDim itemData(0 To iValue - 1)
    For nrItems = 0 To iValue - 1
    
        ' retrieve the icon/cursor ID for each one that belongs to
        ' the selected Directory. They will always be numeric
        ' and can be found as the last 2 bytes within that 16 byte header
        Get #FileNum, fOffset + 6 + (nrItems * 14) + 12, iValue
        
        ' is the ID one of the ones we are looking for?
        For lValue = 0 To UBound(m_Resources(rtIcon).Data)
            If m_Resources(rtIcon).Data(lValue).ID = iValue Then
                ' yep, cache its RT_ICON/CURSOR child reference
                itemData(nrItems) = lValue
                Exit For
            End If
        Next
        ' go on to the next icon/cursor as needed
    Next
    
    ' ok, now add up the amount of bytes each one will use
    For nrItems = 0 To UBound(itemData)
        arraySize = arraySize + m_Resources(rtIcon).Data(itemData(nrItems)).DataLength
    Next
    
    ' here we are building the 6 byte header
    iValue = UBound(itemData) + 1
    ReDim rtnArray(0 To arraySize - 1)
    If whichGroup = RT_CURSOR Then
        CopyMemory rtnArray(2), 2, 2&   ' cursor
    Else
        CopyMemory rtnArray(2), 1, 2&   ' icon
    End If
    CopyMemory rtnArray(4), iValue, 2&  ' nr of icons/cursors
    lValue = iValue * 16 + 6            ' offset for first item
    
    For nrItems = 0 To iValue - 1
        
        ReDim fileDat(1 To 14) ' directories are 16 bytes, but in resources they are 14
        Get #FileNum, fOffset + 6 + (nrItems * 14), fileDat()   ' get next icondirectory entry
        If rtnArray(2) = 1 Then 'icon
            fileDat(2) = fileDat(2) \ 2 ' height is expressed * 2
            iValue = 0                  ' file offset
        Else                    ' cursor
            ' not quite as expected. Whereas the std struture uses bytes for these
            ' properties, the res structure uses Integers; shift array values
            fileDat(2) = fileDat(3) ' height
            fileDat(3) = fileDat(5) ' color count if used
            fileDat(4) = fileDat(6) ' reserved
            
            ' the 1st 4 bytes of the res data for cursors will be the X,Y hotspots.
            ' again unique. Get those values & put them in their place within our array
            Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset, iValue
            CopyMemory fileDat(5), iValue, 2&   ' X hotspot
            Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset + 2, iValue
            CopyMemory fileDat(7), iValue, 2&   ' Y hotspot
            ' the data length is really 4 more than needed now; we offset
            iValue = 4
            ' copy the new data length into the cursor directory
            CopyMemory fileDat(9), (m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue), 4&
        End If
        
        ' copy the icon/cursor directory to our array
        CopyMemory rtnArray(6 + (nrItems * 16)), fileDat(1), 14&
        ' add the image offset value to the directory
        CopyMemory rtnArray(6 + (nrItems * 16) + 12), lValue, 4&
        
        ' now copy the item data
        ReDim fileDat(1 To m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue)
        Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset + iValue, fileDat()
        CopyMemory rtnArray(lValue), fileDat(1), m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue
        
        ' calculate next position in our array for the next icon/cursor
        lValue = lValue + m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue
        
    Next
    Erase fileDat

End Sub

Private Function ByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  IsArrayEmpty = (lArrayPointer = -1&)
End Function

Public Function CreateStdPicFromArray(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty when calling class' LoadStream was called
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), CreateStdPicFromArray)
            End If
        End If
    End If

End Function

