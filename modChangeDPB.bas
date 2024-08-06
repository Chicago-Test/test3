Attribute VB_Name = "modChangeDPB"
Option Explicit

Private Function load_binary_file(fileFullPath As String) As Byte()

    Dim nFileLen As Long
    nFileLen = FileLen(fileFullPath)
    If nFileLen = 0 Then
        End
    End If

    'Get file number
    Dim iFile As Integer
    iFile = FreeFile

    Open fileFullPath For Binary As #iFile

    Dim bData() As Byte
    ReDim bData(0 To nFileLen - 1)

    Get #iFile, , bData
    Close #iFile
    load_binary_file = bData

End Function

Private Sub create_binary_file(fileFullPath As String, bData() As Byte)
    Dim iFile As Integer
    iFile = FreeFile
    Open fileFullPath For Binary As #iFile
    Put #iFile, , bData
    Close #iFile
End Sub

Private Function get_startIDX_of_DPB_from_vbaProject(bData() As Byte) As Long
    'https://www.vbforums.com/showthread.php?797047-Search-byte-pattern-in-byte-array
    Dim bytesToFind(0 To 5) As Byte
    bytesToFind(0) = &HD
    bytesToFind(1) = &HA
    bytesToFind(2) = Asc("D") '0x44
    bytesToFind(3) = Asc("P") '0x50
    bytesToFind(4) = Asc("B") '0x42
    bytesToFind(5) = Asc("=") '0x3D
    
    
    Dim output() As Byte
    Dim x As Long
    For x = 0 To UBound(bData) - UBound(bytesToFind) - 1
        If bData(x) = bytesToFind(0) And bData(x + 1) = bytesToFind(1) And bData(x + 2) = bytesToFind(2) And bData(x + 3) = bytesToFind(3) _
           And bData(x + 4) = bytesToFind(4) And bData(x + 5) = bytesToFind(5) Then
            
            'Debug.Print x & ":" & "0x0" & Hex(bData(x)) & " 0x0" & Hex(bData(x + 1)) & " " & Chr(bData(x + 2)) & Chr(bData(x + 3)) & Chr(bData(x + 4))
            GoTo continue1
        
        End If
    Next
continue1:
    get_startIDX_of_DPB_from_vbaProject = x + 7
End Function

Private Sub change_VBA_password()
    Dim in_fileFullPath As String
    Dim out_fileFullPath As String
    in_fileFullPath = "C:\Users\Administrator\source\repos\CompoundFormat1\vbaProject3.bin"
    out_fileFullPath = "C:\Users\Administrator\source\repos\CompoundFormat1\vbaProject3a.bin"

    Call change_VBA_password_of_VBAProject(in_fileFullPath, out_fileFullPath)
End Sub

Public Function change_VBA_password_of_VBAProject(IN_VBAProject_fullPath As String, OUT_VBAProject_fullPath As String) As Integer
    ' 0:OK, -1: Error
    On Error GoTo ErrHandler
    change_VBA_password_of_VBAProject = -1

    Dim i As Long
    Dim str As String

    Dim bData() As Byte
    bData = load_binary_file(IN_VBAProject_fullPath)

    Dim passA_0 As String, passA_1 As String, passA_2 As String, passA_3 As String
    ' password 'a'
    passA_0 = "0103ADB2CAB2CA4D36B3CABA3C61B3AAB7155276839396528188C489F937118CF4E5A3ED" 'len=72
    passA_1 = "0200AE17743474348BCC75347CD2A32DE42DCFD830F9D51C900FCA4A4B7F759BCA7A27392F" 'len=74
    passA_2 = "0406A809D8D9F5D9F5260BDAF5D113766A91E03C0DADCC5A216D0AAD4D74622AA68101F0B0F8" 'len=76
    passA_3 = "0604AA0BDE7D849A849A7B66859A8C4CB3A3D4A7FF426093A566A0D1FA943B2965E1FAC477737F" 'len=78

    Dim idx As Long
    idx = get_startIDX_of_DPB_from_vbaProject(bData)
    
    ' Extract DPB string.Extract characters until '"'
    Dim c As String
    Dim strDPB As String
    strDPB = ""
    For i = idx To idx + 80
        c = chr(bData(i))
        If c = chr(34) Then GoTo continue2       'chr(34)='"'
        strDPB = strDPB & c
    Next
continue2:
    Debug.Print "DPB=" & strDPB
    
    Dim newstrDPB As String
    If Len(strDPB) = 72 Then
        newstrDPB = passA_0
    ElseIf Len(strDPB) = 74 Then
        newstrDPB = passA_1
    ElseIf Len(strDPB) = 76 Then
        newstrDPB = passA_2
    ElseIf Len(strDPB) = 78 Then
        newstrDPB = passA_3
    Else
        'Debug.Print "no VBA password or error in strDPB"
        change_VBA_password_of_VBAProject = -1
        Exit Function
    End If
    
    For i = idx To idx + Len(strDPB) - 1
        'bData(i) = Asc(Mid(strDPB, i - idx + 1, 1))
        bData(i) = Asc(Mid(newstrDPB, i - idx + 1, 1))
    Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    'fso.DeleteFile IN_VBAProject_fullPath
    Call create_binary_file(OUT_VBAProject_fullPath, bData)
    '''''fso.GetFile(OUT_VBAProject_fullPath).Name = "vbaProject.bin"
    change_VBA_password_of_VBAProject = 0
Exit Function
ErrHandler:
    change_VBA_password_of_VBAProject = -1
End Function

