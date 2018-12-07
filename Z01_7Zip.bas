'modified from https://www.rondebruin.nl/win/winfiles/7zip_zipexamples.txt

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#End If

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

Function Susutan_Buka(ByVal NamLokZip As String, ByVal LokasiBuka As String, Optional ByVal Sandi As String)

    If Sandi = "" Then
        B_UnZip_Zip_File_Fixed NamLokZip, LokasiBuka
    Else
        Sandi = Sembunyikan(Sandi)
        Sandi = Left(Sandi, 48)
        B_UnZip_Zip_File_Fixed NamLokZip, LokasiBuka, Sandi
    End If
    
End Function

Function Berkas_BuatZip(ByVal NamLokBerkas As String, ByVal NamLokZipTanpaEkstensi As String, Optional ByVal Sandi As String)
    
    If Sandi = "" Then
        Susutkan NamLokBerkas, NamLokZipTanpaEkstensi
    Else
        Sandi = Sembunyikan(Sandi)
        Sandi = Left(Sandi, 48)
        Susutkan NamLokBerkas, NamLokZipTanpaEkstensi, 4, Sandi
    End If
    
End Function

Private Function Susutkan(ByVal NamLokBerkas As String, ByVal NamLokZipTanpaEkstensi As String, Optional ByVal Mode As Integer, Optional ByVal EkstensiAtauKataKunci As String)
    Dim PathZipProgram As String, NameZipFile As String, FolderName As String
    Dim ShellStr As String, strDate As String, DefPath As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        Debug.Print "Please find your copy of 7z.exe and try again"
        Exit Function
    End If

    DefPath = NamLokZipTanpaEkstensi
    
    If Not NamLokBerkas = "" Then
        FolderName = NamLokBerkas
        
        Select Case Mode
            Case 0  'Zip all the files in the folder and subfolders, -r is Include subfolders
                NameZipFile = DefPath & ".zip"
                ShellStr = PathZipProgram & "7z.exe a -r" _
                         & " " & Chr(34) & NameZipFile & Chr(34) _
                         & " " & Chr(34) & FolderName & "\" & "*.*" & Chr(34)
            Case 1  'Zip the Ekstensi files in the folder and subfolders
                NameZipFile = DefPath & ".zip"
                ShellStr = PathZipProgram & "7z.exe a -r" _
                                  & " " & Chr(34) & NameZipFile & Chr(34) _
                                  & " " & Chr(34) & FolderName & "\" & "*." & EkstensiAtauKataKunci & "*" & Chr(34)
            Case 2  'Zip all files in the folder and subfolders with a name that start with Week
                NameZipFile = DefPath & ".zip"
                ShellStr = PathZipProgram & "7z.exe a -r" _
                                  & " " & Chr(34) & NameZipFile & Chr(34) _
                                  & " " & Chr(34) & FolderName & "\" & "*" & EkstensiAtauKataKunci & "*.*" & Chr(34)
            Case 3  'Zip every file with the name ron.xlsx in the folder and subfolders
                NameZipFile = DefPath & ".zip"
                ShellStr = PathZipProgram & "7z.exe a -r" _
                                  & " " & Chr(34) & NameZipFile & Chr(34) _
                                  & " " & Chr(34) & FolderName & "\" & EkstensiAtauKataKunci & Chr(34)
            Case 4  'Add -ppassword -mhe of you want to add a password to the zip file(only .7z files)
                NameZipFile = DefPath & ".7z"
                ShellStr = PathZipProgram & "7z.exe a -r -p" & EkstensiAtauKataKunci & " -mhe" _
                                                   & " " & Chr(34) & NameZipFile & Chr(34) _
                                                   & " " & Chr(34) & FolderName & "\" & "*.*" & Chr(34)
            Case 5  'Zip only a file with password -ppassword -mhe of you want to add a password to the zip file(only .7z files)
                NameZipFile = DefPath & ".7z"
                ShellStr = PathZipProgram & "7z.exe a -r -p" & EkstensiAtauKataKunci & " -mhe" _
                                                   & " " & Chr(34) & NameZipFile & Chr(34) _
                                                   & " " & Chr(34) & FolderName & Chr(34)
        End Select
        
        ShellAndWait ShellStr, vbHide
        
'        Debug.Print "You will find the zip file here: " & NamLokZipTanpaEkstensi
        Berkas_Susutkan = True
    End If
End Function

'With this example you zip the ActiveWorkbook
'The name of the zip file will be the name of the workbook + Date/Time
'The zip file will be saved in: DefPath = Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'You can change this folder to this if you want to use another folder
'DefPath = "C:\Users\Ron\ZipFolder"
'There is no need to change the code before you test it

Private Sub E_Zip_ActiveWorkbook()
    Dim PathZipProgram As String, NameZipFile As String
    Dim ShellStr As String, strDate As String, DefPath As String
    Dim FileNameXls As String, TempFilePath As String, TempFileName As String
    Dim MyWb As Workbook, FileExtStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If
    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'Build the path and name for the new xls? file
    Set MyWb = ActiveWorkbook
    If ActiveWorkbook.Path = "" Then Exit Sub

    TempFilePath = Environ$("temp") & "\"
    FileExtStr = "." & LCase(Right(MyWb.Name, _
                                   Len(MyWb.Name) - InStrRev(MyWb.Name, ".", , 1)))
    TempFileName = Left(MyWb.Name, Len(MyWb.Name) - Len(FileExtStr))

    'Use SaveCopyAs to make a copy of the file
    FileNameXls = TempFilePath & TempFileName & FileExtStr
    MyWb.SaveCopyAs FileNameXls

    'Build the path and name for the new zip file
    'The name of the zip file will be the name of the workbook + Date/Time
    'The zip file will be saved in: DefPath = Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder.
    'You can change this folder to this if you want to use another folder
    'DefPath = "C:\Users\Ron\ZipFolder"
    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If
    strDate = Format(Now, "yyyy-mm-dd h-mm-ss")
    NameZipFile = DefPath & TempFileName & " " & strDate & ".zip"

    'Zip FileNameXls (copy of the ActiveWorkbook)
    ShellStr = PathZipProgram & "7z.exe a" _
             & " " & Chr(34) & NameZipFile & Chr(34) _
             & " " & Chr(34) & FileNameXls & Chr(34)
    ShellAndWait ShellStr, vbHide

    'Delete the file that you saved with SaveCopyAs and add to the zip file
    Kill TempFilePath & TempFileName & FileExtStr

    MsgBox "You will find the zip file here: " & NameZipFile
End Sub

Private Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub

Private Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function

'With this example you browse to the zip or 7z file you want to unzip
'The zip file will be unzipped in a new folder in: Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'The name of the folder that the code create in this folder is the Date/Time
'You can change this folder to this if you want to use a fixed folder:
'NameUnZipFolder = "C:\Users\Ron\TestFolder\"
'Read the comments in the code about the commands/Switches in the ShellStr
'There is no need to change the code before you test it

Private Sub A_UnZip_Zip_File_Browse()
    Dim PathZipProgram As String, NameUnZipFolder As String
    Dim FileNameZip As Variant, ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        Debug.Print "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'Create path and name of the normal folder to unzip the files in
    'In this example we use: Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder
    'The name of the folder that the code create in this folder is the Date/Time
    NameUnZipFolder = Application.DefaultFilePath & "\" & Format(Now, "yyyy-mm-dd h-mm-ss")
    'You can also use a fixed path like
    'NameUnZipFolder = "C:\Users\Ron\TestFolder"

    'Select the zip file (.zip or .7z files)
    FileNameZip = Application.GetOpenFilename(filefilter:="Zip Files, *.zip, 7z Files, *.7z", _
                                              MultiSelect:=False, Title:="Select the file that you want to unzip")

    'Unzip the files/folders from the zip file in the NameUnZipFolder folder
    If FileNameZip = False Then
        'do nothing
    Else
        'There are a few commands/Switches that you can change in the ShellStr
        'We use x command now to keep the folder stucture, replace it with e if you want only the files
        '-aoa Overwrite All existing files without prompt.
        '-aos Skip extracting of existing files.
        '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
        '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
        'Use -r if you also want to unzip the subfolders from the zip file
        'You can add -ppassword if you want to unzip a zip file with password (only 7zip files)
        'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
        'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
        ShellStr = PathZipProgram & "7z.exe x -aoa -r" _
                 & " " & Chr(34) & FileNameZip & Chr(34) _
                 & " -o" & Chr(34) & NameUnZipFolder & Chr(34) & " " & "*.*"

        ShellAndWait ShellStr, vbHide
        Debug.Print "Look in " & NameUnZipFolder & " for extracted files"

    End If
End Sub

'With this example you unzip a fixed zip file: FileNameZip = "C:\Users\Ron\Test.zip"
'Note this file must exist, this is the only thing that you must change before you test it
'The zip file will be unzipped in a new folder in: Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'The name of the folder that the code create in this folder is the Date/Time
'You can change this folder to this if you want to use a fixed folder:
'NameUnZipFolder = "C:\Users\Ron\TestFolder\"
'Read the comments in the code about the commands/Switches in the ShellStr

Private Sub B_UnZip_Zip_File_Fixed(ByVal NamLokZip As String, ByVal LokasiBuka As String, _
    Optional ByVal Sandi As String)
    
    Dim PathZipProgram As String, NameUnZipFolder As String
    Dim FileNameZip As Variant, ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        Debug.Print "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'Create path and name of the normal folder to unzip the files in
    'In this example we use: Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder
    'The name of the folder that the code create in this folder is the Date/Time
    NameUnZipFolder = LokasiBuka
    'You can also use a fixed path like
    'NameUnZipFolder = "C:\Users\Ron\TestFolder\"

    'Name of the zip file that you want to unzip (.zip or .7z files)
    FileNameZip = NamLokZip
    
    'There are a few commands/Switches that you can change in the ShellStr
    'We use x command now to keep the folder stucture, replace it with e if you want only the files
    '-aoa Overwrite All existing files without prompt.
    '-aos Skip extracting of existing files.
    '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
    '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
    'Use -r if you also want to unzip the subfolders from the zip file
    'You can add -ppassword if you want to unzip a zip file with password (only .7z files)
    'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
    'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
    ShellStr = PathZipProgram & "7z.exe x -aoa -r -p" & Sandi _
             & " " & Chr(34) & FileNameZip & Chr(34) _
             & " -o" & Chr(34) & NameUnZipFolder & Chr(34) & " " & "*.*"

    ShellAndWait ShellStr, vbHide
    Debug.Print "Look in " & NameUnZipFolder & " for extracted files"

End Sub
