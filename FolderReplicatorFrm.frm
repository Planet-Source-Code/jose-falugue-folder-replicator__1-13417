VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FolderReplicatorFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Replicator"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7350
   Icon            =   "FolderReplicatorFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   100
      TabIndex        =   8
      Top             =   630
      Visible         =   0   'False
      Width           =   7125
      Begin ComctlLib.ListView ListView1 
         Height          =   2355
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4154
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "file"
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "size"
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   "from"
            Object.Tag             =   ""
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   "to"
            Object.Tag             =   ""
            Text            =   "To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   "action"
            Object.Tag             =   ""
            Text            =   "Action"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   100
      TabIndex        =   4
      Top             =   630
      Width           =   7125
      Begin VB.CommandButton BorrarCmd 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   6120
         Picture         =   "FolderReplicatorFrm.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete folder from the update list"
         Top             =   990
         Width           =   1000
      End
      Begin VB.CommandButton AgregarCmd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   6120
         Picture         =   "FolderReplicatorFrm.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add folders to the update list"
         Top             =   0
         Width           =   1000
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         ItemData        =   "FolderReplicatorFrm.frx":1CAC
         Left            =   90
         List            =   "FolderReplicatorFrm.frx":1CAE
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   0
         Width           =   5895
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   3130
      TabIndex        =   16
      Top             =   4140
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton SalirCmd 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   6280
      Picture         =   "FolderReplicatorFrm.frx":1CB0
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exit Folder Replicator"
      Top             =   3420
      Width           =   1000
   End
   Begin VB.CommandButton ReplicarCmd 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1060
      Picture         =   "FolderReplicatorFrm.frx":223F
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Update Folders from the Update List"
      Top             =   3420
      Width           =   1725
   End
   Begin VB.CommandButton AboutCmd 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   70
      Picture         =   "FolderReplicatorFrm.frx":30D9
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3420
      Width           =   1000
   End
   Begin VB.CheckBox CheckUPdate 
      Caption         =   "Update newer files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3130
      TabIndex        =   11
      Top             =   3420
      Value           =   1  'Checked
      Width           =   2445
   End
   Begin VB.CheckBox CheckAdd 
      Caption         =   "Add non existing files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3130
      TabIndex        =   10
      Top             =   3780
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton NuevoGrupoCmd 
      Caption         =   "New..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      ToolTipText     =   "Create a new folder's group"
      Top             =   90
      Width           =   1170
   End
   Begin VB.CommandButton EliminarGrupoCmd 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6030
      TabIndex        =   2
      ToolTipText     =   "Deletes group from the update list"
      Top             =   90
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FolderReplicatorFrm.frx":37B3
      Left            =   2160
      List            =   "FolderReplicatorFrm.frx":37BA
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   2640
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3165
      Left            =   68
      TabIndex        =   0
      Top             =   180
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5583
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   "general"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preview"
            Key             =   "preview"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Preview all the files before they will be updated"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   4500
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   7938
            MinWidth        =   7937
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   4939
            MinWidth        =   176
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5670
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
   End
End
Attribute VB_Name = "FolderReplicatorFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Private Result As Long
Private Cancelo As Long 'Sirve para saber si cuando estoy copiando los archivos
                        ' El usuario canceló la operación, y pueda salir de la función
                        'ReplicarPath que es recursiva.
Private TotalBytes As Long
Const IniFileName = "Paths.ini"
Private Const ACT_UPDATEFILES = 1, ACT_UPDATELISTVIEW = 2

'Variable usada en UpdateFiles para guardar los datos de los archivos a actualizar
Private Type FILELISTTYPE 'Usado en UpdateFiles para guardar los datos del archivo
    FileName As String
    ListIndexFrom As Long
    ListIndexTo As Long
End Type






Private Sub AboutCmd_Click()
About FolderReplicatorFrm
End Sub

Private Sub AgregarCmd_Click()
Dim PathStr As String
PathStr = BuscarCarpeta(CSIDL_DRIVES, BIF_RETURNFSANCESTORS)
'Si elijo Cancel devuelve vbNullString
'Para no repetir la carpeta en el LisBox verifico que no exista previamente
'en la función IsInListBox
If PathStr <> vbNullString And (Not IsInListBox(List1, PathStr)) Then
    List1.AddItem (PathStr)
    List1.Selected(List1.ListCount - 1) = True 'Selecciono el CheckBox
    GuardarPaths
End If

End Sub

Private Function IsInListBox(ListCtrl As Control, Cad As String) As Boolean
Dim x As Long
IsInListBox = False
For x = 0 To ListCtrl.ListCount - 1
    If ListCtrl.List(x) Like Cad Then
        IsInListBox = True
        Exit For
    End If
Next x
        

End Function
Private Sub BorrarCmd_Click()
If List1.ListIndex <> -1 Then
    List1.RemoveItem (List1.ListIndex)
    GuardarPaths
End If

End Sub


Private Sub CheckAdd_Click()
If CheckUPdate.Value = vbUnchecked Then
    CheckAdd.Value = vbChecked
End If
End Sub

Private Sub CheckUPdate_Click()
If CheckAdd.Value = vbUnchecked Then
    CheckUPdate.Value = vbChecked
End If
End Sub

Private Sub Combo1_Click()
Dim NombreCombinac As String, StrBuffer As String
Dim Cont As Long, Posic As Long
List1.Clear

StrBuffer = Space$(MAX_PATH)


NombreCombinac = Combo1.List(Combo1.ListIndex)

StrBuffer = Space$(MAX_PATH)
Cont = 1
Do
    Result = GetPrivateProfileString(NombreCombinac, Str$(Cont), "", StrBuffer, MAX_PATH - 1, CompletarPath(App.Path) & IniFileName)
    If Result > 0 Then
        Posic = InStr(3, StrBuffer, ":")
        If Posic <> 0 Then 'está aclarado en el archivo ini si está chequeado el item
            List1.AddItem (Left$(StrBuffer, Posic - 1))
            List1.Selected(Cont - 1) = Mid$(StrBuffer, Posic + 1, 1) = "1"
        Else
            List1.AddItem (Left$(StrBuffer, Result))
            List1.Selected(Cont - 1) = True
        End If
    End If
    Cont = Cont + 1
Loop While Result > 0
If TabStrip1.SelectedItem.Index = 2 Then
    CopiarDesdeListBox ACT_UPDATELISTVIEW
End If
End Sub


Private Sub EliminarGrupoCmd_Click()
Dim NombreGrupo As String
Dim IniPath As String
Dim Msg As String

NombreGrupo = Combo1.List(Combo1.ListIndex)
If Combo1.ListCount = 1 Then
    Msg = "Sorry, I can't delete " & NombreGrupo & "." & vbCrLf
    Msg = Msg & "Must exist at least one combination."
    MsgBox Msg, vbInformation, "Delete Combination"
    Exit Sub
End If
    
    
IniPath = CompletarPath(App.Path) & IniFileName

If Combo1.ListIndex <> -1 Then
    Combo1.RemoveItem (Combo1.ListIndex)
    Combo1.ListIndex = Combo1.ListCount - 1
    Result = WritePrivateProfileSection(NombreGrupo, vbNullString, IniPath)
    GuardarCombinaciones
End If
End Sub

Private Sub Form_Load()
Dim StrBuffer As String, Cont As Long
Dim Posic As Long
Dim NombreCombinac As String
Dim IniPath As String

With Picture1
    .Width = Screen.TwipsPerPixelX * 20
    .Height = Screen.TwipsPerPixelX * 20
    .BackColor = vbWindowBackground
End With

IniPath = CompletarPath(App.Path) & IniFileName
StrBuffer = Space$(MAX_PATH)
Result = GetPrivateProfileString("Options", "AddFiles", "1", StrBuffer, 2, IniPath)
CheckAdd.Value = Val(Left$(StrBuffer, Result))
Result = GetPrivateProfileString("Options", "UpdateFiles", "1", StrBuffer, 2, IniPath)
CheckUPdate.Value = Val(Left$(StrBuffer, Result))
'Combo1.List(0) = "Grupo1"
'Combo1.ListIndex = 0
'Leo las combinaciones y las guardo en el combo1
Cont = 1
Do
    Result = GetPrivateProfileString("Combinaciones", Str$(Cont), "", StrBuffer, MAX_PATH - 1, IniPath)
    If Result > 0 Then
        Combo1.List(Cont - 1) = (Left$(StrBuffer, Result))
    End If
    Cont = Cont + 1
Loop While Result > 0
If Cont = 2 Then 'No había guardada ninguna combinación
    Combo1.List(0) = "Grupo1"
    Combo1.ListIndex = 0
End If
Result = GetPrivateProfileString("Combinaciones", "Selected", "1", StrBuffer, MAX_PATH - 1, IniPath)
Combo1.ListIndex = Val(Left$(StrBuffer, Result)) - 1


End Sub


Private Sub List1_Click()
Dim SIZE As Long
Dim PathSelected As String
StatusBar1.SimpleText = ""
PathSelected = List1.List(List1.ListIndex)
    MousePointer = vbArrowHourglass
    SIZE = GetSizeOfPath(PathSelected)
    If FileExist(PathSelected) And SIZE <> -1 Then
        StatusBar1.SimpleText = "Size: " & Format$(SIZE, "#,#;;0") & " Bytes"
    Else
        Beep
        StatusBar1.SimpleText = "Sorry, I can't find " & PathSelected
    End If
MousePointer = vbDefault


End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    BorrarCmd_Click
End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
ListView1.SortOrder = IIf(ListView1.SortOrder = lvwAscending, lvwDescending, lvwAscending)
ListView1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub NuevoGrupoCmd_Click()
Dim NombreGrupo As String

NombreGrupo = InputBox("Enter the name for the new group:", "New Group", Combo1.List(Combo1.ListIndex))
If NombreGrupo <> "" Then
    If Not (IsInListBox(Combo1, NombreGrupo)) Then
        Combo1.AddItem (NombreGrupo)
        Combo1.ListIndex = Combo1.ListCount - 1
        GuardarCombinaciones
    End If
    
End If
End Sub

Private Sub ReplicarCmd_Click()
Select Case TabStrip1.SelectedItem.Index
    Case 1
        CopiarDesdeListBox ACT_UPDATEFILES
    Case 2
        'Para no tener que recorrer todo el arbol
        'Veo los archivos que ya tengo en listview
        CopiarDesdeListView
End Select
ListView1.ListItems.Clear

End Sub

Private Sub SalirCmd_Click()
GuardarPaths
End

End Sub
Private Sub GuardarPaths()
Dim x As Long, Path As String, IniPath As String
Dim Info As String
Dim NombreSection As String
IniPath = App.Path & "\" & IniFileName

Result = WritePrivateProfileString("Options", "AddFiles", Str$(CheckAdd.Value), IniPath)
Result = WritePrivateProfileString("Options", "UpdateFiles", Str$(CheckUPdate.Value), IniPath)

GuardarCombinaciones 'Agrego el nombre de la combinación en Paths.ini
NombreSection = Combo1.List(Combo1.ListIndex)
'Borro todas las claves anteriores porque si existía un número mayor de entradas
'y grabo un número menor las otras seguirán existiendo si no las borro
Result = WritePrivateProfileSection(NombreSection, vbNullString, IniPath)
    For x = 0 To List1.ListCount - 1
        Path = List1.List(x)
        Info = Path & ":" & IIf(List1.Selected(x), "1", "0")
        Result = WritePrivateProfileString(NombreSection, Str$(x + 1), Info, IniPath)
        DoEvents
    Next x
End Sub
Public Function LlenarListView(ByVal PathOrig As String, ByVal PathDest As String) As Long
Dim x As Long
Dim TempAct As Long, TempCopy As Long 'Guardo temporalmente la cantidad de archivos
Dim hFindFile As Long, Sigo As Long
Dim FileData As WIN32_FIND_DATA
Dim FileName As String, FileOrig As String, FileDest As String
TempAct = 0
TempCopy = 0

    
PathOrig = CompletarPath(PathOrig): PathDest = CompletarPath(PathDest)
hFindFile = FindFirstFile(PathOrig & "*.*", FileData)
If hFindFile = INVALID_HANDLE_VALUE Or Cancelo <> 0 Then
    LlenarListView = 0
    Exit Function
End If
With FileData
Do
    FileName = Left$(.cFileName, InStr(.cFileName, vbNullChar) - 1)
    FileDest = PathDest & FileName
    FileOrig = PathOrig & FileName
    
    If (.dwFileAttributes And vbDirectory) Then 'Es una carpeta
        If FileName <> ".." And FileName <> "." Then  'Es un nombre de carpeta válido
            LlenarListView FileOrig, FileDest
        End If
    Else 'es un archivo
       'veo si el archivo existe en el destino
       On Error Resume Next
        If FileExist(FileDest) Then 'El archivo existe
        'El archivo existe pero si el origen es mas nuevo lo copio al destino
            If FileDateTime(FileOrig) > FileDateTime(FileDest) And CheckUPdate.Value = vbChecked Then
                LLenarLineaListView PathOrig, PathDest, FileName, "Update"
                TempAct = TempAct + 1
                TotalBytes = TotalBytes + FileLen(FileOrig)
            End If
        ElseIf CheckAdd.Value = vbChecked Then 'el archivo no existe en el destino
            LLenarLineaListView PathOrig, PathDest, FileName, "Copy"
            TempCopy = TempCopy + 1
            TotalBytes = TotalBytes + FileLen(FileOrig)
        End If
           
    End If
    Sigo = FindNextFile(hFindFile, FileData)
Loop While Sigo <> 0
StatusBar1.SimpleText = ""
End With
NroArchivosActualizados = NroArchivosActualizados + TempAct
NroArchivosCopiados = NroArchivosCopiados + TempCopy
FindClose hFindFile
StatusBar1.SimpleText = ""

End Function

Public Function ReplicarPath(ByVal PathOrig As String, ByVal PathDest As String) As Long  ' , ByVal SubFolders As Boolean, ByVal ClusterSize&, AllocSize&) As Long
Dim x As Long
Dim TempAct As Long, TempCopy As Long 'Guardo temporalmente la cantidad de archivos
Dim hFindFile As Long, Sigo As Long
Dim FileData As WIN32_FIND_DATA
Dim FileName As String, FileOrig As String, FileDest As String
Dim FilesToCopy As String

FilesToCopy = vbNullString
TempAct = 0
TempCopy = 0
PathOrig = CompletarPath(PathOrig)
PathDest = CompletarPath(PathDest)
hFindFile = FindFirstFile(PathOrig & "*.*", FileData)
If hFindFile = INVALID_HANDLE_VALUE Or Cancelo <> 0 Then
    ReplicarPath = 0
    Exit Function
End If
With FileData
Do
    FileName = Left$(.cFileName, InStr(.cFileName, vbNullChar) - 1)
    FileDest = PathDest & FileName
    FileOrig = PathOrig & FileName
    If (.dwFileAttributes And vbDirectory) Then 'Es una carpeta
        If FileName <> ".." And FileName <> "." Then  'Es un nombre de carpeta válido
            ReplicarPath FileOrig, FileDest
            If Cancelo <> 0 Then Exit Function
        End If
    Else 'es un archivo
        'veo si el archivo existe en el destino
        If Dir(FileDest, vbHidden Or vbReadOnly Or vbSystem) <> "" Then
        'El archivo existe pero si el origen es mas nuevo lo copio al destino
            If FileDateTime(FileOrig) > FileDateTime(FileDest) And CheckUPdate.Value = vbChecked Then
                    FilesToCopy = FilesToCopy & FileOrig & vbNullChar
                    StatusBar1.SimpleText = "Updating " & FileDest
                    TempAct = TempAct + 1
                    TotalBytes = TotalBytes + FileLen(FileOrig)
            End If
        ElseIf CheckAdd.Value = vbChecked Then 'el archivo no existe en el destino
                FilesToCopy = FilesToCopy & FileOrig & vbNullChar
                StatusBar1.SimpleText = "Copying " & FileDest
                TempCopy = TempCopy + 1
                TotalBytes = TotalBytes + FileLen(FileOrig)
        End If
    End If
    Sigo = FindNextFile(hFindFile, FileData)
Loop While Sigo <> 0
StatusBar1.SimpleText = ""
End With
If FilesToCopy <> vbNullString Then
        Cancelo = ShellCopyFile(FilesToCopy, PathDest)
        If Cancelo = 0 Then 'Todo O.K.
            NroArchivosActualizados = NroArchivosActualizados + TempAct
            NroArchivosCopiados = NroArchivosCopiados + TempCopy
        End If
End If
    
FindClose hFindFile
StatusBar1.SimpleText = ""
End Function

Private Sub GuardarCombinaciones()
Dim IniPath As String, x As Long
IniPath = App.Path & "\" & IniFileName
Result = WritePrivateProfileSection("Combinaciones", vbNullString, IniPath)
For x = 0 To Combo1.ListCount - 1
    Result = WritePrivateProfileString("Combinaciones", Str$(x + 1), Combo1.List(x), IniPath)
Next x

'Guarddo el nro item seleccionado en el combo1
Result = WritePrivateProfileString("Combinaciones", "Selected", Combo1.ListIndex + 1, IniPath)


End Sub

Private Sub TabStrip1_Click()
Static PreviousSelected As Long
Dim TabElegido As Long
TabElegido = TabStrip1.SelectedItem.Index - 1
If TabElegido = PreviousSelected Then Exit Sub
Frame1(PreviousSelected).Visible = False

PreviousSelected = TabElegido
Frame1(PreviousSelected).Visible = True
If TabElegido = 1 Then
    CopiarDesdeListBox ACT_UPDATELISTVIEW
End If
End Sub

Private Sub LLenarLineaListView(ByVal PathOrig$, ByVal PathDest$, ByVal FileName, ByVal Accion)
Const LargoString = 40
Dim FileSize As Long
Dim FileOrig As String, FileDest As String
Dim imgX As ListImage
Dim itemX As ListItem
Dim hIcon As Long

FileOrig = PathOrig & FileName
FileSize = FileLen(PathOrig & FileName) \ 1024
FileSize = IIf(FileSize = 0, 1, FileSize)

hIcon = jlfExtractAssociatedIcon(FileOrig, 0)
With Picture1
    .Picture = LoadPicture()
    DrawIconEx .hdc, 1, 1, hIcon, 16, 16, 0, 0, DI_NORMAL
    .Refresh
End With

Set imgX = ImageList1.ListImages.Add(, FileOrig, Picture1.Image)
Set itemX = ListView1.ListItems.Add(, FileOrig, FileName, FileOrig, FileOrig)
With itemX
    .SubItems(1) = Format$(FileSize, "#,0 Kb")
    .SubItems(2) = PathOrig
    .SubItems(3) = PathDest
    .SubItems(4) = Accion
End With
DoEvents

Result = DestroyIcon(hIcon)
End Sub
Private Sub CopiarDesdeListView()
Dim itmX As ListItem
Dim DestPath As String
Dim FilesToCopy As String
Dim DestFolder As String, NextDestFolder As String
Dim TempArray() As String
Dim x As Long, y As Long, NroItems As Long

' itmX.SubItems(2) = Folder Origen : itmX.SubItems(3) = Folder Destino
If ListView1.ListItems.Count = 0 Then Exit Sub

Set itmX = ListView1.ListItems.Item(1)
ReDim TempArray(ListView1.ListItems.Count - 1)

For Each itmX In ListView1.ListItems
With itmX   'Pongo primero los folders destino en TempArray para ordenarlos
            'y luego mandar varios archivos al mismo folder
        TempArray(itmX.Index - 1) = .SubItems(3) + vbNullChar + .SubItems(2) + .Text
End With
Next itmX
'Ordeno el array segun los folders destino
SortStringArray TempArray

NroItems = UBound(TempArray)
With ProgressBar1
    .Visible = True
    .Max = NroItems + 1
End With

x = 0
Do
    FilesToCopy = ""
    DestFolder = GetStrNro(TempArray(x), 1, vbNullChar)
    Do
        ProgressBar1.Value = x + 1
        FilesToCopy = FilesToCopy + GetStrNro(TempArray(x), 2, vbNullChar) + vbNullChar
        x = x + 1
        If x <= NroItems Then
            NextDestFolder = GetStrNro(TempArray(x), 1, vbNullChar)
        Else
            NextDestFolder = vbNullChar
        End If
    Loop While NextDestFolder = DestFolder And x <= NroItems
    Result = ShellCopyFile(FilesToCopy, DestFolder)
Loop While x <= NroItems

ProgressBar1.Visible = False
StatusBar1.SimpleText = ""
End Sub

Public Function UpdateFolders(ByVal PathOrigen As String, ByVal PathDestino As String) As Long
'En SubFolders le voy pasando las subcarpetas comunes a todos los paths que guardo
'en el ListBox

Dim TempAct As Long, TempCopy As Long 'Guardo temporalmente la cantidad de archivos
Dim hFindFile As Long, Sigo As Long, FileData As WIN32_FIND_DATA
Dim FileName As String, FilesToCopy As String
Dim ListPathName As String
Dim x As Long
Dim FileArray() As FILELISTTYPE
Dim FileOrig As String, FileDest As String, ExistDest As Boolean

UpdateFolders = 1: FilesToCopy = vbNullString
TempAct = 0: TempCopy = 0
PathOrigen = CompletarPath(PathOrigen): PathDestino = CompletarPath(PathDestino)
hFindFile = FindFirstFile(PathOrigen & "*.*", FileData)
If hFindFile = INVALID_HANDLE_VALUE Or Cancelo <> 0 Then
    UpdateFolders = 0
    Exit Function
End If

Do
    FileName = Left$(FileData.cFileName, InStr(FileData.cFileName, vbNullChar) - 1)
    FileOrig = PathOrigen + FileName
    FileDest = PathDestino + FileName: ExistDest = FileExist(FileDest)
    If (FileData.dwFileAttributes And vbDirectory) Then 'Es una carpeta
        If FileName <> ".." And FileName <> "." Then  'Es un nombre de carpeta válido
            'Paso como parámetro sólo las Subcarpetas
            If Not ExistDest Then
                MakeAllDir (FileDest)
            End If
            Result = UpdateFolders(FileOrig, FileDest)
        End If
        
    Else 'Es un archivo
        If ExistDest Then
        'El archivo existe pero si el origen es mas nuevo lo copio al destino
            If FileDateTime(FileOrig) > FileDateTime(FileDest) And CheckUPdate.Value = vbChecked Then
                    FilesToCopy = FilesToCopy & FileOrig & vbNullChar
                    StatusBar1.SimpleText = "Updating " & FileDest
                    TempAct = TempAct + 1
                    TotalBytes = TotalBytes + FileLen(FileOrig)
            End If
        ElseIf CheckAdd.Value = vbChecked Then 'el archivo no existe en el destino
                FilesToCopy = FilesToCopy & FileOrig & vbNullChar
                StatusBar1.SimpleText = "Copying " & FileDest
                TempCopy = TempCopy + 1
                TotalBytes = TotalBytes + FileLen(FileOrig)
        End If
        
    End If
    Sigo = FindNextFile(hFindFile, FileData)
Loop While Sigo <> 0
FindClose hFindFile
If FilesToCopy <> vbNullString Then
    Cancelo = ShellCopyFile(FilesToCopy, PathDestino)
    If Cancelo = 0 Then 'Todo O.K.
        NroArchivosActualizados = NroArchivosActualizados + TempAct
        NroArchivosCopiados = NroArchivosCopiados + TempCopy
    End If
End If

StatusBar1.SimpleText = ""

End Function
Public Sub CopiarDesdeListBox(ByVal Action As Long)
Dim Msg As String, CopioFolder As Boolean
Dim NroItemsEnLista As Long, imgX As ListImage
Dim SrcFolder As Long, DestFolder As Long
NroArchivosCopiados = 0: NroArchivosActualizados = 0: TotalBytes = 0
If Action = ACT_UPDATELISTVIEW Then
With ListView1
   .ListItems.Clear
   .SmallIcons = Nothing: .Icons = Nothing
   ImageList1.ListImages.Clear
   Set imgX = ImageList1.ListImages.Add(, , Picture1.Image)
   .Icons = ImageList1: .SmallIcons = ImageList1
End With
End If

NroItemsEnLista = List1.ListCount
Cancelo = 0
MousePointer = vbArrowHourglass
With ProgressBar1
    .Visible = True
    .Max = NroItemsEnLista ^ 2 'IIf(NroItemsEnLista > 0, NroItemsEnLista ^ 2, 1)
    .Value = 0
End With

For SrcFolder = 0 To NroItemsEnLista - 2
        For DestFolder = SrcFolder + 1 To NroItemsEnLista - 1
            ProgressBar1.Value = (SrcFolder) * (NroItemsEnLista) + DestFolder + 1
            DoEvents
            CopioFolder = List1.Selected(SrcFolder) And List1.Selected(DestFolder)
            If Cancelo = 0 And (CopioFolder) Then
                If Action = ACT_UPDATEFILES Then
                    UpdateFolders List1.List(SrcFolder), List1.List(DestFolder)
                    UpdateFolders List1.List(DestFolder), List1.List(SrcFolder)

                ElseIf Action = ACT_UPDATELISTVIEW Then
                    LlenarListView List1.List(SrcFolder), List1.List(DestFolder)
                    LlenarListView List1.List(DestFolder), List1.List(SrcFolder)
               End If
                               
            End If
        Next DestFolder
Next SrcFolder
If Action = ACT_UPDATEFILES Then
    Msg = "Copied Files  : " & NroArchivosCopiados
    Msg = Msg & "---Updated Files : " & NroArchivosActualizados
    Msg = Msg & "---TOTAL: " & NroArchivosActualizados + NroArchivosCopiados
ElseIf Action = ACT_UPDATELISTVIEW Then
    Msg = "Files to Copy : " & NroArchivosCopiados
    Msg = Msg & "-- Files to update: " & NroArchivosActualizados
    Msg = Msg & "-- TOTAL: " & NroArchivosActualizados + NroArchivosCopiados
    Msg = Msg & "-- TOTAL Bytes: " & Format$(TotalBytes, "#,#")

End If
If Cancelo <> 0 Then
        Msg = Msg & "--- Operation Canceled!!" & vbCrLf
End If
StatusBar1.SimpleText = Msg
ProgressBar1.Visible = False
MousePointer = vbDefault





End Sub
