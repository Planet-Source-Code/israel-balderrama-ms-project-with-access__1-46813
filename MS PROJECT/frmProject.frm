VERSION 5.00
Begin VB.Form frmProject 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3375
   Begin VB.CommandButton cmd 
      Caption         =   "&TAREAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "MS PROJECT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'''CODIGO GENERADO POR ISRAEL BALDERRAMA'''
'''EMAIL: IBALDERRAMA@HOTMAIL.COM       '''
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Dim cnn As New ADODB.Connection
Dim comando As ADODB.Command

Dim query As ADODB.Recordset
Dim QUERYNIVEL As ADODB.Recordset
Dim QUERYTIPO As ADODB.Recordset
Dim QUERYFASE As ADODB.Recordset
Dim QUERYACT As ADODB.Recordset

Dim queryFec As ADODB.Recordset

Dim P As Object

Dim task As Integer
'Dim Dif As Integer

Dim fini As Date
Dim ffin As Date

Dim BCap As Boolean
Dim BEd As Boolean
Dim BPar As Boolean

Private Sub cmd_Click()

On Error GoTo ErrorProject

BCap = False
BEd = False
BPar = False

Set P = CreateObject("MSProject.Application")
P.Visible = True
P.FileNew

Call Conexion

Call GProject

Exit Sub

ErrorProject:
MsgBox Err.Description

End Sub

Private Sub GProject()

comando.CommandText = "SELECT MIN(fini) AS FIni, MAX(ffin) AS Ffin FROM tbl_Tarea" & _
                    " WHERE IdProyecto=1"
Set queryFec = comando.Execute()

P.ActiveProject.ProjectStart = "" & queryFec!fini
P.ActiveProject.ProjectFinish = "" & queryFec!ffin

task = 1

comando.CommandText = "SELECT IdNivel FROM tbl_Proyecto WHERE IdProyecto=1"
Set QUERYNIVEL = comando.Execute()

Do While QUERYNIVEL.EOF = False
    comando.CommandText = "SELECT MIN(fini) AS FIni, MAX(ffin) AS FFin FROM tbl_Tarea" & _
                        " WHERE IdNivel='" & Trim(QUERYNIVEL!IdNivel) & "'"
    Set queryFec = comando.Execute()
    
    P.ActiveProject.Tasks.Add Name:="NIVEL: " & UCase(Trim(QUERYNIVEL!IdNivel))
    P.ActiveProject.Tasks(task).Start = "" & queryFec!fini
    P.ActiveProject.Tasks(task).Finish = "" & queryFec!ffin
    
    comando.CommandText = "SELECT IdTipo, (SELECT Desc " & _
                "FROM tbl_Tipo WHERE Id_Tipo=tbl_Tarea.IdTipo) AS Nom " & _
                "FROM tbl_Tarea WHERE IdNivel='" & QUERYNIVEL!IdNivel & _
                "' GROUP BY IdTipo"
    Set QUERYTIPO = comando.Execute()
    
    Do While QUERYTIPO.EOF = False
        
        comando.CommandText = "SELECT MIN(fini) AS FIni, MAX(ffin) AS FFin FROM tbl_Tarea" & _
                        " WHERE IdNivel='" & Trim(QUERYNIVEL!IdNivel) & "' " & _
                        "AND IdTipo=" & QUERYTIPO!IdTipo
        Set queryFec = comando.Execute()
                                        
        task = task + 1
        P.ActiveProject.Tasks.Add Name:="TIPO : " & UCase(QUERYTIPO!Nom)
        P.ActiveProject.Tasks(task).Start = "" & queryFec!fini
        P.ActiveProject.Tasks(task).Finish = "" & queryFec!ffin
'        MsgBox QUERYTIPO!Id_Capitulo
        If BCap = True Then
            P.ActiveProject.Tasks(task).OutlineOutdent
            P.ActiveProject.Tasks(task).OutlineOutdent
            'P.ActiveProject.Tasks(task).OutlineOutdent
            'P.ActiveProject.Tasks(task).OutlineOutdent
            BCap = False
        End If
        
        If i = 0 Then
            P.ActiveProject.Tasks(task).OutlineIndent
            i = 1
        End If
        
        comando.CommandText = "SELECT IdFase, (SELECT Desc FROM " & _
                    "tbl_Fase WHERE Id_Fase=tbl_Tarea.IdFase) AS NomEd" & _
                    " FROM tbl_Tarea WHERE IdNivel='" & QUERYNIVEL!IdNivel & _
                    "' AND IdTipo=" & QUERYTIPO!IdTipo & " GROUP BY IdFase"
        Set QUERYFASE = comando.Execute()
        BEd = False
        Do While QUERYFASE.EOF = False
        
            comando.CommandText = "SELECT MIN(fini) AS FIni, MAX(ffin) AS FFin FROM tbl_Tarea" & _
                        " WHERE IdNivel='" & Trim(QUERYNIVEL!IdNivel) & "' " & _
                        "AND IdTipo=" & QUERYTIPO!IdTipo & " AND IdFase='" & _
                        Trim(QUERYFASE!IdFase) & "'"
            Set queryFec = comando.Execute()
            
            task = task + 1
            P.ActiveProject.Tasks.Add Name:="" & UCase(Trim(QUERYFASE!NomEd))
            P.ActiveProject.Tasks(task).Start = "" & queryFec!fini
            P.ActiveProject.Tasks(task).Finish = "" & queryFec!ffin
    '        MsgBox QUERYTIPO!Id_Capitulo
                
            If BEd = True Then
                P.ActiveProject.Tasks(task).OutlineOutdent
                P.ActiveProject.Tasks(task).OutlineOutdent
                'P.ActiveProject.Tasks(task).OutlineOutdent
                BEd = False
            End If
        
            If j = 0 Then
                P.ActiveProject.Tasks(task).OutlineIndent
                j = 1
            End If
            
            comando.CommandText = "SELECT IdActividad, (SELECT Desc FROM " & _
                        "tbl_Actividad WHERE Id_Act=tbl_Tarea.IdActividad) AS NomP" & _
                        " FROM tbl_Tarea WHERE IdNivel='" & QUERYNIVEL!IdNivel & _
                        "' AND IdTipo=" & QUERYTIPO!IdTipo & " AND IdFase='" & _
                        Trim(QUERYFASE!IdFase) & "' GROUP BY IdActividad"
            Set QUERYACT = comando.Execute()
            
            BPar = False
            Do While QUERYACT.EOF = False
                comando.CommandText = "SELECT MIN(fini) AS FIni, MAX(ffin) AS FFin FROM tbl_Tarea" & _
                            " WHERE IdNivel='" & Trim(QUERYNIVEL!IdNivel) & "' " & _
                            "AND IdTipo=" & QUERYTIPO!IdTipo & " AND IdFase='" & _
                            Trim(QUERYFASE!IdFase) & "' AND IdActividad=" & QUERYACT!IdActividad
    
                Set queryFec = comando.Execute()
                
                task = task + 1
                P.ActiveProject.Tasks.Add Name:="" & Trim(QUERYACT!NomP)
                P.ActiveProject.Tasks(task).Start = "" & queryFec!fini
                P.ActiveProject.Tasks(task).Finish = "" & queryFec!ffin
        '        MsgBox QUERYTIPO!Id_Capitulo
                
                If BPar = True Then
                    P.ActiveProject.Tasks(task).OutlineOutdent
                    BPar = False
                End If
                
                If k = 0 Then
                    P.ActiveProject.Tasks(task).OutlineIndent
                    k = 1
                End If
                
                comando.CommandText = "SELECT Id_Tarea, fini, ffin FROM tbl_Tarea" & _
                                " WHERE IdNivel='" & Trim(QUERYNIVEL!IdNivel) & "' " & _
                                "AND IdTipo=" & QUERYTIPO!IdTipo & " AND IdFase='" & _
                                Trim(QUERYFASE!IdFase) & "' AND IdActividad=" & QUERYACT!IdActividad
                Set query = comando.Execute
                
                Do While query.EOF = False
                    task = task + 1
                    
                    P.ActiveProject.Tasks.Add Name:="" & query!Id_Tarea
                    P.ActiveProject.Tasks(task).Start = "" & query!fini
                    P.ActiveProject.Tasks(task).Finish = "" & query!ffin
                
                    If l = 0 Then
                        P.ActiveProject.Tasks(task).OutlineIndent
                        'P.ActiveProject.Tasks(task).OutlineIndent
                    l = 1
                    End If
                    query.MoveNext
                Loop
                l = 0
                BPar = True
                QUERYACT.MoveNext
            Loop
            k = 0
            BEd = True
            QUERYFASE.MoveNext
        Loop
        j = 0
        BCap = True
        QUERYTIPO.MoveNext
    Loop
    
QUERYNIVEL.MoveNext
Loop

'SELECCIONA COLUMNA
'P.ActiveProject.SelectColumn 3
'AJUSTA EL ANCHO DE LA COLUMNA
P.ColumnBestFit 3
'ELIMINA COLUMNA
P.ColumnDelete
'ENCABEZADO DE PAGINA
P.FilePageSetupHeader , pjCenter, "TRABAJO: 1011"
'PROPIEDADES DE VISTA
P.FilePageSetupView , False, , False, False, False
'PROPIEDADES DE LEGENDA
P.FilePageSetupLegend , , pjNoLegend

cnn.Close

End Sub


Private Sub Conexion()

With cnn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source") = App.Path & "\BD.mdb"
    '.Properties("Jet OLEDB:Database Password") = "2822con"
    .CursorLocation = adUseClient
    .Open
End With

Set comando = New ADODB.Command
Set query = New ADODB.Recordset

comando.ActiveConnection = cnn
comando.CommandType = adCmdText

End Sub

