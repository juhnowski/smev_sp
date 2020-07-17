VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SudebniyPrikaz 
   Caption         =   "Судебный приказ"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Судебный участок"
      TabPicture(0)   =   "SudebniyPrikaz.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "mirovoy_sudiya"
      Tab(0).Control(1)=   "site"
      Tab(0).Control(2)=   "address"
      Tab(0).Control(3)=   "polnoe_naimenovanie"
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "Label3"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Взыскатель"
      TabPicture(1)   =   "SudebniyPrikaz.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SchetPoluchatelya"
      Tab(1).Control(1)=   "KBK"
      Tab(1).Control(2)=   "BIK"
      Tab(1).Control(3)=   "OKTMO"
      Tab(1).Control(4)=   "OKATO"
      Tab(1).Control(5)=   "l_s"
      Tab(1).Control(6)=   "ogrn"
      Tab(1).Control(7)=   "okpo"
      Tab(1).Control(8)=   "kpp"
      Tab(1).Control(9)=   "inn"
      Tab(1).Control(10)=   "kor_schet"
      Tab(1).Control(11)=   "vz_mesto_nahozhdeniya"
      Tab(1).Control(12)=   "vz_naimenovanie"
      Tab(1).Control(13)=   "Label31"
      Tab(1).Control(14)=   "Label30"
      Tab(1).Control(15)=   "Label29"
      Tab(1).Control(16)=   "Label28"
      Tab(1).Control(17)=   "Label27"
      Tab(1).Control(18)=   "Label15"
      Tab(1).Control(19)=   "Label14"
      Tab(1).Control(20)=   "Label13"
      Tab(1).Control(21)=   "Label12"
      Tab(1).Control(22)=   "Label11"
      Tab(1).Control(23)=   "Label10"
      Tab(1).Control(24)=   "Label9"
      Tab(1).Control(25)=   "Label8"
      Tab(1).Control(26)=   "Label7"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Должник"
      TabPicture(2)   =   "SudebniyPrikaz.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label19"
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(5)=   "Label21"
      Tab(2).Control(6)=   "fio_dolzhnika"
      Tab(2).Control(7)=   "mesto_zhitelstva"
      Tab(2).Control(8)=   "identificator"
      Tab(2).Control(9)=   "data_rozhdeniya"
      Tab(2).Control(10)=   "mesto_rozhdeniya"
      Tab(2).Control(11)=   "mesto_raboti"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Судебный приказ"
      TabPicture(3)   =   "SudebniyPrikaz.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label22"
      Tab(3).Control(1)=   "Label24"
      Tab(3).Control(2)=   "Label25"
      Tab(3).Control(3)=   "Label26"
      Tab(3).Control(4)=   "period"
      Tab(3).Control(5)=   "summa"
      Tab(3).Control(6)=   "AddV"
      Tab(3).Control(7)=   "DelV"
      Tab(3).Control(8)=   "summa_list"
      Tab(3).Control(9)=   "VidV"
      Tab(3).Control(10)=   "EditV"
      Tab(3).Control(11)=   "SaveV"
      Tab(3).Control(12)=   "UndoV"
      Tab(3).Control(13)=   "period_razmernost"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Выгрузка XML"
      TabPicture(4)   =   "SudebniyPrikaz.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "CommandButtonValidate"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "CommandButtonOpen"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "CommandButtonGenerate"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "CommandButtonSign"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "CommandButtonSend"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "log"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Command3"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      Begin VB.ComboBox period_razmernost 
         Height          =   315
         ItemData        =   "SudebniyPrikaz.frx":008C
         Left            =   -68040
         List            =   "SudebniyPrikaz.frx":0099
         TabIndex        =   72
         Text            =   "месяцев"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox SchetPoluchatelya 
         Height          =   375
         Left            =   -69480
         TabIndex        =   71
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox KBK 
         Height          =   405
         Left            =   -69960
         TabIndex        =   69
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox BIK 
         Height          =   405
         Left            =   -69960
         TabIndex        =   68
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox OKTMO 
         Height          =   405
         Left            =   -69960
         TabIndex        =   67
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox OKATO 
         Height          =   405
         Left            =   -69960
         TabIndex        =   66
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton UndoV 
         Caption         =   "Отменить"
         Height          =   375
         Left            =   -66840
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton SaveV 
         Caption         =   "Сохранить"
         Height          =   375
         Left            =   -68520
         TabIndex        =   60
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton EditV 
         Caption         =   "Редактировать"
         Height          =   375
         Left            =   -73200
         TabIndex        =   59
         Top             =   4320
         Width           =   1575
      End
      Begin VB.ComboBox VidV 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "SudebniyPrikaz.frx":00B1
         Left            =   -69480
         List            =   "SudebniyPrikaz.frx":00C7
         TabIndex        =   58
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox mirovoy_sudiya 
         Height          =   315
         Left            =   -72840
         TabIndex        =   55
         Top             =   1920
         Width           =   7575
      End
      Begin VB.ComboBox site 
         Height          =   315
         Left            =   -72840
         TabIndex        =   54
         Top             =   1440
         Width           =   7575
      End
      Begin VB.ComboBox address 
         Height          =   315
         Left            =   -72840
         TabIndex        =   53
         Top             =   960
         Width           =   7575
      End
      Begin VB.ComboBox polnoe_naimenovanie 
         Height          =   315
         Left            =   -72840
         TabIndex        =   52
         Top             =   480
         Width           =   7575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Сохранить как шаблон"
         Height          =   495
         Left            =   6480
         TabIndex        =   51
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox log 
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   1080
         Width           =   9735
      End
      Begin VB.CommandButton CommandButtonSend 
         Caption         =   "Отправить"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6120
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton CommandButtonSign 
         Caption         =   "Подписать"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5880
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton CommandButtonGenerate 
         Caption         =   "Сгенерировать"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3960
         TabIndex        =   47
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton CommandButtonOpen 
         Caption         =   "Открыть"
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton CommandButtonValidate 
         Caption         =   "Проверить"
         Height          =   495
         Left            =   2040
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox summa_list 
         Height          =   2985
         Left            =   -74880
         TabIndex        =   44
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CommandButton DelV 
         Caption         =   "Удалить"
         Height          =   375
         Left            =   -71520
         TabIndex        =   43
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton AddV 
         Caption         =   "Добавить"
         Height          =   375
         Left            =   -74880
         TabIndex        =   42
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox summa 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69480
         TabIndex        =   41
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox period 
         Height          =   375
         Left            =   -69600
         TabIndex        =   39
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox mesto_raboti 
         Height          =   375
         Left            =   -73320
         TabIndex        =   37
         Top             =   3120
         Width           =   8055
      End
      Begin VB.TextBox mesto_rozhdeniya 
         Height          =   405
         Left            =   -72360
         TabIndex        =   35
         Top             =   2640
         Width           =   7095
      End
      Begin VB.TextBox data_rozhdeniya 
         Height          =   375
         Left            =   -72360
         TabIndex        =   33
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox identificator 
         Height          =   375
         Left            =   -73320
         TabIndex        =   31
         Top             =   1680
         Width           =   8055
      End
      Begin VB.TextBox mesto_zhitelstva 
         Height          =   375
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   9615
      End
      Begin VB.TextBox fio_dolzhnika 
         Height          =   375
         Left            =   -73080
         TabIndex        =   27
         Top             =   480
         Width           =   7815
      End
      Begin VB.TextBox l_s 
         Height          =   375
         Left            =   -73320
         TabIndex        =   25
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox ogrn 
         Height          =   375
         Left            =   -73320
         TabIndex        =   23
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox okpo 
         Height          =   375
         Left            =   -73320
         TabIndex        =   21
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox kpp 
         Height          =   375
         Left            =   -73320
         TabIndex        =   19
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox inn 
         Height          =   375
         Left            =   -73320
         TabIndex        =   17
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox kor_schet 
         Height          =   375
         Left            =   -73320
         TabIndex        =   15
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox vz_mesto_nahozhdeniya 
         Height          =   375
         Left            =   -73320
         TabIndex        =   12
         Top             =   960
         Width           =   8175
      End
      Begin VB.TextBox vz_naimenovanie 
         Height          =   375
         Left            =   -73320
         TabIndex        =   10
         Top             =   480
         Width           =   8175
      End
      Begin VB.Label Label31 
         Caption         =   "Счет получателя"
         Height          =   375
         Left            =   -70920
         TabIndex        =   70
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "КБК"
         Height          =   255
         Left            =   -70560
         TabIndex        =   65
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "БИК"
         Height          =   255
         Left            =   -70560
         TabIndex        =   64
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "ОКТМО"
         Height          =   255
         Left            =   -70800
         TabIndex        =   63
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "ОКАТО"
         Height          =   255
         Left            =   -70800
         TabIndex        =   62
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Сумма"
         Height          =   375
         Left            =   -70200
         TabIndex        =   57
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Вид расходов"
         Height          =   375
         Left            =   -70680
         TabIndex        =   56
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Взыскания:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label22 
         Caption         =   "Период, за который образовалась взыскиваемая задолженность"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label21 
         Caption         =   "Место работы:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Место рождения / регистрации"
         Height          =   375
         Left            =   -74880
         TabIndex        =   34
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "Дата рождения / регистрации"
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Идентификатор"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Место жительства или место пребывания:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label16 
         Caption         =   "Наименование / ФИО"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Лицевой счет:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "ОГРН"
         Height          =   255
         Left            =   -73920
         TabIndex        =   22
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "ОКПО"
         Height          =   255
         Left            =   -73920
         TabIndex        =   20
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "КПП"
         Height          =   255
         Left            =   -73920
         TabIndex        =   18
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "ИНН"
         Height          =   375
         Left            =   -73920
         TabIndex        =   16
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Корсчет"
         Height          =   255
         Left            =   -74160
         TabIndex        =   14
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Реквизиты банковского счета:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Место нахождения"
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Наименование:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Мировой судья"
         Height          =   255
         Left            =   -74160
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Сайт"
         Height          =   375
         Left            =   -73440
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Почтовый адрес"
         Height          =   255
         Left            =   -74280
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Полное наименование"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox proizvodstvo_nomer 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox data_vineseniya 
      Height          =   405
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Производство №"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Дата вынесения"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "SudebniyPrikaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dictSud As New Collection
Dim sD As Sud
Dim dictSudya As New Collection

Dim dictVD As New Collection
Dim vd As Vziskaniya
Dim idxVD As Integer
Dim vdEdit As Boolean

Const docDir As String = "D:\смэв\судебные_приказы\"
Const dd1 As String = "D:\"
Const dd2 As String = "D:\смэв\"

Const templateFileName = "template"

Private Sub address_Click()
      Dim i As Integer
      Dim selectedSD As Sud
      
      If (address.Text = "") Then
        mirovoy_sudiya.Clear
          For Each ds In dictSudya
            mirovoy_sudiya.AddItem ds
  Next ds
      Else
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
          If (selectedSD.address = address.Text) Then
              polnoe_naimenovanie.Text = selectedSD.name
              mirovoy_sudiya.Clear
              For j = 1 To selectedSD.sudyaName.Count
                mirovoy_sudiya.AddItem selectedSD.sudyaName(j)
              Next j
              
              If mirovoy_sudiya.ListCount = 1 Then
                mirovoy_sudiya.Text = selectedSD.sudyaName(j - 1)
                Else
                mirovoy_sudiya.Text = ""
              End If
              Exit For
          End If
      Next i
      site.Text = selectedSD.site
      End If
End Sub

Private Sub Command1_Click()
    If (summa.Text <> "") Then
        summa_list.AddItem summa.Text
        summa.Text = ""
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Integer
          For i = 0 To summa_list.ListCount - 1
          If summa_list.Selected(i) Then
              summa_list.RemoveItem (i)
              Exit For
          End If
      Next i
End Sub

Private Sub AddV_Click()
    SaveV.Visible = True
    UndoV.Visible = True
    
    AddV.Visible = False
    EditV.Visible = False
    DelV.Visible = False
    
    Set vd = New Vziskaniya
    
    VidV.Text = ""
    summa.Text = ""
    
    enableV
End Sub

Private Sub enableV()
    VidV.Enabled = True
    summa.Enabled = True
End Sub

Private Sub disableV()
    VidV.Enabled = False
    summa.Enabled = False
End Sub

Private Sub Command3_Click()
 Save templateFileName
End Sub

Private Sub Command4_Click()
    If (statiya.Text <> "") Then
        statiy_list.AddItem statiya.Text
        statiya.Text = ""
    End If
End Sub

Private Sub Command5_Click()
    Dim i As Integer
          For i = 0 To statiy_list.ListCount - 1
          If statiy_list.Selected(i) Then
              statiy_list.RemoveItem (i)
              Exit For
          End If
      Next i
End Sub

Private Sub CommandButtonGenerate_Click()
    filename = proizvodstvo_nomer.Text
    filename = Replace(filename, "-", "_")
    filename = Replace(filename, "\", "_")
    filename = Replace(filename, "/", "_")
    filename = Replace(filename, ".", "_")
    Save (filename)
End Sub
Private Sub Save(ByVal filename As String)

  filename1 = docDir + "sp_" + filename + ".xml"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set out = fso.CreateTextFile(filename1, True, True)
  
  out.WriteLine ("<?xml version='1.0'?>")
  out.WriteLine ("<sudebniy_prikaz>")
  
    out.WriteLine ("<polnoe_naimenovanie>")
        out.WriteLine (polnoe_naimenovanie.Text)
    out.WriteLine ("</polnoe_naimenovanie>")
    
    out.WriteLine ("<address>")
        out.WriteLine (address.Text)
    out.WriteLine ("</address>")

    out.WriteLine ("<site>")
        out.WriteLine (site.Text)
    out.WriteLine ("</site>")
    
    out.WriteLine ("<data_vineseniya>")
     out.WriteLine (data_vineseniya.Text)
    out.WriteLine ("</data_vineseniya>")
    
    out.WriteLine ("<proizvodstvo_nomer>")
        out.WriteLine (proizvodstvo_nomer.Text)
    out.WriteLine ("</proizvodstvo_nomer>")
    
    out.WriteLine ("<mirovoy_sudiya>")
        out.WriteLine (mirovoy_sudiya.Text)
    out.WriteLine ("</mirovoy_sudiya>")
    
    out.WriteLine ("<vz_naimenovanie>")
        out.WriteLine (vz_naimenovanie.Text)
    out.WriteLine ("</vz_naimenovanie>")

    out.WriteLine ("<vz_mesto_nahozhdeniya>")
        out.WriteLine (vz_mesto_nahozhdeniya.Text)
    out.WriteLine ("</vz_mesto_nahozhdeniya>")
    
    out.WriteLine ("<kor_schet>")
        out.WriteLine (kor_schet.Text)
    out.WriteLine ("</kor_schet>")

    out.WriteLine ("<inn>")
        out.WriteLine (inn.Text)
    out.WriteLine ("</inn>")
    
    out.WriteLine ("<kpp>")
        out.WriteLine (kpp.Text)
    out.WriteLine ("</kpp>")

    out.WriteLine ("<okpo>")
        out.WriteLine (okpo.Text)
    out.WriteLine ("</okpo>")
    
    out.WriteLine ("<ogrn>")
        out.WriteLine (ogrn.Text)
    out.WriteLine ("</ogrn>")
    
    out.WriteLine ("<l_s>")
        out.WriteLine (l_s.Text)
    out.WriteLine ("</l_s>")

    out.WriteLine ("<fio_dolzhnika>")
        out.WriteLine (fio_dolzhnika.Text)
    out.WriteLine ("</fio_dolzhnika>")
    
    out.WriteLine ("<mesto_zhitelstva>")
        out.WriteLine (mesto_zhitelstva.Text)
    out.WriteLine ("</mesto_zhitelstva>")
    
    out.WriteLine ("<identificator>")
        out.WriteLine (identificator.Text)
    out.WriteLine ("</identificator>")
    
    out.WriteLine ("<data_rozhdeniya>")
        out.WriteLine (data_rozhdeniya.Text)
    out.WriteLine ("</data_rozhdeniya>")

    out.WriteLine ("<mesto_rozhdeniya>")
        out.WriteLine (mesto_rozhdeniya.Text)
    out.WriteLine ("</mesto_rozhdeniya>")

    out.WriteLine ("<mesto_raboti>")
        out.WriteLine (mesto_raboti.Text)
    out.WriteLine ("</mesto_raboti>")
    
    out.WriteLine ("<period>")
        out.WriteLine (period.Text)
    out.WriteLine ("</period>")
    
    out.WriteLine ("<period_razmernost>")
        out.WriteLine (period_razmernost.Text)
    out.WriteLine ("</period_razmernost>")
    
    out.WriteLine ("<OKATO>")
        out.WriteLine (OKATO.Text)
    out.WriteLine ("</OKATO>")
    
    out.WriteLine ("<OKTMO>")
        out.WriteLine (OKTMO.Text)
    out.WriteLine ("</OKTMO>")
    
    out.WriteLine ("<BIK>")
        out.WriteLine (BIK.Text)
    out.WriteLine ("</BIK>")
    
    out.WriteLine ("<KBK>")
        out.WriteLine (KBK.Text)
    out.WriteLine ("</KBK>")
    
    out.WriteLine ("<SchetPoluchatelya>")
        out.WriteLine (SchetPoluchatelya.Text)
    out.WriteLine ("</SchetPoluchatelya>")
    
    out.WriteLine ("<vziskaniya_list>")
        For Each dVD In dictVD
            out.WriteLine ("<vziskanie>")
                out.WriteLine ("<vid>")
                    out.WriteLine (dVD.VidRashodov)
                out.WriteLine ("</vid>")
                
                out.WriteLine ("<summa>")
                    out.WriteLine (dVD.summa)
                out.WriteLine ("</summa>")
            out.WriteLine ("</vziskanie>")
        Next
    out.WriteLine ("</vziskaniya_list>")

    
    
  out.WriteLine ("</sudebniy_prikaz>")

  out.Close
  log.Text = log.Text + vbNewLine + "Документ сгенерирован:" + filename1
  CommandButtonSign.Enabled = True
 


End Sub

Private Sub CommandButtonOpen_Click()
CommonDialog1.Filter = "Суд.приказ (*.xml)|*.xml|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.InitDir = "D:\смэв"
CommonDialog1.ShowOpen

openDoc CommonDialog1.filename

End Sub

Private Sub CommandButtonSign_Click()
    CommandButtonSend.Enabled = True

End Sub

Private Sub CommandButtonValidate_Click()
    If ((data_vineseniya.Text = "") Or (proizvodstvo_nomer.Text = "")) Then
        log.Text = "Проверка не прошла"
        CommandButtonGenerate.Enabled = False
    Else
        log.Text = "Проверка прошла"
        CommandButtonGenerate.Enabled = True
    End If


End Sub

Private Sub DelV_Click()
    SaveV.Visible = False
    UndoV.Visible = False
    
        Dim i As Integer
          For i = 0 To summa_list.ListCount - 1
          If summa_list.Selected(i) Then
              Set vd = dictVD(i + 1)
              Exit For
          End If
      Next i
      

      If i + 1 < dictVD.Count Then
       For j = i + 1 To dictVD.Count - 1
        Set vd1 = dictVD(j + 1)

        dictVD(j).VidRashodov = vd1.VidRashodov
        dictVD(j).summa = vd1.summa

       Next j
      End If
      
      summa_list.RemoveItem (i)
      
      If summa_list.ListCount > 0 Then
        dictVD.Remove (dictVD.Count)
        summa_list.Selected(0) = True
      Else

    VidV.Text = ""
    summa.Text = ""


      End If
      
      disableV
End Sub

Private Sub EditV_Click()
    SaveV.Visible = True
    UndoV.Visible = True
    
    AddV.Visible = False
    EditV.Visible = False
    DelV.Visible = False
    vdEdit = True
    enableV
End Sub

Private Sub Form_Load()
    summa_list.Clear
    disableV
    readSud
    initCheck
    openDoc docDir + "sp_" + templateFileName + ".xml"
End Sub

Sub openDoc(ByVal filename As String)
    Dim XDoc As Object
    On Error GoTo error_open_doc
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (filename)
    
    summa_list.Clear
    
    'Get Document Elements
    Set Lists = XDoc.DocumentElement
    
    For Each ListNode In Lists.ChildNodes
        'Debug.Print "----" & listNode.BaseName & "----" & listNode.Text
        Select Case ListNode.BaseName
                Case "polnoe_naimenovanie"
                    polnoe_naimenovanie.Text = ListNode.Text
                Case "address"
                    address.Text = ListNode.Text
                Case "site"
                    site.Text = ListNode.Text
                Case "data_vineseniya"
                    data_vineseniya.Text = ListNode.Text
                Case "proizvodstvo_nomer"
                    proizvodstvo_nomer.Text = ListNode.Text
                Case "mirovoy_sudiya"
                    mirovoy_sudiya.Text = ListNode.Text
                Case "vz_naimenovanie"
                    vz_naimenovanie.Text = ListNode.Text
                Case "vz_mesto_nahozhdeniya"
                    vz_mesto_nahozhdeniya.Text = ListNode.Text
                Case "kor_schet"
                    kor_schet.Text = ListNode.Text
                Case "inn"
                    inn.Text = ListNode.Text
                Case "kpp"
                    kpp.Text = ListNode.Text
                Case "okpo"
                    okpo.Text = ListNode.Text
                Case "ogrn"
                    ogrn.Text = ListNode.Text
                Case "l_s"
                    l_s.Text = ListNode.Text
                Case "fio_dolzhnika"
                    fio_dolzhnika.Text = ListNode.Text
                Case "mesto_zhitelstva"
                    mesto_zhitelstva.Text = ListNode.Text
                Case "identificator"
                    identificator.Text = ListNode.Text
                Case "data_rozhdeniya"
                    data_rozhdeniya.Text = ListNode.Text
                Case "mesto_rozhdeniya"
                    mesto_rozhdeniya.Text = ListNode.Text
                Case "mesto_raboti"
                    mesto_raboti.Text = ListNode.Text
                Case "period"
                    period.Text = ListNode.Text
                Case "period_razmernost"
                    period_razmernost.Text = ListNode.Text
                Case "rashodi_gosposhliny"
                    rashodi_gosposhliny.Text = ListNode.Text
                    
                Case "OKATO"
                    OKATO.Text = ListNode.Text
                Case "OKTMO"
                    OKTMO.Text = ListNode.Text
                Case "BIK"
                    BIK.Text = ListNode.Text
                Case "KBK"
                    KBK.Text = ListNode.Text
                Case "SchetPoluchatelya"
                    SchetPoluchatelya.Text = ListNode.Text
                    
                Case "vziskaniya_list"
                    For Each ListNode1 In ListNode.ChildNodes
                        Set vd = New Vziskaniya
                        For Each ListNode2 In ListNode1.ChildNodes
                            Select Case ListNode2.BaseName
                                Case "summa"
                                    vd.summa = ListNode2.Text
                                Case "vid"
                                    vd.VidRashodov = ListNode2.Text
                            End Select
                        Next ListNode2
                        
                        dictVD.Add vd
                        summa_list.AddItem vd.summa
                        
                    Next ListNode1
        End Select
    Next ListNode
        
    log.Text = "Открыто судебное решение: " & filename
error_open_doc:
End Sub

Private Sub initCheck()
    
    Dim sFilePath As String
 
    If Dir(dd1, vbDirectory) = "" Then
          MsgBox "Отсутствует диск D:\ - обратитесь в службу поддержки"
          GoTo finishInitCheck
    End If
    
    If Dir(dd2, vbDirectory) = "" Then
          MkDir dd2
    End If
    
    If Dir(docDir, vbDirectory) = "" Then
          MkDir docDir
    End If
        
    
    sFilePath = docDir + "sp_" + templateFileName + ".xml"
    
    If Dir(sFilePath) = "" Then
        Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set out1 = fso1.CreateTextFile(sFilePath, True, True)
        out1.Close
    End If

finishInitCheck:
End Sub

Private Sub readSud()
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim the_array() As String
Dim R As Long
Dim C As Long
Dim sName As String
Dim sudyaName As String
Dim saddr As String

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = _
        file_name & "\"
    file_name = file_name & "data.csv"

    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ";")
    num_cols = UBound(one_line)

sName = ""
sudyaName = ""
saddr = ""
    ' Copy the data into the array.
    For R = 1 To num_rows
    
        If Len(lines(R)) > 0 Then
            one_line = Split(lines(R), ";")
            For C = 0 To num_cols
                Select Case C
                    Case 1
                        If (one_line(C) <> sName) Then
                             If sName <> "" Then
                                dictSud.Add sD
                             End If
                             
                             Set sD = New Sud
                             sName = one_line(C)
                             sD.name = sName
                             polnoe_naimenovanie.AddItem sName
                        End If
                    Case 2
                        If (sD.address = "") Then
                            sD.address = one_line(C)
                            If (sD.address <> saddr) Then
                                saddr = sD.address
                                address.AddItem saddr
                                
                            End If
                        End If
                    Case 3
                        If (one_line(C) <> "") Then
                            sudyaName = one_line(C)
                            dictSudya.Add sudyaName
                            mirovoy_sudiya.AddItem sudyaName
                        End If
                        sD.sudyaName.Add sudyaName
                    Case 7
                        If (one_line(C) <> "") Then
                            sD.site = one_line(C)
                            site.AddItem sD.site
                        End If
                    End Select
            
            Next C
        End If
    Next R


End Sub

Private Sub mirovoy_sudiya_Click()
      Dim i As Integer
      Dim selectedSD As Sud
      Dim flg As Boolean
      If (mirovoy_sudiya.Text <> "") Then

        flg = False
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
            
              For j = 1 To selectedSD.sudyaName.Count
                If selectedSD.sudyaName(j) = mirovoy_sudiya.Text Then
                    flg = True
                    Exit For
                End If
              Next j
        If flg Then
            Exit For
        End If
      Next i
      
      If flg Then
        polnoe_naimenovanie.Text = selectedSD.name
        address.Text = selectedSD.address
        site.Text = selectedSD.site
      End If
      
      End If
End Sub

Private Sub polnoe_naimenovanie_Click()
      Dim i As Integer
      Dim selectedSD As Sud
      
      If (polnoe_naimenovanie.Text = "") Then
        mirovoy_sudiya.Clear
          For Each ds In dictSudya
            mirovoy_sudiya.AddItem ds
  Next ds
      Else
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
          If (selectedSD.name = polnoe_naimenovanie.Text) Then
              address.Text = selectedSD.address
              mirovoy_sudiya.Clear
              For j = 1 To selectedSD.sudyaName.Count
                mirovoy_sudiya.AddItem selectedSD.sudyaName(j)
              Next j
              
              If mirovoy_sudiya.ListCount = 1 Then
                mirovoy_sudiya.Text = selectedSD.sudyaName(j - 1)
                Else
                mirovoy_sudiya.Text = ""
              End If
              Exit For
          End If
      Next i
      site.Text = selectedSD.site
      End If
End Sub

Private Sub SaveV_Click()
    Dim i As Integer
    If (vdEdit) Then
      For i = 0 To summa_list.ListCount - 1
          If summa_list.Selected(i) Then
              Set vd = dictVD(i + 1)
              Exit For
          End If
      Next i
        vd.VidRashodov = VidV.Text
        vd.summa = summa.Text
        
        vdEdit = False
    Else
        vd.VidRashodov = VidV.Text
        vd.summa = summa.Text
    
    
    summa_list.AddItem vd.summa
    dictVD.Add vd
    idxVD = idxVD + 1
    End If
    
    SaveV.Visible = False
    UndoV.Visible = False
    
    AddV.Visible = True
    EditV.Visible = True
    DelV.Visible = True
    disableV
End Sub

Private Sub site_Click()
      Dim i As Integer
      Dim selectedSD As Sud
      
      If (site.Text = "") Then
        mirovoy_sudiya.Clear
          For Each ds In dictSudya
            mirovoy_sudiya.AddItem ds
  Next ds
      Else
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
          If (selectedSD.site = site.Text) Then
              address.Text = selectedSD.address
              polnoe_naimenovanie.Text = selectedSD.name
              mirovoy_sudiya.Clear
              For j = 1 To selectedSD.sudyaName.Count
                mirovoy_sudiya.AddItem selectedSD.sudyaName(j)
              Next j
              
              If mirovoy_sudiya.ListCount = 1 Then
                mirovoy_sudiya.Text = selectedSD.sudyaName(j - 1)
                Else
                mirovoy_sudiya.Text = ""
              End If
              Exit For
          End If
      Next i

      End If
End Sub

Private Sub summa_list_Click()
    su_change
End Sub


Private Sub su_change()
      Dim i As Integer
      Dim selectedVD As Vziskaniya
      
      If summa_list.ListCount > 0 Then
      
      For i = 0 To summa_list.ListCount - 1
          If summa_list.Selected(i) Then
              Set selectedVD = dictVD(i + 1)
              Exit For
          End If
      Next i
       
        VidV.Text = selectedVD.VidRashodov
        summa.Text = selectedVD.summa
      End If
End Sub

Private Sub summa_list_KeyPress(KeyAscii As Integer)
    su_change
End Sub
