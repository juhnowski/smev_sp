VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SmevIspDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������������� ��������"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "�����"
      TabPicture(0)   =   "IspDoc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(12)=   "Label23"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "srok_predyavleniya_k_ispolneniyu_znachenie"
      Tab(0).Control(15)=   "srok_predyavleniya_k_ispolneniyu_razmernost"
      Tab(0).Control(16)=   "nomer_ekz_ID"
      Tab(0).Control(17)=   "ob_data_vidachi"
      Tab(0).Control(18)=   "data_sudebnogo_acta"
      Tab(0).Control(19)=   "dublicat"
      Tab(0).Control(20)=   "vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti"
      Tab(0).Control(21)=   "data_vsupleniya_v_zs"
      Tab(0).Control(22)=   "podl_nemedl_isp"
      Tab(0).Control(23)=   "summa_dolga"
      Tab(0).Control(24)=   "valyuta_dolga"
      Tab(0).Control(25)=   "FIO_sudiy"
      Tab(0).Control(26)=   "vid_sushnosti_ispolneniya_ID"
      Tab(0).Control(27)=   "SolidarnoeVziskanie"
      Tab(0).Control(28)=   "Gosposhlina"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "�������� ���"
      TabPicture(1)   =   "IspDoc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "ustanovochnaya_chast_sudebnogo_acta"
      Tab(1).Control(3)=   "rezolyutativnaya_chast_sudebnogo_acta"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "����������"
      TabPicture(2)   =   "IspDoc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "�������"
      TabPicture(3)   =   "IspDoc.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "SSTab3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "���������"
      TabPicture(4)   =   "IspDoc.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label76"
      Tab(4).Control(1)=   "Label77"
      Tab(4).Control(2)=   "Label78"
      Tab(4).Control(3)=   "mesto_rassmotreniya_dela"
      Tab(4).Control(4)=   "naimenovanie_suda_vidayushego_ispolnitelniy_document"
      Tab(4).Control(5)=   "adres_suda_vidayushego_ispolnitelniy_document"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "�������"
      TabPicture(5)   =   "IspDoc.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label79"
      Tab(5).Control(1)=   "Label80"
      Tab(5).Control(2)=   "SignatureValue"
      Tab(5).Control(3)=   "X509Certificate"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "��������"
      TabPicture(6)   =   "IspDoc.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command1"
      Tab(6).Control(1)=   "log"
      Tab(6).Control(2)=   "CommandButtonOpen"
      Tab(6).Control(3)=   "CommandButtonSend"
      Tab(6).Control(4)=   "CommandButtonSign"
      Tab(6).Control(5)=   "CommandButtonGenerate"
      Tab(6).Control(6)=   "CommandButtonValidate"
      Tab(6).ControlCount=   7
      Begin VB.TextBox Gosposhlina 
         Height          =   405
         Left            =   -67080
         TabIndex        =   201
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ComboBox SolidarnoeVziskanie 
         Height          =   315
         ItemData        =   "IspDoc.frx":00C4
         Left            =   -66960
         List            =   "IspDoc.frx":00CE
         TabIndex        =   199
         Text            =   "���"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox vid_sushnosti_ispolneniya_ID 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   197
         Top             =   2760
         Width           =   9015
      End
      Begin VB.ComboBox adres_suda_vidayushego_ispolnitelniy_document 
         Height          =   315
         Left            =   -74760
         TabIndex        =   194
         Top             =   1680
         Width           =   8775
      End
      Begin VB.ComboBox naimenovanie_suda_vidayushego_ispolnitelniy_document 
         Height          =   315
         Left            =   -74760
         TabIndex        =   193
         Top             =   840
         Width           =   8775
      End
      Begin VB.ComboBox FIO_sudiy 
         Height          =   315
         Left            =   -72960
         TabIndex        =   192
         Top             =   2040
         Width           =   6855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��������� ��� ������"
         Height          =   375
         Left            =   -68640
         TabIndex        =   180
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox log 
         Height          =   4095
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   179
         Top             =   1200
         Width           =   8655
      End
      Begin VB.CommandButton CommandButtonOpen 
         Caption         =   "�������"
         Height          =   375
         Left            =   -74640
         TabIndex        =   178
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CommandButtonSend 
         Caption         =   "���������"
         Height          =   375
         Left            =   -69120
         TabIndex        =   177
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CommandButtonSign 
         Caption         =   "���������"
         Height          =   375
         Left            =   -69600
         TabIndex        =   176
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CommandButtonGenerate 
         Caption         =   "�������������"
         Height          =   375
         Left            =   -71280
         TabIndex        =   175
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CommandButtonValidate 
         Caption         =   "���������"
         Height          =   375
         Left            =   -72960
         TabIndex        =   174
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox X509Certificate 
         Height          =   3855
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   173
         Top             =   1560
         Width           =   9015
      End
      Begin VB.TextBox SignatureValue 
         Height          =   375
         Left            =   -74880
         TabIndex        =   171
         Top             =   720
         Width           =   9015
      End
      Begin VB.TextBox mesto_rassmotreniya_dela 
         Height          =   375
         Left            =   -74760
         TabIndex        =   169
         Top             =   2520
         Width           =   8775
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   5055
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "����� ��������"
         TabPicture(0)   =   "IspDoc.frx":00DB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label39"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label40"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label41"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label42"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label43"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label44"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label45"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dolzhnik_status_lica"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "dolzhnik_dolzhnik"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "dolzhnik_adres"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "dolzhnik_kpp"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "dolzhnik_ogrn"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "dolzhnik_data_registracii"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "strana_grazhdanstva_ili_registracii"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Frame1"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "���������"
         TabPicture(1)   =   "IspDoc.frx":00F7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label46"
         Tab(1).Control(1)=   "Label47"
         Tab(1).Control(2)=   "Label48"
         Tab(1).Control(3)=   "Label49"
         Tab(1).Control(4)=   "Label50"
         Tab(1).Control(5)=   "Label51"
         Tab(1).Control(6)=   "Label52"
         Tab(1).Control(7)=   "Label53"
         Tab(1).Control(8)=   "Label54"
         Tab(1).Control(9)=   "UdostDocument"
         Tab(1).Control(10)=   "vid"
         Tab(1).Control(11)=   "seriya"
         Tab(1).Control(12)=   "nomer"
         Tab(1).Control(13)=   "fio"
         Tab(1).Control(14)=   "data_rozhdeniya"
         Tab(1).Control(15)=   "data_vidachi"
         Tab(1).Control(16)=   "kod_podrazdeleniya"
         Tab(1).Control(17)=   "mesto_rozhdeniya"
         Tab(1).Control(18)=   "addUdostDocument"
         Tab(1).Control(19)=   "EditUdostDocument"
         Tab(1).Control(20)=   "DeleteUdostDocument"
         Tab(1).Control(21)=   "SaveUdostDocument"
         Tab(1).Control(22)=   "UndoUdostDocument"
         Tab(1).Control(23)=   "pol"
         Tab(1).ControlCount=   24
         TabCaption(2)   =   "������������"
         TabPicture(2)   =   "IspDoc.frx":0113
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label55"
         Tab(2).Control(1)=   "Label56"
         Tab(2).Control(2)=   "Label57"
         Tab(2).Control(3)=   "Label58"
         Tab(2).Control(4)=   "Label59"
         Tab(2).Control(5)=   "Label60"
         Tab(2).Control(6)=   "Label61"
         Tab(2).Control(7)=   "Nedvizhimost"
         Tab(2).Control(8)=   "Actualnost"
         Tab(2).Control(9)=   "Naimenovanie"
         Tab(2).Control(10)=   "Ploshad"
         Tab(2).Control(11)=   "UslNomer"
         Tab(2).Control(12)=   "InvNomer"
         Tab(2).Control(13)=   "KadastrNomer"
         Tab(2).Control(14)=   "TochAdres"
         Tab(2).Control(15)=   "AddNedvizhimost"
         Tab(2).Control(16)=   "EditNedvizhimost"
         Tab(2).Control(17)=   "DeleteNedvizhimost"
         Tab(2).Control(18)=   "SaveNedvizhimost"
         Tab(2).Control(19)=   "UndoNedvizhimost"
         Tab(2).ControlCount=   20
         TabCaption(3)   =   "����� ������"
         TabPicture(3)   =   "IspDoc.frx":012F
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label62"
         Tab(3).Control(1)=   "Label63"
         Tab(3).Control(2)=   "Label64"
         Tab(3).Control(3)=   "Label65"
         Tab(3).Control(4)=   "mr_actualnost"
         Tab(3).Control(5)=   "naimenovanie_organizacii_fio_ip"
         Tab(3).Control(6)=   "jur_address"
         Tab(3).Control(7)=   "fact_address"
         Tab(3).ControlCount=   8
         TabCaption(4)   =   "���������"
         TabPicture(4)   =   "IspDoc.frx":014B
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label66"
         Tab(4).Control(1)=   "Label67"
         Tab(4).Control(2)=   "Label68"
         Tab(4).Control(3)=   "Label69"
         Tab(4).Control(4)=   "Label70"
         Tab(4).Control(5)=   "Label71"
         Tab(4).Control(6)=   "Label72"
         Tab(4).Control(7)=   "Label73"
         Tab(4).Control(8)=   "Label74"
         Tab(4).Control(9)=   "Label75"
         Tab(4).Control(10)=   "TransSredstva"
         Tab(4).Control(11)=   "TS_Actualnost"
         Tab(4).Control(12)=   "Kategoriya"
         Tab(4).Control(13)=   "Marka"
         Tab(4).Control(14)=   "Model"
         Tab(4).Control(15)=   "Cvet"
         Tab(4).Control(16)=   "GosZnak"
         Tab(4).Control(17)=   "VIN"
         Tab(4).Control(18)=   "NDvig"
         Tab(4).Control(19)=   "KodPodr"
         Tab(4).Control(20)=   "GodVipuska"
         Tab(4).Control(21)=   "AddTransSredstva"
         Tab(4).Control(22)=   "EditTransSredstva"
         Tab(4).Control(23)=   "DeleteTransSredstva"
         Tab(4).Control(24)=   "SaveTransSredstva"
         Tab(4).Control(25)=   "UndoTransSredstva"
         Tab(4).ControlCount=   26
         Begin VB.ComboBox pol 
            Height          =   315
            ItemData        =   "IspDoc.frx":0167
            Left            =   -67920
            List            =   "IspDoc.frx":0171
            TabIndex        =   195
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Frame Frame1 
            Caption         =   "��� ���.���"
            Height          =   1935
            Left            =   2400
            TabIndex        =   181
            Top             =   1920
            Width           =   6495
            Begin VB.TextBox dolzhnik_mesto_rozhdeniya 
               Height          =   375
               Left            =   120
               TabIndex        =   191
               Top             =   1440
               Width           =   6255
            End
            Begin VB.TextBox dolzhnik_snils 
               Height          =   375
               Left            =   3480
               TabIndex        =   189
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox dolzhnik_inn 
               Height          =   375
               Left            =   840
               TabIndex        =   187
               Top             =   840
               Width           =   1695
            End
            Begin VB.TextBox dolzhnik_data_rozhdeniya 
               Height          =   405
               Left            =   3120
               TabIndex        =   185
               Top             =   360
               Width           =   1215
            End
            Begin VB.ComboBox dolzhnik_pol 
               Height          =   315
               ItemData        =   "IspDoc.frx":017B
               Left            =   840
               List            =   "IspDoc.frx":0185
               TabIndex        =   183
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label85 
               Caption         =   "����� ��������"
               Height          =   255
               Left            =   120
               TabIndex        =   190
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label84 
               Caption         =   "�����"
               Height          =   255
               Left            =   2760
               TabIndex        =   188
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label83 
               Caption         =   "���"
               Height          =   255
               Left            =   120
               TabIndex        =   186
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label82 
               Caption         =   "���� ��������:"
               Height          =   375
               Left            =   1800
               TabIndex        =   184
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label81 
               Caption         =   "���"
               Height          =   375
               Left            =   120
               TabIndex        =   182
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.CommandButton UndoTransSredstva 
            Caption         =   "��������"
            Height          =   375
            Left            =   -67800
            TabIndex        =   165
            Top             =   4440
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton SaveTransSredstva 
            Caption         =   "���������"
            Height          =   375
            Left            =   -69480
            TabIndex        =   164
            Top             =   4440
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton DeleteTransSredstva 
            Caption         =   "�������"
            Height          =   375
            Left            =   -71640
            TabIndex        =   163
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton EditTransSredstva 
            Caption         =   "�������������"
            Height          =   375
            Left            =   -73200
            TabIndex        =   162
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton AddTransSredstva 
            Caption         =   "��������"
            Height          =   375
            Left            =   -74880
            TabIndex        =   161
            Top             =   4440
            Width           =   1575
         End
         Begin VB.TextBox GodVipuska 
            Height          =   375
            Left            =   -67680
            TabIndex        =   160
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox KodPodr 
            Height          =   375
            Left            =   -70560
            TabIndex        =   158
            Top             =   3840
            Width           =   1575
         End
         Begin VB.TextBox NDvig 
            Height          =   375
            Left            =   -68520
            TabIndex        =   156
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox VIN 
            Height          =   375
            Left            =   -71520
            TabIndex        =   154
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox GosZnak 
            Height          =   375
            Left            =   -69960
            TabIndex        =   152
            Top             =   2880
            Width           =   3735
         End
         Begin VB.TextBox Cvet 
            Height          =   375
            Left            =   -69960
            TabIndex        =   150
            Top             =   2400
            Width           =   3735
         End
         Begin VB.TextBox Model 
            Height          =   375
            Left            =   -69960
            TabIndex        =   148
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox Marka 
            Height          =   375
            Left            =   -69960
            TabIndex        =   146
            Top             =   1440
            Width           =   3735
         End
         Begin VB.TextBox Kategoriya 
            Height          =   375
            Left            =   -69960
            TabIndex        =   144
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox TS_Actualnost 
            Height          =   375
            Left            =   -69960
            TabIndex        =   142
            Top             =   480
            Width           =   1815
         End
         Begin VB.ListBox TransSredstva 
            Height          =   3765
            Left            =   -74880
            TabIndex        =   140
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox fact_address 
            Height          =   375
            Left            =   -74760
            TabIndex        =   139
            Top             =   3240
            Width           =   8535
         End
         Begin VB.TextBox jur_address 
            Height          =   375
            Left            =   -74760
            TabIndex        =   137
            Top             =   2280
            Width           =   8535
         End
         Begin VB.TextBox naimenovanie_organizacii_fio_ip 
            Height          =   375
            Left            =   -74760
            TabIndex        =   135
            Top             =   1320
            Width           =   8535
         End
         Begin VB.TextBox mr_actualnost 
            Height          =   375
            Left            =   -72120
            TabIndex        =   133
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton UndoNedvizhimost 
            Caption         =   "��������"
            Height          =   375
            Left            =   -67680
            TabIndex        =   131
            Top             =   4560
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton SaveNedvizhimost 
            Caption         =   "���������"
            Height          =   375
            Left            =   -69360
            TabIndex        =   130
            Top             =   4560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton DeleteNedvizhimost 
            Caption         =   "�������"
            Height          =   375
            Left            =   -72120
            TabIndex        =   129
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton EditNedvizhimost 
            Caption         =   "�������������"
            Height          =   375
            Left            =   -73560
            TabIndex        =   128
            Top             =   4560
            Width           =   1335
         End
         Begin VB.CommandButton AddNedvizhimost 
            Caption         =   "��������"
            Height          =   375
            Left            =   -74880
            TabIndex        =   127
            Top             =   4560
            Width           =   1215
         End
         Begin VB.TextBox TochAdres 
            Height          =   735
            Left            =   -71640
            TabIndex        =   126
            Top             =   3720
            Width           =   5295
         End
         Begin VB.TextBox KadastrNomer 
            Height          =   375
            Left            =   -69840
            TabIndex        =   124
            Top             =   2880
            Width           =   3495
         End
         Begin VB.TextBox InvNomer 
            Height          =   375
            Left            =   -69840
            TabIndex        =   122
            Top             =   2400
            Width           =   3495
         End
         Begin VB.TextBox UslNomer 
            Height          =   375
            Left            =   -69840
            TabIndex        =   120
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox Ploshad 
            Height          =   375
            Left            =   -69840
            TabIndex        =   118
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Naimenovanie 
            Height          =   375
            Left            =   -69840
            TabIndex        =   116
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox Actualnost 
            Height          =   375
            Left            =   -69840
            TabIndex        =   114
            Top             =   480
            Width           =   2055
         End
         Begin VB.ListBox Nedvizhimost 
            Height          =   3960
            Left            =   -74880
            TabIndex        =   112
            Top             =   480
            Width           =   3015
         End
         Begin VB.CommandButton UndoUdostDocument 
            Caption         =   "��������"
            Height          =   375
            Left            =   -67920
            TabIndex        =   111
            Top             =   4560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton SaveUdostDocument 
            Caption         =   "���������"
            Height          =   375
            Left            =   -69600
            TabIndex        =   110
            Top             =   4560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton DeleteUdostDocument 
            Caption         =   "�������"
            Height          =   375
            Left            =   -71520
            TabIndex        =   109
            Top             =   4560
            Width           =   1695
         End
         Begin VB.CommandButton EditUdostDocument 
            Caption         =   "�������������"
            Height          =   375
            Left            =   -73200
            TabIndex        =   108
            Top             =   4560
            Width           =   1575
         End
         Begin VB.CommandButton addUdostDocument 
            Caption         =   "��������"
            Height          =   375
            Left            =   -74880
            TabIndex        =   107
            Top             =   4560
            Width           =   1575
         End
         Begin VB.TextBox mesto_rozhdeniya 
            Height          =   375
            Left            =   -69480
            TabIndex        =   106
            Top             =   3960
            Width           =   3375
         End
         Begin VB.TextBox kod_podrazdeleniya 
            Height          =   375
            Left            =   -69480
            TabIndex        =   104
            Top             =   3480
            Width           =   855
         End
         Begin VB.TextBox data_vidachi 
            Height          =   375
            Left            =   -70080
            TabIndex        =   102
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox data_rozhdeniya 
            Height          =   375
            Left            =   -70080
            TabIndex        =   99
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox fio 
            Height          =   615
            Left            =   -70080
            MultiLine       =   -1  'True
            TabIndex        =   97
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox nomer 
            Height          =   405
            Left            =   -70080
            TabIndex        =   95
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox seriya 
            Height          =   375
            Left            =   -70080
            TabIndex        =   93
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox vid 
            Height          =   315
            Left            =   -70080
            TabIndex        =   91
            Top             =   480
            Width           =   3975
         End
         Begin VB.ListBox UdostDocument 
            Height          =   4155
            Left            =   -74880
            TabIndex        =   89
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox strana_grazhdanstva_ili_registracii 
            Height          =   315
            Left            =   240
            TabIndex        =   88
            Top             =   4320
            Width           =   8535
         End
         Begin VB.TextBox dolzhnik_data_registracii 
            Height          =   405
            Left            =   720
            TabIndex        =   86
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox dolzhnik_ogrn 
            Height          =   405
            Left            =   720
            TabIndex        =   84
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox dolzhnik_kpp 
            Height          =   375
            Left            =   720
            TabIndex        =   82
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox dolzhnik_adres 
            Height          =   375
            Left            =   1320
            TabIndex        =   80
            Top             =   1440
            Width           =   7455
         End
         Begin VB.TextBox dolzhnik_dolzhnik 
            Height          =   405
            Left            =   1920
            TabIndex        =   78
            Top             =   960
            Width           =   6855
         End
         Begin VB.ComboBox dolzhnik_status_lica 
            Height          =   315
            Left            =   1320
            TabIndex        =   76
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label75 
            Caption         =   "��� ���. ��"
            Height          =   375
            Left            =   -68760
            TabIndex        =   159
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label74 
            Caption         =   "��� ����. ���. ��"
            Height          =   375
            Left            =   -72120
            TabIndex        =   157
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label73 
            Caption         =   "� ����."
            Height          =   255
            Left            =   -69120
            TabIndex        =   155
            Top             =   3360
            Width           =   615
         End
         Begin VB.Label Label72 
            Caption         =   "VIN"
            Height          =   255
            Left            =   -72000
            TabIndex        =   153
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label Label71 
            Caption         =   "��������������� ���.����"
            Height          =   255
            Left            =   -72120
            TabIndex        =   151
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Label70 
            Caption         =   "����"
            Height          =   255
            Left            =   -70800
            TabIndex        =   149
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label69 
            Caption         =   "������"
            Height          =   375
            Left            =   -70800
            TabIndex        =   147
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label68 
            Caption         =   "�����"
            Height          =   255
            Left            =   -70680
            TabIndex        =   145
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label67 
            Caption         =   "��������� ��"
            Height          =   255
            Left            =   -71160
            TabIndex        =   143
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label66 
            Caption         =   "������������ ��������"
            Height          =   375
            Left            =   -71880
            TabIndex        =   141
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label65 
            Caption         =   "����������� �����"
            Height          =   255
            Left            =   -74760
            TabIndex        =   138
            Top             =   2880
            Width           =   2775
         End
         Begin VB.Label Label64 
            Caption         =   "����������� �����"
            Height          =   255
            Left            =   -74760
            TabIndex        =   136
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label63 
            Caption         =   "������������ ����������� / ��� ���.������."
            Height          =   255
            Left            =   -74760
            TabIndex        =   134
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label62 
            Caption         =   "������������ ��������"
            Height          =   375
            Left            =   -74760
            TabIndex        =   132
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label61 
            Caption         =   "������ ����� (��������������)"
            Height          =   255
            Left            =   -71640
            TabIndex        =   125
            Top             =   3360
            Width           =   4575
         End
         Begin VB.Label Label60 
            Caption         =   "����������� �����"
            Height          =   255
            Left            =   -71640
            TabIndex        =   123
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label59 
            Caption         =   "����������� �����"
            Height          =   255
            Left            =   -71640
            TabIndex        =   121
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label58 
            Caption         =   "�������� �����"
            Height          =   255
            Left            =   -71280
            TabIndex        =   119
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label57 
            Caption         =   "�������, �2"
            Height          =   255
            Left            =   -71040
            TabIndex        =   117
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "������������ �������"
            Height          =   255
            Left            =   -71760
            TabIndex        =   115
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label55 
            Caption         =   "������������ ����."
            Height          =   375
            Left            =   -71640
            TabIndex        =   113
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label54 
            Caption         =   "����� ��������"
            Height          =   375
            Left            =   -70920
            TabIndex        =   105
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label53 
            Caption         =   "��� �������������"
            Height          =   255
            Left            =   -71160
            TabIndex        =   103
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label Label52 
            Caption         =   "���� ������"
            Height          =   255
            Left            =   -71400
            TabIndex        =   101
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label51 
            Caption         =   "���"
            Height          =   255
            Left            =   -68280
            TabIndex        =   100
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label50 
            Caption         =   "���� ��������"
            Height          =   255
            Left            =   -71520
            TabIndex        =   98
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label49 
            Caption         =   "������� ��� ��������"
            Height          =   495
            Left            =   -71520
            TabIndex        =   96
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label48 
            Caption         =   "�����"
            Height          =   255
            Left            =   -70920
            TabIndex        =   94
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label47 
            Caption         =   "�����"
            Height          =   255
            Left            =   -70920
            TabIndex        =   92
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label46 
            Caption         =   "��� ���������"
            Height          =   255
            Left            =   -71520
            TabIndex        =   90
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label45 
            Caption         =   "������ ����������� ��� �����������"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   3960
            Width           =   2895
         End
         Begin VB.Label Label44 
            Caption         =   "���� �����������"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label43 
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label42 
            Caption         =   "���"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label41 
            Caption         =   "�����"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label40 
            Caption         =   "��� / ������������"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label39 
            Caption         =   "������ ����"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   32
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "�������� � ����������"
         TabPicture(0)   =   "IspDoc.frx":018F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "mesto_registracii"
         Tab(0).Control(1)=   "data_registracii"
         Tab(0).Control(2)=   "ogrn"
         Tab(0).Control(3)=   "kpp"
         Tab(0).Control(4)=   "inn"
         Tab(0).Control(5)=   "adres"
         Tab(0).Control(6)=   "vziskatel"
         Tab(0).Control(7)=   "status_lica"
         Tab(0).Control(8)=   "Label86"
         Tab(0).Control(9)=   "Label25"
         Tab(0).Control(10)=   "Label24"
         Tab(0).Control(11)=   "Label22"
         Tab(0).Control(12)=   "Label21"
         Tab(0).Control(13)=   "Label20"
         Tab(0).Control(14)=   "Label19"
         Tab(0).Control(15)=   "Label18"
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "��������� ��� ������������"
         TabPicture(1)   =   "IspDoc.frx":01AB
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label26"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label27"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label28"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label29"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label30"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label31"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label32"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label33"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label34"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label35"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label36"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label37"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label38"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "naimenovanie_poluchatelya"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "schet_poluchatelya"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "licevoy_schet"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "summa"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "okato"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "oktmo"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "inn_poluchatelya"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "kpp_poluchatelya"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "naimenovanie_banka_poluchatelya"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "korschet_banka_poluchatelya"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "bik_banka_poluchatelya"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "pokazatel_tipa_platezha"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "kbk"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).ControlCount=   26
         Begin VB.TextBox kbk 
            Height          =   405
            Left            =   2760
            TabIndex        =   73
            Top             =   4440
            Width           =   2055
         End
         Begin VB.ComboBox pokazatel_tipa_platezha 
            Height          =   315
            Left            =   2760
            TabIndex        =   70
            Top             =   3960
            Width           =   3135
         End
         Begin VB.TextBox bik_banka_poluchatelya 
            Height          =   375
            Left            =   2760
            TabIndex        =   69
            Top             =   3480
            Width           =   6015
         End
         Begin VB.TextBox korschet_banka_poluchatelya 
            Height          =   375
            Left            =   2760
            TabIndex        =   67
            Top             =   3000
            Width           =   6015
         End
         Begin VB.TextBox naimenovanie_banka_poluchatelya 
            Height          =   375
            Left            =   2760
            TabIndex        =   65
            Top             =   2520
            Width           =   6015
         End
         Begin VB.TextBox kpp_poluchatelya 
            Height          =   375
            Left            =   4800
            TabIndex        =   63
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox inn_poluchatelya 
            Height          =   405
            Left            =   1680
            TabIndex        =   61
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox oktmo 
            Height          =   405
            Left            =   4320
            TabIndex        =   59
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox okato 
            Height          =   375
            Left            =   720
            TabIndex        =   57
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox summa 
            Height          =   375
            Left            =   7440
            TabIndex        =   55
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox licevoy_schet 
            Height          =   375
            Left            =   4680
            TabIndex        =   53
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox schet_poluchatelya 
            Height          =   375
            Left            =   1440
            TabIndex        =   51
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox naimenovanie_poluchatelya 
            Height          =   375
            Left            =   2640
            TabIndex        =   49
            Top             =   480
            Width           =   6135
         End
         Begin VB.TextBox mesto_registracii 
            Height          =   285
            Left            =   -73200
            TabIndex        =   47
            Top             =   3960
            Width           =   5535
         End
         Begin VB.TextBox data_registracii 
            Height          =   285
            Left            =   -73200
            TabIndex        =   45
            Top             =   3600
            Width           =   1935
         End
         Begin VB.TextBox ogrn 
            Height          =   285
            Left            =   -73200
            TabIndex        =   43
            Top             =   3240
            Width           =   2175
         End
         Begin VB.TextBox kpp 
            Height          =   285
            Left            =   -73200
            TabIndex        =   42
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox inn 
            Height          =   285
            Left            =   -73200
            TabIndex        =   40
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox adres 
            Height          =   315
            Left            =   -74880
            TabIndex        =   38
            Top             =   2040
            Width           =   8655
         End
         Begin VB.TextBox vziskatel 
            Height          =   375
            Left            =   -74880
            TabIndex        =   36
            Top             =   1320
            Width           =   8655
         End
         Begin VB.ComboBox status_lica 
            Height          =   315
            Left            =   -74880
            TabIndex        =   34
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label86 
            Caption         =   "����"
            Height          =   255
            Left            =   -73920
            TabIndex        =   196
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label38 
            Caption         =   "���"
            Height          =   255
            Left            =   1680
            TabIndex        =   72
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label37 
            Caption         =   "���������� ���� �������"
            Height          =   255
            Left            =   600
            TabIndex        =   71
            Top             =   3960
            Width           =   2055
         End
         Begin VB.Label Label36 
            Caption         =   "��� ����� ����������"
            Height          =   255
            Left            =   840
            TabIndex        =   68
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label35 
            Caption         =   "������� ����� ����������"
            Height          =   255
            Left            =   600
            TabIndex        =   66
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label34 
            Caption         =   "������������ ����� ����������"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   2520
            Width           =   2535
         End
         Begin VB.Label Label33 
            Caption         =   "��� ����������"
            Height          =   255
            Left            =   4320
            TabIndex        =   62
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label32 
            Caption         =   "��� ����������"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "�����"
            Height          =   255
            Left            =   3480
            TabIndex        =   58
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "�����"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "�����"
            Height          =   375
            Left            =   6840
            TabIndex        =   54
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "������� ����"
            Height          =   375
            Left            =   3600
            TabIndex        =   52
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label27 
            Caption         =   "���� ����������"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "������������ ����������"
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label25 
            Caption         =   "����� �����������"
            Height          =   255
            Left            =   -74880
            TabIndex        =   46
            Top             =   3960
            Width           =   1575
         End
         Begin VB.Label Label24 
            Caption         =   "���� �����������"
            Height          =   255
            Left            =   -74760
            TabIndex        =   44
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "���"
            Height          =   255
            Left            =   -73800
            TabIndex        =   41
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label21 
            Caption         =   "���"
            Height          =   255
            Left            =   -73800
            TabIndex        =   39
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "�����"
            Height          =   255
            Left            =   -74880
            TabIndex        =   37
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "����������"
            Height          =   255
            Left            =   -74880
            TabIndex        =   35
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "������ ����"
            Height          =   255
            Left            =   -74880
            TabIndex        =   33
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox rezolyutativnaya_chast_sudebnogo_acta 
         Height          =   3735
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   1560
         Width           =   9015
      End
      Begin VB.TextBox ustanovochnaya_chast_sudebnogo_acta 
         Height          =   285
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   9015
      End
      Begin VB.ComboBox valyuta_dolga 
         Height          =   315
         Left            =   -70320
         TabIndex        =   27
         Text            =   "���������� �����"
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox summa_dolga 
         Height          =   375
         Left            =   -73560
         TabIndex        =   25
         Top             =   5040
         Width           =   2295
      End
      Begin VB.ComboBox podl_nemedl_isp 
         Height          =   315
         Left            =   -74760
         TabIndex        =   22
         Text            =   "���"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox data_vsupleniya_v_zs 
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti 
         Height          =   315
         Left            =   -66600
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox dublicat 
         Height          =   315
         Left            =   -70320
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox data_sudebnogo_acta 
         Height          =   285
         Left            =   -73080
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox ob_data_vidachi 
         Height          =   375
         Left            =   -68880
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox nomer_ekz_ID 
         Height          =   375
         Left            =   -69960
         TabIndex        =   9
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox srok_predyavleniya_k_ispolneniyu_razmernost 
         Height          =   315
         Left            =   -72000
         TabIndex        =   7
         Text            =   "����"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox srok_predyavleniya_k_ispolneniyu_znachenie 
         Height          =   375
         Left            =   -72600
         TabIndex        =   6
         Text            =   "3"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "����������"
         Height          =   255
         Left            =   -68280
         TabIndex        =   200
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "C��������� ���������"
         Height          =   255
         Left            =   -67680
         TabIndex        =   198
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label80 
         Caption         =   "X509Certificate"
         Height          =   255
         Left            =   -74880
         TabIndex        =   172
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label79 
         Caption         =   "SignatureValue"
         Height          =   255
         Left            =   -74880
         TabIndex        =   170
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label78 
         Caption         =   "����� ������������ ����"
         Height          =   255
         Left            =   -74760
         TabIndex        =   168
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label77 
         Caption         =   "����� ����, ��������� �������������� ��������"
         Height          =   255
         Left            =   -74760
         TabIndex        =   167
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label76 
         Caption         =   "������������ ����, ��������� �������������� ��������"
         Height          =   375
         Left            =   -74760
         TabIndex        =   166
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label17 
         Caption         =   "������������ ����� ��������� ����:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "������������ ����� ��������� ����:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label15 
         Caption         =   "������"
         Height          =   255
         Left            =   -71040
         TabIndex        =   26
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "����� �����"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "��� �������� ���������� ��"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "����.������.���."
         Height          =   255
         Left            =   -74400
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "������� ��� �������� ����� (��� ����������), ��������� ��"
         Height          =   255
         Left            =   -72960
         TabIndex        =   20
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label10 
         Caption         =   "���� ���������� � �/�"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "�� ����. ����. � ���� ��������"
         Height          =   255
         Left            =   -69240
         TabIndex        =   16
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "��������"
         Height          =   255
         Left            =   -71160
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "���� ��������� ����"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "���� ������"
         Height          =   255
         Left            =   -68880
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "� ���. ��"
         Height          =   255
         Left            =   -69960
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "���� ������������ � ����������"
         Height          =   255
         Left            =   -72720
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox po_delu_nomer 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox ispolnitelniy_document_nomer 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "�� ���� �:"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�������������� �������� �:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "SmevIspDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dictUD As New Collection
Dim ud As UdostDocument
Dim idxUD As Integer
Dim udEdit As Boolean

Dim dictND As New Collection
Dim nd As Nedvizhimost
Dim idxND As Integer
Dim ndEdit As Boolean

Dim dictTD As New Collection
Dim td As TransSredstva
Dim idxTD As Integer
Dim tdEdit As Boolean

Dim dictSud As New Collection
Dim sD As Sud
Dim dictSudya As New Collection

Const docDir As String = "D:\����\��������������_���������\"
Const dd1 As String = "D:\"
Const dd2 As String = "D:\����\"

Const templateFileName = "template"

Private Sub AddNedvizhimost_Click()
    SaveNedvizhimost.Visible = True
    UndoNedvizhimost.Visible = True
    
    AddNedvizhimost.Visible = False
    EditNedvizhimost.Visible = False
    DeleteNedvizhimost.Visible = False
    
    Set nd = New Nedvizhimost
    
    Actualnost.Text = ""
    Naimenovanie.Text = ""
    Ploshad.Text = ""
    UslNomer.Text = ""
    InvNomer.Text = ""
    KadastrNomer.Text = ""
    TochAdres.Text = ""
    
    enableNedvizhimost
End Sub

Private Sub AddTransSredstva_Click()
    SaveTransSredstva.Visible = True
    UndoTransSredstva.Visible = True
    AddTransSredstva.Visible = False
    EditTransSredstva.Visible = False
    DeleteTransSredstva.Visible = False
    
    Set td = New TransSredstva
    
    TS_Actualnost.Text = ""
    Kategoriya.Text = ""
    Marka.Text = ""
    Model.Text = ""
    Cvet.Text = ""
    GosZnak.Text = ""
    VIN.Text = ""
    NDvig.Text = ""
    KodPodr.Text = ""
    GodVipuska.Text = ""
    
    enableTransSredstva
End Sub



Private Sub adres_suda_vidayushego_ispolnitelniy_document_Click()
    'dictSud.
    
      Dim i As Integer
      Dim selectedSD As Sud
      
      If (adres_suda_vidayushego_ispolnitelniy_document.Text = "") Then
        FIO_sudiy.Clear
          For Each ds In dictSudya
            FIO_sudiy.AddItem ds
  Next ds
      Else
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
          If (selectedSD.address = adres_suda_vidayushego_ispolnitelniy_document.Text) Then
              naimenovanie_suda_vidayushego_ispolnitelniy_document.Text = selectedSD.name
              FIO_sudiy.Clear
              For j = 1 To selectedSD.sudyaName.Count
                FIO_sudiy.AddItem selectedSD.sudyaName(j)
              Next j
              
              If FIO_sudiy.ListCount = 1 Then
                FIO_sudiy.Text = selectedSD.sudyaName(j - 1)
                Else
                FIO_sudiy.Text = ""
              End If
              Exit For
          End If
      Next i
      End If
      
      mesto_rassmotreniya_dela.Text = adres_suda_vidayushego_ispolnitelniy_document.Text
End Sub

Private Sub Command1_Click()
    Save templateFileName
End Sub

Private Sub CommandButtonOpen_Click()
CommonDialog1.Filter = "���.���. (*.xml)|*.xml|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.InitDir = docDir
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
SmevIspDoc.openDoc (CommonDialog1.filename)
End Sub

Private Sub DeleteNedvizhimost_Click()
    SaveNedvizhimost.Visible = False
    UndoNedvizhimost.Visible = False
    Dim i As Integer
          For i = 0 To Nedvizhimost.ListCount - 1
          If Nedvizhimost.Selected(i) Then
              Set nd = dictND(i + 1)
              Exit For
          End If
      Next i


      If i + 1 < dictND.Count Then
       For j = i + 1 To dictND.Count - 1
        Set nd1 = dictND(j + 1)

dictND(j).Actualnost = nd1.Actualnost
dictND(j).Naimenovanie = nd1.Naimenovanie
dictND(j).Ploshad = nd1.Ploshad
dictND(j).UslNomer = nd1.UslNomer
dictND(j).InvNomer = nd1.InvNomer
dictND(j).KadastrNomer = nd1.KadastrNomer
dictND(j).TochAdres = nd1.TochAdres

       Next j
      End If
      
      Nedvizhimost.RemoveItem (i)
      
      If Nedvizhimost.ListCount > 0 Then
        dictND.Remove (dictND.Count)
        Nedvizhimost.Selected(0) = True
      Else
        Actualnost.Text = ""
        Naimenovanie.Text = ""
        Ploshad.Text = ""
        UslNomer.Text = ""
        InvNomer.Text = ""
        KadastrNomer.Text = ""
        TochAdres.Text = ""
      End If
      
      disableNedvizhimost
End Sub

Private Sub DeleteTransSredstva_Click()
    SaveTransSredstva.Visible = False
    UndoTransSredstva.Visible = False
    Dim i As Integer
          For i = 0 To TransSredstva.ListCount - 1
          If TransSredstva.Selected(i) Then
              Set td = dictTD(i + 1)
              Exit For
          End If
      Next i
      

      If i + 1 < dictTD.Count Then
       For j = i + 1 To dictTD.Count - 1
        Set td1 = dictTD(j + 1)

dictTD(j).Actualnost = td1.Actualnost
dictTD(j).Kategoriya = td1.Kategoriya
dictTD(j).Marka = td1.Marka
dictTD(j).Model = td1.Model
dictTD(j).Cvet = td1.Cvet
dictTD(j).GosZnak = td1.GosZnak
dictTD(j).VIN = td1.VIN
dictTD(j).NDvig = td1.NDvig
dictTD(j).KodPodr = td1.KodPodr
dictTD(j).GodVipuska = td1.GodVipuska

       Next j
      End If
      
      TransSredstva.RemoveItem (i)
      
      If TransSredstva.ListCount > 0 Then
        dictTD.Remove (dictTD.Count)
        TransSredstva.Selected(0) = True
      Else
TS_Actualnost.Text = ""
Kategoriya.Text = ""
Marka.Text = ""
Model.Text = ""
Cvet.Text = ""
GosZnak.Text = ""
VIN.Text = ""
NDvig.Text = ""
KodPodr.Text = ""
GodVipuska.Text = ""

      End If
      
      disableTransSredstva
End Sub

Private Sub EditNedvizhimost_Click()
    SaveNedvizhimost.Visible = True
    UndoNedvizhimost.Visible = True
    
    AddNedvizhimost.Visible = False
    EditNedvizhimost.Visible = False
    DeleteNedvizhimost.Visible = False
    
    ndEdit = True
    
    enableNedvizhimost
End Sub

Private Sub EditTransSredstva_Click()
    SaveTransSredstva.Visible = True
    UndoTransSredstva.Visible = True
    AddTransSredstva.Visible = False
    EditTransSredstva.Visible = False
    DeleteTransSredstva.Visible = False
    tdEdit = True
    enableTransSredstva
End Sub

Private Sub FIO_sudiy_Click()
      Dim i As Integer
      Dim selectedSD As Sud
      Dim flg As Boolean
      If (FIO_sudiy.Text <> "") Then

        flg = False
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
            
              For j = 1 To selectedSD.sudyaName.Count
                If selectedSD.sudyaName(j) = FIO_sudiy.Text Then
                    flg = True
                    Exit For
                End If
              Next j
        If flg Then
            Exit For
        End If
      Next i
      
      If flg Then
        naimenovanie_suda_vidayushego_ispolnitelniy_document.Text = selectedSD.name
        adres_suda_vidayushego_ispolnitelniy_document.Text = selectedSD.address
        mesto_rassmotreniya_dela.Text = adres_suda_vidayushego_ispolnitelniy_document.Text
      End If
      
      End If
End Sub

Private Sub naimenovanie_suda_vidayushego_ispolnitelniy_document_Click()

      Dim i As Integer
      Dim selectedSD As Sud
      
      If (naimenovanie_suda_vidayushego_ispolnitelniy_document.Text = "") Then
        FIO_sudiy.Clear
          For Each ds In dictSudya
            FIO_sudiy.AddItem ds
  Next ds
      Else
      For i = 1 To dictSud.Count
        Set selectedSD = dictSud(i)
          If (selectedSD.name = naimenovanie_suda_vidayushego_ispolnitelniy_document.Text) Then
              adres_suda_vidayushego_ispolnitelniy_document.Text = selectedSD.address
              mesto_rassmotreniya_dela.Text = adres_suda_vidayushego_ispolnitelniy_document.Text
              FIO_sudiy.Clear
              For j = 1 To selectedSD.sudyaName.Count
                FIO_sudiy.AddItem selectedSD.sudyaName(j)
              Next j
              
              If FIO_sudiy.ListCount = 1 Then
                FIO_sudiy.Text = selectedSD.sudyaName(j - 1)
                Else
                FIO_sudiy.Text = ""
              End If
              Exit For
          End If
      Next i
      End If
End Sub

Private Sub Nedvizhimost_Click()
    Nedvizhimost_Change
End Sub

Private Sub Nedvizhimost_KeyPress(KeyAscii As Integer)
    Nedvizhimost_Change
End Sub

Private Sub SaveTransSredstva_Click()
    Dim i As Integer
    If (tdEdit) Then
      For i = 0 To TransSredstva.ListCount - 1
          If TransSredstva.Selected(i) Then
              Set td = dictTD(i + 1)
              Exit For
          End If
      Next i
        td.Actualnost = TS_Actualnost.Text
        td.Kategoriya = Kategoriya.Text
        td.Marka = Marka.Text
        td.Model = Model.Text
        td.Cvet = Cvet.Text
        td.GosZnak = GosZnak.Text
        td.VIN = VIN.Text
        td.NDvig = NDvig.Text
        td.KodPodr = KodPodr.Text
        td.GodVipuska = GodVipuska.Text
        
        tdEdit = False
    Else
    
        td.Actualnost = TS_Actualnost.Text
        td.Kategoriya = Kategoriya.Text
        td.Marka = Marka.Text
        td.Model = Model.Text
        td.Cvet = Cvet.Text
        td.GosZnak = GosZnak.Text
        td.VIN = VIN.Text
        td.NDvig = NDvig.Text
        td.KodPodr = KodPodr.Text
        td.GodVipuska = GodVipuska.Text
    
    TransSredstva.AddItem td.Kategoriya
    dictTD.Add td
    idxTD = idxTD + 1
    End If
    
    SaveTransSredstva.Visible = False
    UndoTransSredstva.Visible = False
    
    AddTransSredstva.Visible = True
    EditTransSredstva.Visible = True
    DeleteTransSredstva.Visible = True
    disableTransSredstva
End Sub

Private Sub TransSredstva_Change()
      Dim i As Integer
      Dim selectedTD As TransSredstva
      
      If TransSredstva.ListCount > 0 Then
      
      For i = 0 To TransSredstva.ListCount - 1
          If TransSredstva.Selected(i) Then
              Set selectedTD = dictTD(i + 1)
              Exit For
          End If
      Next i
      
     
      TS_Actualnost.Text = selectedTD.Actualnost
      Kategoriya.Text = selectedTD.Kategoriya
      Marka.Text = selectedTD.Marka
      Model.Text = selectedTD.Model
      Cvet.Text = selectedTD.Cvet
      GosZnak.Text = selectedTD.GosZnak
      VIN.Text = selectedTD.VIN
      NDvig.Text = selectedTD.NDvig
      KodPodr.Text = selectedTD.KodPodr
      GodVipuska.Text = selectedTD.GodVipuska
      End If
      
End Sub

Private Sub Nedvizhimost_Change()
      Dim i As Integer
      Dim selectedND As Nedvizhimost
      
      If Nedvizhimost.ListCount > 0 Then
      
      For i = 0 To Nedvizhimost.ListCount - 1
          If Nedvizhimost.Selected(i) Then
              Set selectedND = dictND(i + 1)
              Exit For
          End If
      Next i
      
      Actualnost.Text = selectedND.Actualnost
      Naimenovanie.Text = selectedND.Naimenovanie
      Ploshad.Text = selectedND.Ploshad
      UslNomer.Text = selectedND.UslNomer
      InvNomer.Text = selectedND.InvNomer
      KadastrNomer.Text = selectedND.KadastrNomer
      TochAdres.Text = selectedND.TochAdres
      End If
End Sub

Private Sub SaveNedvizhimost_Click()
    Dim i As Integer
    If (ndEdit) Then
      For i = 0 To Nedvizhimost.ListCount - 1
          If Nedvizhimost.Selected(i) Then
              Set nd = dictND(i + 1)
              Exit For
          End If
      Next i
        nd.Actualnost = Actualnost.Text
        nd.Naimenovanie = Naimenovanie.Text
        nd.Ploshad = Ploshad.Text
        nd.UslNomer = UslNomer.Text
        nd.InvNomer = InvNomer.Text
        nd.KadastrNomer = KadastrNomer.Text
        nd.TochAdres = TochAdres.Text
        
        ndEdit = False
    Else
    
        nd.Actualnost = Actualnost.Text
        nd.Naimenovanie = Naimenovanie.Text
        nd.Ploshad = Ploshad.Text
        nd.UslNomer = UslNomer.Text
        nd.InvNomer = InvNomer.Text
        nd.KadastrNomer = KadastrNomer.Text
        nd.TochAdres = TochAdres.Text
    
    Nedvizhimost.AddItem nd.Naimenovanie
    dictND.Add nd
    idxND = idxND + 1
    End If
    
    SaveNedvizhimost.Visible = False
    UndoNedvizhimost.Visible = False

    AddNedvizhimost.Visible = True
    EditNedvizhimost.Visible = True
    DeleteNedvizhimost.Visible = True
    
    disableNedvizhimost
End Sub

Private Sub UdostDocument_Change()
      Dim i As Integer
      Dim selectedUD As UdostDocument
      
      If UdostDocument.ListCount > 0 Then
      
      For i = 0 To UdostDocument.ListCount - 1
          If UdostDocument.Selected(i) Then
              Set selectedUD = dictUD(i + 1)
              Exit For
          End If
      Next i
      
        vid.Text = selectedUD.vid
        seriya.Text = selectedUD.seriya
        nomer.Text = selectedUD.nomer
        fio.Text = selectedUD.fio
        pol.Text = selectedUD.pol
        data_rozhdeniya.Text = selectedUD.data_rozhdeniya
        data_vidachi.Text = selectedUD.data_vidachi
        kod_podrazdeleniya.Text = selectedUD.kod_podrazdeleniya
        mesto_rozhdeniya.Text = selectedUD.mesto_rozhdeniya
      End If
End Sub

Private Sub addUdostDocument_Click()
    SaveUdostDocument.Visible = True
    UndoUdostDocument.Visible = True
    
    addUdostDocument.Visible = False
    EditUdostDocument.Visible = False
    DeleteUdostDocument.Visible = False
    
    Set ud = New UdostDocument
    
    vid.Text = ""
    seriya.Text = ""
    nomer.Text = ""
    fio.Text = ""
    pol.Text = ""
    data_rozhdeniya.Text = ""
    data_vidachi.Text = ""
    kod_podrazdeleniya.Text = ""
    mesto_rozhdeniya.Text = ""
    
    enableUdostDocument
End Sub

Private Sub CommandButtonGenerate_Click()
    Dim filename As String

    filename = ispolnitelniy_document_nomer.Text + "_" + po_delu_nomer.Text
    filename = Replace(filename, "-", "_")
    filename = Replace(filename, "\", "_")
    filename = Replace(filename, "/", "_")
    filename = Replace(filename, ".", "_")
      
    Save filename

End Sub

Private Sub Save(ByVal filename As String)
  
  filename2 = docDir + "id_" + filename + ".xml"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set out = fso.CreateTextFile(filename2, True, True)
  out.WriteLine ("<?xml version='1.0'?>")
  
  out.WriteLine ("<ispolnitelniy_document")
  out.WriteLine ("xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""")
  out.WriteLine ("xsi:noNamespaceSchemaLocation=""data_id.xsd"">")
  
   out.WriteLine ("<ispolnitelniy_document_nomer>")
   out.WriteLine (ispolnitelniy_document_nomer.Text)
   out.WriteLine ("</ispolnitelniy_document_nomer>")

  out.WriteLine ("<po_delu_nomer>")
    out.WriteLine (po_delu_nomer.Text)
  out.WriteLine ("</po_delu_nomer>")
  
  out.WriteLine ("<srok_predyavleniya_k_ispolneniyu_znachenie>")
    out.WriteLine (srok_predyavleniya_k_ispolneniyu_znachenie.Text)
  out.WriteLine ("</srok_predyavleniya_k_ispolneniyu_znachenie>")
  
  out.WriteLine ("<srok_predyavleniya_k_ispolneniyu_razmernost>")
    out.WriteLine (srok_predyavleniya_k_ispolneniyu_razmernost.Text)
  out.WriteLine ("</srok_predyavleniya_k_ispolneniyu_razmernost>")

  out.WriteLine ("<SolidarnoeVziskanie>")
    out.WriteLine (SolidarnoeVziskanie.Text)
  out.WriteLine ("</SolidarnoeVziskanie>")
  
  out.WriteLine ("<nomer_ekz_ID>")
    out.WriteLine (nomer_ekz_ID.Text)
  out.WriteLine ("</nomer_ekz_ID>")
  
  out.WriteLine ("<Gosposhlina>")
    out.WriteLine (Gosposhlina.Text)
  out.WriteLine ("</Gosposhlina>")
  
  out.WriteLine ("<data_vidachi>")
    out.WriteLine (ob_data_vidachi.Text)
  out.WriteLine ("</data_vidachi>")
  
  out.WriteLine ("<data_sudebnogo_acta>")
    out.WriteLine (data_sudebnogo_acta.Text)
  out.WriteLine ("</data_sudebnogo_acta>")
  
  out.WriteLine ("<dublicat>")
    out.WriteLine (dublicat.Text)
  out.WriteLine ("</dublicat>")
  
  out.WriteLine ("<vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti>")
    out.WriteLine (vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti.Text)
  out.WriteLine ("</vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti>")
  
  out.WriteLine ("<data_vsupleniya_v_zs>")
    out.WriteLine (data_vsupleniya_v_zs.Text)
  out.WriteLine ("</data_vsupleniya_v_zs>")
  
  out.WriteLine ("<FIO_sudiy>")
    out.WriteLine (FIO_sudiy.Text)
  out.WriteLine ("</FIO_sudiy>")
  
  out.WriteLine ("<podl_nemedl_isp>")
    out.WriteLine (podl_nemedl_isp.Text)
  out.WriteLine ("</podl_nemedl_isp>")
  
  out.WriteLine ("<vid_sushnosti_ispolneniya_ID>")
    out.WriteLine (vid_sushnosti_ispolneniya_ID.Text)
  out.WriteLine ("</vid_sushnosti_ispolneniya_ID>")
  
  out.WriteLine ("<summa_dolga>")
    out.WriteLine (summa_dolga.Text)
  out.WriteLine ("</summa_dolga>")
  
  out.WriteLine ("<valyuta_dolga>")
    out.WriteLine (valyuta_dolga.Text)
  out.WriteLine ("</valyuta_dolga>")
  
  out.WriteLine ("<ustanovochnaya_chast_sudebnogo_acta>")
    out.WriteLine (ustanovochnaya_chast_sudebnogo_acta.Text)
  out.WriteLine ("</ustanovochnaya_chast_sudebnogo_acta>")
  
  out.WriteLine ("<rezolyutativnaya_chast_sudebnogo_acta>")
    out.WriteLine (rezolyutativnaya_chast_sudebnogo_acta.Text)
  out.WriteLine ("</rezolyutativnaya_chast_sudebnogo_acta>")
  
  out.WriteLine ("<dolzhnik_status_lica>")
    out.WriteLine (dolzhnik_status_lica.Text)
  out.WriteLine ("</dolzhnik_status_lica>")
  
  out.WriteLine ("<vziskatel>")
    out.WriteLine (vziskatel.Text)
  out.WriteLine ("</vziskatel>")
  
  out.WriteLine ("<adres>")
    out.WriteLine (adres.Text)
  out.WriteLine ("</adres>")
  
  out.WriteLine ("<inn>")
    out.WriteLine (inn.Text)
  out.WriteLine ("</inn>")
  
  out.WriteLine ("<kpp>")
    out.WriteLine (kpp.Text)
  out.WriteLine ("</kpp>")
  
  out.WriteLine ("<ogrn>")
    out.WriteLine (ogrn.Text)
  out.WriteLine ("</ogrn>")
  
  out.WriteLine ("<data_registracii>")
    out.WriteLine (data_registracii.Text)
  out.WriteLine ("</data_registracii>")
  
  out.WriteLine ("<mesto_registracii>")
    out.WriteLine (mesto_registracii.Text)
  out.WriteLine ("</mesto_registracii>")

'TODO - ���� ������� �� �����
  out.WriteLine ("<data_rozhdeniya></data_rozhdeniya>")
  out.WriteLine ("<snils></snils>")
  out.WriteLine ("<mesto_rozhdeniya></mesto_rozhdeniya>")
  
  out.WriteLine ("<naimenovanie_poluchatelya>")
    out.WriteLine (naimenovanie_poluchatelya.Text)
  out.WriteLine ("</naimenovanie_poluchatelya>")
  
  out.WriteLine ("<schet_poluchatelya>")
    out.WriteLine (schet_poluchatelya.Text)
  out.WriteLine ("</schet_poluchatelya>")
  
  out.WriteLine ("<licevoy_schet>")
    out.WriteLine (licevoy_schet.Text)
  out.WriteLine ("</licevoy_schet>")
  
  out.WriteLine ("<summa>")
    out.WriteLine (summa.Text)
  out.WriteLine ("</summa>")
  
  out.WriteLine ("<okato>")
    out.WriteLine (okato.Text)
  out.WriteLine ("</okato>")
  
  out.WriteLine ("<oktmo>")
    out.WriteLine (oktmo.Text)
  out.WriteLine ("</oktmo>")
  
  out.WriteLine ("<inn_poluchatelya>")
    out.WriteLine (inn_poluchatelya.Text)
  out.WriteLine ("</inn_poluchatelya>")
  
  out.WriteLine ("<kpp_poluchatelya>")
    out.WriteLine (kpp_poluchatelya.Text)
  out.WriteLine ("</kpp_poluchatelya>")
  
  out.WriteLine ("<naimenovanie_banka_poluchatelya>")
    out.WriteLine (naimenovanie_banka_poluchatelya.Text)
  out.WriteLine ("</naimenovanie_banka_poluchatelya>")
  
  out.WriteLine ("<korschet_banka_poluchatelya>")
    out.WriteLine (korschet_banka_poluchatelya.Text)
  out.WriteLine ("</korschet_banka_poluchatelya>")
  
  out.WriteLine ("<bik_banka_poluchatelya>")
    out.WriteLine (bik_banka_poluchatelya.Text)
  out.WriteLine ("</bik_banka_poluchatelya>")
  
  out.WriteLine ("<pokazatel_tipa_platezha>")
    out.WriteLine (pokazatel_tipa_platezha.Text)
  out.WriteLine ("</pokazatel_tipa_platezha>")
  
  out.WriteLine ("<kbk>")
    out.WriteLine (kbk.Text)
  out.WriteLine ("</kbk>")
  
  out.WriteLine ("<dolzhnik_status_lica>")
    out.WriteLine (dolzhnik_status_lica.Text)
  out.WriteLine ("</dolzhnik_status_lica>")
  
  out.WriteLine ("<dolzhnik>")
    out.WriteLine (dolzhnik_dolzhnik.Text)
  out.WriteLine ("</dolzhnik>")
  
  out.WriteLine ("<dolzhnik_adres>")
    out.WriteLine (dolzhnik_adres.Text)
  out.WriteLine ("</dolzhnik_adres>")
  
  out.WriteLine ("<dolzhnik_kpp>")
    out.WriteLine (dolzhnik_kpp.Text)
  out.WriteLine ("</dolzhnik_kpp>")
  
  out.WriteLine ("<dolzhnik_ogrn>")
    out.WriteLine (dolzhnik_ogrn.Text)
  out.WriteLine ("</dolzhnik_ogrn>")
  
  out.WriteLine ("<dolzhnik_data_registracii>")
    out.WriteLine (dolzhnik_data_registracii.Text)
  out.WriteLine ("</dolzhnik_data_registracii>")
  
  out.WriteLine ("<strana_grazhdanstva_ili_registracii>")
    out.WriteLine (strana_grazhdanstva_ili_registracii.Text)
  out.WriteLine ("</strana_grazhdanstva_ili_registracii>")
  
  out.WriteLine ("<dolzhnik_pol>")
    out.WriteLine (dolzhnik_pol.Text)
  out.WriteLine ("</dolzhnik_pol>")
  
  out.WriteLine ("<dolzhnik_data_rozhdeniya>")
    out.WriteLine (dolzhnik_data_rozhdeniya.Text)
  out.WriteLine ("</dolzhnik_data_rozhdeniya>")
  
  out.WriteLine ("<dolzhnik_inn>")
    out.WriteLine (dolzhnik_inn.Text)
  out.WriteLine ("</dolzhnik_inn>")
  
  out.WriteLine ("<dolzhnik_snils>")
    out.WriteLine (dolzhnik_snils.Text)
  out.WriteLine ("</dolzhnik_snils>")
  
  out.WriteLine ("<dolzhnik_mesto_rozhdeniya>")
    out.WriteLine (dolzhnik_mesto_rozhdeniya.Text)
  out.WriteLine ("</dolzhnik_mesto_rozhdeniya>")
  
  out.WriteLine ("<UdostDocumentList>")
  ' ���� �� ����������
     For Each dUD In dictUD
    out.WriteLine ("<UdostDocument>")
        
        out.WriteLine ("<vid>")
            out.WriteLine (dUD.vid)
        out.WriteLine ("</vid>")
        
        out.WriteLine ("<seriya>")
            out.WriteLine (dUD.seriya)
        out.WriteLine ("</seriya>")
        
        out.WriteLine ("<nomer>")
            out.WriteLine (dUD.nomer)
        out.WriteLine ("</nomer>")
        
        out.WriteLine ("<fio>")
            out.WriteLine (dUD.fio)
        out.WriteLine ("</fio>")
        
        out.WriteLine ("<pol>")
            out.WriteLine (dUD.pol)
        out.WriteLine ("</pol>")
        
        out.WriteLine ("<data_rozhdeniya>")
            out.WriteLine (dUD.data_rozhdeniya)
        out.WriteLine ("</data_rozhdeniya>")
        
        out.WriteLine ("<data_vidachi>")
            out.WriteLine (dUD.data_vidachi)
        out.WriteLine ("</data_vidachi>")
        
        out.WriteLine ("<kod_podrazdeleniya>")
            out.WriteLine (dUD.kod_podrazdeleniya)
        out.WriteLine ("</kod_podrazdeleniya>")
        
        out.WriteLine ("<mesto_rozhdeniya>")
            out.WriteLine (dUD.mesto_rozhdeniya)
        out.WriteLine ("</mesto_rozhdeniya>")
        
    out.WriteLine ("</UdostDocument>")
    Next
  out.WriteLine ("</UdostDocumentList>")
  
  out.WriteLine ("<NedvizhimostList>")
  ' ���� �� ������������
  For Each dND In dictND
    out.WriteLine ("<Nedvizhimost>")
        out.WriteLine ("<Actualnost>")
            out.WriteLine (dND.Actualnost)
        out.WriteLine ("</Actualnost>")
        
        out.WriteLine ("<Naimenovanie>")
            out.WriteLine (dND.Naimenovanie)
        out.WriteLine ("</Naimenovanie>")
        
        out.WriteLine ("<Ploshad>")
            out.WriteLine (dND.Ploshad)
        out.WriteLine ("</Ploshad>")
        
        out.WriteLine ("<UslNomer>")
            out.WriteLine (dND.UslNomer)
        out.WriteLine ("</UslNomer>")
        
        out.WriteLine ("<InvNomer>")
            out.WriteLine (dND.InvNomer)
        out.WriteLine ("</InvNomer>")
        
        out.WriteLine ("<KadastrNomer>")
            out.WriteLine (dND.KadastrNomer)
        out.WriteLine ("</KadastrNomer>")
        
        out.WriteLine ("<TochAdres>")
            out.WriteLine (dND.TochAdres)
        out.WriteLine ("</TochAdres>")
    out.WriteLine ("</Nedvizhimost>")
    Next
  out.WriteLine ("</NedvizhimostList>")
  
  out.WriteLine ("<mr_actualnost>")
  out.WriteLine (mr_actualnost.Text)
  out.WriteLine ("</mr_actualnost>")
  
  out.WriteLine ("<naimenovanie_organizacii_fio_ip>")
    out.WriteLine (naimenovanie_organizacii_fio_ip.Text)
  out.WriteLine ("</naimenovanie_organizacii_fio_ip>")
  
  out.WriteLine ("<jur_address>")
    out.WriteLine (jur_address.Text)
  out.WriteLine ("</jur_address>")
  
  out.WriteLine ("<fact_address>")
    out.WriteLine (fact_address.Text)
  out.WriteLine ("</fact_address>")
  
  out.WriteLine ("<TransSredstvaList>")
  
  ' ���� �� ����������
  For Each dTD In dictTD
    out.WriteLine ("<TransSredstva>")
        out.WriteLine ("<Actualnost>")
            out.WriteLine (dTD.Actualnost)
        out.WriteLine ("</Actualnost>")
        
        out.WriteLine ("<Kategoriya>")
            out.WriteLine (dTD.Kategoriya)
        out.WriteLine ("</Kategoriya>")
        
        out.WriteLine ("<Marka>")
            out.WriteLine (dTD.Marka)
        out.WriteLine ("</Marka>")
        
        out.WriteLine ("<Model>")
            out.WriteLine (dTD.Model)
        out.WriteLine ("</Model>")
        
        out.WriteLine ("<Cvet>")
            out.WriteLine (dTD.Cvet)
        out.WriteLine ("</Cvet>")
        
        out.WriteLine ("<GosZnak>")
            out.WriteLine (dTD.GosZnak)
        out.WriteLine ("</GosZnak>")
        
        out.WriteLine ("<VIN>")
            out.WriteLine (dTD.VIN)
        out.WriteLine ("</VIN>")
        
        out.WriteLine ("<NDvig>")
            out.WriteLine (dTD.NDvig)
        out.WriteLine ("</NDvig>")
        
        out.WriteLine ("<KodPodr>")
            out.WriteLine (dTD.KodPodr)
        out.WriteLine ("</KodPodr>")
        
        out.WriteLine ("<GodVipuska>")
            out.WriteLine (dTD.GodVipuska)
        out.WriteLine ("</GodVipuska>")
        
    out.WriteLine ("</TransSredstva>")
    Next
  out.WriteLine ("</TransSredstvaList>")
  
  out.WriteLine ("<naimenovanie_suda_vidayushego_ispolnitelniy_document>")
    out.WriteLine (naimenovanie_suda_vidayushego_ispolnitelniy_document.Text)
  out.WriteLine ("</naimenovanie_suda_vidayushego_ispolnitelniy_document>")
  
  out.WriteLine ("<adres_suda_vidayushego_ispolnitelniy_document>")
    out.WriteLine (adres_suda_vidayushego_ispolnitelniy_document.Text)
  out.WriteLine ("</adres_suda_vidayushego_ispolnitelniy_document>")
  
  out.WriteLine ("<mesto_rassmotreniya_dela>")
    out.WriteLine (mesto_rassmotreniya_dela.Text)
  out.WriteLine ("</mesto_rassmotreniya_dela>")

out.WriteLine ("<ds:Signature xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"">")
  out.WriteLine ("<ds:SignedInfo>")
    out.WriteLine ("<ds:CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    out.WriteLine ("<ds:SignatureMethod Algorithm=""urn:ietf:params:xml:ns:cpxmlsec:algorithms:gostr34102001-gostr3411""/>")
    out.WriteLine ("<ds:Reference Type=""xml"" URI="""">")
        out.WriteLine ("<ds:Transforms>")
            out.WriteLine ("<ds:Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
            out.WriteLine ("<ds:Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments""/>")
        out.WriteLine ("</ds:Transforms>")
        out.WriteLine ("<ds:DigestMethod Algorithm=""urn:ietf:params:xml:ns:cpxmlsec:algorithms:gostr3411""/>")
        out.WriteLine ("<ds:DigestValue>bnhnxBTRSHuT1RODPY/6wWvB9pG2g8aImNeLPeYgnoE=</ds:DigestValue>")
    out.WriteLine ("</ds:Reference>")
  out.WriteLine ("</ds:SignedInfo>")

out.WriteLine ("<ds:SignatureValue>")
out.WriteLine (SignatureValue.Text)
out.WriteLine ("</ds:SignatureValue>")


out.WriteLine ("<ds:KeyInfo>")
    out.WriteLine ("<ds:X509Data>")
        out.WriteLine ("<ds:X509Certificate>")
            out.WriteLine (X509Certificate.Text)
        out.WriteLine ("</ds:X509Certificate>")
    out.WriteLine ("</ds:X509Data>")
out.WriteLine ("</ds:KeyInfo>")
out.WriteLine ("</ds:Signature>")

  out.WriteLine ("</ispolnitelniy_document>")
  
  out.Close
  
CommandButtonSign.Enabled = True
log.Text = log.Text + vbNewLine + "�������� ������������ " + filename2
End Sub

Private Sub CommandButtonSend_Click()
log.Text = log.Text + vbNewLine + "�������� ���������"

End Sub

Private Sub CommandButtonSign_Click()

CommandButtonSend.Enabled = True
log.Text = log.Text + vbNewLine + "�������� ��������"
End Sub

Private Sub CommandButtonValidate_Click()
Rem ���������

If ispolnitelniy_document_nomer.Text = "" Then
MsgBox "�������������� �������� � - ������������ ����", 48
GoTo error
End If


If po_delu_nomer.Text = "" Then
MsgBox "�� ���� � - ������������ ����", 48
GoTo error
End If

log.Text = "�������� ��������"
CommandButtonGenerate.Enabled = True

GoTo finish2
error:
finish2:
End Sub

Private Sub DeleteUdostDocument_Click()
    SaveUdostDocument.Visible = False
    UndoUdostDocument.Visible = False
    Dim i As Integer
          For i = 0 To UdostDocument.ListCount - 1
          If UdostDocument.Selected(i) Then
              Set ud = dictUD(i + 1)
              Exit For
          End If
      Next i

      If i + 1 < dictUD.Count Then
       For j = i + 1 To dictUD.Count - 1
        Set ud1 = dictUD(j + 1)
        dictUD(j).vid = ud1.vid
        dictUD(j).seriya = ud1.seriya
        dictUD(j).nomer = ud1.nomer
        dictUD(j).fio = ud1.fio
        dictUD(j).data_rozhdeniya = ud1.data_rozhdeniya
        dictUD(j).pol = ud1.pol
        dictUD(j).data_vidachi = ud1.data_vidachi
        dictUD(j).kod_podrazdeleniya = ud1.kod_podrazdeleniya
        
       Next j
      End If
      
      UdostDocument.RemoveItem (i)
      
      If UdostDocument.ListCount > 0 Then
        dictUD.Remove (dictUD.Count)
        UdostDocument.Selected(0) = True
      Else
        vid.Text = ""
        seriya.Text = ""
        nomer.Text = ""
        fio.Text = ""
        data_rozhdeniya.Text = ""
        pol.Text = ""
        data_vidachi.Text = ""
        kod_podrazdeleniya.Text = ""
        mesto_rozhdeniya.Text = ""
      End If
      
      disableUdostDocument
End Sub

Private Sub EditUdostDocument_Click()
    SaveUdostDocument.Visible = True
    UndoUdostDocument.Visible = True
    
    addUdostDocument.Visible = False
    EditUdostDocument.Visible = False
    DeleteUdostDocument.Visible = False
    
    udEdit = True
    enableUdostDocument
End Sub

Private Sub SaveUdostDocument_Click()
    Dim i As Integer
    If (udEdit) Then
      For i = 0 To UdostDocument.ListCount - 1
          If UdostDocument.Selected(i) Then
              Set ud = dictUD(i + 1)
              Exit For
          End If
      Next i
        ud.data_rozhdeniya = data_rozhdeniya.Text
        ud.data_vidachi = data_vidachi.Text
        ud.fio = fio.Text
        ud.kod_podrazdeleniya = kod_podrazdeleniya.Text
        ud.mesto_rozhdeniya = mesto_rozhdeniya.Text
        ud.nomer = nomer.Text
        ud.pol = pol.Text
        ud.seriya = seriya.Text
        
        udEdit = False
    Else
    
    ud.data_rozhdeniya = data_rozhdeniya.Text
    ud.data_vidachi = data_vidachi.Text
    ud.fio = fio.Text
    ud.kod_podrazdeleniya = kod_podrazdeleniya.Text
    ud.mesto_rozhdeniya = mesto_rozhdeniya.Text
    ud.nomer = nomer.Text
    ud.pol = pol.Text
    ud.seriya = seriya.Text
    
    ud.vid = vid.Text
    
    UdostDocument.AddItem vid.Text
    dictUD.Add ud
    idxUD = idxUD + 1
    End If
    
    SaveUdostDocument.Visible = False
    UndoUdostDocument.Visible = False
    
    addUdostDocument.Visible = True
    EditUdostDocument.Visible = True
    DeleteUdostDocument.Visible = True
    
    disableUdostDocument
End Sub

Private Sub Form_Load()
idxUD = 0
udEdit = False

dublicat.AddItem ("��")
dublicat.AddItem ("���")

vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti.AddItem ("��")
vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti.AddItem ("���")

podl_nemedl_isp.AddItem ("��")
podl_nemedl_isp.AddItem ("���")

valyuta_dolga.AddItem ("���������� �����")
valyuta_dolga.AddItem ("������������ ������")
valyuta_dolga.AddItem ("����")

status_lica.AddItem ("����������� ����")
status_lica.AddItem ("���������� ����")

pokazatel_tipa_platezha.AddItem ("�� ���������� 1")
pokazatel_tipa_platezha.AddItem ("�� ���������� 2")

dolzhnik_status_lica.AddItem ("����������� ����")
dolzhnik_status_lica.AddItem ("���������� ����")

strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (������������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����������� (��������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� �������� ��������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (��������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� � ������� (������� � �������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (������������ ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (��������� ���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� ������� (����������� ��������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (�������� ���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (�����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������������� ����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ � ����������� (������ � �����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (������������ ���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (����������� ������-����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������-���� (�������-����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������������� (���������� ����������� �������������� � �������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (�������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (�������������� ���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� ����� (��������������� ���������� ��������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������������� ���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������������� ���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������-����� (���������� ������-�����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (������������ ���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (�������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������������� ���������� (������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (�������� ���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ����������� �����������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (��������� ���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����-����� (���������� ����-�����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ���������� (���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� �������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���� ��������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� ����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������������� ���������� ����� (��������������� ���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (��������� �������-��������������� ���������� (�������� �����))")
strana_grazhdanstva_ili_registracii.AddItem ("���������� ����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("�����-���� (���������� �����-����)")
strana_grazhdanstva_ili_registracii.AddItem ("���-������ (���������� ���-������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (����������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (�������� �������-��������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (����������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (�����)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����������� (��������� �����������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (������� ���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (��������� ���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� ������� (���������� ���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (������������ ����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (��������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ���� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (������������ ��������������� ���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (������������ ���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (����������� �����������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� �������� (����� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("��� (����������� �������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (�������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������� ���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� � ����� ������ (����������� ����������� ����� � ����� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (�������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���-���������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (����������� ����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("���-������ (���������� ���-������)")
strana_grazhdanstva_ili_registracii.AddItem ("���-���� � �������� (��������������� ���������� ���-���� � ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� ������ (����������� ���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� ��������� (���������� �������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("����������� ������� (���������� ����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("����-������� � ��������� (����-������� � ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("����-���� � ����� (��������� ����-���� � �����)")
strana_grazhdanstva_ili_registracii.AddItem ("����-����� (����-�����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� �������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("��� (���������� ����� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� ������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������������ ���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������-����� (���������� ������-�����)")
strana_grazhdanstva_ili_registracii.AddItem ("����������� (���������� �����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (����������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� � ������ (���������� �������� � ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (������������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (�������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (���������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (���������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (�������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (��������� ���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("������������ ����� ���������� (������������ ����� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (���������� �������� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (���������� ���������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (���������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("��������������������� ���������� (��������������������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("��� (���������� ���)")
strana_grazhdanstva_ili_registracii.AddItem ("���������� (����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� (������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("���� (���������� ����)")
strana_grazhdanstva_ili_registracii.AddItem ("��������� (����������� ������������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (����������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("���-����� (��������������� ���������������� ���������� ���-�����)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������������� ������ (���������� �������������� ������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (����������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("�������� (����������� ��������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (��������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("������� (������������ ��������������� ���������� �������)")
strana_grazhdanstva_ili_registracii.AddItem ("��� (����-����������� ����������)")
strana_grazhdanstva_ili_registracii.AddItem ("����� ����� (���������� ����� �����)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������)")
strana_grazhdanstva_ili_registracii.AddItem ("������ (������)")


srok_predyavleniya_k_ispolneniyu_razmernost.AddItem ("����")
srok_predyavleniya_k_ispolneniyu_razmernost.AddItem ("������")
srok_predyavleniya_k_ispolneniyu_razmernost.AddItem ("���")

vid.AddItem (" ")
vid.AddItem ("01-������� ���������� ����� ��������� ���������������� ���������")
vid.AddItem ("02-������������� ���������� ����� ��������� ���������������� ���������")
vid.AddItem ("03-������������� � ��������")
vid.AddItem ("04-������������� �������� �������")
vid.AddItem ("05-������� �� ������������ �� ����� ������� �������")
vid.AddItem ("06-������� ����������� ����")
vid.AddItem ("07-������� ����� ������� (�������, ��������, ��������)")
vid.AddItem ("08-��������� �������������, �������� ������ �������� ������")
vid.AddItem ("09-��������������� ������� ���������� ���������� ���������")
vid.AddItem ("10-����������� �������")
vid.AddItem ("11-������������� � ������������ ����������� � ��������� �������� �� ���������� ���������� ��������� �� ��������")
vid.AddItem ("12-��� �� ���������� ���� ��� �����������")
vid.AddItem ("13-������������� ������� � ���������� ���������")
vid.AddItem ("14-��������� ������������� �������� ���������� ���������� ���������")
vid.AddItem ("19-���������� �� ��������� ���������� � ���������� ���������")
vid.AddItem ("20-������������� � �������������� ���������� ������� �� ���������� ���������� ���������")
vid.AddItem ("21-������� ���������� ���������� ���������")
vid.AddItem ("22-����������� ������� ���������� ���������� ���������")
vid.AddItem ("23-������������� � ��������, �������� �������������� ������� ������������ �����������")
vid.AddItem ("24-������������� �������� ��������������� ���������� ���������")
vid.AddItem ("26-������� ������")
vid.AddItem ("27-������� ����� ������� ������")
vid.AddItem ("60-���������, �������������� ���� ����������� �� ����� ����������")
vid.AddItem ("91-���� ���������, ��������������� ����������������� ���������� ���������")

readSud

  For Each ds In dictSudya
    FIO_sudiy.AddItem ds
  Next ds

CommandButtonGenerate.Enabled = False
CommandButtonSend.Enabled = False
CommandButtonSign.Enabled = False
CommandButtonValidate.Enabled = True

disableTransSredstva
disableNedvizhimost
disableUdostDocument

    initCheck
    openDoc docDir + "id_" + templateFileName + ".xml"
    
End Sub

Sub openDoc(ByVal filename As String)
    Dim XDoc As Object
    On Error GoTo error_open_doc
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (filename)
    
    TransSredstva.Clear
    Nedvizhimost.Clear
    UdostDocument.Clear
    
    'Get Document Elements
    Set Lists = XDoc.DocumentElement
    
    For Each ListNode In Lists.ChildNodes
        'Debug.Print "----" & listNode.BaseName & "----" & listNode.Text
        Select Case ListNode.BaseName
                Case "ispolnitelniy_document_nomer"
                    ispolnitelniy_document_nomer.Text = ListNode.Text
                Case "po_delu_nomer"
                    po_delu_nomer.Text = ListNode.Text
                Case "srok_predyavleniya_k_ispolneniyu_znachenie"
                    srok_predyavleniya_k_ispolneniyu_znachenie.Text = ListNode.Text
                Case "srok_predyavleniya_k_ispolneniyu_razmernost"
                    srok_predyavleniya_k_ispolneniyu_razmernost.Text = ListNode.Text
                Case "SolidarnoeVziskanie"
                    SolidarnoeVziskanie.Text = ListNode.Text
                Case "nomer_ekz_ID"
                    nomer_ekz_ID.Text = ListNode.Text
                Case "Gosposhlina"
                    Gosposhlina.Text = ListNode.Text
                Case "data_vidachi"
                    data_vidachi.Text = ListNode.Text
                Case "data_sudebnogo_acta"
                    data_sudebnogo_acta.Text = ListNode.Text
                Case "dublicat"
                    dublicat.Text = ListNode.Text
                Case "vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti"
                    vidan_na_osnovanii_sud_acta_ne_podl_razm_v_seti.Text = ListNode.Text
                Case "data_vsupleniya_v_zs"
                    data_vsupleniya_v_zs.Text = ListNode.Text
                Case "FIO_sudiy"
                    FIO_sudiy.Text = ListNode.Text
                Case "podl_nemedl_isp"
                    podl_nemedl_isp.Text = ListNode.Text
                Case "vid_sushnosti_ispolneniya_ID"
                    vid_sushnosti_ispolneniya_ID.Text = ListNode.Text
                Case "summa_dolga"
                    summa_dolga.Text = ListNode.Text
                Case "valyuta_dolga"
                    valyuta_dolga.Text = ListNode.Text
                Case "ustanovochnaya_chast_sudebnogo_acta"
                    ustanovochnaya_chast_sudebnogo_acta.Text = ListNode.Text
                Case "rezolyutativnaya_chast_sudebnogo_acta"
                    rezolyutativnaya_chast_sudebnogo_acta.Text = ListNode.Text
                Case "status_lica"
                    status_lica.Text = ListNode.Text
                Case "vziskatel"
                    vziskatel.Text = ListNode.Text
                Case "adres"
                    adres.Text = ListNode.Text
                Case "inn"
                    inn.Text = ListNode.Text
                Case "kpp"
                    kpp.Text = ListNode.Text
                Case "ogrn"
                    ogrn.Text = ListNode.Text
                Case "data_registracii"
                    data_registracii.Text = ListNode.Text
                Case "mesto_registracii"
                    mesto_registracii.Text = ListNode.Text
                Case "data_rozhdeniya"
                    data_rozhdeniya.Text = ListNode.Text
'----snils----
'----mesto_rozhdeniya----
                Case "naimenovanie_poluchatelya"
                    naimenovanie_poluchatelya.Text = ListNode.Text
                Case "schet_poluchatelya"
                    schet_poluchatelya.Text = ListNode.Text
                Case "licevoy_schet"
                    licevoy_schet.Text = ListNode.Text
                Case "summa"
                    summa.Text = ListNode.Text
                Case "okato"
                    okato.Text = ListNode.Text
                Case "oktmo"
                    oktmo.Text = ListNode.Text
                Case "inn_poluchatelya"
                    inn_poluchatelya.Text = ListNode.Text
                Case "kpp_poluchatelya"
                    kpp_poluchatelya.Text = ListNode.Text
                Case "naimenovanie_banka_poluchatelya"
                    naimenovanie_banka_poluchatelya.Text = ListNode.Text
                Case "korschet_banka_poluchatelya"
                    korschet_banka_poluchatelya.Text = ListNode.Text
                Case "bik_banka_poluchatelya"
                    bik_banka_poluchatelya.Text = ListNode.Text
                Case "pokazatel_tipa_platezha"
                    pokazatel_tipa_platezha.Text = ListNode.Text
                Case "kbk"
                    kbk.Text = ListNode.Text
                Case "dolzhnik_status_lica"
                    dolzhnik_status_lica.Text = ListNode.Text
                Case "dolzhnik"
                    dolzhnik_dolzhnik.Text = ListNode.Text
                Case "dolzhnik_adres"
                    dolzhnik_adres.Text = ListNode.Text
                Case "dolzhnik_kpp"
                    dolzhnik_kpp.Text = ListNode.Text
                Case "dolzhnik_ogrn"
                    dolzhnik_ogrn.Text = ListNode.Text
                Case "dolzhnik_data_registracii"
                    dolzhnik_data_registracii.Text = ListNode.Text
                Case "strana_grazhdanstva_ili_registracii"
                    strana_grazhdanstva_ili_registracii.Text = ListNode.Text
                Case "dolzhnik_pol"
                    dolzhnik_pol.Text = ListNode.Text
                Case "dolzhnik_data_rozhdeniya"
                    dolzhnik_data_rozhdeniya.Text = ListNode.Text
                Case "dolzhnik_inn"
                    dolzhnik_inn.Text = ListNode.Text
                Case "dolzhnik_mesto_rozhdeniya"
                    dolzhnik_mesto_rozhdeniya.Text = ListNode.Text
                Case "dolzhnik_snils"
                    dolzhnik_snils.Text = ListNode.Text
                Case "mr_actualnost"
                    mr_actualnost.Text = ListNode.Text
                Case "naimenovanie_organizacii_fio_ip"
                    naimenovanie_organizacii_fio_ip.Text = ListNode.Text
                Case "jur_address"
                    jur_address.Text = ListNode.Text
                Case "fact_address"
                    fact_address.Text = ListNode.Text
                Case "naimenovanie_suda_vidayushego_ispolnitelniy_document"
                    naimenovanie_suda_vidayushego_ispolnitelniy_document.Text = ListNode.Text
                Case "adres_suda_vidayushego_ispolnitelniy_document"
                    adres_suda_vidayushego_ispolnitelniy_document.Text = ListNode.Text
                Case "mesto_rassmotreniya_dela"
                    mesto_rassmotreniya_dela.Text = ListNode.Text
                Case "UdostDocumentList"
                        For Each listNode1 In ListNode.ChildNodes
                            Set ud = New UdostDocument
                            For Each ListNode2 In listNode1.ChildNodes
                            Select Case ListNode2.BaseName
                                Case "vid"
                                    ud.vid = ListNode2.Text
                                Case "seriya"
                                    ud.seriya = ListNode2.Text
                                Case "nomer"
                                    ud.nomer = ListNode2.Text
                                Case "fio"
                                    ud.fio = ListNode2.Text
                                Case "data_rozhdeniya"
                                    ud.data_rozhdeniya = ListNode2.Text
                                Case "pol"
                                    ud.pol = ListNode2.Text
                                Case "data_vidachi"
                                    ud.data_vidachi = ListNode2.Text
                                Case "kod_podrazdeleniya"
                                    ud.kod_podrazdeleniya = ListNode2.Text
                                Case "mesto_rozhdeniya"
                                    ud.mesto_rozhdeniya = ListNode2.Text
                            End Select
                            Next ListNode2
                            dictUD.Add ud
                            UdostDocument.AddItem ud.vid
                        Next listNode1
                    Case "NedvizhimostList"
                        For Each listNode1 In ListNode.ChildNodes
                            Set nd = New Nedvizhimost
                            For Each ListNode2 In listNode1.ChildNodes
                            Select Case ListNode2.BaseName
                                Case "Actualnost"
                                    nd.Actualnost = ListNode2.Text
                                Case "Naimenovanie"
                                    nd.Naimenovanie = ListNode2.Text
                                Case "Ploshad"
                                    nd.Ploshad = ListNode2.Text
                                Case "UslNomer"
                                    nd.UslNomer = ListNode2.Text
                                Case "InvNomer"
                                    nd.InvNomer = ListNode2.Text
                                Case "KadastrNomer"
                                    nd.KadastrNomer = ListNode2.Text
                                Case "TochAdres"
                                    nd.TochAdres = ListNode2.Text
                            End Select
                            Next ListNode2
                            dictND.Add nd
                            Nedvizhimost.AddItem nd.Naimenovanie
                        Next listNode1
                    Case "TransSredstvaList"
                        For Each listNode1 In ListNode.ChildNodes
                            Set td = New TransSredstva
                            For Each ListNode2 In listNode1.ChildNodes
                            Select Case ListNode2.BaseName
                                Case "Actualnost"
                                    td.Actualnost = ListNode2.Text
                                Case "Kategoriya"
                                    td.Kategoriya = ListNode2.Text
                                Case "Marka"
                                    td.Marka = ListNode2.Text
                                Case "Model"
                                    td.Model = ListNode2.Text
                                Case "Cvet"
                                    td.Cvet = ListNode2.Text
                                Case "GosZnak"
                                    td.GosZnak = ListNode2.Text
                                Case "VIN"
                                    td.VIN = ListNode2.Text
                                Case "NDvig"
                                    td.NDvig = ListNode2.Text
                                Case "KodPodr"
                                    td.KodPodr = ListNode2.Text
                                Case "GodVipuska"
                                    td.GodVipuska = ListNode2.Text
                            End Select
                            Next ListNode2
                            dictTD.Add td
                            TransSredstva.AddItem td.Kategoriya
                        Next listNode1


        End Select
    Next ListNode
    
    If (UdostDocument.ListCount > 0) Then
        UdostDocument.Selected(0) = True
    End If
    
    If (Nedvizhimost.ListCount > 0) Then
        Nedvizhimost.Selected(0) = True
    End If
    
    If (TransSredstva.ListCount > 0) Then
        TransSredstva.Selected(0) = True
    End If
        
    Call Show
error_open_doc:
End Sub

Private Sub setVid(ByVal value As String)
vid.Text = value
'Dim selvalue As String
'For Each l In vid.List
' If InStr(l, value) > 0 Then
'    selvalue = l
'    Exit For
' End If
' Next l
' vid.Text = selvalue
End Sub

Private Sub TransSredstva_Click()
    TransSredstva_Change
End Sub

Private Sub TransSredstva_KeyPress(KeyAscii As Integer)
    TransSredstva_Change
End Sub

Private Sub UdostDocument_Click()
    UdostDocument_Change
End Sub

Private Sub UdostDocument_KeyPress(KeyAscii As Integer)
    UdostDocument_Change
End Sub

Private Sub UndoNedvizhimost_Click()
    ndEdit = False
    UndoNedvizhimost.Visible = False
    SaveNedvizhimost.Visible = False
    
    AddNedvizhimost.Visible = True
    EditNedvizhimost.Visible = True
    DeleteNedvizhimost.Visible = True
    
    If Nedvizhimost.ListCount > 0 Then
        Nedvizhimost.Selected(0) = True
    Else
        Actualnost.Text = ""
        Naimenovanie.Text = ""
        Ploshad.Text = ""
        UslNomer.Text = ""
        InvNomer.Text = ""
        KadastrNomer.Text = ""
        TochAdres.Text = ""
    End If
    
    disableNedvizhimost
End Sub

Private Sub disableNedvizhimost()
    Actualnost.Enabled = False
    Naimenovanie.Enabled = False
    Ploshad.Enabled = False
    UslNomer.Enabled = False
    InvNomer.Enabled = False
    KadastrNomer.Enabled = False
    TochAdres.Enabled = False
End Sub

Private Sub enableNedvizhimost()
    Actualnost.Enabled = True
    Naimenovanie.Enabled = True
    Ploshad.Enabled = True
    UslNomer.Enabled = True
    InvNomer.Enabled = True
    KadastrNomer.Enabled = True
    TochAdres.Enabled = True
End Sub

Private Sub UndoTransSredstva_Click()
    tdEdit = False
    UndoTransSredstva.Visible = False
    SaveTransSredstva.Visible = False
    
    AddTransSredstva.Visible = True
    EditTransSredstva.Visible = True
    DeleteTransSredstva.Visible = True
    
    If TransSredstva.ListCount > 0 Then
        TransSredstva.Selected(0) = True
    Else
        TS_Actualnost.Text = ""
        Kategoriya.Text = ""
        Marka.Text = ""
        Model.Text = ""
        Cvet.Text = ""
        GosZnak = ""
        VIN = ""
        NDvig = ""
        KodPodr = ""
        GodVipuska = ""
    End If
    
    disableTransSredstva
End Sub

Private Sub disableTransSredstva()
    TS_Actualnost.Enabled = False
    Kategoriya.Enabled = False
    Marka.Enabled = False
    Model.Enabled = False
    Cvet.Enabled = False
    GosZnak.Enabled = False
    VIN.Enabled = False
    NDvig.Enabled = False
    KodPodr.Enabled = False
    GodVipuska.Enabled = False
End Sub

Private Sub enableTransSredstva()
    TS_Actualnost.Enabled = True
    Kategoriya.Enabled = True
    Marka.Enabled = True
    Model.Enabled = True
    Cvet.Enabled = True
    GosZnak.Enabled = True
    VIN.Enabled = True
    NDvig.Enabled = True
    KodPodr.Enabled = True
    GodVipuska.Enabled = True
End Sub

Private Sub UndoUdostDocument_Click()
    udEdit = False
    UndoUdostDocument.Visible = False
    SaveUdostDocument.Visible = False
    If UdostDocument.ListCount > 0 Then
        UdostDocument.Selected(0) = True
    Else
        vid.Text = ""
        seriya.Text = ""
        nomer.Text = ""
        fio.Text = ""
        data_rozhdeniya.Text = ""
        pol.Text = ""
        data_vidachi.Text = ""
        kod_podrazdeleniya.Text = ""
        mesto_rozhdeniya.Text = ""
    End If
    
    disableUdostDocument
End Sub

Private Sub disableUdostDocument()
    vid.Enabled = False
    seriya.Enabled = False
    nomer.Enabled = False
    fio.Enabled = False
    data_rozhdeniya.Enabled = False
    pol.Enabled = False
    data_vidachi.Enabled = False
    kod_podrazdeleniya.Enabled = False
    mesto_rozhdeniya.Enabled = False
    
    addUdostDocument.Enabled = True
    EditUdostDocument.Enabled = True
    DeleteUdostDocument.Enabled = True
    
End Sub

Private Sub enableUdostDocument()
    vid.Enabled = True
    seriya.Enabled = True
    nomer.Enabled = True
    fio.Enabled = True
    data_rozhdeniya.Enabled = True
    pol.Enabled = True
    data_vidachi.Enabled = True
    kod_podrazdeleniya.Enabled = True
    mesto_rozhdeniya.Enabled = True
End Sub

Private Sub initCheck()
    
    Dim sFilePath As String
 
    If Dir(dd1, vbDirectory) = "" Then
          MsgBox "����������� ���� D:\ - ���������� � ������ ���������"
          GoTo finishInitCheck
    End If
    
    If Dir(dd2, vbDirectory) = "" Then
          MkDir dd2
    End If
    
    If Dir(docDir, vbDirectory) = "" Then
          MkDir docDir
    End If
        
    
    sFilePath = docDir + "id_" + templateFileName + ".xml"
    
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
                             naimenovanie_suda_vidayushego_ispolnitelniy_document.AddItem sName
                        End If
                    Case 2
                        If (sD.address = "") Then
                            sD.address = one_line(C)
                            If (sD.address <> saddr) Then
                                saddr = sD.address
                                adres_suda_vidayushego_ispolnitelniy_document.AddItem saddr
                                
                            End If
                        End If
                    Case 3
                        If (one_line(C) <> "") Then
                            sudyaName = one_line(C)
                            dictSudya.Add sudyaName
                            FIO_sudiy.AddItem sudyaName
                        End If
                        sD.sudyaName.Add sudyaName
                    End Select
            
            Next C
        End If
    Next R


End Sub
