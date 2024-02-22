VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAsientosIDDOC 
   Caption         =   "Asientos diferidos"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16470
   Icon            =   "frmAsientosIDDOC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   16470
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   825
      Left            =   2085
      TabIndex        =   2
      Top             =   90
      Width           =   990
      _extentx        =   1746
      _extenty        =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid gAsientos 
      Height          =   10770
      Left            =   120
      TabIndex        =   1
      Top             =   1065
      Width           =   8280
      _cx             =   14605
      _cy             =   18997
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   150
      Picture         =   "frmAsientosIDDOC.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   915
   End
   Begin VSFlex7LCtl.VSFlexGrid gSumariza 
      Height          =   10830
      Left            =   8520
      TabIndex        =   3
      Top             =   1020
      Width           =   7815
      _cx             =   13785
      _cy             =   19103
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Gestion.ucXls ucXls2 
      Height          =   825
      Left            =   8535
      TabIndex        =   4
      Top             =   75
      Width           =   975
      _extentx        =   1720
      _extenty        =   1455
   End
End
Attribute VB_Name = "frmAsientosIDDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cWhereExterna As String

Public Function MostrarDif(sWhere As String)
cWhereExterna = sWhere
Me.Show
cmdMostrar_Click
End Function


Private Sub cmdMostrar_Click()
Dim rsAsientos As New ADODB.Recordset, i As Long, cConsul As String, pActual As String, pPrueba As String, rInicio As Long, gEncontro As Boolean
Dim rsMayor As New ADODB.Recordset, x As Long
g_ini
g_ini2
cConsul = "SELECT * FROM ASIENTOS WHERE ACTIVO=1 AND " & cWhereExterna & " "
'cConsul = "SELECT * FROM ASIENTOS WHERE IDDOC in " _
& " (30476,30461,30462,30608,30455,30369,30514,30572,30508,30518,30517,30672,30529,30530,30531,30581,30550,30621,30549,30680,30646,30751,30512,30555,30556,30557,30558,30561,30570,30580,30559,30456,30560,30688,30585,30578,30750,30644,30587,30586,30603,30643,30601,30604,30605,30602,30674,30614,30588,30589,30551,30610,30645,30611,30609,30622,30618,30619,30832,30620,30642,30763,30648,30635,30636,30639,30673,30649,30981,30670,30809,30669,30573,30574,30575,30730,30769,30770,30771,30768,30682,30703,30683,30753,30695,30698,30686,30689,30691,30692,30766,30693,30690,30699,30694,30726,30767,30697,30758,30700,30749,30679,30951,31150,30800,30704,30705,30752,30710,30708,30754,30747,30748,30755,30756,30757,30860,30861,30863,30864,30696,30728,30787,30761,30762,30950,30764,30836,30772,30773,30775,30731,30774,30781,30702,30890,30782,30783,30784,30785,30786,30838,30949,30777,30845,30846,30847,30778,30779,30789,30780,30790,30834,30801,30908,30929,30795,30788,30822,30881,30803,30843,30848,30831, " _
& "  30840,30793,30796,30797,30853,30802,30810,30804,30805,30814,30815,30792,30818,30816,30817,30825,30844,30842,30866,30867,30865,30819,30823,30841,30849,30883,30824,30915,30857,31149,30851,30859,30852,31029,30854,30855,30858,30889,30893,30876,30877,30878,30879,30888,31102,30972,30912,30920,30907,30969,30982,30906,30909,30899,30896,30897,30903,30910,30900,30913,30946,30928,30898,30902,30904,30894,30895,30924,30901,30984,30919,30943,30905,30917,30918,30926,31004,30916,30927,30914,30925,30965,30921,30923,30922,30941,30940,30933,30934,30935,30936,30937,30938,30974,31101,30966,30968,30963,30964,30955,30975,31018,30962,30971,30960,30961,30986,31017,31016,30983,30930,30932,30931,31100,30996,30978,30979,30980,30997,30991,30976,30977,31015,30994,31014,31003,30987,31001,31002,30992,31013,31050,30999,31000,30995,30993,30989,31022,30998,31010,31011,31012,31148,31008,31007,31009,31019,31006,31068,31027,31005,31025,31020,31024,31026,31051,31089,31092,31055,31070,31056,31057,31235,31047, " _
& "  31048,31117,31128,31112,31113,31114,31116,31103,31110,31111,31084,31082,31088,31168,31183,31211,31091,31085,31086,31087,31079,31081,31107,31207,31175,31176,31177,31205,31094,31122,31095,31098,31097,31108,31109,31104,31105,31106,31173,31178,31096,31143,31284,31285,31138,31244,31121,31129,31118,31120,31130,31125,31126,31127,31323,31136,31137,31132,31133,31134,31151,31135,31184,31144,31263,31264,31152,31153,31322,31142,31262,31182,31987,31239,31162,31140,31321,31384,31268,31455,31637,31989,31648,31793,31216,31202,31218,31174,31160,31155,31171,31169,31318,31320,31154,31283,31163,31167,31166,31156,31157,31158,31159,31232,31229,31161,31164,31165,31181,31988,31170,31172,31246,31412,31131,31647,31217,31203,31315,31180,31230,31214,31228,31266,31986,31227,31231,31204,31212,31213,31206,31215,31209,31208,31201,31949,31951,31146,31145,31185,31147,31331,31330,31329,31247,31226,31270,31282,31237,31294,31236,31238,31245,31233,31265,31241,31261,31281,31242,31253,31260,31308,31316,31257, " _
& "  31259,31289,31258,31256,31280,31314,31240,31277,31278,31279,31291,31312,31557,31465,31500,31292,31304,31305,31274,31272,31463,31464,31566,31288,31275,31313,31568,31567,31559,31309,31326,31310,31307,31298,31548,31414,31541,31299,31301,31286,31302,31324,31300,31325,31327,31336,31539,31332,31601,31333,31334,31335,31337,31403,31395,31394,31538,31565,31545,31564,31341,31342,31311,31549,31535,31400,31344,31353,31401,31569,31570,31359,31534,31405,31583,31584,31343,31346,32006,31345,31582,31588,31424,31410,31537,31352,31347,31379,31402,31369,31371,31723,31725,31729,31381,31732,31276,31362,31366,31503,31416,31370,31724,31562,31387,31542,31586,31372,31374,31726,31486,31485,31544,31476,31380,31437,31439,31438,31377,31510,31508,31386,31552,31585,31358,31805,31407,31385,31720,31721,31722,31587,31391,31392,31511,31393,31417,31409,31399,31502,31396,31561,31719,31413,31397,31404,31115,31406,31432,31509,31711,31457,31452,31428,31408,31442,31415,31578,31579,31440,31713,31577,31717,31478, " _
& "  31513,31431,31514,31512,31448,31430,31434,31450,31558,31451,31712,31714,31715,31716,31435,31436,31580,31581,31441,31446,31444,31718,31704,31622,31449,31575,31576,31477,31475,31458,31491,31459,31490,31499,31573,31574,31804,31592,31487,31474,31480,31520,31488,31976,31484,31563,31471,31613,31614,31615,31470,31469,31522,31489,31703,31494,31572,31630,31536,31516,31505,31685,31669,31533,31657,31498,31519,31523,31521,31598,31555,31526,31553,31554,31525,31629,31527,31529,31737,31599,31530,31506,31653,31702,31641,31531,31532,31571,31662,31602,31633,31611,31603,31604,31600,32005,31612,31626,31606,31610,31608,31643,31689,31642,31650,31639,31640,31736,31644,31758,31759,31734,31832,31649,31651,31652,31806,31881,31682,31654,31655,31670,31699,31770,31733,31663,31661,31656,31664,31665,31680,31679,31744,31773,31678,31697,31677,31681,31690,31903,31952,31742,31688,31692,31691,31693,31695,31694,31698,31735,31700,31701,31750,32004,31902,31755,31749,31905,31795,31780,31756,31761,31754,31753, " _
& "  31796,31867,31760,31775,31768,31815,31816,31817,31769,31776,31774,31790,31807,31880,31784,31789,31856,31786,31777,31781,31785,31783,31792,31787,31788,31822,31829,32330,32331,31797,31813,31799,31798,31890,31812,31811,31897,31895,31896,31810,31808,31809,31886,31821,31882,31839,31824,31827,31828,31830,31898,31814,31823,32157,32158,32159,31831,31836,31838,31858,31953,31879,31874,31906,31859,31862,32329,31914,31924,31871,31878,31861,31873,31870,31954,31875,31872,31876,31877,31889,31884,31885,31888,31913,31887,31901,31900,31910,32013,31982,31916,31907,31909,31908,31936,31919,31920,31960,31961,31921,31934,31935,31962,31967,31942,31973,31946,31941,31981,32143,31983,32096,31943,32087,32042,31958,31998,31970,31944,31947,31948,31990,31956,31950,31957,31959,32016,32108,32095,32109,31972,32135,31964,31971,31984,31996,31979,31974,31975,31977,31978,32349,32101,31993,31994,31995,32130,32128,31997,32000,32001,32011,32019,32008,32007,32002,32046,32003,32020,32197,32882,32050,32018,32017, " _
& "  33858,32131,32132,32149,32043,32023,32070,32027,32025,32026,32086,32024,32044,32048,32047,32049,32145,32092,32328,32091,32153,32069,32068,32045,32066,32072,32073,32163,32088,32247,32084,32097,32076,32085,32121,32106,32090,32089,32094,32113,32112,32334,32102,32119,32111,32093,32156,32110,32107,32134,32105,32161,32115,32127,32116,32129,32125,32126,32133,32137,32138,32142,32147,32336,32146,32150,32235,32233,32234,32236,32151,32258,32152,32154,32168,32171,32172,32232,32248,32295,32380,32241,32881,32883,32217,32239,32185,32238,32195,32198,32187,32201,32184,32335,32224,32199,32202,32204,32200,32212,32213,32222,32226,32214,32205,32219,32218,32229,32215,32216,32173,32220,32243,32223,32221,32225,32262,32350,32245,32332,32260,32327,32242,32246,32244,32261,32259,32256,32271,32337,32384,32265,32333,32340,32338,32339,32313,32266,32268,32269,32267,32272,32270,32296,33455,32323,32354,32364,32290,32425,32293,32283,32289,32288,32724,32725,32424,32419,32426,32398,32409,32410,32352,32427, " _
& "  32415,32723,32294,32291,32292,32574,32320,32321,32322,32355,32307,32406,32308,32465,32346,32353,32357,32347,32382,32356,32376,32728,32729,32366,32386,32367,32368,32475,32374,32375,32377,32478,32524,32453,32416,32461,32463,32462,32378,32413,32399,32396,32405,32456,32474,32404,32401,32402,32397,32407,32408,32460,32497,32726,32727,32414,32439,32559,32502,33452,32306,32486,32411,32487,32429,32430,32437,32480,32428,32412,32457,32436,32458,32443,32452,32454,32455,32494,32500,32459,32495,32464,32481,32482,32721,32722,32493,32477,32479,32492,32736,32718,32720,32521,32499,32490,32491,32503,32523,32592,32553,32552,32551,32719,32554,32536,32522,32509,32508,32520,32516,32517,32518,32519,32678,32679,32768,32525,32529,32530,32526,32589,32556,32537,32533,32532,32531,32527,32716,32587,32567,32563,32534,32566,32569,32568,32535,32539,32582,32540,32541,32572,32615,32555,32715,32564,32579,32580,32581,32570,32605,32573,32571,32600,32584,32585,32606,32575,32607,32634,32612,32588,32645,32662, " _
& "  32664,32663,32661,32597,32598,32591,32616,32599,32595,32602,32830,32608,32604,32596,32601,32633,32613,32614,32610,32611,32621,32677,32636,32714,32676,32623,32622,32617,32643,32652,32734,32735,32626,32627,32628,32629,32624,32649,32625,32658,32630,32660,32641,32632,32635,32638,32656,32639,32642,32640,32745,32653,32620,32647,32646,32654,32708,32651,32650,32655,32854,32713,32668,33280,32669,32670,32671,32672,32712,32733,32704,32706,32843,32696,32702,32731,32828,33045,33068,33073,32730,32732,32689,32697,32701,32852,32707,32709,32710,32711,32764,32743,32919,32769,32749,32832,32758,32755,32744,32748,32737,32738,32786,32747,32752,32754,32746,32750,32751,32756,32789,32788,32845,32772,32853,32759,32766,32765,32767,32787,32778,32775,32780,32466,32773,32774,32833,32860,32801,32779,32797,32790,32806,32798,32848,32849,32847,32850,32793,32799,32791,32794,32795,32792,32815,32834,32808,32796,32800,32807,32814,32809,32810,32811,32805,32861,32915,32916,32812,32831,32827,32822,33024,32920, " _
& "  32835,32846,32876,32837,32858,32836,32840,32839,32974,32838,32842,32841,32855,32877,33044,32864,32862,32857,32851,32859,32863,32856,32886,32898,32914,32892,32885,32887,33052,32956,32893,32894,32897,32911,32951,32910,32913,32929,33018,32912,32930,32931,32917,32921,32928,32927,32973,32961,32944,32958,32957,32959,32932,32936,32939,32935,32954,32953,32960,32952,32995,32955,32993,32991,32996,32997,33012,33013,32967,32965,32988,32968,32969,32970,32975,32990,32979,32978,33006,32971,32976,33027,33047,32985,32984,32980,32981,32983,32998,32989,32987,32982,33017,32999,33011,33002,33003,33000,33009,33001,33010,33107,33133,33126,33128,33129,33016,33025,33019,33043,33077,33091,33020,33021,33081,33080,33023,33026,33036,33037,33076,33097,33098,33099,33096,33075,33131,33079,33078,33094,33088,33090,33203,33089,33105,33102,33127,33103,33095,33093,33111,33092,33101,33100,33120,33104,33108,33147,33146,33174,33159,33109,33110,33113,33137,33118,33106,33140,33112,33156,33150,33195,33117,33190, " _
& "  33191,33148,33149,33125,33122,33123,33124,33130,33132,33144,33136,33135,33138,33139,33141,33142,33145,33206,33207,33209,33214,33211,33134,33189,33342,33157,33143,33154,33153,33160,33161,33269,33152,33151,33172,33155,33175,33173,33176,33230,33210,33177,33184,33188,33186,33227,33496,33497,33498,33276,33204,33208,33219,33198,33221,33196,33194,33193,33158,33502,33503,33504,33505,33506,33507,33508,33509,33510,33500,33200,33199,33202,33201,33217,33218,33225,33224,33222,33286,33226,33285,33228,33251,33252,33253,33256,33345,33255,33257,33264,33294,33293,33292,33284,33259,33271,33262,33295,33296,33281,33263,33344,33291,33283,33341,33287,33322,33334,33389,33290,33288,33289,33323,33298,33324,33436,33327,33326,33330,33351,33325,33331,33328,33335,33353,33337,33338,33339,33340,33343,33350,33380,33347,33346,33423,33379,33349,33368,33367,33415,33382,33356,33361,33355,33354,33501,33525,33373,33422,33417,33418,33357,33360,33499,33419,33421,33371,33370,33362,33372,33439,33475,33474,33473, " _
& "  33411,33420,33442,33390,33391,33408,33403,33407,33461,33412,33409,33413,33414,33387,33405,33406,33410,33433,33460,33416,33480,33481,33512,33427,33426,33454,33435,33429,33428,33386,33444,33434,33443,33487,33440,33479,33430,33449,33462,33459,33445,33446,33495,33456,33458,33467,33469,33470,33463,33437,33478,33516,33494,33514,33515,33517,33518,33751,33484,33492,33485,33438,33486,33489,33488,33493,33490,33511,33524,33521,33513,33519,33522,33530,33520,33523,33546,33566,33567,33564,33565,33607,33541,33545,33619,33563,33551,33543,33724,33697,33528,33526,33527,33581,33582,33605,33544,33568,33595,33557,33529,33531,33532,33540,33548,33547,33649,33651,33630,33569,33571,33570,33558,33559,33549,33555,33604,33553,33556,33561,33560 " _
& ") ORDER BY FECHA"

With rsAsientos
    .Open cConsul, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            g_add !Fecha, " (" & !NroAsiento & ")", !concepto, "", "", ""
            cConsul = "SELECT C.DESCRIPCION,M.* FROM MAYOR M INNER JOIN CUENTAS C ON M.CUENTA=C.CUENTA WHERE IDASIENTO=" & !idAsiento
            rsMayor.Open cConsul, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            If rsMayor.EOF And rsMayor.BOF Then GoTo salto
            rsMayor.MoveFirst
            For x = 0 To rsMayor.RecordCount - 1
                g_add "", rsMayor!Cuenta, rsMayor!DESCRIPCION, Format(s2n(rsMayor!Debe), "0.00"), Format(s2n(rsMayor!haber), "0.00"), ""
                rsMayor.MoveNext
            Next
salto:
            Set rsMayor = Nothing
            g_add "", "", "", "", "", ""

            .MoveNext
        Next
    End If
End With
Set rsAsientos = Nothing

With gAsientos
    pActual = ""
    rInicio = 0
    For i = 1 To .rows - 1
        If Trim(.TextMatrix(i, 0)) > "" Then
            pPrueba = Format(CDate(.TextMatrix(i, 0)), "MMMM")
            If pPrueba <> pActual Then
                pActual = pPrueba
                g_add2 "", "", "", "", ""
                g_add2 "PERIODO", "", Nombre_Mes(CDate(.TextMatrix(i, 0))), "", ""
                rInicio = gSumariza.rows - 1
            End If
        End If
        If InStr(.TextMatrix(i, 1), "(") Then
        Else
            If Trim(.TextMatrix(i, 1)) > "" Then
                gEncontro = False
                For x = rInicio To gSumariza.rows - 1
                    If Trim(.TextMatrix(i, 1)) = Trim(gSumariza.TextMatrix(x, 1)) Then
                        gSumariza.TextMatrix(x, 3) = s2n(gSumariza.TextMatrix(x, 3)) + s2n(.TextMatrix(i, 3))
                        gSumariza.TextMatrix(x, 4) = s2n(gSumariza.TextMatrix(x, 4)) + s2n(.TextMatrix(i, 4))
                        gEncontro = True
                        Exit For
                    End If
                Next
                If gEncontro Then
                Else
                    g_add2 "", .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4)
                End If
            End If
        End If
        
    Next
End With
End Sub

Private Function Nombre_Mes(f As Date, Optional limitador As String = "") As String
Dim Mes, Anio, mes_Letra As String
    Anio = Mid(Year(f), 3, 2)
    Mes = Month(f)
    Select Case Mes
        Case 1:
            mes_Letra = "Enero"
        Case 2:
            mes_Letra = "Febrero"
        Case 3:
            mes_Letra = "Marzo"
        Case 4:
            mes_Letra = "Abril"
        Case 5:
            mes_Letra = "Mayo"
        Case 6:
            mes_Letra = "Junio"
        Case 7:
            mes_Letra = "Julio"
        Case 8:
            mes_Letra = "Agosto"
        Case 9:
            mes_Letra = "Septiembre"
        Case 10:
            mes_Letra = "Octubre"
        Case 11:
            mes_Letra = "Noviembre"
        Case 12:
            mes_Letra = "Diciembre"
    End Select
    Nombre_Mes = mes_Letra & " / " & Anio
    If limitador > "" Then Nombre_Mes = mes_Letra & limitador & Anio
End Function

Private Function g_ini()
With gAsientos
    .rows = 1
    .cols = 0
    .cols = 7
    .TextMatrix(0, 0) = " FECHA "
    .TextMatrix(0, 1) = " CUENTA "
    .TextMatrix(0, 2) = " DESCRIPCION "
    .TextMatrix(0, 3) = " DEBE "
    .TextMatrix(0, 4) = " HABER "
    .TextMatrix(0, 5) = " OBS "
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 0
    .ColWidth(6) = 0
End With
End Function

Private Function g_ini2()
With gSumariza
    .rows = 1
    .cols = 0
    .cols = 5
    .TextMatrix(0, 0) = " PERIODO "
    .TextMatrix(0, 1) = " CUENTA "
    .TextMatrix(0, 2) = " DESCRIPCION "
    .TextMatrix(0, 3) = " DEBE "
    .TextMatrix(0, 4) = " HABER "
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
End With
End Function

Private Function g_add(a As String, b As String, C As String, d As String, e As String, f As String)
Dim rou As Long
With gAsientos
    .AddItem " "
    rou = .rows - 1
    .TextMatrix(rou, 0) = a
    .TextMatrix(rou, 1) = b
    .TextMatrix(rou, 2) = C
    If s2n(d) = 0 Then d = ""
    .TextMatrix(rou, 3) = d
    If s2n(e) = 0 Then e = ""
    .TextMatrix(rou, 4) = e
    .TextMatrix(rou, 5) = f
    If a > "" And C > "" And InStr(b, "(") Then  '   B = "" Then
        .cell(flexcpFontBold, rou, 0, rou, 4) = True
    End If
    If a > "" And b > "" And C > "" And d > "" And e > "" And f > "" Then
        .cell(flexcpFontItalic, rou, 0, rou, 4) = True
    End If
    
End With
End Function

Private Function g_add2(a As String, b As String, C As String, d As String, e As String)
Dim rou As Long
With gSumariza
    .AddItem " "
    rou = .rows - 1
    .TextMatrix(rou, 0) = a
    .TextMatrix(rou, 1) = b
    .TextMatrix(rou, 2) = C
    If s2n(d) = 0 Then d = ""
    .TextMatrix(rou, 3) = d
    If s2n(e) = 0 Then e = ""
    .TextMatrix(rou, 4) = e
    If a > "" And C > "" And b = "" Then
        .cell(flexcpFontBold, rou, 0, rou, 4) = True
    End If
End With
End Function

Private Sub Form_Load()
ucXls1.ini gAsientos, "C:\AsientosDiferidos.xls"
ucXls2.ini gSumariza, "C:\AsientosDiferidosPeriodo.xls"
End Sub


Private Sub gAsientos_DblClick()
Dim rr As Long, cc As Long
Dim vAsiento As Long, tmp
rr = gAsientos.Row
cc = gAsientos.Col
If InStr(gAsientos.TextMatrix(rr, cc), "(") Then
    If InStr(gAsientos.TextMatrix(rr, cc), ")") Then
        tmp = Replace(gAsientos.TextMatrix(rr, cc), "(", "")
        tmp = Replace(tmp, ")", "")
        vAsiento = s2n(tmp)
        frmAsientoManual.mostrar s2n(obtenerDeSQL("select iddoc from asientos where activo=1 and nroasiento=" & vAsiento & " order by ejercicio desc"))
    End If
End If


End Sub
