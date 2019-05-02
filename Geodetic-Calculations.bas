Attribute VB_Name = "SurveyFunctions"
''''''''''''''''''''''''''''''''''''''''''''''''
'''''''   Code written by Kent Wheeler    ''''''
'''''''  kent.wheeler@landgate.wa.gov.au  ''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Function DistSDev(DIST)
DistSDev = (5 + 0.005 * DIST) / 1000
End Function
Function AngSDev(DIST)
If DIST > 50 Then
    AngSDev = (((WorksheetFunction.Degrees(Atn(0.0045 / DIST)) * 3600) ^ 2) + 4) ^ 0.5
Else
    AngSDev = WorksheetFunction.Degrees(Atn(0.003 / DIST)) * 3600
End If
End Function
Function FindJobNumber(str)
FindJobNumber = ""
For i = 1 To Len(str) - 8
    tstStr = Mid(str, i, 8)
    If IsNumeric(tstStr) = True And InStr(tstStr, ".") = 0 Then FindJobNumber = tstStr
Next i
End Function
Function strLen(ByVal str, ByVal StringLength)
'either truncate the string to a length that can be used or lengthen it to the requested length with spaces
If Len(str) > Abs(StringLength) Then str = Right(str, Abs(StringLength))
If StringLength < 0 Then
    strLen = Right("                                                                             " & str, Abs(StringLength))
Else
    strLen = Left(str & "                                                                             ", StringLength)
End If
End Function

Function HMS(Ang, ds)
'Input dd.ddddd...
'Output dd mm s.sss -  Number of decimal places set by 'ds'
M = Int(Abs(Ang) * 3600 / 60)
M = M Mod 60
s = (Abs(Ang) * 3600) - (Int(Abs(Ang)) * 3600) - (M * 60)
If Ang < 0 Then HMS = -Int(Abs(Ang)) & Format(M, " 00") & Format(s, " 00." & Left("0000000000000000000000000000", ds))
If Ang >= 0 Then HMS = Int(Ang) & Format(M, " 00") & Format(s, " 00." & Left("0000000000000000000000000000", ds))
End Function
Function hms2hp(Ang)
Dim a
'Input dd mm s.sss
'Output dd.mmsssss
Ang = Replace(Ang, "S", "-")
Ang = Replace(Ang, "W", "-")
Ang = Replace(Ang, "E", "")
Ang = Replace(Ang, "N", "")
a = Split(Trim(Ang), " ")
If Ang <> "" Then hms2hp = a(0) & "." & a(1) & Replace(a(2), ".", "")
End Function
Function HP2dd(Ang)
'Input  dd.mmsssss
'Output dd mm s.sss
a = Split(Ang, ".")
mm = Val(Left(a(1), 2))
SS = Val(Mid(a(1), 3, 2) & "." & Mid(a(1), 5, Len(a(1)) - 4))
If Ang <> "" And Ang < 0 Then HP2dd = Val(a(0)) - mm / 60 - SS / 3600
If Ang <> "" And Ang >= 0 Then HP2dd = Val(a(0)) + mm / 60 + SS / 3600
End Function
Function HP2hms(Ang)
'Input  dd.mmsssss
'Output dd mm SS.sss
a = Split(Ang, ".")
mm = Left(a(1), 2)
SS = Mid(a(1), 3, 2) & "." & Mid(a(1), 5, Len(a(1)) - 4)
If Ang <> "" Then HP2hms = a(0) & " " & mm & " " & SS
End Function
Function PreZs(num, zs)
While Len(num) < zs
    num = "0" & num
Wend
PreZs = num
End Function
Function DD(Ang)
'Input dd mm s.sss
'Output dd.ddddd...

While InStr(Ang, "  ") <> 0
    Ang = Trim(Replace(Ang, "  ", " "))
Wend
Ang = Replace(Ang, "°", " ")
Ang = Replace(Ang, "'", " ")
Ang = Replace(Ang, """", "")
Ang = Replace(Ang, "E", "")
Ang = Replace(Ang, "N", "")
If Ang = "" Then GoTo ed
i = 1
While IsNumeric(Mid(Ang, i, 1)) = False And i <= Len(Ang)
    i = i + 1
Wend
aAng = Split(Right(Ang, Len(Ang) - i + 1))
deg = Abs(Val(aAng(0))) + Val(aAng(1)) / 60 + Val(aAng(2)) / 3600
If InStr(Ang, "S") > 0 Or InStr(Ang, "-") > 0 Or InStr(Ang, "S") > 0 Then deg = -deg
ed:
DD = deg
End Function

Function hmsPlus(Ang1, Ang2)
Ang1 = DD(Ang1)
Ang2 = DD(Ang2)
hmsPlus = HMS(Ang1 + Ang2, 2)
End Function
Function hmsMinus(Ang1, Ang2)
Ang1 = DD(Ang1)
Ang2 = DD(Ang2)
hmsMinus = HMS(Ang1 - Ang2, 2)
End Function
Function Brg(Est1, Nth1, Est2, Nth2)
dEst = Est2 - Est1
dNth = Nth2 - Nth1
If dNth = 0 And dEst = 0 Then
    b = "N/A"
    GoTo ed
End If

b = Application.WorksheetFunction.Degrees(WorksheetFunction.Atan2(dNth, dEst))
While b < 0
    b = b + 360
Wend
While b > 360
    b = b - 360
Wend
ed:
Brg = HMS(b, 2)

End Function

Function DsplNme(ByVal Nme)
'Input - AAA##T#
'Output- AAA ##T #
Nme = Replace(Nme, " ", "")
Nmes = Mid(Nme, 1, 1)
If Not IsNull(Nme) Then
    For i = 2 To Len(Nme)
        spac = ""
        If IsNumeric(Mid(Nme, i, 1)) Then
            If Asc(Mid(Nme, i - 1, 1)) >= 65 And Asc(Mid(Nme, i - 1, 1)) <= 90 Or Asc(Mid(Nme, i - 1, 1)) >= 97 And Asc(Mid(Nme, i - 1, 1)) <= 122 Then
                spac = " "
            End If
        End If
        Nmes = Nmes & spac & Mid(Nme, i, 1)
    Next i
End If
DsplNme = Replace(UCase(Trim(Nmes)), "-", " - ")
End Function

Function xmlDuration2Time(ByVal xmlformat)
'Input - <Duration>P0Y0M30DT3H1M0S</Duration>
'Output- HHmmss
xmlformat = Replace(xmlformat, "Duration", "")
durationYear = Val(Mid(xmlformat, InStr(xmlformat, "P") + 1, InStr(xmlformat, "Y") - InStr(xmlformat, "P")))
durationMth = Val(Mid(xmlformat, InStr(xmlformat, "Y") + 1, InStr(xmlformat, "M") - InStr(xmlformat, "Y")))
durationDay = Val(Mid(xmlformat, InStr(xmlformat, "M") + 1, InStr(xmlformat, "D") - InStr(xmlformat, "M")))
i = InStr(xmlformat, "M") + 1
DurationHr = Val(Mid(xmlformat, InStr(xmlformat, "T") + 1, InStr(xmlformat, "H") - InStr(xmlformat, "T")))
DurationMin = Val(Mid(xmlformat, InStr(xmlformat, "H") + 1, InStr(i, xmlformat, "M") - InStr(xmlformat, "H")))
DurationSec = Val(Mid(xmlformat, InStr(i, xmlformat, "M") + 1, InStr(xmlformat, "S") - InStr(i, xmlformat, "M")))
xmlDuration2Time = TimeSerial(durationMth * 30 * 24 + durationDay * 24 + DurationHr, DurationMin, DurationSec)
End Function

Sub LookupMapSheet(MapsheetNames, MapSheetCodes)
MapsheetNames = Array("ABYSSINIA", "ALLANSON", "APPLECROSS", "ARMADALE", "ASHENDON", "AUSTRALIND", "BALGA", "BANKSIA", "BARTONS", "BAYSWATER", "BECHER", "BERRIJUP", "BICKLEY", "BALD", "BAILUP", "BELLEVUE", "BAILEY", "BALMORAL", "BOORONGING", "BOONGARRA", "BOUVARD", "BLUE ROCK", "BRIDGETOWN NE", "BRIDGETOWN NW", _
        "BRIDGETOWN SE", "BRIDGETOWN SW", "BULLSBROOK", "BURNS", "BELVIDERE", "BYFORD", "BUNBURY NORTH", "BUNBURY SOUTH", "CALISTA", "CANNINGTON", "CARILLA", "CAVERSHAM", "COBBLER", "CARCOOLA", "CHIDLOW", "CLIFFORD", "CHANDALA", "CHEDARING", "CHURCHMAN", "CITY", "COOKE", "CLAREMONT", "CLENTON", "CLOVER", _
        "CUTLER", "COCKMAN", "CHINGANNING", "COOGEE", "CURTIS", "CANNINGVALE", "DALE", "DARKIN", "DAWESVILLE", "DEEFOR", "DANDALUP", "DARDANUP", "DOCONING", "DALYELLUP", "EAGLE", "EAST YANCHEP", "FERGUS", "FERGUSON", "FLORIDA", "FORRESTDALE", "FREMANTLE", "FISHER", "GIBBS", "GLENCAIRN", _
        "GLEN FORREST", "GUILDFORD", "GIDGEGANNUP", "GARDEN ISLAND", "GREENLANDS", "GOLDMINE", "GNANGARA", "GOBBY", "GORRY", "GOSNELLS", "GULL", "HAMERSLEY", "HANCOCK", "HADDRILL", "HELENA EAST", "HERNE HILL", "HODD", "HERRON", "HERRON WEST", "HARRIS", "HASTIES", "HUNTLY", "JANDAKOT", "JARRAHDALE", _
        "JIB", "JULIMAR", "JUMPERKINE", "KALAMUNDA", "KARNUP", "KOOMBANA", "KINGSBURY", "KELMSCOTT", "KENNEDY", "KEWDALE", "KEYSBROOK", "KARNET", "KRONIN", "KEATING", "KWINANA", "LAKE ADAMS", "LAWLEY", "LIGHTBODY", "LAKE CARABOODA", "LEONA", "LESMURDIE", "LAKE JANDABUP", "LESCHENAULT", "LOWLANDS", _
        "MAGENUP", "MANDOGALUP", "MANDURAH NE", "MANDURAH NW", "MARSH", "MANDURAH SE", "MANDURAH SW", "MUNDARING", "MUNDIJONG", "MEALUP", "MELVILLE", "MUNGALUP", "MT HELENA", "MIAMI", "MALKUP", "MOONDYNE", "MORANGUP", "MADDERN", "MARRIOTT", "MARYVILLE", "MULLALOO", "MUCHEA NORTH", "MUCHEA SOUTH", "MAIDA VALE", _
        "MYARA", "NAMBEELUP", "NAVAL BASE", "NEERABUP", "MC NESS", "NETTLETON", "NORTHAM EAST", "NORTHAM WEST", "NOCKINE", "OAKLEY", "OBRIEN", "OMEO", "PARKERVILLE", "PERON", "PAGE", "PICKERING", "PIESSE", "PINJAR EAST", "PINJAR WEST", "PEEL", "PEELHURST", "PHILLIPS", "POISON", "PUNRACK", _
        "QUININE", "QUINNS NORTH", "QUINNS SOUTH", "RANDALL", "RUEN", "ROCKINGHAM", "ROLAND", "ROLEYSTONE", "ROSSMOYNE", "ROTTNEST ISLAND", "RAVENSWOOD", "SAFETY BAY", "SCARBOROUGH", "SOUTH CHITTERING", "SERPENTINE", "SINGLETON", "SMITHS MILL", "SOLUS", "SPRINGVALE", "SAWPIT", "SPEARWOOD", "SCARP", "STONY", "SUNSET", _
        "STONEVILLE", "TOODYAY", "THOMPSON", "THORNLIE", "TURTLE", "TOMAR", "THORPE", "TWO ROCKS NORTH", "TWO ROCKS SOUTH", "UPPER SWAN", "VENN", "VINCENT", "WAHRONGA", "WALYUNGA", "WANNEROO", "WARNBRO", "WARBROOK", "WEST BYFORD", "WUNDOWIE", "WEIR", "WELLARD", "WEMBLEY", "WERRIBEE", "WESTCOTT", _
        "WHITFORDS", "WILTON", "WATERLOO", "WOONGANING", "WOODSOME", "WOOROLOO", "WOTNOT", "WANDARRAH", "WARIIN", "WEST SWAN", "WORSLEY", "WEST TOODYAY", "WESTONS", "WUNGONG", "WEST YANCHEP", "YAKAMINDI", "YANGEDI", "YEAL", "YETAR", "YALGORUP", "YALGORUP WEST", "YOKINE", "YUNDERUP", "AJANA", _
        "ALBANY", "ANKETELL", "ASHTON", "AUGUSTA", "BALLADONIA", "BOORABBIN", "BREMER BAY", "BENCUBBIN", "BEDOUT ISLAND", "BELELE", "BENTLEY", "BALFOUR DOWNS", "BILLILUNA", "BARLEE", "BULLEN", "BURNABBIE", "BROWNE", "BROOME", "BROWSE ISLAND", "BARROW ISLAND", "BUSSELTON", "BYRO", "CAMBRIDGE GULF", "CAMDEN SOUND", _
        "CUNDEELEE", "CROSSLAND", "CHARNLEY", "CORNISH", "COBB", "COCOS ISLAND", "COLLIE", "COLLIER", "CORRIGIN", "CAPE ARID", "COOPER", "CHRISTMAS", "CUE", "CULVER", "DAMPIER", "DERBY", "DIXON RANGE", "DUMMER", "DONGARA", "DRYSDALE", "DUKETON", "DUMBLEYUNG", "EDJUDINA", "EDEL", _
        "EDMUND", "ESPERANCE", "EUCLA", "FORREST", "GORDON DOWNS", "GLENGARRY", "GLENBURGH", "GERALDTON", "GUNANYA", "HELENA", "HERBERT", "HILL RIVER", "HYDEN", "IRWIN INLET", "JACKSON", "JOANNA SPRING", "JUBILEE", "KALGOORLIE", "KELLERBERRIN", "KIRKALOCKA", "KENNEDY RANGE", "KINGSTON", "KURNALPI", "LAGRANGE", _
        "LANSDOWNE", "LAVERTON", "LENNARD RIVER", "LEONORA", "LISSADELL", "LAKE JOHNSTON", "LOONGANA", "LENNIS", "LONDONDERRY", "LUCAS", "MARBLE BAR", "MADLEY", "MALCOLM", "MOUNT BRUCE", "MOUNT BARKER", "MACDONALD", "MCLARTY HILLS", "MANDORA", "MONDRAIN ISLAND", "MADURA", "MEDUSA BANKS", "MINIGWAL", "MINILYA", "MENZIES", _
        "MONTAGUE SOUND", "MOUNT PHILLIPS", "MOORA", "MOUNT RAMSAY", "MURGOO", "MUNRO", "MORRIS", "MASON", "MOUNT ANDERSON", "MOUNT BANNERMAN", "MOUNT ELIZABETH", "MOUNT EGERTON", "NABBERU", "NARETHA", "NEWDEGATE", "NINGALOO", "NINGHAN", "NORTH KEELING", "NEALE", "NOONKANBAH", "NORSEMAN", "NOONAERA", "NULLAGINE", "NEWMAN", _
        "ONSLOW", "PERCIVAL", "PEMBERTON", "PENDER", "PORT HEDLAND", "PINJARRA", "PEAK HILL", "PLUMRIDGE", "PATERSON RANGE", "PERENJORI", "PRINCE REGENT", "PERTH", "PYRAMID", "QUOBBA", "RAVENSTHORPE", "RAWLINSON", "ROBERT", "ROY HILL", "ROBERTSON", "ROEBOURNE", "ROBINSON RANGE", "RASON", "RUDALL", "RUNTON", _
        "RYAN", "SAHARA", "SIR SAMUEL", "SANDSTONE", "SOUTHERN CROSS", "SCOTT", "SEEMORE", "SHARK BAY", "STANLEY", "STANSMORE", "TALBOT", "TABLETOP", "TRAINOR", "THROSSELL", "TUREE CREEK", "URAL", "VERNON", "WEBB", "WAIGEN", "WIDGIEMOOLTHA", "WOORAMEL", "WILUNA", "WANNA", "WINNING POOL", _
        "WARRI", "WILSON", "WESTWOOD", "WYLOO", "YALGOO", "YAMPI", "YARINGA", "YARRALOOLA", "YOUANMI", "YOWALGA", "YARRIE", "YANREY", "ZANTHUS", "ASHMORE REEF", "HIBERNIA REEF", "LONG REEF", "HOUTMAN ABROLHOS")

MapSheetCodes = Array("ABY", "ALN", "APP", "ARM", "ASH", "AUS", "BAL", "BAN", "BAR", "BAYS", "BEC", "BER", "BIC", "BLD", "BLP", "BLV", "BLY", "BML", "BOO", "BOON", "BOU", "BRK", "BTNE", "BTNW", _
        "BTSE", "BTSW", "BUL", "BURN", "BVR", "BYF", "BYN", "BYS", "CAL", "CAN", "CAR", "CAV", "CBR", "CCA", "CDW", "CFD", "CHA", "CHE", "CHUR", "CITY", "CKE", "CLAR", "CLN", "CLO", _
        "CLR", "CMN", "CNG", "COO", "CUR", "CVAL", "DAL", "DAR", "DAW", "DEE", "DLP", "DNP", "DOC", "DYP", "EAG", "EYAN", "FER", "FGN", "FLO", "FOR", "FRE", "FSR", "GBS", "GCN", _
        "GF", "GFD", "GID", "GIS", "GLS", "GM", "GNAN", "GOB", "GOR", "GOS", "GUL", "HAM", "HAN", "HDL", "HLE", "HNH", "HOD", "HRN", "HRNW", "HRS", "HTS", "HTY", "JAN", "JAR", _
        "JIB", "JUL", "JUM", "KAL", "KAR", "KBA", "KBY", "KELM", "KEN", "KEW", "KEY", "KNT", "KRN", "KTG", "KWIN", "LAD", "LAW", "LBY", "LCAR", "LEO", "LES", "LJAN", "LNT", "LOW", _
        "MAG", "MAN", "MANE", "MANW", "MAR", "MASE", "MASW", "MDG", "MDJ", "MEA", "MEL", "MGP", "MHEL", "MIA", "MKP", "MOO", "MOR", "MRN", "MRT", "MRY", "MULL", "MUN", "MUS", "MV", _
        "MYA", "NAM", "NAV", "NEE", "NESS", "NET", "NME", "NMW", "NOC", "OAK", "OBR", "OME", "PAR", "PER", "PGE", "PIC", "PIE", "PINE", "PINW", "PL", "PLH", "PLP", "PSN", "PUN", _
        "QNE", "QUN", "QUS", "RDL", "REN", "ROC", "ROD", "ROL", "ROSS", "ROT", "RWD", "SB", "SCA", "SCH", "SER", "SGTN", "SMM", "SOL", "SPR", "SPT", "SPW", "SRP", "STY", "SUN", _
        "SVLE", "TDY", "THOM", "THOR", "TLE", "TOM", "TPE", "TRN", "TRS", "USW", "VEN", "VIN", "WAH", "WAL", "WAN", "WAR", "WBK", "WBYF", "WDE", "WEIR", "WEL", "WEM", "WER", "WES", _
        "WHIT", "WIL", "WLO", "WNG", "WOO", "WOR", "WOT", "WRH", "WRN", "WSW", "WSY", "WTDY", "WTN", "WUN", "WYAN", "YAK", "YAN", "YEAL", "YET", "YGP", "YGPW", "YOK", "YUN", "AJA", _
        "ALB", "ANK", "ATN", "AUG", "BALA", "BBN", "BBY", "BCN", "BED", "BEL", "BEN", "BFD", "BIL", "BLE", "BLN", "BNB", "BRN", "BRO", "BRS", "BRW", "BUS", "BYR", "CAG", "CAM", _
        "CDE", "CLD", "CLY", "CNH", "COB", "COC", "COL", "COLR", "COR", "CPA", "CPR", "CRI", "CUE", "CUL", "DAM", "DBY", "DIX", "DMR", "DON", "DRY", "DUK", "DUM", "EDA", "EDL", _
        "EMD", "ESP", "EUC", "FST", "GDNS", "GGY", "GLB", "GN", "GUN", "HEL", "HER", "HLR", "HYD", "IRW", "JAK", "JSP", "JUB", "KB", "KE", "KIRK", "KRE", "KTN", "KURN", "LAG", _
        "LAN", "LAV", "LDR", "LEN", "LIS", "LJ", "LNA", "LNS", "LON", "LUC", "MAB", "MAD", "MAL", "MBC", "MBR", "MCD", "MCH", "MDA", "MDN", "MDU", "MED", "MGL", "MIN", "MNZ", _
        "MON", "MPH", "MRA", "MRAM", "MRG", "MRO", "MRS", "MSN", "MTA", "MTB", "MTE", "MTG", "NAB", "NAR", "NEW", "NGL", "NIN", "NK", "NLE", "NOON", "NOR", "NRA", "NUL", "NWM", _
        "ONS", "PCL", "PEM", "PEN", "PH", "PIN", "PKH", "PLUM", "PRA", "PRJ", "PRT", "PTH", "PYR", "QUO", "RAV", "RAW", "RBT", "RHL", "ROB", "ROE", "RRA", "RSN", "RUD", "RUN", _
        "RYN", "SAH", "SAM", "SAN", "SC", "SCT", "SEE", "SKB", "SLY", "SMR", "TAL", "TLT", "TNR", "TSL", "TUR", "URL", "VER", "WEB", "WGN", "WID", "WML", "WNA", "WNN", "WPL", _
        "WRR", "WSN", "WWD", "WYL", "YAL", "YAM", "YAR", "YLA", "YOU", "YOW", "YRE", "YRY", "ZAN", "ARF", "HRF", "LOR", "ABR")
End Sub
Public Function AbbName(MkName)
MkName = UCase(Trim(MkName))
AbbName = MkName
'''' Search to find the first numeric in name to get the mapsheet name '''
If Not IsNull(MkName) Then
    For i = 2 To Len(MkName)
        If IsNumeric(Mid(MkName, i, 1)) Then
            MapsheetName = Trim(Left(MkName, i - 1))
            i = Len(MkName)
        End If
    Next i
End If

'''' Search the mapsheet lookup array to find the abbreviation code ''''
Call LookupMapSheet(MapsheetNames, MapSheetCodes)
For i = 0 To UBound(MapsheetNames) - 1
    If MapsheetNames(i) = MapsheetName Then
        AbbName = Replace(MkName, MapsheetName, MapSheetCodes(i))
        i = UBound(MapsheetNames) - 1
    End If
Next i

End Function
Function FullGESName(ByVal MkName)

MkName = UCase(MkName)
FullGESName = MkName
'''' Search to find the first numeric in name to get the mapsheet code '''
If Not IsNull(MkName) Then
    For i = 2 To Len(MkName)
        If IsNumeric(Mid(MkName, i, 1)) Then
            MapSheetCode = Trim(Left(MkName, i - 1))
            i = Len(MkName)
        End If
    Next i
End If

'''' Search the mapsheet lookup array to find the abbreviation code ''''
Call LookupMapSheet(MapsheetNames, MapSheetCodes)
For i = 0 To UBound(MapsheetNames) - 1
    If MapSheetCodes(i) = MapSheetCode Then
        FullGESName = Replace(MkName, MapSheetCode, MapsheetNames(i))
        i = UBound(MapsheetNames) - 1
    End If
Next i

End Function

Function EllipsoidalParameters(Datum, a, b, OneOnF)
If InStr(1, Datum, "84") > 0 Then
    '''ANS parameters'''
    a = 6378160
    OneOnF = 298.25
End If
'If InStr(1, Datum, "94") > 0 Or UCase(Datum) = "GRS80" Then
    '''GRS80 parameters'''
    a = 6378137
    OneOnF = 298.257222101
'End If
'If InStr(1, Datum, "2020") > 0 Then
'    '''GRS80 parameters'''
'    a = 6378137
'    OneOnF = 298.257222101
'End If
b = (a * (1 - (1 / OneOnF)))
End Function
Function ProjectionParameters(proj, FalseEasting, FalseNorthing, CSF, ZoneWidth, CM)
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)

If UCase(proj) = "MGA94" Or UCase(proj) = "MGA2020" Or UCase(proj) = "AMG84" Then
    FalseEasting = 500000
    FalseNorthing = 10000000
    CSF = 0.9996
    ZoneWidth = 6
    CM = -177
Else
    ZoneWidth = 0.641666666666
    If UCase(proj) = "ALB84" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.000012
        CM = 117.916666666667
    End If
    If UCase(proj) = "ALB94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.0000044
        CM = 117.883333333334
    End If
    If UCase(proj) = "ALB2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.0000044
        CM = 117.883333333334
    End If
    If UCase(proj) = "BIO94" Then
        FalseEasting = 60000
        FalseNorthing = 2600000
        CSF = 1.0000022
        CM = 115.25
    End If
    If UCase(proj) = "BIO2020" Then
        FalseEasting = 60000
        FalseNorthing = 2700000
        CSF = 1.0000022
        CM = 115.25
    End If
    If UCase(proj) = "BRO84" Then
        FalseEasting = 50000
        FalseNorthing = 2200000
        CSF = 1.000003
        CM = 122.333333333334
    End If
    If UCase(proj) = "BRO94" Then
        FalseEasting = 50000
        FalseNorthing = 2200000
        CSF = 1.00000298
        CM = 122.333333333334
    End If
    If UCase(proj) = "BRO2020" Then
        FalseEasting = 50000
        FalseNorthing = 2300000
        CSF = 1.00000298
        CM = 122.333333333334
    End If
    If UCase(proj) = "BCG84" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.000007
        CM = 115.433333333334
    End If
    If UCase(proj) = "BCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 0.99999592
        CM = 115.433333333334
    End If
    If UCase(proj) = "BCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 0.99999592
        CM = 115.433333333334
    End If
    If UCase(proj) = "CARN84" Then
        FalseEasting = 50000
        FalseNorthing = 3050000
        CSF = 1.000005
        CM = 113.666666666667
    End If
    If UCase(proj) = "CARN94" Then
        FalseEasting = 50000
        FalseNorthing = 2950000
        CSF = 0.99999796
        CM = 113.666666666667
    End If
    If UCase(proj) = "CARN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3050000
        CSF = 0.99999796
        CM = 113.666666666667
    End If
    If UCase(proj) = "CIG84" Then
        FalseEasting = 50000
        FalseNorthing = 1300000
        CSF = 1.000024
        CM = 105.625
    End If
    If UCase(proj) = "CIG94" Then
        FalseEasting = 50000
        FalseNorthing = 1300000
        CSF = 1.00002514
        CM = 105.625
    End If
    If UCase(proj) = "CIG2020" Then
        FalseEasting = 50000
        FalseNorthing = 1400000
        CSF = 1.00002514
        CM = 105.625
    End If
    If UCase(proj) = "CKIG84" Then
        FalseEasting = 50000
        FalseNorthing = 1400000
        CSF = 1
        CM = 96.875
    End If
    If UCase(proj) = "CKIG94" Then
        FalseEasting = 50000
        FalseNorthing = 1500000
        CSF = 0.99999387
        CM = 96.875
    End If
    If UCase(proj) = "CKIG2020" Then
        FalseEasting = 50000
        FalseNorthing = 1600000
        CSF = 0.99999387
        CM = 96.875
    End If
    If UCase(proj) = "COL94" Then
        FalseEasting = 40000
        FalseNorthing = 4000000
        CSF = 1.000019
        CM = 115.933333333334
    End If
    If UCase(proj) = "COL2020" Then
        FalseEasting = 40000
        FalseNorthing = 4100000
        CSF = 1.000019
        CM = 115.933333333334
    End If
    If UCase(proj) = "ESP84" Then
        FalseEasting = 50000
        FalseNorthing = 3950000
        CSF = 1.000012
        CM = 121.883333333334
    End If
    If UCase(proj) = "ESP94" Then
        FalseEasting = 50000
        FalseNorthing = 3950000
        CSF = 1.0000055
        CM = 121.883333333334
    End If
    If UCase(proj) = "ESP2020" Then
        FalseEasting = 50000
        FalseNorthing = 4050000
        CSF = 1.0000055
        CM = 121.883333333334
    End If
    If UCase(proj) = "EXM84" Then
        FalseEasting = 60000
        FalseNorthing = 2750000
        CSF = 1.000009
        CM = 114.066666666667
    End If
    If UCase(proj) = "EXM94" Then
        FalseEasting = 50000
        FalseNorthing = 2650000
        CSF = 1.00000236
        CM = 114.066666666667
    End If
    If UCase(proj) = "EXM2020" Then
        FalseEasting = 50000
        FalseNorthing = 2750000
        CSF = 1.00000236
        CM = 114.066666666667
    End If
    If UCase(proj) = "GCG84" Then
        FalseEasting = 50000
        FalseNorthing = 3350000
        CSF = 1.000016
        CM = 114.666666666667
    End If
    If UCase(proj) = "GCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3350000
        CSF = 1.00000628
        CM = 114.583333333334
    End If
    If UCase(proj) = "GCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3450000
        CSF = 1.00000628
        CM = 114.583333333334
    End If
    If UCase(proj) = "GG84" Then
        FalseEasting = 60000
        FalseNorthing = 4000000
        CSF = 1.000057
        CM = 121.45
    End If
    If UCase(proj) = "GOLD84" Then
        FalseEasting = 60000
        FalseNorthing = 4000000
        CSF = 1.000057
        CM = 121.45
    End If
    If UCase(proj) = "GOLD94" Then
        FalseEasting = 60000
        FalseNorthing = 3700000
        CSF = 1.00004949
        CM = 121.5
    End If
    If UCase(proj) = "GOLD2020" Then
        FalseEasting = 60000
        FalseNorthing = 3800000
        CSF = 1.00004949
        CM = 121.5
    End If
    If UCase(proj) = "JCG84" Then
        FalseEasting = 50000
        FalseNorthing = 3550000
        CSF = 1.00001
        CM = 114.983333333334
    End If
    If UCase(proj) = "JCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3550000
        CSF = 1.00000314
        CM = 114.983333333334
    End If
    If UCase(proj) = "JCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3650000
        CSF = 1.00000314
        CM = 114.983333333334
    End If
    If UCase(proj) = "KALB94" Then
        FalseEasting = 55000
        FalseNorthing = 3600000
        CSF = 1.000014
        CM = 114.3152777778
    End If
    If UCase(proj) = "KALB2020" Then
        FalseEasting = 55000
        FalseNorthing = 3700000
        CSF = 1.000014
        CM = 114.3152777778
    End If
    If UCase(proj) = "KAR84" Then
        FalseEasting = 50000
        FalseNorthing = 2450000
        CSF = 1.000004
        CM = 116.933333333334
    End If
    If UCase(proj) = "KAR94" Then
        FalseEasting = 50000
        FalseNorthing = 2450000
        CSF = 0.9999989
        CM = 116.933333333334
    End If
    If UCase(proj) = "KAR2020" Then
        FalseEasting = 50000
        FalseNorthing = 2550000
        CSF = 0.9999989
        CM = 116.933333333334
    End If
    If UCase(proj) = "KUN84" Then
        FalseEasting = 50000
        FalseNorthing = 2000000
        CSF = 1.000014
        CM = 128.75
    End If
    If UCase(proj) = "KUN94" Then
        FalseEasting = 50000
        FalseNorthing = 2000000
        CSF = 1.0000165
        CM = 128.75
    End If
    If UCase(proj) = "KUN2020" Then
        FalseEasting = 50000
        FalseNorthing = 2100000
        CSF = 1.0000165
        CM = 128.75
    End If
    If UCase(proj) = "LCG84" Then
        FalseEasting = 50000
        FalseNorthing = 3650000
        CSF = 1.000008
        CM = 115.366666666667
    End If
    If UCase(proj) = "LCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3650000
        CSF = 1.00000157
        CM = 115.366666666667
    End If
    If UCase(proj) = "LCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3750000
        CSF = 1.00000157
        CM = 115.366666666667
    End If
    If UCase(proj) = "MRCG84" Then
        FalseEasting = 50000
        FalseNorthing = 4050000
        CSF = 1.000014
        CM = 115.1
    End If
    If UCase(proj) = "MRCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3950000
        CSF = 1.0000055
        CM = 115.166666666667
    End If
    If UCase(proj) = "MRCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 4050000
        CSF = 1.0000055
        CM = 115.166666666667
    End If
    If UCase(proj) = "PCG84" Then
        FalseEasting = 40000
        FalseNorthing = 3800000
        CSF = 1.000006
        CM = 115.833333333334
    End If
    If UCase(proj) = "PCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 0.99999906
        CM = 115.816666666667
    End If
    If UCase(proj) = "PCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 0.99999906
        CM = 115.816666666667
    End If
    If UCase(proj) = "PHG84" Then
        FalseEasting = 50000
        FalseNorthing = 2400000
        CSF = 1.000004
        CM = 118.583333333334
    End If
    If UCase(proj) = "PHG94" Then
        FalseEasting = 50000
        FalseNorthing = 2400000
        CSF = 1.00000135
        CM = 118.6
    End If
    If UCase(proj) = "PHG2020" Then
        FalseEasting = 50000
        FalseNorthing = 2500000
        CSF = 1.00000135
        CM = 118.6
    End If
    If UCase(proj) = "AGLIME2020" Then
        FalseEasting = 50000
        FalseNorthing = 3850000
        CSF = 1.00003174
        CM = 116.416666666667
    End If
    If UCase(proj) = "AGLIME94" Then
        FalseEasting = 50000
        FalseNorthing = 3750000
        CSF = 1.00003174
        CM = 116.416666666667
    End If
    If UCase(proj) = "AJANA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3384000
        CSF = 1.00002763
        CM = 114.675
    End If
    If UCase(proj) = "AJANA94" Then
        FalseEasting = 50000
        FalseNorthing = 3284000
        CSF = 1.00002763
        CM = 114.675
    End If
    If UCase(proj) = "AMBANIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3550000
        CSF = 1.0000308
        CM = 115.25
    End If
    If UCase(proj) = "AMBANIA94" Then
        FalseEasting = 50000
        FalseNorthing = 3450000
        CSF = 1.0000308
        CM = 115.25
    End If
    If UCase(proj) = "AMELUP2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.0000396
        CM = 118.316666666667
    End If
    If UCase(proj) = "AMELUP94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.0000396
        CM = 118.316666666667
    End If
    If UCase(proj) = "ATLEY2020" Then
        FalseEasting = 50000
        FalseNorthing = 3437000
        CSF = 1.00008153
        CM = 119.051111111111
    End If
    If UCase(proj) = "ATLEY94" Then
        FalseEasting = 50000
        FalseNorthing = 3337000
        CSF = 1.00008153
        CM = 119.051111111111
    End If
    If UCase(proj) = "AVONRIVER2020" Then
        FalseEasting = 50000
        FalseNorthing = 3862000
        CSF = 1.00002799
        CM = 116.811111111111
    End If
    If UCase(proj) = "AVONRIVER94" Then
        FalseEasting = 50000
        FalseNorthing = 3762000
        CSF = 1.00002799
        CM = 116.811111111111
    End If
    If UCase(proj) = "BAKER2020" Then
        FalseEasting = 50000
        FalseNorthing = 3244000
        CSF = 1.00006076
        CM = 126.128611111111
    End If
    If UCase(proj) = "BAKER94" Then
        FalseEasting = 50000
        FalseNorthing = 3144000
        CSF = 1.00006076
        CM = 126.128611111111
    End If
    If UCase(proj) = "BALLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00001901
        CM = 123.883333333333
    End If
    If UCase(proj) = "BALLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00001901
        CM = 123.883333333333
    End If
    If UCase(proj) = "BCG2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 0.99999592
        CM = 115.433333333333
    End If
    If UCase(proj) = "BCG94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 0.99999592
        CM = 115.433333333333
    End If
    If UCase(proj) = "BENNETT2020" Then
        FalseEasting = 50000
        FalseNorthing = 2900000
        CSF = 1.00010144
        CM = 117.903333333333
    End If
    If UCase(proj) = "BENNETT94" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00010144
        CM = 117.903333333333
    End If
    If UCase(proj) = "BIDYADANGA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2387000
        CSF = 1.00000696
        CM = 121.93
    End If
    If UCase(proj) = "BIDYADANGA94" Then
        FalseEasting = 50000
        FalseNorthing = 2287000
        CSF = 1.00000696
        CM = 121.93
    End If
    If UCase(proj) = "BOHEMIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2406000
        CSF = 1.00003859
        CM = 126.121388888889
    End If
    If UCase(proj) = "BOHEMIA94" Then
        FalseEasting = 50000
        FalseNorthing = 2306000
        CSF = 1.00003859
        CM = 126.121388888889
    End If
    If UCase(proj) = "BOOLARDY2020" Then
        FalseEasting = 100000
        FalseNorthing = 3253000
        CSF = 1.00005465
        CM = 116.583333333333
    End If
    If UCase(proj) = "BOOLARDY94" Then
        FalseEasting = 100000
        FalseNorthing = 3153000
        CSF = 1.00005465
        CM = 116.583333333333
    End If
    If UCase(proj) = "BOONDI2020" Then
        FalseEasting = 50000
        FalseNorthing = 3790000
        CSF = 1.00006387
        CM = 120.283333333333
    End If
    If UCase(proj) = "BOONDI94" Then
        FalseEasting = 50000
        FalseNorthing = 3690000
        CSF = 1.00006387
        CM = 120.283333333333
    End If
    If UCase(proj) = "BOWGADA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3592000
        CSF = 1.00004215
        CM = 116.147377777778
    End If
    If UCase(proj) = "BOWGADA94" Then
        FalseEasting = 50000
        FalseNorthing = 3492000
        CSF = 1.00004215
        CM = 116.147377777778
    End If
    If UCase(proj) = "BOXWOOD2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.000011
        CM = 118.75
    End If
    If UCase(proj) = "BOXWOOD94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.000011
        CM = 118.75
    End If
    If UCase(proj) = "BRDGETWN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00003425
        CM = 116.15
    End If
    If UCase(proj) = "BRDGETWN94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00003425
        CM = 116.15
    End If
    If UCase(proj) = "BROOKTON2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004132
        CM = 117
    End If
    If UCase(proj) = "BROOKTON94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00004132
        CM = 117
    End If
    If UCase(proj) = "BROOME2020" Then
        FalseEasting = 50000
        FalseNorthing = 2300000
        CSF = 1.00000298
        CM = 122.333333333333
    End If
    If UCase(proj) = "BROOME94" Then
        FalseEasting = 50000
        FalseNorthing = 2200000
        CSF = 1.00000298
        CM = 122.333333333333
    End If
    If UCase(proj) = "BULDANIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3880000
        CSF = 1.00004305
        CM = 122.116666666667
    End If
    If UCase(proj) = "BULDANIA94" Then
        FalseEasting = 50000
        FalseNorthing = 3780000
        CSF = 1.00004305
        CM = 122.116666666667
    End If
    If UCase(proj) = "BULLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3770000
        CSF = 1.00006235
        CM = 120.866666666667
    End If
    If UCase(proj) = "BULLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3670000
        CSF = 1.00006235
        CM = 120.866666666667
    End If
    If UCase(proj) = "CARNOT2020" Then
        FalseEasting = 50000
        FalseNorthing = 2245000
        CSF = 1.00002641
        CM = 122.475
    End If
    If UCase(proj) = "CARNOT94" Then
        FalseEasting = 50000
        FalseNorthing = 2145000
        CSF = 1.00002641
        CM = 122.475
    End If
    If UCase(proj) = "CHALLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3454000
        CSF = 1.00007041
        CM = 118.387777777778
    End If
    If UCase(proj) = "CHALLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3354000
        CSF = 1.00007041
        CM = 118.387777777778
    End If
    If UCase(proj) = "CHEELA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2900204
        CSF = 1.00005059
        CM = 117
    End If
    If UCase(proj) = "CHEELA94" Then
        FalseEasting = 50000
        FalseNorthing = 2800204
        CSF = 1.00005059
        CM = 117
    End If
    If UCase(proj) = "CHICHESTER2020" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00006663
        CM = 118.858333333333
    End If
    If UCase(proj) = "CHICHESTER94" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00006663
        CM = 118.858333333333
    End If
    If UCase(proj) = "COOLAWANYA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2720000
        CSF = 1.00004604
        CM = 117.333333333333
    End If
    If UCase(proj) = "COOLAWANYA94" Then
        FalseEasting = 50000
        FalseNorthing = 2620000
        CSF = 1.00004604
        CM = 117.333333333333
    End If
    If UCase(proj) = "COOROW2020" Then
        FalseEasting = 50000
        FalseNorthing = 3700000
        CSF = 1.00003504
        CM = 115.916666666667
    End If
    If UCase(proj) = "COOROW94" Then
        FalseEasting = 50000
        FalseNorthing = 3600000
        CSF = 1.00003504
        CM = 115.916666666667
    End If
    If UCase(proj) = "CORRIG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3950000
        CSF = 1.0000407
        CM = 117.583333333333
    End If
    If UCase(proj) = "CORRIG94" Then
        FalseEasting = 50000
        FalseNorthing = 3850000
        CSF = 1.0000407
        CM = 117.583333333333
    End If
    If UCase(proj) = "COSMO2020" Then
        FalseEasting = 50000
        FalseNorthing = 3440000
        CSF = 1.00007692006
        CM = 122.683333333333
    End If
    If UCase(proj) = "COSMO94" Then
        FalseEasting = 50000
        FalseNorthing = 3340000
        CSF = 1.00007692
        CM = 122.683333333333
    End If
    If UCase(proj) = "CRANBK2020" Then
        FalseEasting = 50000
        FalseNorthing = 4200000
        CSF = 1.00003582
        CM = 117.516666666667
    End If
    If UCase(proj) = "CRANBK94" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00003582
        CM = 117.516666666667
    End If
    If UCase(proj) = "CUEMAG2020" Then
        FalseEasting = 50000
        FalseNorthing = 3380000
        CSF = 1.00006203
        CM = 117.914602777778
    End If
    If UCase(proj) = "CUEMAG94" Then
        FalseEasting = 50000
        FalseNorthing = 3280000
        CSF = 1.00006203
        CM = 117.914602777778
    End If
    If UCase(proj) = "CULLINIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3875000
        CSF = 1.00004448
        CM = 123.128888888889
    End If
    If UCase(proj) = "CULLINIA94" Then
        FalseEasting = 50000
        FalseNorthing = 3775000
        CSF = 1.00004448
        CM = 123.128888888889
    End If
    If UCase(proj) = "DANDARAGAN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3725517
        CSF = 1.00003205
        CM = 115.625
    End If
    If UCase(proj) = "DANDARAGAN94" Then
        FalseEasting = 50000
        FalseNorthing = 3625517
        CSF = 1.00003205
        CM = 115.625
    End If
    If UCase(proj) = "DEPOT2020" Then
        FalseEasting = 50000
        FalseNorthing = 3436500
        CSF = 1.00006615
        CM = 120.061666666667
    End If
    If UCase(proj) = "DEPOT94" Then
        FalseEasting = 50000
        FalseNorthing = 3336500
        CSF = 1.00006615
        CM = 120.061666666667
    End If
    If UCase(proj) = "DERBY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2300000
        CSF = 1.00000393
        CM = 123.683333333333
    End If
    If UCase(proj) = "DERBY94" Then
        FalseEasting = 50000
        FalseNorthing = 2200000
        CSF = 1.00000393
        CM = 123.683333333333
    End If
    If UCase(proj) = "DOCKERRIVER2020" Then
        FalseEasting = 50000
        FalseNorthing = 3105500
        CSF = 1.00009986
        CM = 129.113333333333
    End If
    If UCase(proj) = "DOCKERRIVER94" Then
        FalseEasting = 50000
        FalseNorthing = 3005500
        CSF = 1.00009986
        CM = 129.113333333333
    End If
    If UCase(proj) = "DOOLGUNNA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3600000
        CSF = 1.00008411
        CM = 119.5
    End If
    If UCase(proj) = "DOOLGUNNA94" Then
        FalseEasting = 50000
        FalseNorthing = 3500000
        CSF = 1.00008411
        CM = 119.5
    End If
    If UCase(proj) = "DOWERIN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00004109
        CM = 116.85
    End If
    If UCase(proj) = "DOWERIN94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00004109
        CM = 116.85
    End If
    If UCase(proj) = "DURACK2020" Then
        FalseEasting = 50000
        FalseNorthing = 2100000
        CSF = 1.00002487
        CM = 127.8
    End If
    If UCase(proj) = "DURACK94" Then
        FalseEasting = 50000
        FalseNorthing = 2000000
        CSF = 1.00002487
        CM = 127.8
    End If
    If UCase(proj) = "EASTALB2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00001271
        CM = 118.333333333333
    End If
    If UCase(proj) = "EASTALB94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00001271
        CM = 118.333333333333
    End If
    If UCase(proj) = "EDGAR2020" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00004144931
        CM = 120.227777777778
    End If
    If UCase(proj) = "EDGAR94" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00004145
        CM = 120.227777777778
    End If
    If UCase(proj) = "ELLENBRAE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2100000
        CSF = 1.00006254
        CM = 126.891944444444
    End If
    If UCase(proj) = "ELLENBRAE94" Then
        FalseEasting = 50000
        FalseNorthing = 2000000
        CSF = 1.00006254
        CM = 126.891944444444
    End If
    If UCase(proj) = "EUCLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3850000
        CSF = 1.00000583
        CM = 128.75
    End If
    If UCase(proj) = "EUCLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3750000
        CSF = 1.00000583
        CM = 128.75
    End If
    If UCase(proj) = "EXMXT2020" Then
        FalseEasting = 50000
        FalseNorthing = 3100000
        CSF = 1.00000569
        CM = 114.016666666667
    End If
    If UCase(proj) = "EXMXT94" Then
        FalseEasting = 50000
        FalseNorthing = 3000000
        CSF = 1.00000569
        CM = 114.016666666667
    End If
    If UCase(proj) = "FITZGERALD2020" Then
        FalseEasting = 50000
        FalseNorthing = 4086000
        CSF = 1.00003561
        CM = 119.516666666667
    End If
    If UCase(proj) = "FITZGERALD94" Then
        FalseEasting = 50000
        FalseNorthing = 3986000
        CSF = 1.00003561
        CM = 119.516666666667
    End If
    If UCase(proj) = "FITZROY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2400000
        CSF = 1.00002247
        CM = 125.566666666667
    End If
    If UCase(proj) = "FITZROY94" Then
        FalseEasting = 50000
        FalseNorthing = 2300000
        CSF = 1.00002247
        CM = 125.566666666667
    End If
    If UCase(proj) = "FORTES2020" Then
        FalseEasting = 50000
        FalseNorthing = 2727968
        CSF = 1.00001728
        CM = 115.933333333333
    End If
    If UCase(proj) = "FORTES94" Then
        FalseEasting = 50000
        FalseNorthing = 2627968
        CSF = 1.00001728
        CM = 115.933333333333
    End If
    If UCase(proj) = "FRANKLAND2020" Then
        FalseEasting = 50000
        FalseNorthing = 4167000
        CSF = 1.00001603
        CM = 116.453611111111
    End If
    If UCase(proj) = "FRANKLAND94" Then
        FalseEasting = 50000
        FalseNorthing = 4067000
        CSF = 1.00001603
        CM = 116.453611111111
    End If
    If UCase(proj) = "FRAZER 2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004321
        CM = 122.466666666667
    End If
    If UCase(proj) = "FRAZER 94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00004321
        CM = 122.466666666667
    End If
    If UCase(proj) = "GABYON2020" Then
        FalseEasting = 50000
        FalseNorthing = 3470000
        CSF = 1.00004737
        CM = 116.351088888889
    End If
    If UCase(proj) = "GABYON94" Then
        FalseEasting = 50000
        FalseNorthing = 3370000
        CSF = 1.00004737
        CM = 116.351088888889
    End If
    If UCase(proj) = "GALVANS2020" Then
        FalseEasting = 50000
        FalseNorthing = 2160000
        CSF = 1.00006992
        CM = 125.833333333333
    End If
    If UCase(proj) = "GALVANS94" Then
        FalseEasting = 50000
        FalseNorthing = 2060000
        CSF = 1.00006992
        CM = 125.833333333333
    End If
    If UCase(proj) = "GASCOYNE2020" Then
        FalseEasting = 50000
        FalseNorthing = 3122000
        CSF = 1.00002027
        CM = 115.291666666667
    End If
    If UCase(proj) = "GASCOYNE94" Then
        FalseEasting = 50000
        FalseNorthing = 3022000
        CSF = 1.00002027
        CM = 115.291666666667
    End If
    If UCase(proj) = "GIBB2 2020" Then
        FalseEasting = 50000
        FalseNorthing = 2100000
        CSF = 1.00001571
        CM = 124.55
    End If
    If UCase(proj) = "GIBB2 94" Then
        FalseEasting = 50000
        FalseNorthing = 2000000
        CSF = 1.00001571
        CM = 124.55
    End If
    If UCase(proj) = "GIBBRIVER2020" Then
        FalseEasting = 100000
        FalseNorthing = 2085000
        CSF = 1.00007814
        CM = 126.355833333333
    End If
    If UCase(proj) = "GIBBRIVER94" Then
        FalseEasting = 100000
        FalseNorthing = 1985000
        CSF = 1.00007814
        CM = 126.355833333333
    End If
    If UCase(proj) = "GILES2020" Then
        FalseEasting = 50000
        FalseNorthing = 3132000
        CSF = 1.00008493
        CM = 127.888055555556
    End If
    If UCase(proj) = "GILES94" Then
        FalseEasting = 50000
        FalseNorthing = 3032000
        CSF = 1.00008493
        CM = 127.888055555556
    End If
    If UCase(proj) = "GIRALIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00000424
        CM = 114.716666666667
    End If
    If UCase(proj) = "GIRALIA94" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00000424
        CM = 114.716666666667
    End If
    If UCase(proj) = "H1BANWIL2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00004447
        CM = 116.666666666667
    End If
    If UCase(proj) = "H1BANWIL94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004447
        CM = 116.666666666667
    End If
    If UCase(proj) = "H6BOW2020" Then
        FalseEasting = 50000
        FalseNorthing = 2250000
        CSF = 1.00004761
        CM = 128.166666666667
    End If
    If UCase(proj) = "H6BOW94" Then
        FalseEasting = 50000
        FalseNorthing = 2150000
        CSF = 1.00004761
        CM = 128.166666666667
    End If
    If UCase(proj) = "H6ORD2020" Then
        FalseEasting = 50000
        FalseNorthing = 2350000
        CSF = 1.00005641
        CM = 127.783333333333
    End If
    If UCase(proj) = "H6ORD94" Then
        FalseEasting = 50000
        FalseNorthing = 2250000
        CSF = 1.00005641
        CM = 127.783333333333
    End If
    If UCase(proj) = "H9BRIVER2020" Then
        FalseEasting = 50000
        FalseNorthing = 4230000
        CSF = 1.0000099
        CM = 116.983333333333
    End If
    If UCase(proj) = "H9BRIVER94" Then
        FalseEasting = 50000
        FalseNorthing = 4130000
        CSF = 1.0000099
        CM = 116.983333333333
    End If
    If UCase(proj) = "HAHN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.0000462
        CM = 120.333333333333
    End If
    If UCase(proj) = "HAHN94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.0000462
        CM = 120.333333333333
    End If
    If UCase(proj) = "HAMERSLEY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2780000
        CSF = 1.00007118
        CM = 117.7
    End If
    If UCase(proj) = "HAMERSLEY94" Then
        FalseEasting = 50000
        FalseNorthing = 2680000
        CSF = 1.00007118
        CM = 117.7
    End If
    If UCase(proj) = "HAMPTON2020" Then
        FalseEasting = 100000
        FalseNorthing = 3838000
        CSF = 1.00001246
        CM = 128.163888888889
    End If
    If UCase(proj) = "HAMPTON94" Then
        FalseEasting = 100000
        FalseNorthing = 3738000
        CSF = 1.00001246
        CM = 128.163888888889
    End If
    If UCase(proj) = "HARMS2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00002781
        CM = 123.233333333333
    End If
    If UCase(proj) = "HARMS94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00002781
        CM = 123.233333333333
    End If
    If UCase(proj) = "HOMEVALLEY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2087000
        CSF = 1.00005221
        CM = 127.4025
    End If
    If UCase(proj) = "HOMEVALLEY94" Then
        FalseEasting = 50000
        FalseNorthing = 1987000
        CSF = 1.00005221
        CM = 127.4025
    End If
    If UCase(proj) = "HOPETOUN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00000866
        CM = 120.116666666667
    End If
    If UCase(proj) = "HOPETOUN94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00000866
        CM = 120.116666666667
    End If
    If UCase(proj) = "HUG2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00003472
        CM = 116.233333333333
    End If
    If UCase(proj) = "HUG94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00003472
        CM = 116.233333333333
    End If
    If UCase(proj) = "HYDEN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00004534
        CM = 119
    End If
    If UCase(proj) = "HYDEN94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004534
        CM = 119
    End If
    If UCase(proj) = "ISDELL2020" Then
        FalseEasting = 50000
        FalseNorthing = 2210000
        CSF = 1.00006406
        CM = 125.424722222222
    End If
    If UCase(proj) = "ISDELL94" Then
        FalseEasting = 50000
        FalseNorthing = 2110000
        CSF = 1.00006406
        CM = 125.424722222222
    End If
    If UCase(proj) = "JERRATS2020" Then
        FalseEasting = 50000
        FalseNorthing = 4130000
        CSF = 1.00003472
        CM = 118.9
    End If
    If UCase(proj) = "JERRATS94" Then
        FalseEasting = 50000
        FalseNorthing = 4030000
        CSF = 1.00003472
        CM = 118.9
    End If
    If UCase(proj) = "JILLBUNDA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.000013
        CM = 125.716666666667
    End If
    If UCase(proj) = "JILLBUNDA94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.000013
        CM = 125.716666666667
    End If
    If UCase(proj) = "JOPE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2850000
        CSF = 1.00007102
        CM = 117.483333333333
    End If
    If UCase(proj) = "JOPE94" Then
        FalseEasting = 50000
        FalseNorthing = 2750000
        CSF = 1.00007102
        CM = 117.483333333333
    End If
    If UCase(proj) = "KALBARRI2020" Then
        FalseEasting = 50000
        FalseNorthing = 3400000
        CSF = 1.00001995
        CM = 114.316666666667
    End If
    If UCase(proj) = "KALBARRI94" Then
        FalseEasting = 50000
        FalseNorthing = 3300000
        CSF = 1.00001995
        CM = 114.316666666667
    End If
    If UCase(proj) = "KALUWIRI2020" Then
        FalseEasting = 50000
        FalseNorthing = 3436000
        CSF = 1.00008422
        CM = 119.590555555556
    End If
    If UCase(proj) = "KALUWIRI94" Then
        FalseEasting = 50000
        FalseNorthing = 3336000
        CSF = 1.00008422
        CM = 119.590555555556
    End If
    If UCase(proj) = "KARIJINI2020" Then
        FalseEasting = 50000
        FalseNorthing = 2785000
        CSF = 1.00011302
        CM = 118.35
    End If
    If UCase(proj) = "KARIJINI94" Then
        FalseEasting = 50000
        FalseNorthing = 2685000
        CSF = 1.00011302
        CM = 118.35
    End If
    If UCase(proj) = "KELLER2020" Then
        FalseEasting = 50000
        FalseNorthing = 3810000
        CSF = 1.00004148
        CM = 117.65
    End If
    If UCase(proj) = "KELLER94" Then
        FalseEasting = 50000
        FalseNorthing = 3710000
        CSF = 1.00004148
        CM = 117.65
    End If
    If UCase(proj) = "KILLARA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3210000
        CSF = 1.00009148
        CM = 119.003275
    End If
    If UCase(proj) = "KILLARA94" Then
        FalseEasting = 50000
        FalseNorthing = 3110000
        CSF = 1.00009148
        CM = 119.003275
    End If
    If UCase(proj) = "KOJON2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00003818
        CM = 117.05
    End If
    If UCase(proj) = "KOJON94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00003818
        CM = 117.05
    End If
    If UCase(proj) = "KONDININ2020" Then
        FalseEasting = 50000
        FalseNorthing = 3940000
        CSF = 1.00004305
        CM = 118.2
    End If
    If UCase(proj) = "KONDININ94" Then
        FalseEasting = 50000
        FalseNorthing = 3840000
        CSF = 1.00004305
        CM = 118.2
    End If
    If UCase(proj) = "KOORDA2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00005009
        CM = 117.4
    End If
    If UCase(proj) = "KOORDA94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00005009
        CM = 117.4
    End If
    If UCase(proj) = "KUMARINA2020" Then
        FalseEasting = 100000
        FalseNorthing = 3067000
        CSF = 1.00008818
        CM = 119.480555555556
    End If
    If UCase(proj) = "KUMARINA94" Then
        FalseEasting = 100000
        FalseNorthing = 2967000
        CSF = 1.00008818
        CM = 119.480555555556
    End If
    If UCase(proj) = "LAKEGRACE 2020" Then
        FalseEasting = 50000
        FalseNorthing = 4020000
        CSF = 1.00004274
        CM = 118.283333333333
    End If
    If UCase(proj) = "LAKEGRACE 94" Then
        FalseEasting = 50000
        FalseNorthing = 3920000
        CSF = 1.00004274
        CM = 118.283333333333
    End If
    If UCase(proj) = "LAKEKING2020" Then
        FalseEasting = 50000
        FalseNorthing = 3980000
        CSF = 1.00004601
        CM = 119.689944444444
    End If
    If UCase(proj) = "LAKEKING94" Then
        FalseEasting = 50000
        FalseNorthing = 3880000
        CSF = 1.00004601
        CM = 119.689944444444
    End If
    If UCase(proj) = "LAMBOO2020" Then
        FalseEasting = 50000
        FalseNorthing = 2385000
        CSF = 1.00005831
        CM = 127.253333333333
    End If
    If UCase(proj) = "LAMBOO94" Then
        FalseEasting = 50000
        FalseNorthing = 2285000
        CSF = 1.00005831
        CM = 127.253333333333
    End If
    If UCase(proj) = "LAVERTON2020" Then
        FalseEasting = 50000
        FalseNorthing = 3420000
        CSF = 1.00006488
        CM = 122.1
    End If
    If UCase(proj) = "LAVERTON94" Then
        FalseEasting = 50000
        FalseNorthing = 3320000
        CSF = 1.00006488
        CM = 122.1
    End If
    If UCase(proj) = "LEINSTER2020" Then
        FalseEasting = 50000
        FalseNorthing = 3400000
        CSF = 1.00007778
        CM = 120.642855555556
    End If
    If UCase(proj) = "LEINSTER94" Then
        FalseEasting = 50000
        FalseNorthing = 3300000
        CSF = 1.00007778
        CM = 120.642855555556
    End If
    If UCase(proj) = "LEWIS2020" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00006662
        CM = 128.8
    End If
    If UCase(proj) = "LEWIS94" Then
        FalseEasting = 50000
        FalseNorthing = 2500000
        CSF = 1.00006662
        CM = 128.8
    End If
    If UCase(proj) = "LOGUE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2300000
        CSF = 1.000008
        CM = 123.016666666667
    End If
    If UCase(proj) = "LOGUE94" Then
        FalseEasting = 50000
        FalseNorthing = 2200000
        CSF = 1.000008
        CM = 123.016666666667
    End If
    If UCase(proj) = "LOUISADOWNS2020" Then
        FalseEasting = 50000
        FalseNorthing = 2408000
        CSF = 1.00004533
        CM = 126.6925
    End If
    If UCase(proj) = "LOUISADOWNS94" Then
        FalseEasting = 50000
        FalseNorthing = 2308000
        CSF = 1.00004533
        CM = 126.6925
    End If
    If UCase(proj) = "MADURA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00000097
        CM = 127.033333333333
    End If
    If UCase(proj) = "MADURA94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00000097
        CM = 127.033333333333
    End If
    If UCase(proj) = "MALCOLM2020" Then
        FalseEasting = 50000
        FalseNorthing = 3880000
        CSF = 1.00004965
        CM = 122.8
    End If
    If UCase(proj) = "MALCOLM94" Then
        FalseEasting = 50000
        FalseNorthing = 3780000
        CSF = 1.00004965
        CM = 122.8
    End If
    If UCase(proj) = "MARBLE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00002388
        CM = 119.85
    End If
    If UCase(proj) = "MARBLE94" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00002388
        CM = 119.85
    End If
    If UCase(proj) = "MARVEL 2020" Then
        FalseEasting = 50000
        FalseNorthing = 3760000
        CSF = 1.00006238
        CM = 119.6
    End If
    If UCase(proj) = "MARVEL 94" Then
        FalseEasting = 50000
        FalseNorthing = 3660000
        CSF = 1.00006238
        CM = 119.6
    End If
    If UCase(proj) = "MEDA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2280000
        CSF = 1.00001996
        CM = 124.083333333333
    End If
    If UCase(proj) = "MEDA94" Then
        FalseEasting = 50000
        FalseNorthing = 2180000
        CSF = 1.00001996
        CM = 124.083333333333
    End If
    If UCase(proj) = "MEEKA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3100000
        CSF = 1.00007644
        CM = 118.3
    End If
    If UCase(proj) = "MEEKA94" Then
        FalseEasting = 50000
        FalseNorthing = 3000000
        CSF = 1.00007644
        CM = 118.3
    End If
    If UCase(proj) = "MENZIES2020" Then
        FalseEasting = 50000
        FalseNorthing = 3600000
        CSF = 1.00006489
        CM = 121.066666666667
    End If
    If UCase(proj) = "MENZIES94" Then
        FalseEasting = 50000
        FalseNorthing = 3500000
        CSF = 1.00006489
        CM = 121.066666666667
    End If
    If UCase(proj) = "MERREDIN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00004949
        CM = 118.283333333333
    End If
    If UCase(proj) = "MERREDIN94" Then
        FalseEasting = 50000
        FalseNorthing = 3700000
        CSF = 1.00004949
        CM = 118.283333333333
    End If
    If UCase(proj) = "MILLSTREAM2020" Then
        FalseEasting = 50000
        FalseNorthing = 2670000
        CSF = 1.0000286
        CM = 117.033333333333
    End If
    If UCase(proj) = "MILLSTREAM94" Then
        FalseEasting = 50000
        FalseNorthing = 2570000
        CSF = 1.0000286
        CM = 117.033333333333
    End If
    If UCase(proj) = "MINGENEW2020" Then
        FalseEasting = 50000
        FalseNorthing = 3592000
        CSF = 1.00002497
        CM = 115.333333333333
    End If
    If UCase(proj) = "MINGENEW94" Then
        FalseEasting = 50000
        FalseNorthing = 3492000
        CSF = 1.00002497
        CM = 115.333333333333
    End If
    If UCase(proj) = "MOONERA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3878000
        CSF = 1.00001128
        CM = 126.440833333333
    End If
    If UCase(proj) = "MOONERA94" Then
        FalseEasting = 50000
        FalseNorthing = 3778000
        CSF = 1.00001128
        CM = 126.440833333333
    End If
    If UCase(proj) = "MORLEY2020" Then
        FalseEasting = 50000
        FalseNorthing = 4200000
        CSF = 1.00000126
        CM = 117.466666666667
    End If
    If UCase(proj) = "MORLEY94" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00000126
        CM = 117.466666666667
    End If
    If UCase(proj) = "MTHART2020" Then
        FalseEasting = 50000
        FalseNorthing = 2226000
        CSF = 1.00004971
        CM = 125.140277777778
    End If
    If UCase(proj) = "MTHART94" Then
        FalseEasting = 50000
        FalseNorthing = 2126000
        CSF = 1.00004971
        CM = 125.140277777778
    End If
    If UCase(proj) = "MTMAGS2020" Then
        FalseEasting = 50000
        FalseNorthing = 3500000
        CSF = 1.00006269
        CM = 117.8
    End If
    If UCase(proj) = "MTMAGSTH" Then
        FalseEasting = 50000
        FalseNorthing = 3400000
        CSF = 1.00006269
        CM = 117.8
    End If
    If UCase(proj) = "MUCHEA2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00001813
        CM = 115.95
    End If
    If UCase(proj) = "MUCHEA94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00001813
        CM = 115.95
    End If
    If UCase(proj) = "MUELLER2020" Then
        FalseEasting = 50000
        FalseNorthing = 2500000
        CSF = 1.00006096
        CM = 128.3
    End If
    If UCase(proj) = "MUELLER94" Then
        FalseEasting = 50000
        FalseNorthing = 2400000
        CSF = 1.00006096
        CM = 128.3
    End If
    If UCase(proj) = "MUGGAN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3312000
        CSF = 1.0000677
        CM = 125.428333333333
    End If
    If UCase(proj) = "MUGGAN94" Then
        FalseEasting = 50000
        FalseNorthing = 3212000
        CSF = 1.0000677
        CM = 125.428333333333
    End If
    If UCase(proj) = "MUKINBUDIN 2020" Then
        FalseEasting = 50000
        FalseNorthing = 3740000
        CSF = 1.00004965
        CM = 118.283333333333
    End If
    If UCase(proj) = "MUKINBUDIN 94" Then
        FalseEasting = 50000
        FalseNorthing = 3640000
        CSF = 1.00004965
        CM = 118.283333333333
    End If
    If UCase(proj) = "MULGADOWNS2020" Then
        FalseEasting = 50000
        FalseNorthing = 2803000
        CSF = 1.00006568
        CM = 118.53
    End If
    If UCase(proj) = "MULGADOWNS94" Then
        FalseEasting = 50000
        FalseNorthing = 2703000
        CSF = 1.00006568
        CM = 118.53
    End If
    If UCase(proj) = "MUNDRABILLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3878000
        CSF = 0.99999899
        CM = 127.591666666667
    End If
    If UCase(proj) = "MUNDRABILLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3778000
        CSF = 0.99999899
        CM = 127.591666666667
    End If
    If UCase(proj) = "MUNG2020" Then
        FalseEasting = 50000
        FalseNorthing = 4060000
        CSF = 1.00001587
        CM = 120.616666666667
    End If
    If UCase(proj) = "MUNG94" Then
        FalseEasting = 50000
        FalseNorthing = 3960000
        CSF = 1.00001587
        CM = 120.616666666667
    End If
    If UCase(proj) = "MURCHISON2020" Then
        FalseEasting = 50000
        FalseNorthing = 3304000
        CSF = 1.00003763
        CM = 115.958333333333
    End If
    If UCase(proj) = "MURCHISON94" Then
        FalseEasting = 50000
        FalseNorthing = 3204000
        CSF = 1.00003763
        CM = 115.958333333333
    End If
    If UCase(proj) = "NANNINE2020" Then
        FalseEasting = 100000
        FalseNorthing = 3600000
        CSF = 1.00006708881
        CM = 118.251388888889
    End If
    If UCase(proj) = "NANNINE94" Then
        FalseEasting = 100000
        FalseNorthing = 3500000
        CSF = 1.00006709
        CM = 118.251388888889
    End If
    If UCase(proj) = "NANU2020" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00000786
        CM = 115.416666666667
    End If
    If UCase(proj) = "NANU94" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00000786
        CM = 115.416666666667
    End If
    If UCase(proj) = "NAPIER2020" Then
        FalseEasting = 50000
        FalseNorthing = 2241000
        CSF = 1.00002641
        CM = 124.960555555556
    End If
    If UCase(proj) = "NAPIER94" Then
        FalseEasting = 50000
        FalseNorthing = 2141000
        CSF = 1.00002641
        CM = 124.960555555556
    End If
    If UCase(proj) = "NARROGIN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00004541
        CM = 117.166666666667
    End If
    If UCase(proj) = "NARROGIN94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004541
        CM = 117.166666666667
    End If
    If UCase(proj) = "NEWGATE2020" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00004447
        CM = 119.05
    End If
    If UCase(proj) = "NEWGATE94" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00004447
        CM = 119.05
    End If
    If UCase(proj) = "NEWMAN2020" Then
        FalseEasting = 50000
        FalseNorthing = 2950000
        CSF = 1.00008862
        CM = 119.716666666667
    End If
    If UCase(proj) = "NEWMAN94" Then
        FalseEasting = 50000
        FalseNorthing = 2850000
        CSF = 1.00008862
        CM = 119.716666666667
    End If
    If UCase(proj) = "NULL2020" Then
        FalseEasting = 50000
        FalseNorthing = 2781676
        CSF = 1.00006128
        CM = 120.1
    End If
    If UCase(proj) = "NULL94" Then
        FalseEasting = 50000
        FalseNorthing = 2681676
        CSF = 1.00006128
        CM = 120.1
    End If
    If UCase(proj) = "OAKOVER2020" Then
        FalseEasting = 50000
        FalseNorthing = 2703000
        CSF = 1.00003722
        CM = 121.079444444444
    End If
    If UCase(proj) = "OAKOVER94" Then
        FalseEasting = 50000
        FalseNorthing = 2603000
        CSF = 1.00003722
        CM = 121.079444444444
    End If
    If UCase(proj) = "ONGERUP2020" Then
        FalseEasting = 50000
        FalseNorthing = 4111000
        CSF = 1.00003807
        CM = 118.475
    End If
    If UCase(proj) = "ONGERUP94" Then
        FalseEasting = 50000
        FalseNorthing = 4011000
        CSF = 1.00003807
        CM = 118.475
    End If
    If UCase(proj) = "PAMELIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 2919000
        CSF = 1.00010951543
        CM = 119.425833333333
    End If
    If UCase(proj) = "PAMELIA94" Then
        FalseEasting = 50000
        FalseNorthing = 2819000
        CSF = 1.00010952
        CM = 119.425833333333
    End If
    If UCase(proj) = "PARDOO2020" Then
        FalseEasting = 50000
        FalseNorthing = 2561000
        CSF = 1.00000831
        CM = 119.928333333333
    End If
    If UCase(proj) = "PARDOO94" Then
        FalseEasting = 50000
        FalseNorthing = 2461000
        CSF = 1.00000831
        CM = 119.928333333333
    End If
    If UCase(proj) = "PENDER2020" Then
        FalseEasting = 50000
        FalseNorthing = 2167000
        CSF = 1.00000698
        CM = 122.899444444444
    End If
    If UCase(proj) = "PENDER94" Then
        FalseEasting = 50000
        FalseNorthing = 2067000
        CSF = 1.00000698
        CM = 122.899444444444
    End If
    If UCase(proj) = "PERUP2020" Then
        FalseEasting = 50000
        FalseNorthing = 4107000
        CSF = 1.00003344
        CM = 116.533333333333
    End If
    If UCase(proj) = "PERUP94" Then
        FalseEasting = 50000
        FalseNorthing = 4007000
        CSF = 1.00003344
        CM = 116.533333333333
    End If
    If UCase(proj) = "PINDARTS2020" Then
        FalseEasting = 50000
        FalseNorthing = 3500000
        CSF = 1.00004321
        CM = 115.8
    End If
    If UCase(proj) = "PINDARTS94" Then
        FalseEasting = 50000
        FalseNorthing = 3400000
        CSF = 1.00004321
        CM = 115.8
    End If
    If UCase(proj) = "POINTA2020" Then
        FalseEasting = 50000
        FalseNorthing = 4130000
        CSF = 1.00000466
        CM = 119.45
    End If
    If UCase(proj) = "POINTANN" Then
        FalseEasting = 50000
        FalseNorthing = 4030000
        CSF = 1.00000466
        CM = 119.45
    End If
    If UCase(proj) = "QUANBUN2020" Then
        FalseEasting = 50000
        FalseNorthing = 2337000
        CSF = 1.00002877
        CM = 125.083333333333
    End If
    If UCase(proj) = "QUANBUN94" Then
        FalseEasting = 50000
        FalseNorthing = 2237000
        CSF = 1.00002877
        CM = 125.083333333333
    End If
    If UCase(proj) = "RAVENS2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.0000319
        CM = 120.033333333333
    End If
    If UCase(proj) = "RAVENS94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.0000319
        CM = 120.033333333333
    End If
    If UCase(proj) = "RIPON2020" Then
        FalseEasting = 50000
        FalseNorthing = 2696600
        CSF = 1.00004948
        CM = 120.505
    End If
    If UCase(proj) = "RIPON94" Then
        FalseEasting = 50000
        FalseNorthing = 2596600
        CSF = 1.00004948
        CM = 120.505
    End If
    If UCase(proj) = "SANDFIRE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2466000
        CSF = 1.00000486
        CM = 121.3
    End If
    If UCase(proj) = "SANDFIRE94" Then
        FalseEasting = 50000
        FalseNorthing = 2366000
        CSF = 1.00000486
        CM = 121.3
    End If
    If UCase(proj) = "SCADDAN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4100000
        CSF = 1.00002781
        CM = 121.7
    End If
    If UCase(proj) = "SCADDAN94" Then
        FalseEasting = 50000
        FalseNorthing = 4000000
        CSF = 1.00002781
        CM = 121.7
    End If
    If UCase(proj) = "SCROSS2020" Then
        FalseEasting = 50000
        FalseNorthing = 3760000
        CSF = 1.0000528
        CM = 119.05
    End If
    If UCase(proj) = "SCROSS94" Then
        FalseEasting = 50000
        FalseNorthing = 3660000
        CSF = 1.0000528
        CM = 119.05
    End If
    If UCase(proj) = "SHANNON2020" Then
        FalseEasting = 50000
        FalseNorthing = 4184000
        CSF = 1.00000776
        CM = 116.45
    End If
    If UCase(proj) = "SHANNON94" Then
        FalseEasting = 50000
        FalseNorthing = 4084000
        CSF = 1.00000776
        CM = 116.45
    End If
    If UCase(proj) = "SHAWRVR2020" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00001084
        CM = 119.25
    End If
    If UCase(proj) = "SHAWRVR94" Then
        FalseEasting = 50000
        FalseNorthing = 2500000
        CSF = 1.00001084
        CM = 119.25
    End If
    If UCase(proj) = "SHERWOOD2020" Then
        FalseEasting = 50000
        FalseNorthing = 3281578
        CSF = 1.00008351
        CM = 118.863888888889
    End If
    If UCase(proj) = "SHERWOOD94" Then
        FalseEasting = 50000
        FalseNorthing = 3181578
        CSF = 1.00008351
        CM = 118.863888888889
    End If
    If UCase(proj) = "STOKES 2020" Then
        FalseEasting = 50000
        FalseNorthing = 4080000
        CSF = 1.00000676
        CM = 121.183333333333
    End If
    If UCase(proj) = "STOKES 94" Then
        FalseEasting = 50000
        FalseNorthing = 3980000
        CSF = 1.00000676
        CM = 121.183333333333
    End If
    If UCase(proj) = "SUNDAY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2705744
        CSF = 1.00001446
        CM = 116.4
    End If
    If UCase(proj) = "SUNDAY94" Then
        FalseEasting = 50000
        FalseNorthing = 2605744
        CSF = 1.00001446
        CM = 116.4
    End If
    If UCase(proj) = "THROSSELL2020" Then
        FalseEasting = 50000
        FalseNorthing = 3381000
        CSF = 1.00006116
        CM = 124.03
    End If
    If UCase(proj) = "THROSSELL94" Then
        FalseEasting = 50000
        FalseNorthing = 3281000
        CSF = 1.00006116
        CM = 124.03
    End If
    If UCase(proj) = "TIEBIN2020" Then
        FalseEasting = 50000
        FalseNorthing = 4082000
        CSF = 1.00005075
        CM = 118.366666666667
    End If
    If UCase(proj) = "TIEBIN94" Then
        FalseEasting = 50000
        FalseNorthing = 3982000
        CSF = 1.00005075
        CM = 118.366666666667
    End If
    If UCase(proj) = "TJUKAYIRLA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3345000
        CSF = 1.0000677
        CM = 124.692222222222
    End If
    If UCase(proj) = "TJUKAYIRLA94" Then
        FalseEasting = 50000
        FalseNorthing = 3245000
        CSF = 1.0000677
        CM = 124.692222222222
    End If
    If UCase(proj) = "TOMPRICE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2850000
        CSF = 1.00010103
        CM = 117.733333333333
    End If
    If UCase(proj) = "TOMPRICE94" Then
        FalseEasting = 50000
        FalseNorthing = 2750000
        CSF = 1.00010103
        CM = 117.733333333333
    End If
    If UCase(proj) = "URANDY2020" Then
        FalseEasting = 50000
        FalseNorthing = 2868720
        CSF = 1.0000253
        CM = 115.933333333333
    End If
    If UCase(proj) = "URANDY94" Then
        FalseEasting = 50000
        FalseNorthing = 2768720
        CSF = 1.0000253
        CM = 115.933333333333
    End If
    If UCase(proj) = "VIRGINIA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3880000
        CSF = 1.00001555
        CM = 125.183333333333
    End If
    If UCase(proj) = "VIRGINIA94" Then
        FalseEasting = 50000
        FalseNorthing = 3780000
        CSF = 1.00001555
        CM = 125.183333333333
    End If
    If UCase(proj) = "WALKAWAY2020" Then
        FalseEasting = 50000
        FalseNorthing = 3493000
        CSF = 1.00000479
        CM = 114.816666666667
    End If
    If UCase(proj) = "WALKAWAY94" Then
        FalseEasting = 50000
        FalseNorthing = 3393000
        CSF = 1.00000479
        CM = 114.816666666667
    End If
    If UCase(proj) = "WALLAL2020" Then
        FalseEasting = 100000
        FalseNorthing = 2538000
        CSF = 1.00000547
        CM = 120.576388888889
    End If
    If UCase(proj) = "WALLAL94" Then
        FalseEasting = 100000
        FalseNorthing = 2438000
        CSF = 1.00000547
        CM = 120.576388888889
    End If
    If UCase(proj) = "WALPOLE2020" Then
        FalseEasting = 50000
        FalseNorthing = 4234000
        CSF = 1.00000143
        CM = 116.683333333333
    End If
    If UCase(proj) = "WALPOLE94" Then
        FalseEasting = 50000
        FalseNorthing = 4134000
        CSF = 1.00000143
        CM = 116.683333333333
    End If
    If UCase(proj) = "WANNAN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3158000
        CSF = 1.00008133
        CM = 127.252777777778
    End If
    If UCase(proj) = "WANNAN94" Then
        FalseEasting = 50000
        FalseNorthing = 3058000
        CSF = 1.00008133
        CM = 127.252777777778
    End If
    If UCase(proj) = "WARAKURNA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3108000
        CSF = 1.00009323
        CM = 128.629722222222
    End If
    If UCase(proj) = "WARAKURNA94" Then
        FalseEasting = 50000
        FalseNorthing = 3008000
        CSF = 1.00009323
        CM = 128.629722222222
    End If
    If UCase(proj) = "WARBURTON2020" Then
        FalseEasting = 50000
        FalseNorthing = 3212000
        CSF = 1.00008307146
        CM = 126.726944444444
    End If
    If UCase(proj) = "WARBURTON94" Then
        FalseEasting = 50000
        FalseNorthing = 3112000
        CSF = 1.00008307
        CM = 126.726944444444
    End If
    If UCase(proj) = "WARRAD2020" Then
        FalseEasting = 50000
        FalseNorthing = 3700000
        CSF = 1.00002058
        CM = 115.316666666667
    End If
    If UCase(proj) = "WARRAD94" Then
        FalseEasting = 50000
        FalseNorthing = 3600000
        CSF = 1.00002058
        CM = 115.316666666667
    End If
    If UCase(proj) = "WARREL 2020" Then
        FalseEasting = 50000
        FalseNorthing = 2640000
        CSF = 1.00001697
        CM = 119.65
    End If
    If UCase(proj) = "WARREL 94" Then
        FalseEasting = 50000
        FalseNorthing = 2540000
        CSF = 1.00001697
        CM = 119.65
    End If
    If UCase(proj) = "WEEBO2020" Then
        FalseEasting = 50000
        FalseNorthing = 3449000
        CSF = 1.00006602
        CM = 121
    End If
    If UCase(proj) = "WEEBO94" Then
        FalseEasting = 50000
        FalseNorthing = 3349000
        CSF = 1.00006602
        CM = 121
    End If
    If UCase(proj) = "WESTANGELAS2020" Then
        FalseEasting = 50000
        FalseNorthing = 2900000
        CSF = 1.00010367
        CM = 119.015277777778
    End If
    If UCase(proj) = "WESTANGELAS94" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00010367
        CM = 119.015277777778
    End If
    If UCase(proj) = "WHIMCRK2020" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00000534
        CM = 117.766666666667
    End If
    If UCase(proj) = "WHIMCRK94" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00000534
        CM = 117.766666666667
    End If
    If UCase(proj) = "WHIMEAST2020" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00000456
        CM = 117.933333333333
    End If
    If UCase(proj) = "WHIMEAST94" Then
        FalseEasting = 50000
        FalseNorthing = 2500000
        CSF = 1.00000456
        CM = 117.933333333333
    End If
    If UCase(proj) = "WHIMWEST2020" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00000157
        CM = 117.566666666667
    End If
    If UCase(proj) = "WHIMWEST94" Then
        FalseEasting = 50000
        FalseNorthing = 2600000
        CSF = 1.00000157
        CM = 117.566666666667
    End If
    If UCase(proj) = "WHITECLIFFS2020" Then
        FalseEasting = 50000
        FalseNorthing = 3405000
        CSF = 1.00006706
        CM = 123.353055555556
    End If
    If UCase(proj) = "WHITECLIFFS94" Then
        FalseEasting = 50000
        FalseNorthing = 3305000
        CSF = 1.00006706
        CM = 123.353055555556
    End If
    If UCase(proj) = "WILUNA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3186000
        CSF = 1.00007658
        CM = 120.2
    End If
    If UCase(proj) = "WILUNA94" Then
        FalseEasting = 50000
        FalseNorthing = 3086000
        CSF = 1.00007658
        CM = 120.2
    End If
    If UCase(proj) = "WITTENOOM2020" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00008218
        CM = 118.066666666667
    End If
    If UCase(proj) = "WITTENOOM94" Then
        FalseEasting = 50000
        FalseNorthing = 2700000
        CSF = 1.00008218
        CM = 118.066666666667
    End If
    If UCase(proj) = "WOODSTOCK2020" Then
        FalseEasting = 50000
        FalseNorthing = 2686000
        CSF = 1.00003601
        CM = 118.808333333333
    End If
    If UCase(proj) = "WOODSTOCK94" Then
        FalseEasting = 50000
        FalseNorthing = 2586000
        CSF = 1.00003601
        CM = 118.808333333333
    End If
    If UCase(proj) = "WOORLBA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3900000
        CSF = 1.00001681
        CM = 124.65
    End If
    If UCase(proj) = "WOORLBA94" Then
        FalseEasting = 50000
        FalseNorthing = 3800000
        CSF = 1.00001681
        CM = 124.65
    End If
    If UCase(proj) = "WUBIN2020" Then
        FalseEasting = 50000
        FalseNorthing = 3700000
        CSF = 1.00004572
        CM = 116.7
    End If
    If UCase(proj) = "WUBIN94" Then
        FalseEasting = 50000
        FalseNorthing = 3600000
        CSF = 1.00004572
        CM = 116.7
    End If
    If UCase(proj) = "WYLOO2020" Then
        FalseEasting = 50000
        FalseNorthing = 2900000
        CSF = 1.00003802
        CM = 116.5
    End If
    If UCase(proj) = "WYLOO94" Then
        FalseEasting = 50000
        FalseNorthing = 2800000
        CSF = 1.00003802
        CM = 116.5
    End If
    If UCase(proj) = "WYNDHAM2020" Then
        FalseEasting = 50000
        FalseNorthing = 2070000
        CSF = 1.00002357
        CM = 128.2
    End If
    If UCase(proj) = "WYNDHAM94" Then
        FalseEasting = 50000
        FalseNorthing = 1970000
        CSF = 1.00002357
        CM = 128.2
    End If
    If UCase(proj) = "YALGOO2020" Then
        FalseEasting = 50000
        FalseNorthing = 3470000
        CSF = 1.00004739
        CM = 116.85865
    End If
    If UCase(proj) = "YALGOO94" Then
        FalseEasting = 50000
        FalseNorthing = 3370000
        CSF = 1.00004739
        CM = 116.85865
    End If
    If UCase(proj) = "YANNARIE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2872000
        CSF = 1.000014
        CM = 115.083333333333
    End If
    If UCase(proj) = "YANNARIE94" Then
        FalseEasting = 50000
        FalseNorthing = 2772000
        CSF = 1.000014
        CM = 115.083333333333
    End If
    If UCase(proj) = "YARINGA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3200000
        CSF = 1.00000597
        CM = 114.366666666667
    End If
    If UCase(proj) = "YARINGA94" Then
        FalseEasting = 50000
        FalseNorthing = 3100000
        CSF = 1.00000597
        CM = 114.366666666667
    End If
    If UCase(proj) = "YILGALONG2020" Then
        FalseEasting = 50000
        FalseNorthing = 2706000
        CSF = 1.00003953
        CM = 120.775
    End If
    If UCase(proj) = "YILGALONG94" Then
        FalseEasting = 50000
        FalseNorthing = 2606000
        CSF = 1.00003953
        CM = 120.775
    End If
    If UCase(proj) = "YULE2020" Then
        FalseEasting = 50000
        FalseNorthing = 2765000
        CSF = 1.00004626
        CM = 118.833333333333
    End If
    If UCase(proj) = "YULE94" Then
        FalseEasting = 50000
        FalseNorthing = 2665000
        CSF = 1.00004626
        CM = 118.833333333333
    End If
    If UCase(proj) = "YUNA2020" Then
        FalseEasting = 50000
        FalseNorthing = 3451000
        CSF = 1.00003453
        CM = 114.916666666667
    End If
    If UCase(proj) = "YUNA94" Then
        FalseEasting = 50000
        FalseNorthing = 3351000
        CSF = 1.00003453
        CM = 114.916666666667
    End If
End If

End Function
    
Function SevenParaParameters(FrmEllip, ToEllip, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
'Input: - _
 translations (dX, dY, dZ) all in meters ; _
 X, Y and Z rotations in seconds of arc (X_Rot, Y_Rot, Z_Rot) and scale in ppm (s).
'There are two different ways of applying the sign conventions for the rotations. In both _
 cases a positive rotation is an anti-clockwise rotation, when viewed along the positive axis _
 towards the origin but: _
 1. The International Earth Rotation Service (IERS) assumes the rotations to be of the _
 points around the cartesian axes, while; _
 2. The method historically used in Australia assumes the rotations to be of the _
 cartesian axes around the points _
'This script uses method (2)

If FrmEllip = "AGD84" And ToEllip = "GDA94" Then
    dX = -117.763   'm
    dY = -51.51   'm
    dZ = 139.061   'm
    X_Rot = -0.292  'arc sec
    Y_Rot = -0.443  'arc sec
    Z_Rot = -0.277  'arc sec
    s = -0.191   'ppm
End If

If FrmEllip = "GDA94" And ToEllip = "AGD84" Then
    dX = 117.763   'm
    dY = 51.51   'm
    dZ = -139.061   'm
    X_Rot = 0.292  'arc sec
    Y_Rot = 0.443  'arc sec
    Z_Rot = 0.277  'arc sec
    s = 0.191   'ppm
End If

If FrmEllip = "GDA94" And ToEllip = "GDA2020" Then
    dX = 0.06155   'm
    dY = -0.01087   'm
    dZ = -0.04019   'm
    X_Rot = -0.0394924  'arc sec
    Y_Rot = -0.0327221  'arc sec
    Z_Rot = -0.0328979  'arc sec
    s = -0.009994   'ppm
End If

If FrmEllip = "GDA2020" And ToEllip = "GDA94" Then
    dX = -0.06155   'm
    dY = 0.01087   'm
    dZ = 0.04019   'm
    X_Rot = 0.0394924  'arc sec
    Y_Rot = 0.0327221  'arc sec
    Z_Rot = 0.0328979  'arc sec
    s = 0.009994   'ppm
End If

End Function
''''''''''''''''''''''''''''''''''''''''''''''''
'''''''   Code written by Kent Wheeler    ''''''
'''''''  kent.wheeler@landgate.wa.gov.au  ''''''
''''''''''''''''''''''''''''''''''''''''''''''''
Function GeoAz(ByVal Lat1, ByVal Lng1, ByVal Lat2, ByVal Lng2)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

If InStr(1, Lat1, " ") > 0 Then Lat1 = DD(Lat1)
If InStr(1, Lng1, " ") > 0 Then Lng1 = DD(Lng1)
If InStr(1, Lat2, " ") > 0 Then Lat2 = DD(Lat2)
If InStr(1, Lng2, " ") > 0 Then Lng2 = DD(Lng2)

Lat1 = WorksheetFunction.Radians(Lat1)
Lat2 = WorksheetFunction.Radians(Lat2)
Lng1 = WorksheetFunction.Radians(Lng1)
Lng2 = WorksheetFunction.Radians(Lng2)

orient = 0
PI = WorksheetFunction.PI

If Lat1 = Lat2 And Lng1 = Lng2 Then
    orient = 1
    GoTo skp
End If
If Lat1 < Lat2 And Lng1 = Lng2 Then
    orient = 2
End If
If Lat1 > Lat2 And Lng1 = Lng2 Then
    orient = 3
End If

C15 = a
C17 = 1 / OneOnF
C16 = C15 * (1 - C17)
C18 = Lng2 - Lng1

F15 = C17 + C17 * C17
F16 = F15 / 2
F17 = (C17 * C17) / 2
F18 = (C17 * C17) / 8
F19 = (C17 * C17) / 16

C20 = (1 - C17) * Tan(Lat1)        'tanU1
C21 = Atn(C20)                  'U1
C22 = Sin(C21)                   'sinU1
C23 = Cos(C21)                   'cosU1
C24 = (1 - C17) * Tan(Lat2)    'tanU2
C25 = Atn(C24)                  'U2
C26 = Sin(C25)                   'sinU2
C27 = Cos(C25)                   'cosU2

'formula13
C29 = C18            'L1
'formula14
C31 = (C27 * Sin(C29) * C27 * Sin(C29)) + (C23 * C26 - C22 * C27 * Cos(C29)) * (C23 * C26 - C22 * C27 * Cos(C29))
'formula15
C33 = (C22 * C26) + (C23 * C27 * Cos(C29))
'formula16
C35 = Sqr(C31) / C33
C36 = Atn(C35)
'formula17
If C31 = 0 Then
    C38 = 0
Else
    C38 = C23 * C27 * Sin(C29) / Sqr(C31)
End If
'formula18
If (Cos(Asin(C38)) * Cos(Asin(C38))) = 0 Then
    C40 = 0
Else
    C40 = C33 - 2 * C22 * C26 / (Cos(Asin(C38)) * Cos(Asin(C38)))
End If
C41 = Cos(Asin(C38)) * Cos(Asin(C38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
'formula3
C43 = 1 + C41 / 16384 * (4096 + C41 * (-768 + C41 * (320 - 175 * C41)))
'formula4
C45 = C41 / 1024 * (256 + C41 * (-128 + C41 * (74 - 47 * C41)))
'formula6
C47 = C45 * Sqr(C31) * (C40 + C45 / 4 * (C33 * (-1 + 2 * C40 * C40) - C45 / 6 * C40 * (-3 + 4 * C31) * (-3 + 4 * C40 * C40)))

'formula10
C50 = C17 / 16 * Cos(Asin(C38)) * Cos(Asin(C38)) * (4 + C17 * (4 - 3 * Cos(Asin(C38)) * Cos(Asin(C38))))
'formula11modified
C52 = C18 + (1 - C50) * C17 * C38 * (Acos(C33) + C50 * Sin(Acos(C33)) * (C40 + C50 * C33 * (-1 + 2 * C40 * C40)))
'formula19
If orient = 1 Then
    C54 = 0
Else
    C54 = C16 * C43 * (Atn(C35) - C47)
End If
E29 = 0
chnge = 100
While chnge > 0.0000000001
    E31 = (C27 * Sin(E29) * C27 * Sin(E29)) + (C23 * C26 - C22 * C27 * Cos(E29)) * (C23 * C26 - C22 * C27 * Cos(E29))
    E33 = (C22 * C26) + (C23 * C27 * Cos(E29))
    E35 = Sqr(E31) / E33
    If Atn(E35) < 0 Then
        E36 = Atn(E35) + PI
    Else
        E36 = Atn(E35)
    End If
    If E31 = 0 Then
        E38 = 0
    Else
        E38 = C23 * C27 * Sin(E29) / Sqr(E31)
    End If
    E40 = E33 - 2 * C22 * C26 / (Cos(Asin(E38)) * Cos(Asin(E38)))
    E41 = Cos(Asin(E38)) * Cos(Asin(E38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
    E43 = 1 + E41 / 16384 * (4096 + E41 * (-768 + E41 * (320 - 175 * E41)))
    E45 = E41 / 1024 * (256 + E41 * (-128 + E41 * (74 - 47 * E41)))
    E47 = E45 * Sqr(E31) * (E40 + E45 / 4 * (E33 * (-1 + 2 * E40 * E40) - E45 / 6 * E40 * (-3 + 4 * E31) * (-3 + 4 * E40 * E40)))
    E50 = C17 / 16 * Cos(Asin(E38)) * Cos(Asin(E38)) * (4 + C17 * (4 - 3 * Cos(Asin(E38)) * Cos(Asin(E38))))
    E52 = C18 + (1 - E50) * C17 * E38 * (Acos(E33) + E50 * Sin(Acos(E33)) * (E40 + E50 * E33 * (-1 + 2 * E40 * E40)))
    chnge = Abs(E29 - E52)
    E29 = E52
Wend
If orient = 1 Then
    E54 = 0
Else
    E54 = C16 * E43 * (E36 - E47)
End If
If C23 * C26 - C22 * C27 * Cos(E52) = 0 Then
    E56 = 0
Else
    E56 = C27 * Sin(E52) / (C23 * C26 - C22 * C27 * Cos(E52))
End If
If (-C22 * C27 + C23 * C26 * Cos(C52)) = 0 Then
    E58 = 0
Else
    E58 = C23 * Sin(C52) / (-C22 * C27 + C23 * C26 * Cos(C52))
End If

C63 = Atn(E56)
C64 = (C63 / PI) * 180 + 360
If C64 >= 360 Then
    C65 = C64 - 360
Else
    C65 = C64
End If
F63 = Atn(E58)
F64 = (F63 / PI) * 180 + 180
If F64 >= 360 Then
    F65 = F64 - 360
Else
    F65 = F64
End If
C66 = C65

I63 = WorksheetFunction.Atan2(Lat2 - Lat1, Lng2 - Lng1) * 180 / PI
J63 = WorksheetFunction.Atan2(Lat1 - Lat2, Lng1 - Lng2) * 180 / PI
If I63 < 0 Then
    I64 = C63 * 180 / PI + 360
Else
    I64 = (C63 * 180) / PI
End If
If J63 < 0 Then
    J64 = C63 * 180 / PI + 360
Else
    J64 = (C63 * 180) / PI
End If
I65 = WorksheetFunction.Atan2((C23 * C26 - C22 * C27 * Cos(E52)), C27 * Sin(E52))
J65 = WorksheetFunction.Atan2((-C22 * C27 + C23 * C26 * Cos(E52)), C23 * Sin(E52))
F66 = F65
J66 = J65 * 180 / PI + 180
If I65 < 0 Then
    I66 = I65 * 180 / PI + 360
Else
    I66 = I65 * 180 / PI
End If
skp:
If orient = 1 Then
    C66 = 0
    F66 = 0
    I66 = 0
    J66 = 0
End If
If orient = 2 Then
    C66 = 0
    F66 = 180
    I66 = 0
    J66 = 180
End If
If orient = 3 Then
    C66 = 180
    F66 = 0
    J66 = 0
    I66 = 180
End If

GeoAz = HMS(I66, 2)
Exit Function
Errr:
GeoAz = Err.Description
End Function

Function GeoRevAz(ByVal Lat1, ByVal Lng1, ByVal Lat2, ByVal Lng2)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

If InStr(1, Lat1, " ") > 0 Then Lat1 = DD(Lat1)
If InStr(1, Lng1, " ") > 0 Then Lng1 = DD(Lng1)
If InStr(1, Lat2, " ") > 0 Then Lat2 = DD(Lat2)
If InStr(1, Lng2, " ") > 0 Then Lng2 = DD(Lng2)

Lat1 = WorksheetFunction.Radians(Lat1)
Lat2 = WorksheetFunction.Radians(Lat2)
Lng1 = WorksheetFunction.Radians(Lng1)
Lng2 = WorksheetFunction.Radians(Lng2)

orient = 0
PI = WorksheetFunction.PI

If Lat1 = Lat2 And Lng1 = Lng2 Then
    orient = 1
    GoTo skp
End If
If Lat1 < Lat2 And Lng1 = Lng2 Then
    orient = 2
End If
If Lat1 > Lat2 And Lng1 = Lng2 Then
    orient = 3
End If

C15 = a
C17 = 1 / OneOnF
C16 = C15 * (1 - C17)
C18 = Lng2 - Lng1

F15 = C17 + C17 * C17
F16 = F15 / 2
F17 = (C17 * C17) / 2
F18 = (C17 * C17) / 8
F19 = (C17 * C17) / 16

C20 = (1 - C17) * Tan(Lat1)        'tanU1
C21 = Atn(C20)                  'U1
C22 = Sin(C21)                   'sinU1
C23 = Cos(C21)                   'cosU1
C24 = (1 - C17) * Tan(Lat2)    'tanU2
C25 = Atn(C24)                  'U2
C26 = Sin(C25)                   'sinU2
C27 = Cos(C25)                   'cosU2

'formula13
C29 = C18            'L1
'formula14
C31 = (C27 * Sin(C29) * C27 * Sin(C29)) + (C23 * C26 - C22 * C27 * Cos(C29)) * (C23 * C26 - C22 * C27 * Cos(C29))
'formula15
C33 = (C22 * C26) + (C23 * C27 * Cos(C29))
'formula16
C35 = Sqr(C31) / C33
C36 = Atn(C35)
'formula17
If C31 = 0 Then
    C38 = 0
Else
    C38 = C23 * C27 * Sin(C29) / Sqr(C31)
End If
'formula18
If (Cos(Asin(C38)) * Cos(Asin(C38))) = 0 Then
    C40 = 0
Else
    C40 = C33 - 2 * C22 * C26 / (Cos(Asin(C38)) * Cos(Asin(C38)))
End If
C41 = Cos(Asin(C38)) * Cos(Asin(C38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
'formula3
C43 = 1 + C41 / 16384 * (4096 + C41 * (-768 + C41 * (320 - 175 * C41)))
'formula4
C45 = C41 / 1024 * (256 + C41 * (-128 + C41 * (74 - 47 * C41)))
'formula6
C47 = C45 * Sqr(C31) * (C40 + C45 / 4 * (C33 * (-1 + 2 * C40 * C40) - C45 / 6 * C40 * (-3 + 4 * C31) * (-3 + 4 * C40 * C40)))

'formula10
C50 = C17 / 16 * Cos(Asin(C38)) * Cos(Asin(C38)) * (4 + C17 * (4 - 3 * Cos(Asin(C38)) * Cos(Asin(C38))))
'formula11modified
C52 = C18 + (1 - C50) * C17 * C38 * (Acos(C33) + C50 * Sin(Acos(C33)) * (C40 + C50 * C33 * (-1 + 2 * C40 * C40)))
'formula19
If orient = 1 Then
    C54 = 0
Else
    C54 = C16 * C43 * (Atn(C35) - C47)
End If
E29 = 0
chnge = 100
While chnge > 0.0000000001
    E31 = (C27 * Sin(E29) * C27 * Sin(E29)) + (C23 * C26 - C22 * C27 * Cos(E29)) * (C23 * C26 - C22 * C27 * Cos(E29))
    E33 = (C22 * C26) + (C23 * C27 * Cos(E29))
    E35 = Sqr(E31) / E33
    If Atn(E35) < 0 Then
        E36 = Atn(E35) + PI
    Else
        E36 = Atn(E35)
    End If
    If E31 = 0 Then
        E38 = 0
    Else
        E38 = C23 * C27 * Sin(E29) / Sqr(E31)
    End If
    E40 = E33 - 2 * C22 * C26 / (Cos(Asin(E38)) * Cos(Asin(E38)))
    E41 = Cos(Asin(E38)) * Cos(Asin(E38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
    E43 = 1 + E41 / 16384 * (4096 + E41 * (-768 + E41 * (320 - 175 * E41)))
    E45 = E41 / 1024 * (256 + E41 * (-128 + E41 * (74 - 47 * E41)))
    E47 = E45 * Sqr(E31) * (E40 + E45 / 4 * (E33 * (-1 + 2 * E40 * E40) - E45 / 6 * E40 * (-3 + 4 * E31) * (-3 + 4 * E40 * E40)))
    E50 = C17 / 16 * Cos(Asin(E38)) * Cos(Asin(E38)) * (4 + C17 * (4 - 3 * Cos(Asin(E38)) * Cos(Asin(E38))))
    E52 = C18 + (1 - E50) * C17 * E38 * (Acos(E33) + E50 * Sin(Acos(E33)) * (E40 + E50 * E33 * (-1 + 2 * E40 * E40)))
    chnge = Abs(E29 - E52)
    E29 = E52
Wend
If orient = 1 Then
    E54 = 0
Else
    E54 = C16 * E43 * (E36 - E47)
End If
If C23 * C26 - C22 * C27 * Cos(E52) = 0 Then
    E56 = 0
Else
    E56 = C27 * Sin(E52) / (C23 * C26 - C22 * C27 * Cos(E52))
End If
If (-C22 * C27 + C23 * C26 * Cos(C52)) = 0 Then
    E58 = 0
Else
    E58 = C23 * Sin(C52) / (-C22 * C27 + C23 * C26 * Cos(C52))
End If

C63 = Atn(E56)
C64 = (C63 / PI) * 180 + 360
If C64 >= 360 Then
    C65 = C64 - 360
Else
    C65 = C64
End If
F63 = Atn(E58)
F64 = (F63 / PI) * 180 + 180
If F64 >= 360 Then
    F65 = F64 - 360
Else
    F65 = F64
End If
C66 = C65

I63 = WorksheetFunction.Atan2(Lat2 - Lat1, Lng2 - Lng1) * 180 / PI
J63 = WorksheetFunction.Atan2(Lat1 - Lat2, Lng1 - Lng2) * 180 / PI
If I63 < 0 Then
    I64 = C63 * 180 / PI + 360
Else
    I64 = (C63 * 180) / PI
End If
If J63 < 0 Then
    J64 = C63 * 180 / PI + 360
Else
    J64 = (C63 * 180) / PI
End If
I65 = WorksheetFunction.Atan2((C23 * C26 - C22 * C27 * Cos(E52)), C27 * Sin(E52))
J65 = WorksheetFunction.Atan2((-C22 * C27 + C23 * C26 * Cos(E52)), C23 * Sin(E52))
F66 = F65
J66 = J65 * 180 / PI + 180
If I65 < 0 Then
    I66 = I65 * 180 / PI + 360
Else
    I66 = I65 * 180 / PI
End If
skp:
If orient = 1 Then
    C66 = 0
    F66 = 0
    I66 = 0
    J66 = 0
End If
If orient = 2 Then
    C66 = 0
    F66 = 180
    I66 = 0
    J66 = 180
End If
If orient = 3 Then
    C66 = 180
    F66 = 0
    J66 = 0
    I66 = 180
End If

GeoRevAz = HMS(J66, 2)
Exit Function
Errr:
GeoRevAz = Err.Description
End Function

Function SphDist(ByVal Lat1, ByVal Lng1, ByVal Lat2, ByVal Lng2)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

If InStr(1, Lat1, " ") > 0 Then Lat1 = DD(Lat1)
If InStr(1, Lng1, " ") > 0 Then Lng1 = DD(Lng1)
If InStr(1, Lat2, " ") > 0 Then Lat2 = DD(Lat2)
If InStr(1, Lng2, " ") > 0 Then Lng2 = DD(Lng2)

If Abs(Lat1 - Lat2) < 0.0000000001 And Abs(Lng1 - Lng2) < 0.0000000001 Then
    orient = 1
    GoTo skp
End If
If Lat1 < Lat2 And Abs(Lng1 - Lng2) < 0.0000000001 Then
    orient = 2
End If
If Lat1 > Lat2 And Abs(Lng1 - Lng2)  < 0.0000000001 Then
    orient = 3
End If

Lat1 = WorksheetFunction.Radians(Lat1)
Lat2 = WorksheetFunction.Radians(Lat2)
Lng1 = WorksheetFunction.Radians(Lng1)
Lng2 = WorksheetFunction.Radians(Lng2)

orient = 0
PI = WorksheetFunction.PI

C15 = a
C17 = 1 / OneOnF
C16 = C15 * (1 - C17)
C18 = Lng2 - Lng1

F15 = C17 + C17 * C17
F16 = F15 / 2
F17 = (C17 * C17) / 2
F18 = (C17 * C17) / 8
F19 = (C17 * C17) / 16

C20 = (1 - C17) * Tan(Lat1)        'tanU1
C21 = Atn(C20)                  'U1
C22 = Sin(C21)                   'sinU1
C23 = Cos(C21)                   'cosU1
C24 = (1 - C17) * Tan(Lat2)    'tanU2
C25 = Atn(C24)                  'U2
C26 = Sin(C25)                   'sinU2
C27 = Cos(C25)                   'cosU2

'formula13
C29 = C18            'L1
'formula14
C31 = (C27 * Sin(C29) * C27 * Sin(C29)) + (C23 * C26 - C22 * C27 * Cos(C29)) * (C23 * C26 - C22 * C27 * Cos(C29))
'formula15
C33 = (C22 * C26) + (C23 * C27 * Cos(C29))
'formula16
C35 = Sqr(C31) / C33
C36 = Atn(C35)
'formula17
If C31 = 0 Then
    C38 = 0
Else
    C38 = C23 * C27 * Sin(C29) / Sqr(C31)
End If
'formula18
If (Cos(Asin(C38)) * Cos(Asin(C38))) = 0 Then
    C40 = 0
Else
    C40 = C33 - 2 * C22 * C26 / (Cos(Asin(C38)) * Cos(Asin(C38)))
End If
C41 = Cos(Asin(C38)) * Cos(Asin(C38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
'formula3
C43 = 1 + C41 / 16384 * (4096 + C41 * (-768 + C41 * (320 - 175 * C41)))
'formula4
C45 = C41 / 1024 * (256 + C41 * (-128 + C41 * (74 - 47 * C41)))
'formula6
C47 = C45 * Sqr(C31) * (C40 + C45 / 4 * (C33 * (-1 + 2 * C40 * C40) - C45 / 6 * C40 * (-3 + 4 * C31) * (-3 + 4 * C40 * C40)))

'formula10
C50 = C17 / 16 * Cos(Asin(C38)) * Cos(Asin(C38)) * (4 + C17 * (4 - 3 * Cos(Asin(C38)) * Cos(Asin(C38))))
'formula11modified
C52 = C18 + (1 - C50) * C17 * C38 * (Acos(C33) + C50 * Sin(Acos(C33)) * (C40 + C50 * C33 * (-1 + 2 * C40 * C40)))
'formula19
If orient = 1 Then
    C54 = 0
Else
    C54 = C16 * C43 * (Atn(C35) - C47)
End If
E29 = 0
chnge = 100
While chnge > 0.0000000001
    E31 = (C27 * Sin(E29) * C27 * Sin(E29)) + (C23 * C26 - C22 * C27 * Cos(E29)) * (C23 * C26 - C22 * C27 * Cos(E29))
    E33 = (C22 * C26) + (C23 * C27 * Cos(E29))
    E35 = Sqr(E31) / E33
    If Atn(E35) < 0 Then
        E36 = Atn(E35) + PI
    Else
        E36 = Atn(E35)
    End If
    If E31 = 0 Then
        E38 = 0
    Else
        E38 = C23 * C27 * Sin(E29) / Sqr(E31)
    End If
    E40 = E33 - 2 * C22 * C26 / (Cos(Asin(E38)) * Cos(Asin(E38)))
    E41 = Cos(Asin(E38)) * Cos(Asin(E38)) * (C15 * C15 - C16 * C16) / (C16 * C16)
    E43 = 1 + E41 / 16384 * (4096 + E41 * (-768 + E41 * (320 - 175 * E41)))
    E45 = E41 / 1024 * (256 + E41 * (-128 + E41 * (74 - 47 * E41)))
    E47 = E45 * Sqr(E31) * (E40 + E45 / 4 * (E33 * (-1 + 2 * E40 * E40) - E45 / 6 * E40 * (-3 + 4 * E31) * (-3 + 4 * E40 * E40)))
    E50 = C17 / 16 * Cos(Asin(E38)) * Cos(Asin(E38)) * (4 + C17 * (4 - 3 * Cos(Asin(E38)) * Cos(Asin(E38))))
    E52 = C18 + (1 - E50) * C17 * E38 * (Acos(E33) + E50 * Sin(Acos(E33)) * (E40 + E50 * E33 * (-1 + 2 * E40 * E40)))
    chnge = Abs(E29 - E52)
    E29 = E52
Wend
If orient = 1 Then
    E54 = 0
Else
    E54 = C16 * E43 * (E36 - E47)
End If
If C23 * C26 - C22 * C27 * Cos(E52) = 0 Then
    E56 = 0
Else
    E56 = C27 * Sin(E52) / (C23 * C26 - C22 * C27 * Cos(E52))
End If
If (-C22 * C27 + C23 * C26 * Cos(C52)) = 0 Then
    E58 = 0
Else
    E58 = C23 * Sin(C52) / (-C22 * C27 + C23 * C26 * Cos(C52))
End If

C63 = Atn(E56)
C64 = (C63 / PI) * 180 + 360
If C64 >= 360 Then
    C65 = C64 - 360
Else
    C65 = C64
End If
F63 = Atn(E58)
F64 = (F63 / PI) * 180 + 180
If F64 >= 360 Then
    F65 = F64 - 360
Else
    F65 = F64
End If
C66 = C65

I63 = WorksheetFunction.Atan2(Lat2 - Lat1, Lng2 - Lng1) * 180 / PI
J63 = WorksheetFunction.Atan2(Lat1 - Lat2, Lng1 - Lng2) * 180 / PI
If I63 < 0 Then
    I64 = C63 * 180 / PI + 360
Else
    I64 = (C63 * 180) / PI
End If
If J63 < 0 Then
    J64 = C63 * 180 / PI + 360
Else
    J64 = (C63 * 180) / PI
End If
On Error GoTo skp
I65 = WorksheetFunction.Atan2((C23 * C26 - C22 * C27 * Cos(E52)), C27 * Sin(E52))
J65 = WorksheetFunction.Atan2((-C22 * C27 + C23 * C26 * Cos(E52)), C23 * Sin(E52))
F66 = F65
J66 = J65 * 180 / PI + 180
If I65 < 0 Then
    I66 = I65 * 180 / PI + 360
Else
    I66 = I65 * 180 / PI
End If
skp:
If orient = 1 Then
    C66 = 0
    F66 = 0
    I66 = 0
    J66 = 0
    E54 = 0
End If
If orient = 2 Then
    C66 = 0
    F66 = 180
    I66 = 0
    J66 = 180
End If
If orient = 3 Then
    C66 = 180
    F66 = 0
    J66 = 0
    I66 = 180
End If

SphDist = E54
Exit Function
Errr:
SphDist = Err.Description
End Function
Function FwdLat(ByVal Lat, ByVal Lng, ByVal Az, ByVal G4)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

C10 = a
C12 = 1 / OneOnF
C11 = C10 * (1 - C12)
C13 = (C10 * C10 - C11 * C11) / (C11 * C11)
PI = WorksheetFunction.PI

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
If InStr(1, Az, " ") > 0 Then Az = DD(Az)

F11 = Lat
F12 = Lng
F13 = Az

G11 = F11 * PI / 180
G12 = F12 * PI / 180
G13 = F13 * PI / 180

C17 = (1 - C12) * Tan(G11)
C18 = C17 / Cos(G13)

C20 = Sin(Atn(C17))
C21 = Cos(Atn(C17))
C22 = C21 * Sin(G13)

C15 = Cos(Asin(C22)) * Cos(Asin(C22)) * ((C10 * C10) - (C11 * C11)) / (C11 * C11)

C24 = 1 + C15 / 16384 * (4096 + C15 * (-768 + C15 * (320 - 175 * C15)))

C26 = C15 / 1024 * (256 + C15 * (-128 + C15 * (74 - 47 * C15)))

prevC39 = 100
C39 = 10
While Abs(C39 - prevC39) > 0.00000000000001
    C28 = 2 * Atn(C18) + C39
    
    F28 = 2 * Atn(C18)
    F29 = C26 * Sin(C39)
    F30 = Cos(C28)
    F31 = Cos(C39) * (-1 + 2 * F30 * F30)
    
    F33 = -3 + 4 * Sin(C39) * Sin(C39)
    F34 = -3 + 4 * F30 * F30
    
    F37 = F29 * (F30 + C26 / 4 * (F31 - C26 / 6 * F30 * F33 * F34))
    C37 = C26 * Sin(C39) * (Cos(C28) + (1 / 4) * C26 * (Cos(C39) * (-1 + 2 * Cos(C28) * Cos(C28)) - (1 / 6) * C26 * Cos(C28) * (-3 + 4 * Sin(C39) * Sin(C39)) * (-3 + 4 * Cos(C28) * Cos(C28))))
    
    prevC39 = C39
    C39 = G4 / (C11 * C24) + F37
Wend

C41 = Sin(Atn(C17)) * Cos(C39) + Cos(Atn(C17)) * Sin(C39) * Cos(G13)
C42 = (1 - C12) * Sqr(C22 * C22 + (Sin(Atn(C17)) * Sin(C39) - Cos(Atn(C17)) * Cos(C39) * Cos(G13)) * (Sin(Atn(C17)) * Sin(C39) - Cos(Atn(C17)) * Cos(C39) * Cos(G13)))
C43 = C41 / C42
F43 = Atan2(C42, C41)
C44 = Atn(C43)
C45 = C44 * 180 / PI
F45 = Round(F43 * 180 / PI, 8)

FwdLat = HMS(F45, 5)
Exit Function
Errr:
FwdLat = Err.Description
End Function
Function FwdLong(ByVal Lat, ByVal Lng, ByVal Az, ByVal G4)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

C10 = a
C12 = 1 / OneOnF
C11 = C10 * (1 - C12)
C13 = (C10 * C10 - C11 * C11) / (C11 * C11)
PI = WorksheetFunction.PI

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
If InStr(1, Az, " ") > 0 Then Az = DD(Az)

F11 = Lat
F12 = Lng
F13 = Az

G11 = F11 * PI / 180
G12 = F12 * PI / 180
G13 = F13 * PI / 180

C17 = (1 - C12) * Tan(G11)
C18 = C17 / Cos(G13)

C20 = Sin(Atn(C17))
C21 = Cos(Atn(C17))
C22 = C21 * Sin(G13)

C15 = Cos(Asin(C22)) * Cos(Asin(C22)) * ((C10 * C10) - (C11 * C11)) / (C11 * C11)

C24 = 1 + C15 / 16384 * (4096 + C15 * (-768 + C15 * (320 - 175 * C15)))

C26 = C15 / 1024 * (256 + C15 * (-128 + C15 * (74 - 47 * C15)))

prevC39 = 100
C39 = 10
While Abs(C39 - prevC39) > 0.00000000000001
    C28 = 2 * Atn(C18) + C39
    
    F28 = 2 * Atn(C18)
    F29 = C26 * Sin(C39)
    F30 = Cos(C28)
    F31 = Cos(C39) * (-1 + 2 * F30 * F30)
    
    F33 = -3 + 4 * Sin(C39) * Sin(C39)
    F34 = -3 + 4 * F30 * F30
    
    F37 = F29 * (F30 + C26 / 4 * (F31 - C26 / 6 * F30 * F33 * F34))
    C37 = C26 * Sin(C39) * (Cos(C28) + (1 / 4) * C26 * (Cos(C39) * (-1 + 2 * Cos(C28) * Cos(C28)) - (1 / 6) * C26 * Cos(C28) * (-3 + 4 * Sin(C39) * Sin(C39)) * (-3 + 4 * Cos(C28) * Cos(C28))))
    
    prevC39 = C39
    C39 = G4 / (C11 * C24) + F37
Wend

C47 = Sin(C39) * Sin(G13) / (C21 * Cos(C39) - Sin(Atn(C17)) * Sin(C39) * Cos(G13))
C48 = Atan2((C21 * Cos(C39) - Sin(Atn(C17)) * Sin(C39) * Cos(G13)), Sin(C39) * Sin(G13))

C50 = C12 / 16 * Cos(Asin(C22)) * Cos(Asin(C22)) * (4 + C12 * (4 - 3 * Cos(Asin(C22)) * Cos(Asin(C22))))

C52 = C48 - (1 - C50) * C12 * C22 * (C39 + C50 * Sin(C39) * (Cos(C28) + C50 * Cos(C39) * (-1 + 2 * Cos(C28) * Cos(C28))))
C53 = G12 + C52
If C53 < 0 Then
    C54 = C53 * 180 / PI + 180
Else
    C54 = C53 * 180 / PI
End If

F54 = Round(C53 * 180 / PI, 8)
If F54 < -179.999999999 Then
    F55 = 360 + F54
Else
    If F54 > 179.999999999 Then
        F55 = -360 + F54
    Else
        F55 = F54
    End If
End If

FwdLong = HMS(F55, 5)
Exit Function
Errr:
FwdLong = Err.Description
End Function
Function FwdRevAz(ByVal Lat, ByVal Lng, ByVal Az, ByVal G4)
On Error GoTo Errr
Call EllipsoidalParameters("GRS80", a, b, OneOnF)

C10 = a
C12 = 1 / OneOnF
C11 = C10 * (1 - C12)
C13 = (C10 * C10 - C11 * C11) / (C11 * C11)
PI = WorksheetFunction.PI

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
If InStr(1, Az, " ") > 0 Then Az = DD(Az)

F11 = Lat
F12 = Lng
F13 = Az

G11 = F11 * PI / 180
G12 = F12 * PI / 180
G13 = F13 * PI / 180

C17 = (1 - C12) * Tan(G11)
C18 = C17 / Cos(G13)

C20 = Sin(Atn(C17))
C21 = Cos(Atn(C17))
C22 = C21 * Sin(G13)

C15 = Cos(Asin(C22)) * Cos(Asin(C22)) * ((C10 * C10) - (C11 * C11)) / (C11 * C11)

C24 = 1 + C15 / 16384 * (4096 + C15 * (-768 + C15 * (320 - 175 * C15)))

C26 = C15 / 1024 * (256 + C15 * (-128 + C15 * (74 - 47 * C15)))

prevC39 = 100
C39 = 10
While Abs(C39 - prevC39) > 0.00000000000001
    C28 = 2 * Atn(C18) + C39
    
    F28 = 2 * Atn(C18)
    F29 = C26 * Sin(C39)
    F30 = Cos(C28)
    F31 = Cos(C39) * (-1 + 2 * F30 * F30)
    
    F33 = -3 + 4 * Sin(C39) * Sin(C39)
    F34 = -3 + 4 * F30 * F30
    
    F37 = F29 * (F30 + C26 / 4 * (F31 - C26 / 6 * F30 * F33 * F34))
    C37 = C26 * Sin(C39) * (Cos(C28) + (1 / 4) * C26 * (Cos(C39) * (-1 + 2 * Cos(C28) * Cos(C28)) - (1 / 6) * C26 * Cos(C28) * (-3 + 4 * Sin(C39) * Sin(C39)) * (-3 + 4 * Cos(C28) * Cos(C28))))
    
    prevC39 = C39
    C39 = G4 / (C11 * C24) + F37
Wend

C41 = Sin(Atn(C17)) * Cos(C39) + Cos(Atn(C17)) * Sin(C39) * Cos(G13)

C56 = C22 / (-Sin(Atn(C17)) * Sin(C39) + C21 * Cos(C39) * Cos(G13))
C57 = Atn(C56)
C58 = C57 * 180 / PI + 180

F57 = Atan2((-C20 * Sin(C39) + C21 * Cos(C39) * Cos(G13)), C22)
F58 = F57 * 180 / PI + 180

FwdRevAz = HMS(F58, 2)
Exit Function
Errr:
FwdRevAz = Err.Description
End Function
Function Asin(num)
Asin = WorksheetFunction.Asin(num)
End Function
Function Acos(num)
If Val(num) = 1 Then
    Acos = 0
Else
    Acos = WorksheetFunction.Acos(num)
End If
End Function
Function Atan2(one, two)
Atan2 = WorksheetFunction.Atan2(one, two)
End Function

Function GridEasting(Lat, Lng, proj)
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
PI = WorksheetFunction.PI
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)

'Constants for GRS80 and GDA-MGA
ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)

G4 = WorksheetFunction.Radians(Lat)

Zone = Int((Lng - ConstantC25) / ConstantC23)
C7 = Zone
G7 = WorksheetFunction.Radians((C7 * ConstantC23) + ConstantC26)
M4 = WorksheetFunction.Radians(Lng)
M6 = M4 - G7

E9 = Sin(G4)
E10 = Sin(2 * G4)
E11 = Sin(4 * G4)
E12 = Sin(6 * G4)

K10 = ConstantC8
K11 = K10 * K10
K12 = K10 * K11

M9 = 1 - (K10 / 4) - ((3 * K11) / 64) - ((5 * K12) / 256)
M10 = (3 / 8) * (K10 + (K11 / 4) + ((15 * K12) / 128))
m11 = (15 / 256) * (K11 + ((3 * K12) / 4))
m12 = (35 * K12) / 3072

E15 = (a) * M9 * G4
E16 = -(a) * M10 * E10
E17 = (a) * m11 * E11
E18 = -(a) * m12 * E12
E19 = E15 + E16 + E17 + E18

K15 = a * (1 - K10) / (1 - (K10 * E9 * E9)) ^ 1.5
K16 = a / (1 - (K10 * E9 * E9)) ^ 0.5

E22 = Cos(G4)
E23 = E22 * E22
E24 = E23 * E22
E25 = E23 * E23
E26 = E24 * E23
E27 = E24 * E24
E28 = E27 * E22

H22 = M6
H23 = H22 * H22
H24 = H23 * H22
H25 = H23 * H23
H26 = H24 * H23
H27 = H24 * H24
H28 = H27 * H22
H29 = H25 * H25

K22 = Tan(G4)
K23 = K22 * K22
K25 = K23 * K23
K27 = K23 * K25

m22 = K16 / K15
m23 = m22 * m22
M24 = m22 * m23
M25 = m23 * m23

'Easting
E33 = K16 * H22 * E22
E34 = K16 * H24 * E24 * (m22 - K23) / 6
E35 = K16 * H26 * E26 * (4 * M24 * (1 - 6 * K23) + m23 * (1 + 8 * K23) - m22 * (2 * K23) + K25) / 120
E36 = K16 * H28 * E28 * (61 - 479 * K23 + 179 * K25 - K27) / 5040
E37 = E33 + E34 + E35 + E36
E38 = ConstantC22 * E37
E39 = ConstantC20
GridEasting = E38 + E39
Exit Function
Errr:
GridEasting = Err.Description
End Function
Function GridNorthing(Lat, Lng, proj)
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)

G4 = WorksheetFunction.Radians(Lat)

Zone = Int((Lng - ConstantC25) / ConstantC23)
C7 = Zone
G7 = WorksheetFunction.Radians((C7 * ConstantC23) + ConstantC26)
M4 = WorksheetFunction.Radians(Lng)
M6 = M4 - G7

E9 = Sin(G4)
E10 = Sin(2 * G4)
E11 = Sin(4 * G4)
E12 = Sin(6 * G4)

K10 = ConstantC8
K11 = K10 * K10
K12 = K10 * K11

M9 = 1 - (K10 / 4) - ((3 * K11) / 64) - ((5 * K12) / 256)
M10 = (3 / 8) * (K10 + (K11 / 4) + ((15 * K12) / 128))
m11 = (15 / 256) * (K11 + ((3 * K12) / 4))
m12 = (35 * K12) / 3072

E15 = (a) * M9 * G4
E16 = -(a) * M10 * E10
E17 = (a) * m11 * E11
E18 = -(a) * m12 * E12
E19 = E15 + E16 + E17 + E18

K15 = a * (1 - K10) / (1 - (K10 * E9 * E9)) ^ 1.5
K16 = a / (1 - (K10 * E9 * E9)) ^ 0.5

E22 = Cos(G4)
E23 = E22 * E22
E24 = E23 * E22
E25 = E23 * E23
E26 = E24 * E23
E27 = E24 * E24
E28 = E27 * E22

H22 = M6
H23 = H22 * H22
H24 = H23 * H22
H25 = H23 * H23
H26 = H24 * H23
H27 = H24 * H24
H28 = H27 * H22
H29 = H25 * H25

K22 = Tan(G4)
K23 = K22 * K22
K25 = K23 * K23
K27 = K23 * K25

m22 = K16 / K15
m23 = m22 * m22
M24 = m22 * m23
M25 = m23 * m23

'Northing
K32 = E19
K33 = K16 * E9 * H23 * E22 / 2
K34 = K16 * E9 * H25 * E24 * (4 * m23 + m22 - K23) / 24
K35 = K16 * E9 * H27 * E26 * (8 * M25 * (11 - 24 * K23) - 28 * M24 * (1 - 6 * K23) + m23 * (1 - 32 * K23) - m22 * (2 * K23) + K25) / 720
K36 = K16 * E9 * H29 * E28 * (1385 - 3111 * K23 + 543 * K25 - K27) / 40320
K37 = K32 + K33 + K34 + K35 + K36
K38 = ConstantC22 * K37
K39 = ConstantC21
GridNorthing = K38 + K39
Exit Function
Errr:
GridNorthing = Err.Description
End Function
Function GridZone(Lat, Lng, proj)
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)
'
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)

If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)

GridZone = Int((Lng - ConstantC25) / ConstantC23)
Exit Function
Errr:
GridZone = Err.Description
End Function

Function GridGridCon(Lat, Lng, proj)
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)

G4 = WorksheetFunction.Radians(Lat)

Zone = Int((Lng - ConstantC25) / ConstantC23)
C7 = Zone
G7 = WorksheetFunction.Radians((C7 * ConstantC23) + ConstantC26)
M4 = WorksheetFunction.Radians(Lng)
M6 = M4 - G7

E9 = Sin(G4)
E10 = Sin(2 * G4)
E11 = Sin(4 * G4)
E12 = Sin(6 * G4)

K10 = ConstantC8
K11 = K10 * K10
K12 = K10 * K11

M9 = 1 - (K10 / 4) - ((3 * K11) / 64) - ((5 * K12) / 256)
M10 = (3 / 8) * (K10 + (K11 / 4) + ((15 * K12) / 128))
m11 = (15 / 256) * (K11 + ((3 * K12) / 4))
m12 = (35 * K12) / 3072

E15 = (a) * M9 * G4
E16 = -(a) * M10 * E10
E17 = (a) * m11 * E11
E18 = -(a) * m12 * E12
E19 = E15 + E16 + E17 + E18

K15 = a * (1 - K10) / (1 - (K10 * E9 * E9)) ^ 1.5
K16 = a / (1 - (K10 * E9 * E9)) ^ 0.5

E22 = Cos(G4)
E23 = E22 * E22
E24 = E23 * E22
E25 = E23 * E23
E26 = E24 * E23
E27 = E24 * E24
E28 = E27 * E22

H22 = M6
H23 = H22 * H22
H24 = H23 * H22
H25 = H23 * H23
H26 = H24 * H23
H27 = H24 * H24
H28 = H27 * H22
H29 = H25 * H25

K22 = Tan(G4)
K23 = K22 * K22
K25 = K23 * K23
K27 = K23 * K25

m22 = K16 / K15
m23 = m22 * m22
M24 = m22 * m23
M25 = m23 * m23

'Grid Convergence
G42 = -E9 * H22
G43 = -E9 * H24 * E23 * (2 * m23 - m22) / 3
G44 = -E9 * H26 * E25 * (M25 * (11 - 24 * K23) - M24 * (11 - 36 * K23) + 2 * m23 * (1 - 7 * K23) + m22 * K23) / 15
G45 = E9 * H28 * E27 * (17 - 26 * K23 + 2 * K25) / 315
G47 = G42 + G43 + G44 + G45
GridGridCon = HMS(WorksheetFunction.Degrees(G47), 2)
Exit Function
Errr:
GridGridCon = Err.Description
End Function

Function GridPntScleFact(Lat, Lng, proj)
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)

G4 = WorksheetFunction.Radians(Lat)

Zone = Int((Lng - ConstantC25) / ConstantC23)
C7 = Zone
G7 = WorksheetFunction.Radians((C7 * ConstantC23) + ConstantC26)
M4 = WorksheetFunction.Radians(Lng)
M6 = M4 - G7

E9 = Sin(G4)
E10 = Sin(2 * G4)
E11 = Sin(4 * G4)
E12 = Sin(6 * G4)

K10 = ConstantC8
K11 = K10 * K10
K12 = K10 * K11

M9 = 1 - (K10 / 4) - ((3 * K11) / 64) - ((5 * K12) / 256)
M10 = (3 / 8) * (K10 + (K11 / 4) + ((15 * K12) / 128))
m11 = (15 / 256) * (K11 + ((3 * K12) / 4))
m12 = (35 * K12) / 3072

E15 = (a) * M9 * G4
E16 = -(a) * M10 * E10
E17 = (a) * m11 * E11
E18 = -(a) * m12 * E12
E19 = E15 + E16 + E17 + E18

K15 = a * (1 - K10) / (1 - (K10 * E9 * E9)) ^ 1.5
K16 = a / (1 - (K10 * E9 * E9)) ^ 0.5

E22 = Cos(G4)
E23 = E22 * E22
E24 = E23 * E22
E25 = E23 * E23
E26 = E24 * E23
E27 = E24 * E24
E28 = E27 * E22

H22 = M6
H23 = H22 * H22
H24 = H23 * H22
H25 = H23 * H23
H26 = H24 * H23
H27 = H24 * H24
H28 = H27 * H22
H29 = H25 * H25

K22 = Tan(G4)
K23 = K22 * K22
K25 = K23 * K23
K27 = K23 * K25

m22 = K16 / K15
m23 = m22 * m22
M24 = m22 * m23
M25 = m23 * m23

'Point Scale Factor
K42 = 1 + (H23 * E23 * m22) / 2
K43 = H25 * E25 * (4 * M24 * (1 - 6 * K23) + m23 * (1 + 24 * K23) - 4 * m22 * K23) / 24
K44 = H27 * E27 * (61 - 148 * K23 + 16 * K25) / 720
K45 = K42 + K43 + K44
GridPntScleFact = ConstantC22 * K45
Exit Function
Errr:
GridPntScleFact = Err.Description
End Function

Function DatumLat(East, North, Zone, proj)
'Easting and Northing (East, North) are specified in metres
'Zone is the UTM zone, this value is one if the project grid does not have multiple zones
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)
PI = WorksheetFunction.PI

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)


E5 = ConstantC20
E6 = East - E5
E7 = E6 / ConstantC22

L5 = ConstantC21
L6 = North - L5
L7 = L6 / ConstantC22
L9 = (L7 * PI) / (ConstantC17 * 180)
L10 = L9 * 2
L11 = L9 * 4
L12 = L9 * 6
L13 = L9 * 8

G16 = L9
G17 = ((3 * ConstantC13 / 2) - (27 * ConstantC15 / 32)) * Sin(L10)
G18 = ((21 * ConstantC14 / 16) - (55 * ConstantC16 / 32)) * Sin(L11)
G19 = (151 * ConstantC15) * Sin(L12) / 96
G20 = 1097 * ConstantC16 * Sin(L13) / 512
G21 = G16 + G17 + G18 + G19 + G20
G23 = Sin(G21)
G24 = 1 / Cos(G21)
J28 = Tan(G21)
L16 = a * (1 - ConstantC8) / (1 - ConstantC8 * G23 * G23) ^ 1.5
L17 = a / (1 - ConstantC8 * G23 * G23) ^ 0.5
G25 = J28 / (ConstantC22 * L17)


G28 = ConstantC13
G29 = G28 * G28
G30 = G28 * G29
G31 = G29 * G29

H28 = E7 / L17
H29 = H28 * H28
H30 = H28 * H29
H31 = H29 * H29
H32 = H31 * H28
H33 = H30 * H30
H34 = H28 * H33

I28 = (E7 * E7) / (L16 * L17)
I29 = I28 * I28
I30 = I28 * I29

J29 = J28 * J28
J30 = J28 * J29
J31 = J29 * J29
J32 = J31 * J28
J33 = J30 * J30

K28 = L17 / L16
K29 = K28 * K28
K30 = K29 * K28
K31 = K29 * K29

'Latitude
G37 = G21
G38 = -((J28 / (ConstantC22 * L16)) * H28 * E6 / 2)
G39 = (J28 / (ConstantC22 * L16)) * (H30 * E6 / 24) * (-4 * K29 + 9 * K28 * (1 - J29) + 12 * J29)
G40 = -(J28 / (ConstantC22 * L16)) * (H32 * E6 / 720) * (8 * K31 * (11 - 24 * J29) - 12 * K30 * (21 - 71 * J29) + 15 * K29 * (15 - 98 * J29 + 15 * J31) + 180 * K28 * (5 * J29 - 3 * J31) + 360 * J31)
G41 = (J28 / (ConstantC22 * L16)) * (H34 * E6 / 40320) * (1385 + 3633 * J29 + 4095 * J31 + 1575 * J33)
G43 = G37 + G38 + G39 + G40 + G41
DatumLat = HMS(WorksheetFunction.Degrees(G43), 5)
Exit Function
Errr:
DatumLat = Err.Description
End Function
Function DatumLong(East, North, Zone, proj)
'Easting and Northing (East, North) are specified in metres
'Zone is the UTM zone, this value is one if the project grid does not have multiple zones
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)
PI = WorksheetFunction.PI

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

E5 = ConstantC20
E6 = East - E5
E7 = E6 / ConstantC22

L5 = ConstantC21
L6 = North - L5
L7 = L6 / ConstantC22
L9 = (L7 * PI) / (ConstantC17 * 180)
L10 = L9 * 2
L11 = L9 * 4
L12 = L9 * 6
L13 = L9 * 8

G16 = L9
G17 = ((3 * ConstantC13 / 2) - (27 * ConstantC15 / 32)) * Sin(L10)
G18 = ((21 * ConstantC14 / 16) - (55 * ConstantC16 / 32)) * Sin(L11)
G19 = (151 * ConstantC15) * Sin(L12) / 96
G20 = 1097 * ConstantC16 * Sin(L13) / 512
G21 = G16 + G17 + G18 + G19 + G20
G23 = Sin(G21)
G24 = 1 / Cos(G21)
J28 = Tan(G21)
L16 = a * (1 - ConstantC8) / (1 - ConstantC8 * G23 * G23) ^ 1.5
L17 = a / (1 - ConstantC8 * G23 * G23) ^ 0.5
G25 = J28 / (ConstantC22 * L17)


G28 = ConstantC13
G29 = G28 * G28
G30 = G28 * G29
G31 = G29 * G29

H28 = E7 / L17
H29 = H28 * H28
H30 = H28 * H29
H31 = H29 * H29
H32 = H31 * H28
H33 = H30 * H30
H34 = H28 * H33

I28 = (E7 * E7) / (L16 * L17)
I29 = I28 * I28
I30 = I28 * I29

J29 = J28 * J28
J30 = J28 * J29
J31 = J29 * J29
J32 = J31 * J28
J33 = J30 * J30

K28 = L17 / L16
K29 = K28 * K28
K30 = K29 * K28
K31 = K29 * K29

'Longitude
P37 = (Zone * ConstantC23) + ConstantC24 - ConstantC23
Q37 = (P37 / 180) * PI
Q38 = G24 * H28
Q39 = -G24 * (H30 / 6) * (K28 + 2 * J29)
Q40 = G24 * (H32 / 120) * (-4 * K30 * (1 - 6 * J29) + K29 * (9 - 68 * J29) + 72 * K28 * J29 + 24 * J31)
Q41 = -G24 * (H34 / 5040) * (61 + 662 * J29 + 1320 * J31 + 720 * J33)
Q43 = Q37 + Q38 + Q39 + Q40 + Q41
DatumLong = HMS(WorksheetFunction.Degrees(Q43), 5)
Exit Function
Errr:
DatumLong = Err.Description
End Function
Function DatumGridConv(East, North, Zone, proj)
'Easting and Northing (East, North) are specified in metres
'Zone is the UTM zone, this value is one if the project grid does not have multiple zones
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)
PI = WorksheetFunction.PI

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

E5 = ConstantC20
E6 = East - E5
E7 = E6 / ConstantC22

L5 = ConstantC21
L6 = North - L5
L7 = L6 / ConstantC22
L9 = (L7 * PI) / (ConstantC17 * 180)
L10 = L9 * 2
L11 = L9 * 4
L12 = L9 * 6
L13 = L9 * 8

G16 = L9
G17 = ((3 * ConstantC13 / 2) - (27 * ConstantC15 / 32)) * Sin(L10)
G18 = ((21 * ConstantC14 / 16) - (55 * ConstantC16 / 32)) * Sin(L11)
G19 = (151 * ConstantC15) * Sin(L12) / 96
G20 = 1097 * ConstantC16 * Sin(L13) / 512
G21 = G16 + G17 + G18 + G19 + G20
G23 = Sin(G21)
G24 = 1 / Cos(G21)
J28 = Tan(G21)
L16 = a * (1 - ConstantC8) / (1 - ConstantC8 * G23 * G23) ^ 1.5
L17 = a / (1 - ConstantC8 * G23 * G23) ^ 0.5
G25 = J28 / (ConstantC22 * L17)


G28 = ConstantC13
G29 = G28 * G28
G30 = G28 * G29
G31 = G29 * G29

H28 = E7 / L17
H29 = H28 * H28
H30 = H28 * H29
H31 = H29 * H29
H32 = H31 * H28
H33 = H30 * H30
H34 = H28 * H33

I28 = (E7 * E7) / (L16 * L17)
I29 = I28 * I28
I30 = I28 * I29

J29 = J28 * J28
J30 = J28 * J29
J31 = J29 * J29
J32 = J31 * J28
J33 = J30 * J30

K28 = L17 / L16
K29 = K28 * K28
K30 = K29 * K28
K31 = K29 * K29

'Grid Convergence
G45 = -J28 * H28
G46 = (J28 * H30 / 3) * (-2 * K29 + 3 * K28 + J29)
G47 = -(J28 * H32 / 15) * (K31 * (11 - 24 * J29) - 3 * K30 * (8 - 23 * J29) + 5 * K29 * (3 - 14 * J29) + 30 * K28 * J29 + 3 * J31)
G48 = (J28 * H34 / 315) * (17 + 77 * J29 + 105 * J31 + 45 * J33)
G50 = G45 + G46 + G47 + G48
DatumGridConv = HMS(WorksheetFunction.Degrees(G50), 2)
Exit Function
Errr:
DatumGridConv = Err.Description
End Function
Function DatumPntScle(East, North, Zone, proj)
'Easting and Northing (East, North) are specified in metres
'Zone is the UTM zone, this value is one if the project grid does not have multiple zones
'Projection names (proj) are sprcifed as a string, the names and parameters are commensurate with the project grids publish _
 on the landgate website (https://www0.landgate.wa.gov.au/business-and-government/specialist-services/geodetic/project-grids)
On Error GoTo Errr
Call EllipsoidalParameters(proj, a, b, OneOnF)
Call ProjectionParameters(proj, ConstantC20, ConstantC21, ConstantC22, ConstantC23, ConstantC24)
PI = WorksheetFunction.PI

ConstantC8 = Ecc2(a, OneOnF)
ConstantC13 = n(a, OneOnF)
ConstantC14 = ConstantC13 ^ 2
ConstantC15 = ConstantC13 ^ 3
ConstantC16 = ConstantC13 ^ 4
ConstantC17 = a * (1 - ConstantC13) * (1 - ConstantC14) * (1 + (9 * ConstantC14) / 4 + (225 * ConstantC16) / 64) * PI / 180
ConstantC25 = ConstantC24 - (1.5 * ConstantC23)
ConstantC26 = ConstantC25 + (ConstantC23 / 2)

E5 = ConstantC20
E6 = East - E5
E7 = E6 / ConstantC22

L5 = ConstantC21
L6 = North - L5
L7 = L6 / ConstantC22
L9 = (L7 * PI) / (ConstantC17 * 180)
L10 = L9 * 2
L11 = L9 * 4
L12 = L9 * 6
L13 = L9 * 8

G16 = L9
G17 = ((3 * ConstantC13 / 2) - (27 * ConstantC15 / 32)) * Sin(L10)
G18 = ((21 * ConstantC14 / 16) - (55 * ConstantC16 / 32)) * Sin(L11)
G19 = (151 * ConstantC15) * Sin(L12) / 96
G20 = 1097 * ConstantC16 * Sin(L13) / 512
G21 = G16 + G17 + G18 + G19 + G20
G23 = Sin(G21)
G24 = 1 / Cos(G21)
J28 = Tan(G21)
L16 = a * (1 - ConstantC8) / (1 - ConstantC8 * G23 * G23) ^ 1.5
L17 = a / (1 - ConstantC8 * G23 * G23) ^ 0.5
G25 = J28 / (ConstantC22 * L17)


G28 = ConstantC13
G29 = G28 * G28
G30 = G28 * G29
G31 = G29 * G29

H28 = E7 / L17
H29 = H28 * H28
H30 = H28 * H29
H31 = H29 * H29
H32 = H31 * H28
H33 = H30 * H30
H34 = H28 * H33

I28 = (E7 * E7) / (L16 * L17)
I29 = I28 * I28
I30 = I28 * I29

J29 = J28 * J28
J30 = J28 * J29
J31 = J29 * J29
J32 = J31 * J28
J33 = J30 * J30

K28 = L17 / L16
K29 = K28 * K28
K30 = K29 * K28
K31 = K29 * K29

'Point Scale
Q45 = 1 + I28 / 2
Q46 = (I29 / 24) * (4 * K28 * (1 - 6 * J29) - 3 * (1 - 16 * J29) - 24 * J29 / K28)
Q47 = I30 / 720
Q48 = Q45 + Q46 + Q47
DatumPntScle = Q48 * ConstantC22
Exit Function
Errr:
DatumPntScle = Err.Description
End Function

Function Ecc2(C3, C4)
C5 = 1 / C4
C7 = 1 / C4
C6 = C3 * (1 - C7)
C8 = (2 * C7) - (C7 * C7)
Ecc2 = C8
End Function

Function n(C3, C4)

C7 = 1 / C4
C6 = C3 * (1 - C7)
n = (C3 - C6) / (C3 + C6)

End Function
Function latSevenPara(FrmDatum, ToDatum, Lat, Lng, Ht)
'FrmDatum and ToDatum is specified as "AGD94" "GDA94" or "GDA2020"
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Ht is specified as in metres as an ellipsoidal Ht
'THIS FUNCTION REQUIRES THE "SevenParaParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "EllipsoidalParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_H_to_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "HMS" FUNCTION
'THIS FUNCTION REQUIRES THE "DD" FUNCTION

If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
Call SevenParaParameters(FrmDatum, ToDatum, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
Call EllipsoidalParameters(FrmDatum, a, b, OneOnF)

X = Lat_Long_H_to_X(Lat, Lng, Ht, a, b)
Y = Lat_Long_H_to_Y(Lat, Lng, Ht, a, b)
Z = Lat_H_to_Z(Lat, Ht, a, b)

tX = Helmert_X(FrmDatum, ToDatum, X, Y, Z)
tY = Helmert_Y(FrmDatum, ToDatum, X, Y, Z)
tZ = Helmert_Z(FrmDatum, ToDatum, X, Y, Z)

Call EllipsoidalParameters(ToDatum, a, b, OneOnF)

latSevenPara = HMS(XYZ_to_Lat(tX, tY, tZ, a, b), 5)
End Function
Function lngSevenPara(FrmDatum, ToDatum, Lat, Lng, Ht)
'FrmDatum and ToDatum is specified as "AGD94" "GDA94" or "GDA2020"
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Ht is specified as in metres as an ellipsoidal Ht
'THIS FUNCTION REQUIRES THE "SevenParaParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "EllipsoidalParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_H_to_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "HMS" FUNCTION
'THIS FUNCTION REQUIRES THE "DD" FUNCTION
If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
Call SevenParaParameters(FrmDatum, ToDatum, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
Call EllipsoidalParameters(FrmDatum, a, b, OneOnF)

X = Lat_Long_H_to_X(Lat, Lng, Ht, a, b)
Y = Lat_Long_H_to_Y(Lat, Lng, Ht, a, b)
Z = Lat_H_to_Z(Lat, Ht, a, b)

tX = Helmert_X(FrmDatum, ToDatum, X, Y, Z)
tY = Helmert_Y(FrmDatum, ToDatum, X, Y, Z)
tZ = Helmert_Z(FrmDatum, ToDatum, X, Y, Z)

Call EllipsoidalParameters(ToDatum, a, b, OneOnF)
lngSevenPara = HMS(XYZ_to_Long(tX, tY), 5)
End Function
Function HtSevenPara(FrmDatum, ToDatum, Lat, Lng, Ht)
'FrmDatum and ToDatum is specified as "AGD94" "GDA94" or "GDA2020"
'Latitude and Longitude (Lat, Lng) are specified as either decimal degrees of Hours minutes seconds
'Ht is specified as in metres as an ellipsoidal Ht
'THIS FUNCTION REQUIRES THE "SevenParaParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "EllipsoidalParameters" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_Long_H_to_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Lat_H_to_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_X" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Y" FUNCTION
'THIS FUNCTION REQUIRES THE "Helmert_Z" FUNCTION
'THIS FUNCTION REQUIRES THE "HMS" FUNCTION
'THIS FUNCTION REQUIRES THE "DD" FUNCTION
If InStr(1, Lat, " ") > 0 Then Lat = DD(Lat)
If InStr(1, Lng, " ") > 0 Then Lng = DD(Lng)
Call SevenParaParameters(FrmDatum, ToDatum, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
Call EllipsoidalParameters(FrmDatum, a, b, OneOnF)

X = Lat_Long_H_to_X(Lat, Lng, Ht, a, b)
Y = Lat_Long_H_to_Y(Lat, Lng, Ht, a, b)
Z = Lat_H_to_Z(Lat, Ht, a, b)

tX = Helmert_X(FrmDatum, ToDatum, X, Y, Z)
tY = Helmert_Y(FrmDatum, ToDatum, X, Y, Z)
tZ = Helmert_Z(FrmDatum, ToDatum, X, Y, Z)

Call EllipsoidalParameters(ToDatum, a, b, OneOnF)
HtSevenPara = XYZ_to_H(tX, tY, tZ, a, b)
End Function
Function XYZ_to_Lat(X, Y, Z, a, b)
'Convert XYZ to Latitude (PHI) in Dec Degrees.
'Input: - _
 XYZ cartesian coords (X,Y,Z) and ellipsoid axis dimensions (a & b), all in meters.

'THIS FUNCTION REQUIRES THE "Iterate_XYZ_to_Lat" FUNCTION
'THIS FUNCTION IS CALLED BY THE "XYZ_to_H" FUNCTION

    RootXYSqr = Sqr((X ^ 2) + (Y ^ 2))
    e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
    PHI1 = Atn(Z / (RootXYSqr * (1 - e2)))
    
    PHI = Iterate_XYZ_to_Lat(a, e2, PHI1, Z, RootXYSqr)
    
    PI = WorksheetFunction.PI
    
    XYZ_to_Lat = PHI * (180 / PI)

End Function

Function Iterate_XYZ_to_Lat(a, e2, PHI1, Z, RootXYSqr)
'Iteratively computes Latitude (PHI).
'Input: - _
 ellipsoid semi major axis (a) in meters; _
 eta squared (e2); _
 estimated value for latitude (PHI1) in radians; _
 cartesian Z coordinate (Z) in meters; _
 RootXYSqr computed from X & Y in meters.

'THIS FUNCTION IS CALLED BY THE "XYZ_to_PHI" FUNCTION
'THIS FUNCTION IS ALSO USED ON IT'S OWN IN THE _
 "Projection and Transformation Calculations.xls" SPREADSHEET


    V = a / (Sqr(1 - (e2 * ((Sin(PHI1)) ^ 2))))
    PHI2 = Atn((Z + (e2 * V * (Sin(PHI1)))) / RootXYSqr)
    
    Do While Abs(PHI1 - PHI2) > 0.000000001
        PHI1 = PHI2
        V = a / (Sqr(1 - (e2 * ((Sin(PHI1)) ^ 2))))
        PHI2 = Atn((Z + (e2 * V * (Sin(PHI1)))) / RootXYSqr)
    Loop

    Iterate_XYZ_to_Lat = PHI2

End Function

Function XYZ_to_Long(X, Y)
'Convert XYZ to Longitude (LAM) in Dec Degrees.
'Input: - _
 X and Y cartesian coords in meters.

    PI = WorksheetFunction.PI
    
'tailor the output to fit the equatorial quadrant as determined by the signs of X and Y
    If X >= 0 Then 'longitude is in the W90 thru 0 to E90 hemisphere
        XYZ_to_Long = (Atn(Y / X)) * (180 / PI)
    End If
    
    If X < 0 And Y >= 0 Then 'longitude is in the E90 to E180 quadrant
        XYZ_to_Long = ((Atn(Y / X)) * (180 / PI)) + 180
    End If
    
    If X < 0 And Y < 0 Then 'longitude is in the E180 to W90 quadrant
        XYZ_to_Long = ((Atn(Y / X)) * (180 / PI)) - 180
    End If
    

End Function

Function XYZ_to_H(X, Y, Z, a, b)
'Convert XYZ to Ellipsoidal Height.
'Input: - _
 XYZ cartesian coords (X,Y,Z) and ellipsoid axis dimensions (a & b), all in meters.

'REQUIRES THE "XYZ_to_Lat" FUNCTION

'Compute PHI (Dec Degrees) first
    PHI = XYZ_to_Lat(X, Y, Z, a, b)

'Convert PHI radians
    PI = WorksheetFunction.PI
    RadPHI = PHI * (PI / 180)
    
'Compute H
    RootXYSqr = Sqr((X ^ 2) + (Y ^ 2))
    e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
    V = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
    H = (RootXYSqr / Cos(RadPHI)) - V
    
    XYZ_to_H = H
    
End Function

Function Lat_Long_H_to_X(PHI, LAM, H, a, b)
'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian X coordinate.
'Input: - _
 Latitude (PHI)& Longitude (LAM) both in decimal degrees; _
 Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.

'Convert angle measures to radians
    PI = WorksheetFunction.PI
    RadPHI = PHI * (PI / 180)
    RadLAM = LAM * (PI / 180)

'Compute eccentricity squared and nu
    e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
    V = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

'Compute X
    Lat_Long_H_to_X = (V + H) * (Cos(RadPHI)) * (Cos(RadLAM))

End Function

Function Lat_Long_H_to_Y(PHI, LAM, H, a, b)
'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian Y coordinate.
'Input: - _
 Latitude (PHI)& Longitude (LAM) both in decimal degrees; _
 Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.

'Convert angle measures to radians
    PI = WorksheetFunction.PI
    RadPHI = PHI * (PI / 180)
    RadLAM = LAM * (PI / 180)

'Compute eccentricity squared and nu
    e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
    V = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

'Compute Y
    Lat_Long_H_to_Y = (V + H) * (Cos(RadPHI)) * (Sin(RadLAM))

End Function

Function Lat_H_to_Z(PHI, H, a, b)
'Convert geodetic coord components latitude (PHI) and height (H) to cartesian Z coordinate.
'Input: - _
 Latitude (PHI) decimal degrees; _
 Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.

'Convert angle measures to radians
    PI = WorksheetFunction.PI
    RadPHI = PHI * (PI / 180)

'Compute eccentricity squared and nu
    e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
    V = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

'Compute X
    Lat_H_to_Z = ((V * (1 - e2)) + H) * (Sin(RadPHI))

End Function

'********************************************************************************************* _
 THE FUNCTIONS IN THIS MODULE ARE WRITTEN TO BE "STAND ALONE" WITH AS LITTLE DEPENDANCE _
 ON OTHER FUNCTIONS AS POSSIBLE.  THIS MAKES THEM EASIER TO COPY TO OTHER VB APPLICATIONS. _
 WHERE A FUNCTION DOES CALL ANOTHER FUNCTION THIS IS STATED IN THE COMMENTS AT THE START OF _
 THE FUNCTION. _
 *********************************************************************************************

Function Helmert_X(FromEllip, ToEllip, X, Y, Z)
'Computed Helmert transformed X coordinate.
'Input: - _
 cartesian XYZ coords (X,Y,Z), X translation (dX) all in meters ; _
 Y and Z rotations in seconds of arc (Y_Rot, Z_Rot) and scale in ppm (s).

Call SevenParaParameters(FromEllip, ToEllip, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
'Convert rotations to radians and ppm scale to a factor
    PI = WorksheetFunction.PI
    sfactor = s / 1000000 ' 0.0000001
    RadY_Rot = -1 * (Y_Rot / 3600) * (PI / 180)
    RadZ_Rot = -1 * (Z_Rot / 3600) * (PI / 180)
    
'Compute transformed X coord
    Helmert_X = X + (X * sfactor) - (Y * RadZ_Rot) + (Z * RadY_Rot) + dX
 
End Function

Function Helmert_Y(FromEllip, ToEllip, X, Y, Z)
'Computed Helmert transformed Y coordinate.
'Input: - _
 cartesian XYZ coords (X,Y,Z), Y translation (DY) all in meters ; _
 X and Z rotations in seconds of arc (X_Rot, Z_Rot) and scale in ppm (s).
 
Call SevenParaParameters(FromEllip, ToEllip, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
'Convert rotations to radians and ppm scale to a factor
    PI = WorksheetFunction.PI
    sfactor = s * 0.000001
    RadX_Rot = -1 * (X_Rot / 3600) * (PI / 180)
    RadZ_Rot = -1 * (Z_Rot / 3600) * (PI / 180)
    
'Compute transformed Y coord
    Helmert_Y = (X * RadZ_Rot) + Y + (Y * sfactor) - (Z * RadX_Rot) + dY
 
End Function

Function Helmert_Z(FromEllip, ToEllip, X, Y, Z)
'Computed Helmert transformed Z coordinate.
'Input: - _
 cartesian XYZ coords (X,Y,Z), Z translation (DZ) all in meters ; _
 X and Y rotations in seconds of arc (X_Rot, Y_Rot) and scale in ppm (s).
 
Call SevenParaParameters(FromEllip, ToEllip, dX, dY, dZ, X_Rot, Y_Rot, Z_Rot, s)
'Convert rotations to radians and ppm scale to a factor
    PI = WorksheetFunction.PI
    sfactor = s * 0.000001
    RadX_Rot = -1 * (X_Rot / 3600) * (PI / 180)
    RadY_Rot = -1 * (Y_Rot / 3600) * (PI / 180)
    
'Compute transformed Z coord
    Helmert_Z = (-1 * X * RadY_Rot) + (Y * RadX_Rot) + Z + (Z * sfactor) + dZ
 
End Function

