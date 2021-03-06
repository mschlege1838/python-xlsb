from enum import Enum

def validate_rw(value):
    if value < 0 or value > 1048575:
        raise ValueError(f'Row index must be between 0 and 1048575: {value}')

def validate_col(value):
    if value < 0 or value > 16383:
        raise ValueError(f'Column index must be between 0 and 16383: {value}')

class RelationshipType(Enum):
    WORKBOOK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    SHARED_STRINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
    CALCULATION_CHAIN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'
    STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
    THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
    
    @staticmethod
    def resolve(value):
        try:
            return RelationshipType(value)
        except ValueError:
            return OtherRelationshipType(value)

class OtherRelationshipType:
    def __init__(self, value):
        self.value = value
    
    @property
    def name(self):
        return '<OtherRelationshipType>'
    
    def __str__(self):
        return f'{self.name}: {self.value}'


class ContentType(Enum):
    WORKBOOK = 'application/vnd.ms-excel.sheet.binary.macroEnabled.main'
    RELS = 'application/vnd.openxmlformats-package.relationships+xml'
    WORKSHEET = 'application/vnd.ms-excel.worksheet'
    CALCULATION_CHAIN = 'application/vnd.ms-excel.calcChain'
    STYLES = 'application/vnd.ms-excel.styles'
    THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'
    
    @staticmethod
    def resolve(value):
        try:
            return ContentType(value)
        except ValueError:
            return OtherContentType(value)

class OtherContentType:
    def __init__(self, value):
        self.value = value
    
    @property
    def name(self):
        return '<OtherContentType>'
    
    def __str__(self):
        return f'{self.name}: {self.value}'





class BinaryRecordType(Enum):
    BrtAbsPath15 = 2071
    BrtACBegin = 37
    BrtACEnd = 38
    BrtActiveX = 644
    BrtAFilterDateGroupItem = 175
    BrtArrFmla = 426
    BrtBeginActiveXControls = 643
    BrtBeginAFilter = 161
    BrtBeginAutoSortScope = 459
    BrtBeginBook = 131
    BrtBeginBookViews = 135
    BrtBeginBorders = 613
    BrtBeginBundleShs = 143
    BrtBeginCalcFeatures = 5095
    BrtBeginCellIgnoreECs = 648
    BrtBeginCellIgnoreECs14 = 1169
    BrtBeginCellSmartTag = 590
    BrtBeginCellSmartTags = 592
    BrtBeginCellStyleXFs = 626
    BrtBeginCellWatches = 605
    BrtBeginCellXFs = 617
    BrtBeginCFRule = 463
    BrtBeginCFRule14 = 1048
    BrtBeginColBrk = 394
    BrtBeginColInfos = 390
    BrtBeginColorPalette = 473
    BrtBeginColorScale = 469
    BrtBeginColorScale14 = 1157
    BrtBeginComment = 635
    BrtBeginCommentAuthors = 630
    BrtBeginCommentList = 633
    BrtBeginComments = 628
    BrtBeginConditionalFormatting = 461
    BrtBeginConditionalFormatting14 = 1046
    BrtBeginConditionalFormattings = 1135
    BrtBeginCRErrs = 608
    BrtBeginCsView = 141
    BrtBeginCsViews = 139
    BrtBeginCustomFilters = 172
    BrtBeginCustomFilters14 = 1178
    BrtBeginCustomRichFilters = 5086
    BrtBeginDatabar = 467
    BrtBeginDatabar14 = 1051
    BrtBeginDataFeedPr15 = 2113
    BrtBeginDataModel = 2121
    BrtBeginDbTables15 = 2118
    BrtBeginDCon = 495
    BrtBeginDecoupledPivotCacheIDs = 2048
    BrtBeginDeletedName = 453
    BrtBeginDeletedNames = 451
    BrtBeginDim = 275
    BrtBeginDims = 273
    BrtBeginDRefs = 497
    BrtBeginDVals = 573
    BrtBeginDVals14 = 1054
    BrtBeginDxF14s = 1172
    BrtBeginDXFs = 505
    BrtBeginDXFs15 = 2103
    BrtBeginDynamicArrayPr = 4096
    BrtBeginECDbProps = 203
    BrtBeginECOlapProps = 205
    BrtBeginECParam = 267
    BrtBeginECParams = 265
    BrtBeginECTwFldInfo = 542
    BrtBeginECTwFldInfo15 = 2133
    BrtBeginECTWFldInfoLst = 540
    BrtBeginECTWFldInfoLst15 = 2131
    BrtBeginECTxtWiz = 538
    BrtBeginECTxtWiz15 = 2129
    BrtBeginECWebProps = 261
    BrtBeginEcWpTables = 263
    BrtBeginEsfmd = 339
    BrtBeginEsmdb = 337
    BrtBeginEsmdtinfo = 334
    BrtBeginEsmdx = 372
    BrtBeginEsstr = 380
    BrtBeginExtConn14 = 1068
    BrtBeginExtConn15 = 2109
    BrtBeginExtConnection = 201
    BrtBeginExtConnections = 429
    BrtBeginExternals = 353
    BrtBeginFills = 603
    BrtBeginFilterColumn = 163
    BrtBeginFilters = 165
    BrtBeginFmd = 52
    BrtBeginFmts = 615
    BrtBeginFnGroup = 664
    BrtBeginFonts = 611
    BrtBeginHeaderFooter = 479
    BrtBeginIconSet = 465
    BrtBeginIconSet14 = 1052
    BrtBeginPaletteColors = 565
    BrtBeginISXTHCols = 322
    BrtBeginISXTHRws = 320
    BrtBeginISXVDCols = 311
    BrtBeginISXVDRws = 309
    BrtBeginISXVIs = 388
    BrtBeginItemUniqueNames = 2106
    BrtBeginList = 343
    BrtBeginListCol = 347
    BrtBeginListCols = 345
    BrtBeginListParts = 660
    BrtBeginListXmlCPr = 349
    BrtBeginMap = 492
    BrtBeginMdx = 54
    BrtBeginMdxKPI = 378
    BrtBeginMdxMbrProp = 376
    BrtBeginMdxSet = 374
    BrtBeginMdxTuple = 56
    BrtBeginMergeCells = 177
    BrtBeginMetadata = 332
    BrtBeginMG = 490
    BrtBeginMGMaps = 488
    BrtBeginMgs = 486
    BrtBeginModelRelationships = 2126
    BrtBeginModelTables = 2123
    brtBeginModelTimeGrouping = 2139
    brtBeginModelTimeGroupings = 2137
    BrtBeginMRUColors = 569
    BrtBeginOledbPr15 = 2111
    BrtBeginOleObjects = 638
    BrtBeginPCD14 = 1066
    BrtBeginPCDCalcItem = 245
    BrtBeginPCDCalcItems = 243
    BrtBeginPCDCalcMem = 433
    BrtBeginPCDCalcMem14 = 1038
    BrtBeginPCDCalcMemExt = 1137
    BrtBeginPCDCalcMems = 431
    BrtBeginPCDCalcMemsExt = 1139
    BrtBeginPCDFAtbl = 189
    BrtBeginPCDFGDiscrete = 225
    BrtBeginPCDFGItems = 221
    BrtBeginPCDFGRange = 223
    BrtBeginPCDFGroup = 219
    BrtBeginPCDField = 183
    BrtBeginPCDFields = 181
    BrtBeginPCDHFieldsUsage = 199
    BrtBeginPCDHGLevel = 437
    BrtBeginPCDHGLevels = 435
    BrtBeginPCDHGLGMember = 445
    BrtBeginPCDHGLGMembers = 443
    BrtBeginPCDHGLGroup = 441
    BrtBeginPCDHGLGroups = 439
    BrtBeginPCDHierarchies = 195
    BrtBeginPCDHierarchy = 197
    BrtBeginPCDIRun = 191
    BrtBeginPCDKPI = 271
    BrtBeginPCDKPIs = 269
    BrtBeginPCDSConsol = 207
    BrtBeginPCDSCPage = 211
    BrtBeginPCDSCPages = 209
    BrtBeginPCDSCPItem = 213
    BrtBeginPCDSCSet = 217
    BrtBeginPCDSCSets = 215
    BrtBeginPCDSDTCEMember = 233
    BrtBeginPCDSDTCEMembers = 231
    BrtBeginPCDSDTCEMembersSortBy = 646
    BrtBeginPCDSDTCEntries = 229
    BrtBeginPCDSDTCQueries = 235
    BrtBeginPCDSDTCQuery = 237
    BrtBeginPCDSDTCSet = 241
    BrtBeginPCDSDTCSets = 239
    BrtBeginPCDSDTupleCache = 227
    BrtBeginPcdSFCIEntries = 657
    BrtBeginPCDSource = 185
    BrtBeginPCDSRange = 187
    BrtBeginPivotCacheDef = 179
    BrtBeginPivotCacheID = 386
    BrtBeginPivotCacheIDs = 384
    BrtBeginPivotCacheRecords = 193
    BrtBeginPivotTableRefs = 2051
    BrtBeginPivotTableUISettings = 2072
    BrtBeginPName = 255
    BrtBeginPNames = 253
    BrtBeginPNPair = 259
    BrtBeginPNPairs = 257
    BrtBeginPRFilter = 251
    BrtBeginPRFilter14 = 1165
    BrtBeginPRFilters = 249
    BrtBeginPRFilters14 = 1163
    BrtBeginPRFItem = 382
    BrtBeginPRFItem14 = 1167
    BrtBeginPRule = 247
    BrtBeginPRule14 = 1161
    BrtBeginQSI = 447
    BrtBeginQSIF = 457
    BrtBeginQSIFs = 455
    BrtBeginQSIR = 449
    BrtBeginRichFilterColumn = 5084
    BrtBeginRichFilters = 5081
    BrtBeginRichSortCondition = 5092
    BrtBeginRichValueBlock = 5002
    BrtBeginRRSort = 673
    BrtBeginRwBrk = 392
    BrtBeginScenMan = 500
    BrtBeginSct = 502
    BrtBeginSheet = 129
    BrtBeginSheetData = 145
    BrtBeginSingleCells = 341
    BrtBeginSlicer = 1083
    BrtBeginSlicerCache = 1075
    BrtBeginSlicerCacheDef = 1077
    BrtBeginSlicerCacheID = 1072
    BrtBeginSlicerCacheIDs = 1070
    BrtBeginSlicerCacheLevelData = 1090
    BrtBeginSlicerCacheLevelsData = 1088
    BrtBeginSlicerCacheNative = 1100
    BrtBeginSlicerCacheOlapImpl = 1086
    BrtBeginSlicerCacheSelections = 1097
    BrtBeginSlicerCacheSiRange = 1094
    BrtBeginSlicerCacheSiRanges = 1092
    BrtBeginSlicerCachesPivotCacheID = 1133
    BrtBeginSlicerCachesPivotCacheIDs = 1113
    BrtBeginSlicerEx = 1081
    BrtBeginSlicers = 1115
    BrtBeginSlicersEx = 1079
    BrtBeginSlicerStyle = 1128
    BrtBeginSlicerStyleElements = 1144
    BrtBeginSlicerStyles = 1142
    BrtBeginSmartTags = 594
    BrtBeginSmartTagTypes = 597
    BrtBeginSortCond = 532
    BrtBeginSortCond14 = 1152
    BrtBeginSortState = 530
    BrtBeginSparklineGroup = 1041
    BrtBeginSparklineGroups = 1058
    BrtBeginSparklines = 1056
    BrtBeginSst = 159
    BrtBeginStyles = 619
    BrtBeginStyleSheet = 278
    BrtBeginStyleSheetExt14 = 1131
    BrtBeginSupBook = 360
    BrtBeginSXChange = 1122
    BrtBeginSXChanges = 1124
    BrtBeginSXCondFmt = 558
    BrtBeginSXCondFmt14 = 1147
    BrtBeginSXCondFmts = 560
    BrtBeginSXCondFmts14 = 1149
    BrtBeginSXCrtFormat = 481
    BrtBeginSXCrtFormats = 483
    BrtBeginSXDI = 293
    BrtBeginSXDIs = 295
    BrtBeginSXEdit = 1118
    BrtBeginSXEdits = 1120
    BrtBeginSXFILTER = 601
    BrtBeginSXFilters = 599
    BrtBeginSXFormat = 303
    BrtBeginSXFormats = 305
    BrtBeginSXLI = 297
    BrtBeginSXLICols = 301
    BrtBeginSXLIRws = 299
    BrtBeginSXLocation = 314
    BrtBeginSXPI = 289
    BrtBeginSXPIs = 291
    BrtBeginSxRow = 2057
    BrtBeginSxrules = 641
    BrtBeginSxrules14 = 1159
    BrtBeginSxSelect = 307
    BrtBeginSXTDMP = 326
    BrtBeginSXTDMPS = 324
    BrtBeginSXTH = 318
    BrtBeginSXTHItem = 330
    BrtBeginSXTHItems = 328
    BrtBeginSXTHs = 316
    BrtBeginSXTupleSet = 1026
    BrtBeginSXTupleSetData = 1031
    BrtBeginSXTupleSetHeader = 1028
    BrtBeginSXTupleSetRow = 1033
    BrtBeginSxvcells = 2055
    BrtBeginSXVD = 285
    BrtBeginSXVDs = 287
    BrtBeginSXVI = 282
    BrtBeginSXView = 280
    BrtBeginSxView14 = 1062
    BrtBeginSXView16 = 1064
    BrtBeginSXVIs = 283
    BrtBeginTableSlicerCache = 2077
    BrtBeginTableStyle = 510
    BrtBeginTableStyles = 508
    BrtBeginTimelineCacheID = 2085
    BrtBeginTimelineCacheIDs = 2083
    BrtBeginTimelineCachePivotCacheIDs = 2080
    BrtBeginTimelineEx = 2089
    BrtBeginTimelinesEx = 2087
    BrtBeginTimelineStyle = 2093
    BrtBeginTimelineStyleElements = 2100
    BrtBeginTimelineStyles = 2098
    BrtBeginTimelineStylesheetExt15 = 2096
    BrtBeginUserCsView = 655
    BrtBeginUserCsViews = 653
    BrtBeginUsers = 401
    BrtBeginUserShView = 423
    BrtBeginUserShViews = 422
    BrtBeginVolDeps = 514
    BrtBeginVolMain = 518
    BrtBeginVolTopic = 520
    BrtBeginVolType = 516
    BrtBeginWebExtensions = 2068
    BrtBeginWebPubItem = 556
    BrtBeginWebPubItems = 554
    BrtBeginWsSortMap = 671
    BrtBeginWsView = 137
    BrtBeginWsViews = 133
    BrtBigName = 625
    BrtBkHim = 562
    BrtBookProtection = 534
    BrtBookProtectionIso = 677
    BrtBookView = 158
    BrtBorder = 46
    BrtBrk = 396
    BrtBundleSh = 156
    BrtCalcFeature = 5097
    BrtCalcProp = 157
    BrtCellBlank = 1
    BrtCellBool = 4
    BrtCellError = 3
    BrtCellIgnoreEC = 649
    BrtCellIgnoreEC14 = 1105
    BrtCellIsst = 7
    BrtCellMeta = 49
    BrtCellReal = 5
    BrtCellRk = 2
    BrtCellRString = 62
    BrtCellSmartTagProperty = 589
    BrtCellSt = 6
    BrtCellWatch = 607
    BrtCFIcon = 1112
    BrtCFRuleExt = 1146
    BrtCFVO = 471
    BrtCFVO14 = 1050
    BrtColInfo = 60
    BrtColor = 564
    BrtColor14 = 1055
    BrtColorFilter = 168
    BrtCommentAuthor = 632
    BrtCommentText = 637
    BrtCrashRecErr = 610
    BrtCsPageSetup = 652
    BrtCsProp = 651
    BrtCsProtection = 669
    BrtCsProtectionIso = 679
    BrtCUsr = 399
    BrtCustomFilter = 174
    BrtCustomFilter14 = 1180
    BRTCustomRichFilter = 5088
    BrtDbCommand15 = 2117
    BrtDbTable15 = 2120
    BrtDecoupledPivotCacheID = 2048
    BrtDrawing = 550
    BrtDRef = 499
    BrtDVal = 64
    BrtDVal14 = 1053
    BrtDValList = 681
    BrtDXF = 507
    BrtDXF14 = 1171
    BrtDXF15 = 2102
    BrtDynamicFilter = 171
    BrtDynamicRichFilter = 5090
    BrtEndActiveXControls = 645
    BrtEndAFilter = 162
    BrtEndAutoSortScope = 460
    BrtEndBook = 132
    BrtEndBookViews = 136
    BrtEndBorders = 614
    BrtEndBundleShs = 144
    BrtEndCalcFeatures = 5096
    BrtEndCellIgnoreECs = 650
    BrtEndCellIgnoreECs14 = 1170
    BrtEndCellSmartTag = 591
    BrtEndCellSmartTags = 593
    BrtEndCellStyleXFs = 627
    BrtEndCellWatches = 606
    BrtEndCellXFs = 618
    BrtEndCFRule = 464
    BrtEndCFRule14 = 1049
    BrtEndColBrk = 395
    BrtEndColInfos = 391
    BrtEndColorPalette = 474
    BrtEndColorScale = 470
    BrtEndColorScale14 = 1158
    BrtEndComment = 636
    BrtEndCommentAuthors = 631
    BrtEndCommentList = 634
    BrtEndComments = 629
    BrtEndConditionalFormatting = 462
    BrtEndConditionalFormatting14 = 1047
    BrtEndConditionalFormattings = 1136
    BrtEndCRErrs = 609
    BrtEndCsView = 142
    BrtEndCsViews = 140
    BrtEndCustomFilters = 173
    BrtEndCustomRichFilters = 5087
    BrtEndDatabar = 468
    BrtEndDatabar14 = 1156
    BrtEndDataFeedPr15 = 2114
    BrtEndDataModel = 2122
    BrtEndDbTables15 = 2119
    BrtEndDCon = 496
    BrtEndDecoupledPivotCacheIDs = 2049
    BrtEndDeletedName = 454
    BrtEndDeletedNames = 452
    BrtEndDim = 276
    BrtEndDims = 274
    BrtEndDRefs = 498
    BrtEndDVals = 574
    BrtEndDVals14 = 1154
    BrtEndDxf14s = 1173
    BrtEndDXFs = 506
    BrtEndDXFs15 = 2104
    BrtEndDynamicArrayPr = 4097
    BrtEndECDbProps = 204
    BrtEndECOlapProps = 206
    BrtEndECParam = 268
    BrtEndECParams = 266
    BrtEndECTWFldInfoLst = 541
    BrtEndECTWFldInfoLst15 = 2132
    BrtEndECTxtWiz = 539
    BrtEndECTxtWiz15 = 2130
    BrtEndECWebProps = 262
    BrtEndECWPTables = 264
    BrtEndEsfmd = 340
    BrtEndEsmdb = 338
    BrtEndEsmdtinfo = 336
    BrtEndEsmdx = 373
    BrtEndEsstr = 381
    BrtEndExtConn14 = 1069
    BrtEndExtConn15 = 2110
    BrtEndExtConnection = 202
    BrtEndExtConnections = 430
    BrtEndExternals = 354
    BrtEndFills = 604
    BrtEndFilterColumn = 164
    BrtEndFilters = 166
    BrtEndFmd = 53
    BrtEndFmts = 616
    BrtEndFnGroup = 666
    BrtEndFonts = 612
    BrtEndHeaderFooter = 480
    BrtEndIconSet = 466
    BrtEndIconSet14 = 1155
    BrtEndPaletteColors = 566
    BrtEndISXTHCols = 323
    BrtEndISXTHRws = 321
    BrtEndISXVDCols = 312
    BrtEndISXVDRws = 310
    BrtEndISXVIs = 389
    BrtEndItemUniqueNames = 2107
    BrtEndList = 344
    BrtEndListCol = 348
    BrtEndListCols = 346
    BrtEndListParts = 662
    BrtEndListXmlCPr = 350
    BrtEndMap = 493
    BrtEndMdx = 55
    BrtEndMdxKPI = 379
    BrtEndMdxMbrProp = 377
    BrtEndMdxSet = 375
    BrtEndMdxTuple = 57
    BrtEndMergeCells = 178
    BrtEndMetadata = 333
    BrtEndMG = 491
    BrtEndMGMaps = 489
    BrtEndMGs = 487
    BrtEndModelRelationships = 2127
    BrtEndModelTables = 2124
    brtEndModelTimeGrouping = 2140
    brtEndModelTimeGroupings = 2138
    BrtEndMRUColors = 570
    BrtEndOledbPr15 = 2112
    BrtEndOleObjects = 640
    BrtEndPCD14 = 1067
    BrtEndPCDCalcItem = 246
    BrtEndPCDCalcItems = 244
    BrtEndPCDCalcMem = 434
    BrtEndPCDCalcMem14 = 1039
    BrtEndPCDCalcMemExt = 1138
    BrtEndPCDCalcMems = 432
    BrtEndPCDCalcMemsExt = 1140
    BrtEndPCDFAtbl = 190
    BrtEndPCDFGDiscrete = 226
    BrtEndPCDFGItems = 222
    BrtEndPCDFGRange = 224
    BrtEndPCDFGroup = 220
    BrtEndPCDField = 184
    BrtEndPCDFields = 182
    BrtEndPCDHFieldsUsage = 200
    BrtEndPCDHGLevel = 438
    BrtEndPCDHGLevels = 436
    BrtEndPCDHGLGMember = 446
    BrtEndPCDHGLGMembers = 444
    BrtEndPCDHGLGroup = 442
    BrtEndPCDHGLGroups = 440
    BrtEndPCDHierarchies = 196
    BrtEndPCDHierarchy = 198
    BrtEndPCDIRun = 192
    BrtEndPCDKPI = 272
    BrtEndPCDKPIs = 270
    BrtEndPCDSConsol = 208
    BrtEndPCDSCPage = 212
    BrtEndPCDSCPages = 210
    BrtEndPCDSCPItem = 214
    BrtEndPCDSCSet = 218
    BrtEndPCDSCSets = 216
    BrtEndPCDSDTCEMember = 234
    BrtEndPCDSDTCEMembers = 232
    BrtEndPCDSDTCEntries = 230
    BrtEndPCDSDTCQueries = 236
    BrtEndPCDSDTCQuery = 238
    BrtEndPCDSDTCSet = 242
    BrtEndPCDSDTCSets = 240
    BrtEndPCDSDTupleCache = 228
    BrtEndPCDSFCIEntries = 658
    BrtEndPCDSource = 186
    BrtEndPCDSRange = 188
    BrtEndPivotCacheDef = 180
    BrtEndPivotCacheID = 387
    BrtEndPivotCacheIDs = 385
    BrtEndPivotCacheRecords = 194
    BrtEndPivotTableRefs = 2052
    BrtEndPivotTableUISettings = 2073
    BrtEndPName = 256
    BrtEndPNames = 254
    BrtEndPNPair = 260
    BrtEndPNPairs = 258
    BrtEndPRFilter = 252
    BrtEndPRFilter14 = 1166
    BrtEndPRFilters = 250
    BrtEndPRFilters14 = 1164
    BrtEndPRFItem = 383
    BrtEndPRFItem14 = 1168
    BrtEndPRule = 248
    BrtEndPRule14 = 1162
    BrtEndQSI = 448
    BrtEndQSIF = 458
    BrtEndQSIFs = 456
    BrtEndQSIR = 450
    BrtEndRichFilterColumn = 5085
    BrtEndRichFilters = 5082
    BrtEndRichSortCondition = 5093
    BrtEndRichValueBlock = 5003
    BrtEndRRSort = 674
    BrtEndRwBrk = 393
    BrtEndScenMan = 501
    BrtEndSct = 503
    BrtEndSheet = 130
    BrtEndSheetData = 146
    BrtEndSingleCells = 342
    BrtEndSlicer = 1084
    BrtEndSlicerCache = 1076
    BrtEndSlicerCacheDef = 1078
    BrtEndSlicerCacheID = 1073
    BrtEndSlicerCacheIDs = 1071
    BrtEndSlicerCacheLevelData = 1091
    BrtEndSlicerCacheLevelsData = 1089
    BrtEndSlicerCacheNative = 1101
    BrtEndSlicerCacheOlapImpl = 1087
    BrtEndSlicerCacheSelections = 1099
    BrtEndSlicerCacheSiRange = 1095
    BrtEndSlicerCacheSiRanges = 1093
    BrtEndSlicerCachesPivotCacheID = 1134
    BrtEndSlicerCachesPivotCacheIDs = 1114
    BrtEndSlicerEx = 1082
    BrtEndSlicers = 1116
    BrtEndSlicersEx = 1080
    BrtEndSlicerStyle = 1129
    BrtEndSlicerStyleElements = 1145
    BrtEndSlicerStyles = 1143
    BrtEndSmartTags = 595
    BrtEndSmartTagTypes = 598
    BrtEndSortCond = 533
    BrtEndSortCond14 = 1153
    BrtEndSortState = 531
    BrtEndSparklineGroup = 1042
    BrtEndSparklineGroups = 1059
    BrtEndSparklines = 1057
    BrtEndSst = 160
    BrtEndStyles = 620
    BrtEndStyleSheet = 279
    BrtEndStyleSheetExt14 = 1132
    BrtEndSupBook = 588
    BrtEndSXChange = 1123
    BrtEndSXChanges = 1125
    BrtEndSXCondFmt = 559
    BrtEndSXCondFmt14 = 1148
    BrtEndSXCondFmts = 561
    BrtEndSXCondFmts14 = 1150
    BrtEndSXCrtFormat = 482
    BrtEndSXCrtFormats = 484
    BrtEndSXDI = 294
    BrtEndSXDIs = 296
    BrtEndSXEdit = 1119
    BrtEndSXEdits = 1121
    BrtEndSXFilter = 602
    BrtEndSXFilters = 600
    BrtEndSXFormat = 304
    BrtEndSxFormats = 306
    BrtEndSXLI = 298
    BrtEndSXLICols = 302
    BrtEndSXLIRws = 300
    BrtEndSXLocation = 313
    BrtEndSXPI = 290
    BrtEndSXPIs = 292
    BrtEndSxRow = 2058
    BrtEndSxRules = 642
    BrtEndSxrules14 = 1160
    BrtEndSxSelect = 308
    BrtEndSXTDMP = 327
    BrtEndSXTDMPs = 325
    BrtEndSXTH = 319
    BrtEndSXTHItem = 331
    BrtEndSXTHItems = 329
    BrtEndSXTHs = 317
    BrtEndSXTupleSet = 1027
    BrtEndSXTupleSetData = 1032
    BrtEndSXTupleSetHeader = 1029
    BrtEndSXTupleSetRow = 1034
    BrtEndSxvcells = 2056
    BrtEndSXVD = 286
    BrtEndSXVDs = 288
    BrtEndSXVI = 281
    BrtEndSXView = 315
    BrtEndSxView14 = 1063
    BrtEndSXView16 = 1065
    BrtEndSXVIs = 284
    BrtEndTableSlicerCache = 2078
    BrtEndTableStyle = 511
    BrtEndTableStyles = 509
    BrtEndTimelineCacheID = 2086
    BrtEndTimelineCacheIDs = 2084
    BrtEndTimelineCachePivotCacheIDs = 2081
    BrtEndTimelineEx = 2090
    BrtEndTimelinesEx = 2088
    BrtEndTimelineStyle = 2094
    BrtEndTimelineStyleElements = 2101
    BrtEndTimelineStyles = 2099
    BrtEndTimelineStylesheetExt15 = 2097
    BrtEndUserCsView = 656
    BrtEndUserCsViews = 654
    BrtEndUserShView = 424
    BrtEndUserShViews = 425
    BrtEndVolDeps = 515
    BrtEndVolMain = 519
    BrtEndVolTopic = 521
    BrtEndVolType = 517
    BrtEndWebExtensions = 2069
    BrtEndWebPubItem = 557
    BrtEndWebPubItems = 555
    BrtEndWsSortMap = 672
    BrtEndWsView = 138
    BrtEndWsViews = 134
    BrtEOF = 403
    BrtExternalLinksPr = 5099
    BrtExternCellBlank = 367
    BrtExternCellBool = 369
    BrtExternCellError = 370
    BrtExternCellReal = 368
    BrtExternCellString = 371
    BrtExternRowHdr = 366
    BrtExternSheet = 362
    BrtExternTableEnd = 364
    BrtExternTableStart = 363
    BrtExternValueMeta = 472
    BrtFieldListActiveItem = 2134
    BrtFileRecover = 155
    BrtFileSharing = 548
    BrtFileSharingIso = 676
    BrtFileVersion = 128
    BrtFill = 45
    BrtFilter = 167
    BrtFilter14 = 1177
    BrtFmlaBool = 10
    BrtFmlaError = 11
    BrtFmlaNum = 9
    BrtFmlaString = 8
    BrtFmt = 44
    BrtFnGroup = 665
    BrtFont = 43
    BrtFRTBegin = 35
    BrtFRTEnd = 36
    BrtHLink = 494
    BrtIconFilter = 169
    BrtIconFilter14 = 1181
    BrtIndexBlock = 42
    BrtPaletteColor = 475
    BrtIndexPartEnd = 277
    BrtIndexRowBlock = 40
    BrtInfo = 398
    BrtItemUniqueName = 2108
    BrtKnownFonts = 1025
    BrtLegacyDrawing = 551
    BrtLegacyDrawingHF = 552
    BrtList14 = 1111
    BrtListCCFmla = 351
    BrtListPart = 661
    BrtListTrFmla = 352
    BrtMargins = 476
    BrtMdb = 51
    BrtMdtinfo = 335
    BrtMdxMbrIstr = 58
    BrtMergeCell = 176
    BrtModelRelationship = 2128
    BrtModelTable = 2125
    brtModelTimeGroupingCalcCol = 2141
    BrtMRUColor = 572
    BrtName = 39
    BrtNameExt = 1036
    BrtOleObject = 639
    BrtOleSize = 549
    BrtPageSetup = 478
    BrtPane = 151
    BrtPCDCalcMem15 = 2060
    BrtPCDField14 = 1141
    BrtPCDH14 = 1037
    BrtPCDH15 = 2092
    BrtPCDIABoolean = 29
    BrtPCDIADatetime = 32
    BrtPCDIAError = 30
    BrtPCDIAMissing = 27
    BrtPCDIANumber = 28
    BrtPCDIAString = 31
    BrtPCDIBoolean = 22
    BrtPCDIDatetime = 25
    BrtPCDIError = 23
    BrtPCDIIndex = 26
    BrtPCDIMissing = 20
    BrtPCDINumber = 21
    BrtPCDIString = 24
    BrtPCDSFCIEntry = 659
    BrtPCRRecord = 33
    BrtPCRRecordDt = 34
    BrtPhoneticInfo = 537
    BrtPivotCacheConnectionName = 1182
    BrtPivotCacheIdVersion = 2135
    BrtPivotTableRef = 2053
    BrtPlaceholderName = 361
    BrtPrintOptions = 477
    BrtQsi15 = 2067
    BrtRangePr15 = 2116
    BrtRangeProtection = 536
    BrtRangeProtection14 = 1103
    BrtRangeProtectionIso = 680
    BrtRangeProtectionIso14 = 1104
    brtRevisionPtr = 3073
    BrtRichFilter = 5083
    BrtRichFilterDateGroupItem = 5094
    BrtRowHdr = 0
    BrtRRAutoFmt = 421
    BrtRRChgCell = 409
    BrtRRConflict = 417
    BrtRRDefName = 415
    BrtRREndChgCell = 410
    BrtRREndFormat = 420
    BrtRREndInsDel = 406
    BrtRREndMove = 408
    BrtRRFormat = 419
    BrtRRHeader = 411
    BrtRRInsDel = 405
    BrtRRInsertSh = 414
    BrtRRMove = 407
    BrtRRNote = 416
    BrtRRRenSheet = 413
    BrtRRSortItem = 675
    BrtRRTQSIF = 418
    BrtRRUserView = 412
    BrtRwDescent = 1024
    BrtSel = 152
    BrtSheetCalcProp = 663
    BrtSheetProtection = 535
    BrtSheetProtectionIso = 678
    BrtShrFmla = 427
    BrtSlc = 504
    BrtSlicerCacheBookPivotTables = 2054
    BrtSlicerCacheHideItemsWithNoData = 2105
    BrtSlicerCacheNativeItem = 1102
    BrtSlicerCacheOlapItem = 1096
    BrtSlicerCachePivotTables = 1085
    BrtSlicerCacheSelection = 1098
    BrtSlicerStyleElement = 1130
    BrtSmartTagType = 596
    BrtSparkline = 1043
    BrtSSTItem = 19
    BrtStr = 59
    BrtStyle = 48
    BrtSupAddin = 667
    BrtSupBookSrc = 355
    BrtSupNameBits = 586
    BrtSupNameBool = 584
    BrtSupNameEnd = 587
    BrtSupNameErr = 581
    BrtSupNameFmla = 585
    BrtSupNameNil = 583
    BrtSupNameNum = 580
    BrtSupNameSt = 582
    BrtSupNameStart = 577
    BrtSupNameValueEnd = 579
    BrtSupNameValueStart = 578
    BrtSupSame = 358
    BrtSupSelf = 357
    BrtSupTabs = 359
    BrtSXDI14 = 1044
    BrtSXDI15 = 2136
    BrtSxFilter15 = 2079
    BrtSXTDMPOrder = 668
    BrtSXTH14 = 1040
    BrtSXTupleItems = 1126
    BrtSXTupleSetHeaderItem = 1030
    BrtSXTupleSetRowItem = 1035
    BrtSxvcellBool = 67
    BrtSxvcellDate = 69
    BrtSxvcellErr = 68
    BrtSxvcellNil = 70
    BrtSxvcellNum = 65
    BrtSxvcellStr = 66
    BrtSXVD14 = 1061
    BrtTable = 428
    BrtTableSlicerCacheID = 2076
    BrtTableSlicerCacheIDs = 2075
    BrtTableStyleClient = 513
    BrtTableStyleElement = 512
    BrtTextPr15 = 2115
    BrtTimelineCachePivotCacheID = 2082
    BrtTimelineStyleElement = 2095
    BrtTop10Filter = 170
    BrtTop10RichFilter = 5089
    BrtUCR = 404
    BrtUserBookView = 397
    BrtUsr = 400
    BrtValueMeta = 50
    BrtVolBool = 527
    BrtVolErr = 525
    BrtVolNum = 524
    BrtVolRef = 523
    BrtVolStr = 526
    BrtVolSubtopic = 522
    BrtWbFactoid = 154
    BrtWbProp = 153
    BrtWbProp14 = 1117
    BrtWebExtension = 2070
    BrtWebOpt = 553
    BrtWorkBookPr15 = 2091
    BrtWsDim = 148
    BrtWsFmtInfo = 485
    BrtWsFmtInfoEx14 = 1045
    BrtWsProp = 147
    BrtXF = 47
    BrtUid = 3072


class ExtendedRecordType:
    def __init__(self, tname, tnumber):
        self.tname = tname
        self.tnumber = tnumber
    
    @property
    def name(self):
        return f'<{self.tname}> {self.tnumber}'
    
    @property
    def value(self):
        return self.tnumber
    
    @property
    def rtype(self):
        return type(self)
    
    def __str__(self):
        return self.name

class FutureRecordType(ExtendedRecordType):
    def __init__(self, tnumber):
        super().__init__('FutureRecord', tnumber)

class AlternateContentRecordType(ExtendedRecordType):
    def __init__(self, tnumber):
        super().__init__('AlternateContent', tnumber)



class HorizontalAlignmentType(Enum):
    GENERAL = 0
    LEFT = 1
    CENTER = 2
    RIGHT = 3
    FILL = 4
    JUSTIFY = 5
    CENTER_ACROSS_SELECTION = 6
    DISTRIBUTED = 7

class VerticalAlignmentType(Enum):
    TOP = 0
    CENTER = 1
    BOTTOM = 2
    JUSTIFY = 3
    DISTRIBUTED = 4

class ReadingOrderType(Enum):
    CONTEXT_DEPENDENT = 0
    LEFT_TO_RIGHT = 1
    RIGHT_TO_LEFT = 2

class XFProperty(Enum):
    FMT = 0x0001
    FONT = 0x0002
    ALIGNMENT = 0x0004
    BORDER = 0x0008
    FILL = 0x0010
    PROTECTION = 0x0020

class ColorType(Enum):
    AUTO = 0x00
    PALETTE = 0x01
    RGBA = 0x02
    THEME = 0x03

class PaletteColor(Enum):
    icvBlack = 0x00         # 0x000000FF
    icvWhite = 0x01         # 0xFFFFFFFF
    icvRed = 0x02           # 0xFF0000FF
    icvGreen = 0x03         # 0x00FF00FF
    icvBlue = 0x04          # 0x0000FFFF
    icvYellow = 0x05        # 0xFFFF00FF
    icvMagenta = 0x06       # 0xFF00FFFF
    icvCyan = 0x07          # 0x00FFFFFF
    icvPlt1 = 0x08          # 0x000000FF
    icvPlt2 = 0x09          # 0xFFFFFFFF
    icvPlt3 = 0x0A          # 0xFF0000FF
    icvPlt4 = 0x0B          # 0x00FF00FF
    icvPlt5 = 0x0C          # 0x0000FFFF
    icvPlt6 = 0x0D          # 0xFFFF00FF
    icvPlt7 = 0x0E          # 0xFF00FFFF
    icvPlt8 = 0x0F          # 0x00FFFFFF
    icvPlt9 = 0x10          # 0x800000FF
    icvPlt10 = 0x11         # 0x008000FF
    icvPlt11 = 0x12         # 0x000080FF
    icvPlt12 = 0x13         # 0x808000FF
    icvPlt13 = 0x14         # 0x800080FF
    icvPlt14 = 0x15         # 0x008080FF
    icvPlt15 = 0x16         # 0xC0C0C0FF
    icvPlt16 = 0x17         # 0x808080FF
    icvPlt17 = 0x18         # 0x9999FFFF
    icvPlt18 = 0x19         # 0x993366FF
    icvPlt19 = 0x1A         # 0xFFFFCCFF
    icvPlt20 = 0x1B         # 0xCCFFFFFF
    icvPlt21 = 0x1C         # 0x660066FF
    icvPlt22 = 0x1D         # 0xFF8080FF
    icvPlt23 = 0x1E         # 0x0066CCFF
    icvPlt24 = 0x1F         # 0xCCCCFFFF
    icvPlt25 = 0x20         # 0x000080FF
    icvPlt26 = 0x21         # 0xFF00FFFF
    icvPlt27 = 0x22         # 0xFFFF00FF
    icvPlt28 = 0x23         # 0x00FFFFFF
    icvPlt29 = 0x24         # 0x800080FF
    icvPlt30 = 0x25         # 0x800000FF
    icvPlt31 = 0x26         # 0x008080FF
    icvPlt32 = 0x27         # 0x0000FFFF
    icvPlt33 = 0x28         # 0x00CCFFFF
    icvPlt34 = 0x29         # 0xCCFFFFFF
    icvPlt35 = 0x2A         # 0xCCFFCCFF
    icvPlt36 = 0x2B         # 0xFFFF99FF
    icvPlt37 = 0x2C         # 0x99CCFFFF
    icvPlt38 = 0x2D         # 0xFF99CCFF
    icvPlt39 = 0x2E         # 0xCC99FFFF
    icvPlt40 = 0x2F         # 0xFFCC99FF
    icvPlt41 = 0x30         # 0x3366FFFF
    icvPlt42 = 0x31         # 0x33CCCCFF
    icvPlt43 = 0x32         # 0x99CC00FF
    icvPlt44 = 0x33         # 0xFFCC00FF
    icvPlt45 = 0x34         # 0xFF9900FF
    icvPlt46 = 0x35         # 0xFF6600FF
    icvPlt47 = 0x36         # 0x666699FF
    icvPlt48 = 0x37         # 0x969696FF
    icvPlt49 = 0x38         # 0x003366FF
    icvPlt50 = 0x39         # 0x339966FF
    icvPlt51 = 0x3A         # 0x003300FF
    icvPlt52 = 0x3B         # 0x333300FF
    icvPlt53 = 0x3C         # 0x993300FF
    icvPlt54 = 0x3D         # 0x993366FF
    icvPlt55 = 0x3E         # 0x333399FF
    icvPlt56 = 0x3F         # 0x333333FF
    icvForeground = 0x40    # System color for text in windows.
    icvBackground = 0x41    # System color for window background.
    icvFrame = 0x42         # System color for window frame.
    icv3D = 0x43            # System-defined face color for three-dimensional display elements and for dialog box backgrounds.
    icv3DText = 0x44        # System color for text on push buttons.
    icv3DHilite = 0x45      # System highlight color for three-dimensional display elements (for edges facing the light source).
    icv3DShadow = 0x46      # System shadow color for three-dimensional display elements (for edges facing away from the light source).
    icvHilite = 0x47        # System color for items selected in a control.
    icvCtlText = 0x48       # System color for text in windows.
    icvCtlScrl = 0x49       # System color for scroll bar gray area.
    icvCtlInv = 0x4A        # Bitwise inverse of the RGB value of icvCtlScrl.
    icvCtlBody = 0x4B       # System color for window background.
    icvCtlFrame = 0x4C      # System color for window frame.
    icvCrtFore = 0x4D       # System color for text in windows.
    icvCrtBack = 0x4E       # System color for window background.
    icvCrtNeutral = 0x4F    # 0x000000FF
    icvInfoBk = 0x50        # System background color for tooltip controls.
    icvInfoText = 0x51      # System text color for tooltip controls.
    
    def get_rgba(self):
        if self == PaletteColor.icvBlack: return 0x000000FF
        elif self == PaletteColor.icvWhite: return 0xFFFFFFFF
        elif self == PaletteColor.icvRed: return 0xFF0000FF
        elif self == PaletteColor.icvGreen: return 0x00FF00FF
        elif self == PaletteColor.icvBlue: return 0x0000FFFF
        elif self == PaletteColor.icvYellow: return 0xFFFF00FF
        elif self == PaletteColor.icvMagenta: return 0xFF00FFFF
        elif self == PaletteColor.icvCyan: return 0x00FFFFFF
        elif self == PaletteColor.icvPlt1: return 0x000000FF
        elif self == PaletteColor.icvPlt2: return 0xFFFFFFFF
        elif self == PaletteColor.icvPlt3: return 0xFF0000FF
        elif self == PaletteColor.icvPlt4: return 0x00FF00FF
        elif self == PaletteColor.icvPlt5: return 0x0000FFFF
        elif self == PaletteColor.icvPlt6: return 0xFFFF00FF
        elif self == PaletteColor.icvPlt7: return 0xFF00FFFF
        elif self == PaletteColor.icvPlt8: return 0x00FFFFFF
        elif self == PaletteColor.icvPlt9: return 0x800000FF
        elif self == PaletteColor.icvPlt10: return 0x008000FF
        elif self == PaletteColor.icvPlt11: return 0x000080FF
        elif self == PaletteColor.icvPlt12: return 0x808000FF
        elif self == PaletteColor.icvPlt13: return 0x800080FF
        elif self == PaletteColor.icvPlt14: return 0x008080FF
        elif self == PaletteColor.icvPlt15: return 0xC0C0C0FF
        elif self == PaletteColor.icvPlt16: return 0x808080FF
        elif self == PaletteColor.icvPlt17: return 0x9999FFFF
        elif self == PaletteColor.icvPlt18: return 0x993366FF
        elif self == PaletteColor.icvPlt19: return 0xFFFFCCFF
        elif self == PaletteColor.icvPlt20: return 0xCCFFFFFF
        elif self == PaletteColor.icvPlt21: return 0x660066FF
        elif self == PaletteColor.icvPlt22: return 0xFF8080FF
        elif self == PaletteColor.icvPlt23: return 0x0066CCFF
        elif self == PaletteColor.icvPlt24: return 0xCCCCFFFF
        elif self == PaletteColor.icvPlt25: return 0x000080FF
        elif self == PaletteColor.icvPlt26: return 0xFF00FFFF
        elif self == PaletteColor.icvPlt27: return 0xFFFF00FF
        elif self == PaletteColor.icvPlt28: return 0x00FFFFFF
        elif self == PaletteColor.icvPlt29: return 0x800080FF
        elif self == PaletteColor.icvPlt30: return 0x800000FF
        elif self == PaletteColor.icvPlt31: return 0x008080FF
        elif self == PaletteColor.icvPlt32: return 0x0000FFFF
        elif self == PaletteColor.icvPlt33: return 0x00CCFFFF
        elif self == PaletteColor.icvPlt34: return 0xCCFFFFFF
        elif self == PaletteColor.icvPlt35: return 0xCCFFCCFF
        elif self == PaletteColor.icvPlt36: return 0xFFFF99FF
        elif self == PaletteColor.icvPlt37: return 0x99CCFFFF
        elif self == PaletteColor.icvPlt38: return 0xFF99CCFF
        elif self == PaletteColor.icvPlt39: return 0xCC99FFFF
        elif self == PaletteColor.icvPlt40: return 0xFFCC99FF
        elif self == PaletteColor.icvPlt41: return 0x3366FFFF
        elif self == PaletteColor.icvPlt42: return 0x33CCCCFF
        elif self == PaletteColor.icvPlt43: return 0x99CC00FF
        elif self == PaletteColor.icvPlt44: return 0xFFCC00FF
        elif self == PaletteColor.icvPlt45: return 0xFF9900FF
        elif self == PaletteColor.icvPlt46: return 0xFF6600FF
        elif self == PaletteColor.icvPlt47: return 0x666699FF
        elif self == PaletteColor.icvPlt48: return 0x969696FF
        elif self == PaletteColor.icvPlt49: return 0x003366FF
        elif self == PaletteColor.icvPlt50: return 0x339966FF
        elif self == PaletteColor.icvPlt51: return 0x003300FF
        elif self == PaletteColor.icvPlt52: return 0x333300FF
        elif self == PaletteColor.icvPlt53: return 0x993300FF
        elif self == PaletteColor.icvPlt54: return 0x993366FF
        elif self == PaletteColor.icvPlt55: return 0x333399FF
        elif self == PaletteColor.icvPlt56: return 0x333333FF
        elif self == PaletteColor.icvCrtNeutral: return 0x000000FF
        else: return None

class ThemeColor(Enum):
    DK_1 = 0x00
    LT_1 = 0x01
    DK_2 = 0x02
    LT_2 = 0x03
    ACCENT_1 = 0x04
    ACCENT_2 = 0x05
    ACCENT_3 = 0x06
    ACCENT_4 = 0x07
    ACCENT_5 = 0x08
    ACCENT_6 = 0x09
    HLINK = 0x0A
    FOL_HLINK = 0x0B


class SubscriptType(Enum):
    NONE = 0x00
    SUPERSCRIPT = 0x01
    SUBSCRIPT = 0x02

class UnderlineType(Enum):
    NONE = 0x00
    SINGLE = 0x01
    DOUBLE = 0x02
    ACCOUNTING_SINGLE = 0x21
    ACCOUNTING_DOUBLE = 0x22

class FontFamilyType(Enum):
    NA = 0x00
    ROMAN = 0x01
    SWISS = 0x02
    MODERN = 0x03
    SCRIPT = 0x04
    DECORATIVE = 0x05

class CharacterSetType(Enum):
    ANSI_CHARSET = 0x00
    DEFAULT_CHARSET = 0x01
    SYMBOL_CHARSET = 0x02
    MAC_CHARSET = 0x4D
    SHIFTJIS_CHARSET = 0x80
    HANGUL_CHARSET = 0x81
    JOHAB_CHARSET = 0x82
    GB2312_CHARSET = 0x86
    CHINESEBIG5_CHARSET = 0x88
    GREEK_CHARSET = 0xA1
    TURKISH_CHARSET = 0xA2
    VIETNAMESE_CHARSET = 0xA3
    HEBREW_CHARSET = 0xB1
    ARABIC_CHARSET = 0xB2
    BALTIC_CHARSET = 0xBA
    RUSSIAN_CHARSET = 0xCC
    THAI_CHARSET = 0xDE
    EASTEUROPE_CHARSET = 0xEE
    OEM_CHARSET = 0xFF

class FontSchemeType(Enum):
    NONE = 0x00
    MAJOR = 0x01
    MINOR = 0x02


class FillType(Enum):
    NONE = 0x00
    SOLID = 0x01
    MEDIUM_GRAY = 0x02
    DARK_GRAY = 0x03
    LIGHT_GRAY = 0x04
    HORIZONTAL_STRIPES = 0x05
    VERTICAL_STRIPES = 0x06
    DOWNWARD_DIAGONAL_STRIPES = 0x07
    UPWARD_DIAGONAL_STRIPES = 0x08
    GRID = 0x09
    TRELLIS = 0x0a
    LIGHT_HORIZONTAL_STRIPES = 0x0b
    LIGHT_VERTICAL_STRIPES = 0x0c
    LIGHT_DOWNWARD_DIAGONAL_STRIPES = 0x0d
    LIGHT_UPWARD_DIAGONAL_STRIPES = 0x0e
    LIGHT_GRID = 0x0f
    LIGHT_TRELLIS = 0x10
    GRAYSCALE = 0x11
    LIGHT_GRAYSCALE = 0x12
    GRADIENT = 0x28

class GradientType(Enum):
    LINEAR = 0x00
    RECTANGULAR = 0x01

class BorderType(Enum):
    NONE = 0x00
    THIN = 0x01
    MEDIUM = 0x02
    DASHED = 0x03
    DOTTED = 0x04
    THICK = 0x05
    DOUBLE = 0x06
    HAIRLINE = 0x07
    MEDIUM_DASHED = 0x08
    DASH_DOT = 0x09
    MEDIUM_DASH_DOT = 0x0a
    DASH_DOT_DOT = 0x0b
    MEDIUM_DASH_DOT_DOT = 0x0c
    SLANT_DASH_DOT = 0x0d