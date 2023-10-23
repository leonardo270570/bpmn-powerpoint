Attribute VB_Name = "BPMNEnumeration"
'.Name convention : i + index, eg. i1, i23 (parent shape), i1-1, i1-2, and so on (child shapes)

'.Title convention (split by semicolon)
'1st term : fShapeTypes
'2nd term : coordinate x of center point
'3rd term : coordinate y of center point
'4th term : width
'5th term : height
'6th term : weight of outline
'7th term : dblSizeFactor


'Object Items convention (split by comma)
'1st term : type of activities, events, gateways (parent)
'2nd term : additional shape items to be added to the parent (child)

Public Enum fActivities
    fNormal = 0
    fCall = 1
    fEvent = 2
    fTransactional = 3
End Enum

Public Enum fGateways
    fExclusive = 0
    fParallel = 1
    fInclusive = 2
    fEventBased = 3
    fParallelEventBased = 4
    fComplex = 5
    fExclusiveMarker = 6
End Enum

Public Enum fEvents
    fStart = 0
    fCatch = 1
    fInterrupting = 2
    fNonInterrupting = 3
    fThrow = 4
    fEnd = 5
End Enum

Public Enum fItems
    fTermination = 0
    fTimer = 1
    fConditional = 2
    fLink = 3
    fLinkSolid = 4
    fSignal = 5
    fSignalSolid = 6
    fEscalation = 7
    fEscalationSolid = 8
    fError = 9
    fErrorSolid = 10
    fCancel = 11
    fCancelSolid = 12
    fMultipleParallel = 13
    fMultipleParallelSolid = 14
    fCompensation = 15
    fCompensationSolid = 16
    fMultiple = 17
    fMultipleSolid = 18
    fUser = 19
    fService = 20
    fScript = 21
    fManual = 22
    fBusinessRule = 23
    fSend = 24
    fReceive = 25
    fAdhoc = 26
    fLoop = 27
    fMultiInstance = 28
    fSubProcess = 29
    fSequential = 30
End Enum

Public Enum fConnectors
    fSequence = 0
    fMessage = 1
    fConditional = 2
    fDefault = 3
    fAssociation = 4
    fDirectional = 5
    fBidirectional = 6
    fConversation = 7
End Enum

Public Enum fConnectorTypes
    fStraight = 0
    fElbow = 1
    fCustom = 2
End Enum

Public Enum fShapeSides
    fRight = 1
    fBottom = 2
    fLeft = 3
    fUp = 4
End Enum

Public Enum fPanes
    fPool = 0
    fLane = 1
End Enum

Public Enum fArtifacts
    fDataStore = 0
    fDataObject = 1
    fGroup = 2
    fAnnotation = 3
    fDataInput = 4
    fDataOutput = 5
End Enum

Public Enum fShapeTypes
    fRectangle = 1
    fDiamond = 2
    fCircle = 3
    fCross = 4
    fPentagon = 5
    fTriangle = 6
    fArrow = 7
    fLine = 8
    fPoolLane = 9
    fObject = 10
End Enum

Public Enum fRectangles
    fStandard = 0
    fStandardDash = 1
    fThinOuter = 2
    fThinInner = 3
    fThick = 4
    fSmall = 5
    fGroup = 6
'    fPoolLaneTitle = 7
'    fPoolLaneContent = 8
End Enum

Public Enum fCircles
    fStandard = 0
    fThinOuter = 1
    fThinInner = 2
    fThinInnerInner = 3
    fThick = 4
    fSolidInner = 5
    fDashOuter = 6
    fDashInner = 7
    fSmall = 8
End Enum

Public Enum fDiamonds
    fStandard = 0
    fSmall = 1
End Enum

Public Enum fLines
    fSolid = 0
    fDot = 1
    fDash = 2
    fDouble = 3
End Enum

Public Enum fCrosses
    fSolidPlusSmall = 0
    fSolidPlus = 1
    fSolidPlusLarge = 2
    fSolidMultiply = 3
    fSolidMultiplyLarge = 4
    fOutlinePlus = 5
    fOutlineMultiply = 6
End Enum

Public Enum fEventPositionAtTasks
    fUpperCorner = 0
    fLeftBottom = 1
    fRightBottom = 2
    fCenterBottom = 3
    fLeftUp = 4
    fRightUp = 5
    fCenterUp = 6
End Enum
