Attribute VB_Name = "BPMNTestCode"
Public Sub BusinessProcessExample4()
    dblSizeFactor = 1.5

    'slide metrics
    lSlideHeight = ActivePresentation.PageSetup.SlideHeight 'get slide vertical height
    lSlideWidth = ActivePresentation.PageSetup.SlideWidth 'get slide horizontal width
    
    m_slide = 67
    m_startrow = -25
    m_startcol = 0
    m_rowbetween = 66 * dblSizeFactor '45
    m_colbetween = 70 * dblSizeFactor
    
    Call DeleteAllShapes(m_slide)
    
    Call NewActivity(m_slide, 1, Trim(Str(fActivities.fNormal)), "Procurement Needs", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 2, Trim(Str(fActivities.fNormal)), "Send Request", _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 3, Trim(Str(fActivities.fNormal)), "Provide Criterias", _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 4, Trim(Str(fActivities.fNormal)), "Approve Request for Bidding", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 2)
    Call NewActivity(m_slide, 5, Trim(Str(fActivities.fNormal)), "Add Vendor To Database", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 2)
    Call NewActivity(m_slide, 6, Trim(Str(fActivities.fNormal)), "RFP", _
        dblSizeFactor, m_startcol + m_colbetween * 5, m_startrow + m_rowbetween * 2.35)
    Call NewActivity(m_slide, 7, Trim(Str(fActivities.fNormal)), "Quotation Received", _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 2.35)
    Call NewActivity(m_slide, 8, Trim(Str(fActivities.fNormal)), "Evaluation", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 9, Trim(Str(fActivities.fNormal)), "Technical", _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 10, Trim(Str(fActivities.fNormal)), "Financial", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 11, Trim(Str(fActivities.fNormal)), "Negotiations", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 4)
    Call NewActivity(m_slide, 12, Trim(Str(fActivities.fNormal)), "Issue Contract", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 4)
    Call NewActivity(m_slide, 13, Trim(Str(fActivities.fNormal)), "Sign Agreement", _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 4)
    Call NewActivity(m_slide, 14, Trim(Str(fActivities.fNormal)), "Release Purchase Order", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 5)
    

    Call NewGateway(m_slide, 15, Trim(Str(fGateways.fExclusiveMarker)), "Budget Approved?", True, _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 1)
    Call NewGateway(m_slide, 16, Trim(Str(fGateways.fExclusiveMarker)), "Vendor Shortlisted?", True, _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 2)
    Call NewGateway(m_slide, 17, Trim(Str(fGateways.fExclusiveMarker)), "Qualified Vendor?", True, _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 3)
    Call NewGateway(m_slide, 18, Trim(Str(fGateways.fExclusiveMarker)), "Negotiation Finalized?", True, _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 4)
    
    Call NewPane(m_slide, 19, fPanes.fLane, "Request", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow * -10 / 25 + m_rowbetween * 0.92 * 45 / 66, 900, 100)
    Call NewPane(m_slide, 20, fPanes.fLane, "Bidding", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow * -10 / 25 + m_rowbetween * 2.39 * 45 / 66, 900, 100)
    Call NewPane(m_slide, 21, fPanes.fLane, "Evaluation", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow * -10 / 25 + m_rowbetween * 3.86 * 45 / 66, 900, 100)
    Call NewPane(m_slide, 22, fPanes.fLane, "Negotiation", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow * -10 / 25 + m_rowbetween * 5.33 * 45 / 66, 900, 100)
    Call NewPane(m_slide, 23, fPanes.fLane, "Puchase Order", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow * -10 / 25 + m_rowbetween * 6.8 * 45 / 66, 900, 100)

    Call NewEvent(m_slide, 24, Trim(Str(fEvents.fStart)), "Start", True, _
        dblSizeFactor, m_startcol + m_colbetween * 1, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 25, Trim(Str(fEvents.fEnd)), "End", True, _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 5)

    m_startcon = 26
    Call NewConnector(m_slide, m_startcon + 0, "0, 24, 1, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 1, "0, 1, 2, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 2, "0, 2, 15, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 3, "0, 15, 3, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 4, "0, 3, 4, 2, 4, 0.5, 0.5", "", True, dblSizeFactor, "0,0.65,-1,0.35")
    
    Call NewConnector(m_slide, m_startcon + 5, "0, 4, 16, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 6, "0, 16, 5, 1, 3, 0.5, 0.5", "Yes", True, dblSizeFactor)
    
    Call NewConnector(m_slide, m_startcon + 7, "0, 7, 8, 3, 4, 0.5, 0.5", "", True, dblSizeFactor, "-1,1")
    Call NewConnector(m_slide, m_startcon + 8, "0, 8, 9, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 9, "0, 9, 10, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 10, "0, 10, 17, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)

    Call NewConnector(m_slide, m_startcon + 11, "0, 11, 18, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 12, "0, 18, 12, 1, 3, 0.5, 0.5", "Yes", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 13, "0, 12, 13, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)

    Call NewConnector(m_slide, m_startcon + 14, "0, 13, 14, 2, 4, 0.5, 0.5", "", True, dblSizeFactor, "0,0.65,-1,0.35")
    Call NewConnector(m_slide, m_startcon + 15, "0, 14, 25, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)

    Call NewConnector(m_slide, m_startcon + 16, "0, 17, 11, 2, 4, 0.5, 0.5", "Yes", True, dblSizeFactor, "0,0.65,-1,0.35")
    Call NewConnector(m_slide, m_startcon + 17, "0, 18, 17, 4, 1, 0.5, 0.5", "No", False, dblSizeFactor, "0,-0.35,1.05,-0.65,-0.05")
    Call NewConnector(m_slide, m_startcon + 18, "0, 6, 7, 3, 1, 0.5, 0.5", "", True, dblSizeFactor, "-1")

    Call NewConnector(m_slide, m_startcon + 19, "0, 16, 6, 2, 3, 0.5, 0.3", "No", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 20, "0, 5, 6, 2, 1, 0.5, 0.5", "", True, dblSizeFactor, "0,1,-1")
    Call NewConnector(m_slide, m_startcon + 21, "0, 17, 16, 4, 4, 0.5, 0.5", "No", True, dblSizeFactor, "0,-1.1,-1,0.1")
'    Call NewConnector(m_slide, m_startcon + 22, "1, 6, 3, 4, 2, 0.5, 0.7", "", True, dblSizeFactor, "0,-0.75,1,-0.25")
End Sub

Public Sub BusinessProcessExample3()
    dblSizeFactor = 1.5

    'slide metrics
    lSlideHeight = ActivePresentation.PageSetup.SlideHeight 'get slide vertical height
    lSlideWidth = ActivePresentation.PageSetup.SlideWidth 'get slide horizontal width
    
    m_slide = 65
    m_startrow = 10
    m_startcol = 0
    m_rowbetween = 45 * dblSizeFactor
    m_colbetween = 60 * dblSizeFactor
    
    Call DeleteAllShapes(m_slide)
    
    Call NewActivity(m_slide, 1, Trim(Str(fActivities.fNormal)), "Select a Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 1.85, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 2, Trim(Str(fActivities.fNormal)), "Order a Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 3, Trim(Str(fActivities.fNormal)), "Ask for the Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 2)
    Call NewActivity(m_slide, 4, Trim(Str(fActivities.fNormal)), "Pay the Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 5, Trim(Str(fActivities.fNormal)), "Eat the Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 9.15, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 6, Trim(Str(fActivities.fNormal)), "Call Customer", _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 4)
    Call NewActivity(m_slide, 7, Trim(Str(fActivities.fNormal)), "Bake the Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 5.25)
    Call NewActivity(m_slide, 8, Trim(Str(fActivities.fNormal)), "Deliver the Pizza", _
        dblSizeFactor, m_startcol + m_colbetween * 5, m_startrow + m_rowbetween * 6.5)
    Call NewActivity(m_slide, 9, Trim(Str(fActivities.fNormal)), "Receive Payment", _
        dblSizeFactor, m_startcol + m_colbetween * 6.15, m_startrow + m_rowbetween * 6.5)
    
    Call NewEvent(m_slide, 10, Trim(Str(fEvents.fStart)), "Hungry for Pizza", True, _
        dblSizeFactor, m_startcol + m_colbetween * 1, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 11, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fReceive)), "Pizza Received", True, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 12, Trim(Str(fEvents.fEnd)), "Hunger Satisfied", False, _
        dblSizeFactor, m_startcol + m_colbetween * 10, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 13, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fTimer)), "60 minutes", True, _
        dblSizeFactor, m_startcol + m_colbetween * 5, m_startrow + m_rowbetween * 2)
    Call NewEvent(m_slide, 14, Trim(Str(fEvents.fStart)) & "," & Trim(Str(fItems.fReceive)), "Order Received", False, _
        dblSizeFactor, m_startcol + m_colbetween * 1, m_startrow + m_rowbetween * 4)
    Call NewEvent(m_slide, 15, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fReceive)), "Where is my Pizza", True, _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 4)
    Call NewEvent(m_slide, 16, Trim(Str(fEvents.fEnd)) & "," & Trim(Str(fItems.fTermination)), "", True, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 6.5)

    Call NewGateway(m_slide, 17, Trim(Str(fGateways.fEventBased)), "", True, _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 1)
    Call NewGateway(m_slide, 18, Trim(Str(fGateways.fParallel)), "", True, _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 4)
    
    Call NewPane(m_slide, 19, fPanes.fLane, "Pizza Customer", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 1.5, 900, 180)
    Call NewPane(m_slide, 20, fPanes.fLane, "Delivery Boy", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 6.5, 900, 80)
    Call NewPane(m_slide, 21, fPanes.fLane, "Pizza Chef", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 5.325, 900, 80)
    Call NewPane(m_slide, 22, fPanes.fLane, "Clerk", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 4.005, 900, 100)
    
    m_startcon = 23
    Call NewConnector(m_slide, m_startcon + 0, "0, 10, 1, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 1, "0, 1, 2, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 2, "0, 2, 17, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 3, "0, 17, 11, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 4, "0, 11, 4, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 5, "0, 4, 5, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 6, "0, 5, 12, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 7, "0, 17, 13, 2, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 8, "0, 13, 3, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 9, "0, 14, 18, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 10, "0, 18, 15, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 11, "0, 15, 6, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 12, "0, 18, 7, 2, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 13, "0, 7, 8, 1, 3, 0.5, 0.5", "", True, dblSizeFactor, "0.5,1,0.5")
    Call NewConnector(m_slide, m_startcon + 14, "0, 8, 9, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 15, "0, 9, 16, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 16, "0, 6, 15, 1, 2, 0.5, 0.5", "", True, dblSizeFactor, "0.25,2.25,-1.25,-1.25")
    Call NewConnector(m_slide, m_startcon + 17, "1, 2, 14, 2, 4, 0.5, 0.5", "", True, dblSizeFactor, "0,0.5,-1,0.5")
    Call NewConnector(m_slide, m_startcon + 18, "1, 3, 15, 2, 4, 0.3, 0.5", "", True, dblSizeFactor, "0,0.5,-1,0.5")
    Call NewConnector(m_slide, m_startcon + 19, "1, 8, 11, 4, 2, 0.5, 0.5", "", True, dblSizeFactor, "0,-0.5,1,-0.5")
    Call NewConnector(m_slide, m_startcon + 20, "1, 9, 4, 4, 2, 0.7, 0.7", "", True, dblSizeFactor, "0,-0.5,1,-0.5")
    Call NewConnector(m_slide, m_startcon + 21, "1, 4, 9, 2, 4, 0.3, 0.3", "", True, dblSizeFactor, "0,0.75,-1,0.25")
    Call NewConnector(m_slide, m_startcon + 22, "1, 6, 3, 4, 2, 0.5, 0.7", "", True, dblSizeFactor, "0,-0.75,1,-0.25")
End Sub

Public Sub BusinessProcessExample2()
    dblSizeFactor = 1.5

    'slide metrics
    lSlideHeight = ActivePresentation.PageSetup.SlideHeight 'get slide vertical height
    lSlideWidth = ActivePresentation.PageSetup.SlideWidth 'get slide horizontal width
    
    m_slide = 64
    m_startrow = 100
    m_startcol = 50
    m_rowbetween = 45 * dblSizeFactor
    m_colbetween = 60 * dblSizeFactor
    
    Call DeleteAllShapes(m_slide)
    
    Call NewActivity(m_slide, 1, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fUser)), "Check Availability", _
        dblSizeFactor, m_startcol + m_colbetween * 2, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 2, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fManual)), "Ship Article", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 3, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fUser)) & "," & Trim(Str(fItems.fSubProcess)), "Financial Statement", _
        dblSizeFactor, m_startcol + m_colbetween * 7.15, m_startrow + m_rowbetween * 1)
    Call NewActivity(m_slide, 4, Trim(Str(fActivities.fCall)) & "," & Trim(Str(fItems.fSubProcess)), "Procurement", _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 2, 75)
    Call NewActivity(m_slide, 5, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "Inform Customer", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 6, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "Inform Customer", _
        dblSizeFactor, m_startcol + m_colbetween * 4.85, m_startrow + m_rowbetween * 4)
    Call NewActivity(m_slide, 7, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fUser)), "Remove Article from Catalogue", _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 4)
    
    Call NewEvent(m_slide, 8, Trim(Str(fEvents.fStart)) & "," & Trim(Str(fItems.fReceive)), "Order Received", True, _
        dblSizeFactor, m_startcol + m_colbetween * 1, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 9, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fError)), "Undeliverable", False, _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 1, 4, fEventPositionAtTasks.fLeftBottom)
    Call NewEvent(m_slide, 10, Trim(Str(fEvents.fNonInterrupting)) & "," & Trim(Str(fItems.fEscalation)), "Late Delivery", False, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 2, 4, fEventPositionAtTasks.fRightBottom)
    Call NewEvent(m_slide, 11, Trim(Str(fEvents.fEnd)), "Payment Received", True, _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 12, Trim(Str(fEvents.fEnd)), "Customer Informed", True, _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 3)
    Call NewEvent(m_slide, 13, Trim(Str(fEvents.fEnd)), "Article Removed", True, _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 4)

    Call NewGateway(m_slide, 14, Trim(Str(fGateways.fExclusiveMarker)), "Article Available", True, _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 1)
        
    m_startcon = 15
    Call NewConnector(m_slide, m_startcon + 0, "0, 8, 1, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 1, "0, 1, 14, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 2, "0, 14, 2, 1, 3, 0.5, 0.5", "Yes", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 3, "0, 14, 4, 2, 3, 0.5, 0.5", "No", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 4, "0, 2, 3, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 5, "0, 3, 11, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 6, "0, 5, 12, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 7, "0, 7, 13, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 8, "0, 6, 7, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 9, "0, 6, 11, 4, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 10, "0, 9, 6, 2, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 11, "0, 10, 5, 2, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 12, "0, 4, 2, 1, 2, 0.5, 0.5", "", True, dblSizeFactor, "1,1")
End Sub

Public Sub BusinessProcessExample1()
    dblSizeFactor = 1.5

    'slide metrics
    lSlideHeight = ActivePresentation.PageSetup.SlideHeight 'get slide vertical height
    lSlideWidth = ActivePresentation.PageSetup.SlideWidth 'get slide horizontal width
    
    m_slide = 63
    m_startrow = 0
    m_startcol = -10
    m_rowbetween = 45 * dblSizeFactor
    m_colbetween = 60 * dblSizeFactor
    
    Call DeleteAllShapes(m_slide)
    
    Call NewActivity(m_slide, 1, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fReceive)), "Receive Book Request", _
        dblSizeFactor, m_startcol + m_colbetween * 1.85, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 2, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fService)), "Get Book Status", _
        dblSizeFactor, m_startcol + m_colbetween * 3, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 3, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fUser)), "Checkout Book", _
        dblSizeFactor, m_startcol + m_colbetween * 5, m_startrow + m_rowbetween * 5)
    Call NewActivity(m_slide, 4, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "Checkout Reply", _
        dblSizeFactor, m_startcol + m_colbetween * 6.15, m_startrow + m_rowbetween * 5)
    Call NewActivity(m_slide, 5, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "On Loan Reply", _
        dblSizeFactor, m_startcol + m_colbetween * 5, m_startrow + m_rowbetween * 3)
    Call NewActivity(m_slide, 6, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fService)), "Request Hold", _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 2)
    Call NewActivity(m_slide, 7, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "Hold Reply", _
        dblSizeFactor, m_startcol + m_colbetween * 9.15, m_startrow + m_rowbetween * 2)
    Call NewActivity(m_slide, 8, Trim(Str(fActivities.fNormal)) & "," & Trim(Str(fItems.fSend)), "Cancel Request", _
        dblSizeFactor, m_startcol + m_colbetween * 8, m_startrow + m_rowbetween * 3)
    
    Call NewEvent(m_slide, 9, Trim(Str(fEvents.fStart)), "Start Event", True, _
        dblSizeFactor, m_startcol + m_colbetween * 1, m_startrow + m_rowbetween * 3)
    Call NewEvent(m_slide, 10, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fTimer)), "Two Weeks", True, _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 1)
    Call NewEvent(m_slide, 11, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fReceive)), "Hold Book", True, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 2)
    Call NewEvent(m_slide, 12, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fReceive)), "Decline Hold", True, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 3)
    Call NewEvent(m_slide, 13, Trim(Str(fEvents.fInterrupting)) & "," & Trim(Str(fItems.fTimer)), "One Week", True, _
        dblSizeFactor, m_startcol + m_colbetween * 7, m_startrow + m_rowbetween * 4)
    Call NewEvent(m_slide, 14, Trim(Str(fEvents.fEnd)), "Stop Event", True, _
        dblSizeFactor, m_startcol + m_colbetween * 9, m_startrow + m_rowbetween * 3)

    Call NewGateway(m_slide, 15, Trim(Str(fGateways.fExclusiveMarker)), "", True, _
        dblSizeFactor, m_startcol + m_colbetween * 4, m_startrow + m_rowbetween * 3)
    Call NewGateway(m_slide, 16, Trim(Str(fGateways.fEventBased)), "", True, _
        dblSizeFactor, m_startcol + m_colbetween * 6, m_startrow + m_rowbetween * 3)
    
    Call NewPane(m_slide, 17, fPanes.fLane, "Actuary", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 3, 900, 350)
    Call NewPane(m_slide, 18, fPanes.fLane, "Customer", dblSizeFactor, _
        m_startcol + m_colbetween * 0.5, m_startrow + m_rowbetween * 7, 900, 70)
    
    m_startcon = 19
    Call NewConnector(m_slide, m_startcon + 0, "0, 9, 1, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 1, "0, 1, 2, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 2, "0, 2, 15, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 3, "0, 15, 5, 1, 3, 0.5, 0.5", "Book is On Loan", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 4, "0, 5, 16, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 5, "0, 16, 12, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 6, "0, 12, 8, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 7, "0, 8, 14, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 8, "0, 16, 13, 2, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 9, "0, 16, 11, 4, 3, 0.5, 0.5", "", True, dblSizeFactor, "0,-1,1")
    Call NewConnector(m_slide, m_startcon + 11, "0, 13, 8, 1, 2, 0.5, 0.5", "", True, dblSizeFactor, "1,-1")
    Call NewConnector(m_slide, m_startcon + 12, "0, 15, 3, 2, 3, 0.5, 0.5", "Book is Available", True, dblSizeFactor, "0,1,1")
    Call NewConnector(m_slide, m_startcon + 13, "0, 3, 4, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 14, "0, 4, 14, 1, 2, 0.5, 0.5", "", True, dblSizeFactor, "1,-1")
    Call NewConnector(m_slide, m_startcon + 15, "0, 11, 6, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 16, "0, 6, 7, 1, 3, 0.5, 0.5", "", True, dblSizeFactor)
    Call NewConnector(m_slide, m_startcon + 17, "0, 7, 10, 4, 1, 0.5, 0.5", "", True, dblSizeFactor, "0,-1,-1")
    Call NewConnector(m_slide, m_startcon + 18, "0, 10, 2, 3, 4, 0.5, 0.5", "", True, dblSizeFactor, "-1,1")
    
End Sub

