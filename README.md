# BPMN 2.0 Powerpoint
Create BPMN 2.0 objects for Powerpoint presentation using VBA

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

![sample image](https://github.com/leonardo270570/bpmn-powerpoint/assets/488127/122cdd65-9c70-4c8f-b795-c5e588b0f91d)

