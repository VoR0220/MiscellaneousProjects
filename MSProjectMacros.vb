'******************************************************************************************************************************'
'*****                    Check the task predecessors and Email the appropriate person                                    *****'
'******************************************************************************************************************************'

Sub CheckPreds()
    
    '-----------------------------------------------------------------------'
    '--                 Variable Definition                             ----'
    '-----------------------------------------------------------------------'
    
    Dim ts As Tasks  'Active task selection
    Dim t As Task
    Dim tPred As Task
    Set ts = ActiveSelection.Tasks
    
    '-----------------------------------------------------------------------'
    '--                 Main Loop                                       ----'
    '-----------------------------------------------------------------------'
    For Each t In ts
        If (t Is Nothing) Or (t.Summary) Then
        ' do nothing on blank lines
        Else
            If t.PercentComplete = 100 Then
                'SetTaskField Field:="Text15", Value:="Completed", TaskID:=t.ID
                SetTaskField Field:="Text6", Value:="Completed", TaskID:=t.ID
                SetTaskField Field:="Text5", Value:="SENT", TaskID:=t.ID
            Else
                pcount = 0
                pcompl = 0
                For Each tPred In t.PredecessorTasks  'looping through the predecessor tasks
                    pcount = pcount + 1
                    percomp = tPred.PercentComplete
                    If percomp = "100" Then pcompl = pcompl + 1
                    Next tPred
                    If pcount = 0 Then
                        ready = True
                    Else
                        If pcompl = pcount Then
                            ready = True
                        Else
                            ready = False
                        End If
                    End If
        
                    If ready Then
                        'SetTaskField Field:="Text15", Value:="Ready", TaskID:=t.ID
                        SetTaskField Field:="Text6", Value:="Ready", TaskID:=t.ID
                        
                        '***********************************************************************************
                        ' If the task is ready and the email has not been sent previously: Generating Email
                        '***********************************************************************************
                
                            If (t.Text5 = "No") Or (t.Text5 = "") Then
                                Call Send_Outlook_Email(t)
                                SetTaskField Field:="Text5", Value:="Yes", TaskID:=t.ID
                                SetTaskField Field:="Text6", Value:="In Progress", TaskID:=t.ID
                            End If
                        Else
                            SetTaskField Field:="Text15", Value:="Not ready", TaskID:=t.ID
                            SetTaskField Field:="Text6", Value:="Not Ready", TaskID:=t.ID
                        End If
                End If
            End If
        Next t
    
    'Release Objects from Memory

    Set tPred = Nothing
    Set t = Nothing
    Set ts = Nothing
    
End Sub


Sub Send_Outlook_Email(t As Task)

    '-----------------------------------------------------------------------'
    '--                 Variable Definition                             ----'
    '-----------------------------------------------------------------------'

    Dim objOLApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Const stSubjectStart As String = "Edge Legacy Implementation Project - <"
    Const stSubjectEnd As String = ">"
    
    ' Set the Application object
    Set objOLApp = CreateObject("Outlook.Application")
    Set olMail = objOLApp.CreateItem(olMailItem)
    
    'Outlook mail inputs
    
    olMail.To = t.Text24 'email Responsible in To Category
    olMail.CC = t.Text23 + t.Text26 + t.Text21 'Accountable, Consulted, Informed in that order
    olMail.Subject = stSubjectStart + CStr(t.UniqueID) + "-" + t.Name + stSubjectEnd
    
    Select Case t.Text25
        Case "1-FSD"        'Sample body of email (if you need further help here, just ask me)
                            ' "&" is used for appending to a string
                            ' "_" is used for extending a string when writing it in code (notice how it's always at the end of the statement)
                            ' "vbCr" is essentially an endline statement that will push your text to the next line
                            ' "vbTab"...this one is pretty self explanatory
            olMail.Body = _
            "Dear " & stSubjectStart & ", " & vbCr & vbCr & vbTab & _
            "I'll see you in 2 minutes for our meeting!" & vbCr & vbCr & _
            "Btw: I've added you to my contact list."
        Case "2-TSD"
        
        Case "3-DEV"
        
        Case "4-TUT"
        
        Case "5-FUT"
    End Select
    
    'olMail.Send    'optional call to automate sending of messages...may want to leave this commented out in case of additional notes
    
    olMail.Display
    
    Set objOLApp = Nothing  'release memory
    Set olMail = Nothing
    
End Sub

'******************************************************************************'
'***** This method will refresh the task status of every line in the plan *****'
'******************************************************************************'

Sub RefreshTaskStatus()
    Dim tsks As Tasks
    Dim t As Task
    Dim rgbColor As Long
    Dim predCount As Integer
    Dim predComplete As Integer
    
    OutlineShowAllTasks
    FilterApply "All Tasks"
    
    Set tsks = ActiveProject.Tasks
    
   For Each t In tsks
        ' We do not need to worry about the summary tasks
        If (Not t Is Nothing) And (t.Summary) Then
            SelectRow Row:=t.ID, RowRelative:=False
            Font32Ex CellColor:=&HFFFFFF
        End If
        
        If t.PercentComplete = "100" Then
            'Font32Ex CellColor:=&HCCFFCC
            SetTaskField Field:="Text6", Value:="Completed", TaskID:=t.ID
        End If
        
        ready = False
        If (Not t Is Nothing) And (Not t.Summary) And (t.PercentComplete <> "100") Then
            SelectTaskField Row:=t.ID, Column:="Name", RowRelative:=False
            rgbColor = ActiveCell.CellColorEx
            pcount = 0
            pcompl = 0
        
            For Each tPred In t.PredecessorTasks  'looping through the predecessor tasks
                    pcount = pcount + 1
                    percomp = tPred.PercentComplete
                    If percomp = "100" Then pcompl = pcompl + 1
            Next tPred
                
                If pcount = 0 Then
                        ready = True
                Else
                    If pcompl = pcount Then
                        ready = True
                     Else
                        ready = False
                     End If
                End If
                If (ready) Then
                    'Font32Ex CellColor:=&HF0D9C6
                    SetTaskField Field:="Text6", Value:="Ready", TaskID:=t.ID
                    If (t.Text5 = "Yes") Then
                        SetTaskField Field:="Text6", Value:="In Progress", TaskID:=t.ID
                    End If
                        
                Else
                    'Font32Ex CellColor:=&HFFFFFF
                    SetTaskField Field:="Text6", Value:="Not Ready", TaskID:=t.ID
                End If
            End If
    Next t
  
    
    
End Sub

Private Sub Project_Open(ByVal pj As Project)
    AddHighlightRibbon
End Sub


Private Sub AddHighlightRibbon()
    Dim ribbonXml As String
    
    ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    ribbonXml = ribbonXml + "  <mso:ribbon>"
    ribbonXml = ribbonXml + "    <mso:qat/>"
    ribbonXml = ribbonXml + "    <mso:tabs>"
    ribbonXml = ribbonXml + "      <mso:tab id=""gacTab"" label=""GAC Macros"" insertBeforeQ=""mso:TabFormat"">"
    ribbonXml = ribbonXml + "        <mso:group id=""testGroup"" label=""Run Macros"" autoScale=""false"">"
    ribbonXml = ribbonXml + "          <mso:button id=""refreshtaskStatuss"" label=""Refresh Task Status"" "
    ribbonXml = ribbonXml + "imageMso=""QueryAppend"" onAction=""RefreshTaskStatus""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""generateEmail"" label=""Check Predecessors and Generate Email"" "
    ribbonXml = ribbonXml + "imageMso=""Consolidate"" onAction=""CheckPreds""/>"
    ribbonXml = ribbonXml + "        </mso:group>"
    ribbonXml = ribbonXml + "      </mso:tab>"
    ribbonXml = ribbonXml + "    </mso:tabs>"
    ribbonXml = ribbonXml + "  </mso:ribbon>"
    ribbonXml = ribbonXml + "</mso:customUI>"
    
    ActiveProject.SetCustomUI (ribbonXml)
End Sub


