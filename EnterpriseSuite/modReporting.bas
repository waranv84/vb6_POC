Attribute VB_Name = "modReporting"
Option Explicit

Private Const AD_OPEN_FORWARD_ONLY As Long = 0
Private Const AD_LOCK_READONLY As Long = 1
Private Const AD_CMD_TEXT As Long = 1

Public Function BuildDashboardSummary() As String
    Dim premium As Currency
    Dim openClaims As Currency
    Dim outstandingPremium As Currency
    Dim expiring As Long
    Dim expiredLicenses As Long

    premium = modDatabase.GetTotalPremium()
    openClaims = modDatabase.GetOpenClaimAmount()
    outstandingPremium = modDatabase.GetPastDuePremium()
    expiring = modDatabase.GetExpiringPolicies(DateAdd("m", 2, Date))
    expiredLicenses = modDatabase.GetAgentsWithExpiredLicenses()

    BuildDashboardSummary = "Total Bound Premium: " & modUtilities.FormatCurrencyValue(premium) & vbCrLf & _
                            "Open Claim Exposure: " & modUtilities.FormatCurrencyValue(openClaims) & vbCrLf & _
                            "Outstanding Premium: " & modUtilities.FormatCurrencyValue(outstandingPremium) & vbCrLf & _
                            "Policies Expiring (60 days): " & CStr(expiring) & vbCrLf & _
                            "Agents with expired licenses: " & CStr(expiredLicenses)
End Function

Public Function GetPolicyStatusBreakdown() As Object
    Set GetPolicyStatusBreakdown = modDatabase.CountByStatus("Policies", "Status")
End Function

Public Function GetClaimStatusBreakdown() As Object
    Set GetClaimStatusBreakdown = modDatabase.CountByStatus("Claims", "Status")
End Function

Public Function GetInvoiceStatusBreakdown() As Object
    Set GetInvoiceStatusBreakdown = modDatabase.CountByStatus("Invoices", "Status")
End Function

Public Function GetComplianceAlerts() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT PolicyNumber, DocumentName, ComplianceStatus FROM Documents WHERE ComplianceStatus <> 'Approved' ORDER BY ReceivedDate DESC", _
            modDatabase.gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        results.Add rs.Fields("PolicyNumber").Value & " - " & rs.Fields("DocumentName").Value & " (" & rs.Fields("ComplianceStatus").Value & ")"
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    If results.Count = 0 Then
        results.Add "No outstanding compliance items."
    End If

    Set GetComplianceAlerts = results
End Function

Public Function GetRecentAuditTrail(ByVal topN As Integer) As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT TOP " & topN & " EventTime, Entity, EntityKey, Action FROM AuditTrail ORDER BY EventTime DESC", modDatabase.gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        results.Add Format$(rs.Fields("EventTime").Value, "yyyy-mm-dd hh:nn") & " - " & rs.Fields("Entity").Value & " " & rs.Fields("EntityKey").Value & " (" & rs.Fields("Action").Value & ")"
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    If results.Count = 0 Then
        results.Add "No audit events captured yet."
    End If

    Set GetRecentAuditTrail = results
End Function

Public Function BuildPolicyNarrative(ByVal policy As clsPolicy) As String
    Dim sb As String
    sb = "Policy " & policy.PolicyNumber & " for " & policy.HolderName & " (" & policy.ProductType & ")" & vbCrLf
    sb = sb & "Effective " & modUtilities.FormatDateValue(policy.EffectiveDate) & " through " & modUtilities.FormatDateValue(policy.ExpirationDate) & vbCrLf
    sb = sb & "Premium: " & modUtilities.FormatCurrencyValue(policy.Premium) & vbCrLf
    sb = sb & "Status: " & policy.Status & " | Agent: " & policy.AgentCode
    BuildPolicyNarrative = sb
End Function

Public Function BuildClaimNarrative(ByVal claim As clsClaim) As String
    Dim sb As String
    sb = "Claim " & claim.ClaimNumber & " on policy " & claim.PolicyNumber & vbCrLf
    sb = sb & "Loss Date: " & modUtilities.FormatDateValue(claim.LossDate) & " | Reported: " & modUtilities.FormatDateValue(claim.ReportedDate) & vbCrLf
    sb = sb & "Amount Reserved: " & modUtilities.FormatCurrencyValue(claim.Amount) & vbCrLf
    sb = sb & "Adjuster: " & claim.AdjusterName & vbCrLf
    sb = sb & "Status: " & claim.Status & vbCrLf & claim.Description
    BuildClaimNarrative = sb
End Function

Public Function BuildInvoiceNarrative(ByVal invoice As clsInvoice) As String
    Dim sb As String
    sb = "Invoice " & invoice.InvoiceNumber & " for policy " & invoice.PolicyNumber & vbCrLf
    sb = sb & "Due: " & modUtilities.FormatDateValue(invoice.DueDate) & vbCrLf
    sb = sb & "Amount Due: " & modUtilities.FormatCurrencyValue(invoice.AmountDue) & " | Paid: " & modUtilities.FormatCurrencyValue(invoice.AmountPaid) & vbCrLf
    sb = sb & "Status: " & invoice.Status & " (updated " & modUtilities.FormatDateValue(invoice.LastUpdated) & ")"
    BuildInvoiceNarrative = sb
End Function
