Attribute VB_Name = "modDatabase"
Option Explicit

Private Const DB_FILE_NAME As String = "InsuranceSuite.mdb"
Private Const PROVIDER_STRING As String = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="

Private Const AD_OPEN_FORWARD_ONLY As Long = 0
Private Const AD_OPEN_STATIC As Long = 3
Private Const AD_LOCK_READONLY As Long = 1
Private Const AD_LOCK_OPTIMISTIC As Long = 3
Private Const AD_CMD_TEXT As Long = 1

Public gConn As Object

Private Function SqlEscape(ByVal value As String) As String
    SqlEscape = Replace(value, "'", "''")
End Function

Private Function SqlDate(ByVal value As Date) As String
    SqlDate = "#" & Format$(value, "yyyy-mm-dd hh:nn:ss") & "#"
End Function

Public Sub InitializeDatabase()
    Dim dataFolder As String
    Dim dbPath As String

    dataFolder = modUtilities.AppDataPath()
    modUtilities.EnsureFolderExists dataFolder
    dbPath = modUtilities.CombinePath(dataFolder, DB_FILE_NAME)

    If Not modUtilities.FileExists(dbPath) Then
        CreateDatabaseWithSchema dbPath
    End If

    Set gConn = CreateObject("ADODB.Connection")
    gConn.Open PROVIDER_STRING & dbPath

    SeedDataIfNeeded
End Sub

Private Sub CreateDatabaseWithSchema(ByVal dbPath As String)
    Dim catalog As Object
    Dim conn As Object

    Set catalog = CreateObject("ADOX.Catalog")
    catalog.Create PROVIDER_STRING & dbPath & ";Jet OLEDB:Engine Type=5;"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open PROVIDER_STRING & dbPath

    conn.Execute "CREATE TABLE Policies (" & _
                 "PolicyID AUTOINCREMENT PRIMARY KEY, " & _
                 "PolicyNumber TEXT(20), " & _
                 "HolderName TEXT(100), " & _
                 "ProductType TEXT(50), " & _
                 "EffectiveDate DATETIME, " & _
                 "ExpirationDate DATETIME, " & _
                 "Premium CURRENCY, " & _
                 "Status TEXT(30), " & _
                 "AgentCode TEXT(20))"

    conn.Execute "CREATE TABLE Claims (" & _
                 "ClaimID AUTOINCREMENT PRIMARY KEY, " & _
                 "ClaimNumber TEXT(20), " & _
                 "PolicyNumber TEXT(20), " & _
                 "LossDate DATETIME, " & _
                 "ReportedDate DATETIME, " & _
                 "Amount CURRENCY, " & _
                 "Status TEXT(30), " & _
                 "AdjusterName TEXT(100), " & _
                 "Description MEMO)"

    conn.Execute "CREATE TABLE Agents (" & _
                 "AgentID AUTOINCREMENT PRIMARY KEY, " & _
                 "AgentCode TEXT(20), " & _
                 "FullName TEXT(100), " & _
                 "Region TEXT(50), " & _
                 "Email TEXT(100), " & _
                 "Phone TEXT(30), " & _
                 "LicenseExpiration DATETIME, " & _
                 "Active YESNO)"

    conn.Execute "CREATE TABLE Invoices (" & _
                 "InvoiceID AUTOINCREMENT PRIMARY KEY, " & _
                 "InvoiceNumber TEXT(20), " & _
                 "PolicyNumber TEXT(20), " & _
                 "DueDate DATETIME, " & _
                 "AmountDue CURRENCY, " & _
                 "AmountPaid CURRENCY, " & _
                 "Status TEXT(30), " & _
                 "LastUpdated DATETIME)"

    conn.Execute "CREATE TABLE UnderwritingCases (" & _
                 "CaseID AUTOINCREMENT PRIMARY KEY, " & _
                 "PolicyNumber TEXT(20), " & _
                 "SubmissionDate DATETIME, " & _
                 "RiskScore INTEGER, " & _
                 "AssignedTo TEXT(100), " & _
                 "Status TEXT(30), " & _
                 "Notes MEMO)"

    conn.Execute "CREATE TABLE Documents (" & _
                 "DocumentID AUTOINCREMENT PRIMARY KEY, " & _
                 "PolicyNumber TEXT(20), " & _
                 "DocumentName TEXT(100), " & _
                 "DocumentType TEXT(50), " & _
                 "ReceivedDate DATETIME, " & _
                 "ComplianceStatus TEXT(30))"

    conn.Execute "CREATE TABLE AuditTrail (" & _
                 "AuditID AUTOINCREMENT PRIMARY KEY, " & _
                 "EventTime DATETIME, " & _
                 "Entity TEXT(30), " & _
                 "EntityKey TEXT(30), " & _
                 "Action TEXT(30), " & _
                 "Details MEMO)"

    conn.Close
    Set conn = Nothing
    Set catalog = Nothing
End Sub

Private Sub SeedDataIfNeeded()
    If TableIsEmpty("Agents") Then
        SeedAgents
    End If

    If TableIsEmpty("Policies") Then
        SeedPolicies
    End If

    If TableIsEmpty("Claims") Then
        SeedClaims
    End If

    If TableIsEmpty("Invoices") Then
        SeedInvoices
    End If

    If TableIsEmpty("UnderwritingCases") Then
        SeedUnderwritingCases
    End If

    If TableIsEmpty("Documents") Then
        SeedDocuments
    End If
End Sub

Private Function TableIsEmpty(ByVal tableName As String) As Boolean
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) AS TotalCount FROM [" & tableName & "]", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    TableIsEmpty = (rs.Fields("TotalCount").Value = 0)
    rs.Close
    Set rs = Nothing
End Function

Private Sub SeedAgents()
    gConn.Execute "INSERT INTO Agents (AgentCode, FullName, Region, Email, Phone, LicenseExpiration, Active) VALUES (" & _
                 "'AGT-100', 'Sonia Patel', 'Southwest', 'spatel@brightcoverage.com', '(480) 555-2876', #" & DateAdd("m", 10, Date) & "#, TRUE)"
    gConn.Execute "INSERT INTO Agents (AgentCode, FullName, Region, Email, Phone, LicenseExpiration, Active) VALUES (" & _
                 "'AGT-220', 'Marcus Hsu', 'Pacific Northwest', 'mhsu@brightcoverage.com', '(206) 555-6643', #" & DateAdd("m", 4, Date) & "#, TRUE)"
    gConn.Execute "INSERT INTO Agents (AgentCode, FullName, Region, Email, Phone, LicenseExpiration, Active) VALUES (" & _
                 "'AGT-305', 'Evelyn Wright', 'Midwest', 'ewright@brightcoverage.com', '(312) 555-8921', #" & DateAdd("m", -2, Date) & "#, FALSE)"
End Sub

Private Sub SeedPolicies()
    gConn.Execute "INSERT INTO Policies (PolicyNumber, HolderName, ProductType, EffectiveDate, ExpirationDate, Premium, Status, AgentCode) VALUES (" & _
                 "'POL-10045', 'Cobalt Manufacturing', 'Commercial General Liability', #" & DateSerial(2024, 1, 1) & "#, #" & DateSerial(2024, 12, 31) & "#, 24500.00, 'Active', 'AGT-100')"
    gConn.Execute "INSERT INTO Policies (PolicyNumber, HolderName, ProductType, EffectiveDate, ExpirationDate, Premium, Status, AgentCode) VALUES (" & _
                 "'POL-10078', 'HealthFirst Clinics', 'Professional Liability', #" & DateSerial(2024, 3, 15) & "#, #" & DateSerial(2025, 3, 14) & "#, 78500.00, 'Pending Renewal', 'AGT-220')"
    gConn.Execute "INSERT INTO Policies (PolicyNumber, HolderName, ProductType, EffectiveDate, ExpirationDate, Premium, Status, AgentCode) VALUES (" & _
                 "'POL-10102', 'Northwind Hotels', 'Property', #" & DateSerial(2023, 9, 1) & "#, #" & DateSerial(2024, 8, 31) & "#, 54000.00, 'Under Review', 'AGT-305')"
End Sub

Private Sub SeedClaims()
    gConn.Execute "INSERT INTO Claims (ClaimNumber, PolicyNumber, LossDate, ReportedDate, Amount, Status, AdjusterName, Description) VALUES (" & _
                 "'CLM-7001', 'POL-10045', #" & DateSerial(2024, 2, 20) & "#, #" & DateSerial(2024, 2, 21) & "#, 18450.75, 'Investigation', 'Brian Kelly', 'Water damage due to sprinkler malfunction in primary warehouse.')"
    gConn.Execute "INSERT INTO Claims (ClaimNumber, PolicyNumber, LossDate, ReportedDate, Amount, Status, AdjusterName, Description) VALUES (" & _
                 "'CLM-7022', 'POL-10078', #" & DateSerial(2024, 4, 4) & "#, #" & DateSerial(2024, 4, 4) & "#, 6200.0, 'Submitted', 'Danielle Rose', 'Medical malpractice allegation pending supporting documentation.')"
    gConn.Execute "INSERT INTO Claims (ClaimNumber, PolicyNumber, LossDate, ReportedDate, Amount, Status, AdjusterName, Description) VALUES (" & _
                 "'CLM-7056', 'POL-10102', #" & DateSerial(2023, 12, 18) & "#, #" & DateSerial(2023, 12, 19) & "#, 45320.5, 'Awaiting Settlement', 'Imani Walker', 'Fire suppression system leak impacting banquet hall level.')"
End Sub

Private Sub SeedInvoices()
    gConn.Execute "INSERT INTO Invoices (InvoiceNumber, PolicyNumber, DueDate, AmountDue, AmountPaid, Status, LastUpdated) VALUES (" & _
                 "'INV-9001', 'POL-10045', #" & DateSerial(2024, 3, 31) & "#, 6125.00, 6125.00, 'Paid', #" & DateSerial(2024, 3, 28) & "#)"
    gConn.Execute "INSERT INTO Invoices (InvoiceNumber, PolicyNumber, DueDate, AmountDue, AmountPaid, Status, LastUpdated) VALUES (" & _
                 "'INV-9014', 'POL-10078', #" & DateSerial(2024, 6, 30) & "#, 19500.00, 0, 'Past Due', #" & DateSerial(2024, 7, 5) & "#)"
    gConn.Execute "INSERT INTO Invoices (InvoiceNumber, PolicyNumber, DueDate, AmountDue, AmountPaid, Status, LastUpdated) VALUES (" & _
                 "'INV-9030', 'POL-10102', #" & DateSerial(2024, 5, 15) & "#, 13500.00, 13500.00, 'Paid', #" & DateSerial(2024, 5, 15) & "#)"
End Sub

Private Sub SeedUnderwritingCases()
    gConn.Execute "INSERT INTO UnderwritingCases (PolicyNumber, SubmissionDate, RiskScore, AssignedTo, Status, Notes) VALUES (" & _
                 "'POL-10220', #" & DateSerial(2024, 6, 12) & "#, 72, 'Elena Michaels', 'Awaiting Documents', 'Requesting updated loss runs from broker.')"
    gConn.Execute "INSERT INTO UnderwritingCases (PolicyNumber, SubmissionDate, RiskScore, AssignedTo, Status, Notes) VALUES (" & _
                 "'POL-10345', #" & DateSerial(2024, 6, 18) & "#, 45, 'Jon Rivera', 'In Review', 'Need engineering inspection for new plant addition.')"
    gConn.Execute "INSERT INTO UnderwritingCases (PolicyNumber, SubmissionDate, RiskScore, AssignedTo, Status, Notes) VALUES (" & _
                 "'POL-10372', #" & DateSerial(2024, 6, 25) & "#, 88, 'Tara Singh', 'Escalated', 'High catastrophe exposure due to coastal footprint.')"
End Sub

Private Sub SeedDocuments()
    gConn.Execute "INSERT INTO Documents (PolicyNumber, DocumentName, DocumentType, ReceivedDate, ComplianceStatus) VALUES (" & _
                 "'POL-10045', 'Warehouse Inspection', 'Risk Assessment', #" & DateSerial(2024, 2, 1) & "#, 'Approved')"
    gConn.Execute "INSERT INTO Documents (PolicyNumber, DocumentName, DocumentType, ReceivedDate, ComplianceStatus) VALUES (" & _
                 "'POL-10078', 'Medical Licensing Certificates', 'Compliance', #" & DateSerial(2024, 3, 10) & "#, 'Under Review')"
    gConn.Execute "INSERT INTO Documents (PolicyNumber, DocumentName, DocumentType, ReceivedDate, ComplianceStatus) VALUES (" & _
                 "'POL-10102', 'Fire Safety Plan', 'Safety', #" & DateSerial(2023, 11, 20) & "#, 'Expired')"
End Sub

Public Function GetPolicies() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT PolicyNumber, HolderName, ProductType, EffectiveDate, ExpirationDate, Premium, Status, AgentCode FROM Policies ORDER BY HolderName", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim policy As clsPolicy
        Set policy = New clsPolicy
        policy.PolicyNumber = modUtilities.NullToString(rs.Fields("PolicyNumber").Value)
        policy.HolderName = modUtilities.NullToString(rs.Fields("HolderName").Value)
        policy.ProductType = modUtilities.NullToString(rs.Fields("ProductType").Value)
        policy.EffectiveDate = modUtilities.NullToDate(rs.Fields("EffectiveDate").Value)
        policy.ExpirationDate = modUtilities.NullToDate(rs.Fields("ExpirationDate").Value)
        policy.Premium = modUtilities.NullToCurrency(rs.Fields("Premium").Value)
        policy.Status = modUtilities.NullToString(rs.Fields("Status").Value)
        policy.AgentCode = modUtilities.NullToString(rs.Fields("AgentCode").Value)
        results.Add policy
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetPolicies = results
End Function

Public Function GetClaims() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT ClaimNumber, PolicyNumber, LossDate, ReportedDate, Amount, Status, AdjusterName, Description FROM Claims ORDER BY LossDate DESC", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim claim As clsClaim
        Set claim = New clsClaim
        claim.ClaimNumber = modUtilities.NullToString(rs.Fields("ClaimNumber").Value)
        claim.PolicyNumber = modUtilities.NullToString(rs.Fields("PolicyNumber").Value)
        claim.LossDate = modUtilities.NullToDate(rs.Fields("LossDate").Value)
        claim.ReportedDate = modUtilities.NullToDate(rs.Fields("ReportedDate").Value)
        claim.Amount = modUtilities.NullToCurrency(rs.Fields("Amount").Value)
        claim.Status = modUtilities.NullToString(rs.Fields("Status").Value)
        claim.AdjusterName = modUtilities.NullToString(rs.Fields("AdjusterName").Value)
        claim.Description = modUtilities.NullToString(rs.Fields("Description").Value)
        results.Add claim
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetClaims = results
End Function

Public Function GetAgents() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT AgentCode, FullName, Region, Email, Phone, LicenseExpiration, Active FROM Agents ORDER BY FullName", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim agent As clsAgent
        Set agent = New clsAgent
        agent.AgentCode = modUtilities.NullToString(rs.Fields("AgentCode").Value)
        agent.FullName = modUtilities.NullToString(rs.Fields("FullName").Value)
        agent.Region = modUtilities.NullToString(rs.Fields("Region").Value)
        agent.Email = modUtilities.NullToString(rs.Fields("Email").Value)
        agent.Phone = modUtilities.NullToString(rs.Fields("Phone").Value)
        agent.LicenseExpiration = modUtilities.NullToDate(rs.Fields("LicenseExpiration").Value)
        agent.Active = (rs.Fields("Active").Value = True)
        results.Add agent
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetAgents = results
End Function

Public Function GetInvoices() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT InvoiceNumber, PolicyNumber, DueDate, AmountDue, AmountPaid, Status, LastUpdated FROM Invoices ORDER BY DueDate", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim invoice As clsInvoice
        Set invoice = New clsInvoice
        invoice.InvoiceNumber = modUtilities.NullToString(rs.Fields("InvoiceNumber").Value)
        invoice.PolicyNumber = modUtilities.NullToString(rs.Fields("PolicyNumber").Value)
        invoice.DueDate = modUtilities.NullToDate(rs.Fields("DueDate").Value)
        invoice.AmountDue = modUtilities.NullToCurrency(rs.Fields("AmountDue").Value)
        invoice.AmountPaid = modUtilities.NullToCurrency(rs.Fields("AmountPaid").Value)
        invoice.Status = modUtilities.NullToString(rs.Fields("Status").Value)
        invoice.LastUpdated = modUtilities.NullToDate(rs.Fields("LastUpdated").Value)
        results.Add invoice
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetInvoices = results
End Function

Public Function GetUnderwritingCases() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT PolicyNumber, SubmissionDate, RiskScore, AssignedTo, Status, Notes FROM UnderwritingCases ORDER BY SubmissionDate DESC", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim uw As clsUnderwritingCase
        Set uw = New clsUnderwritingCase
        uw.PolicyNumber = modUtilities.NullToString(rs.Fields("PolicyNumber").Value)
        uw.SubmissionDate = modUtilities.NullToDate(rs.Fields("SubmissionDate").Value)
        uw.RiskScore = rs.Fields("RiskScore").Value
        uw.AssignedTo = modUtilities.NullToString(rs.Fields("AssignedTo").Value)
        uw.Status = modUtilities.NullToString(rs.Fields("Status").Value)
        uw.Notes = modUtilities.NullToString(rs.Fields("Notes").Value)
        results.Add uw
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetUnderwritingCases = results
End Function

Public Function GetDocuments() As Collection
    Dim results As New Collection
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT PolicyNumber, DocumentName, DocumentType, ReceivedDate, ComplianceStatus FROM Documents ORDER BY ReceivedDate DESC", gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim doc As clsDocument
        Set doc = New clsDocument
        doc.PolicyNumber = modUtilities.NullToString(rs.Fields("PolicyNumber").Value)
        doc.DocumentName = modUtilities.NullToString(rs.Fields("DocumentName").Value)
        doc.DocumentType = modUtilities.NullToString(rs.Fields("DocumentType").Value)
        doc.ReceivedDate = modUtilities.NullToDate(rs.Fields("ReceivedDate").Value)
        doc.ComplianceStatus = modUtilities.NullToString(rs.Fields("ComplianceStatus").Value)
        results.Add doc
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set GetDocuments = results
End Function

Public Sub SavePolicy(ByVal policy As clsPolicy)
    gConn.Execute "INSERT INTO Policies (PolicyNumber, HolderName, ProductType, EffectiveDate, ExpirationDate, Premium, Status, AgentCode) VALUES (" & _
                 "'" & SqlEscape(policy.PolicyNumber) & "', " & _
                 "'" & SqlEscape(policy.HolderName) & "', " & _
                 "'" & SqlEscape(policy.ProductType) & "', " & SqlDate(policy.EffectiveDate) & ", " & SqlDate(policy.ExpirationDate) & ", " & _
                 Format$(policy.Premium, "0.00") & ", '" & SqlEscape(policy.Status) & "', '" & SqlEscape(policy.AgentCode) & "')"
    WriteAudit "Policies", policy.PolicyNumber, "Create", "Policy created via UI"
End Sub

Public Sub UpdatePolicyStatus(ByVal policyNumber As String, ByVal newStatus As String, ByVal newExpiration As Date)
    gConn.Execute "UPDATE Policies SET Status='" & SqlEscape(newStatus) & "', ExpirationDate=" & SqlDate(newExpiration) & " WHERE PolicyNumber='" & SqlEscape(policyNumber) & "'"
    WriteAudit "Policies", policyNumber, "Status Update", "Status changed to " & newStatus
End Sub

Public Sub RecordClaimProgress(ByVal claimNumber As String, ByVal newStatus As String, ByVal amount As Currency)
    gConn.Execute "UPDATE Claims SET Status='" & SqlEscape(newStatus) & "', Amount=" & Format$(amount, "0.00") & " WHERE ClaimNumber='" & SqlEscape(claimNumber) & "'"
    WriteAudit "Claims", claimNumber, "Status Update", "Claim moved to " & newStatus
End Sub

Public Sub RecordPayment(ByVal invoiceNumber As String, ByVal amountPaid As Currency, ByVal status As String)
    gConn.Execute "UPDATE Invoices SET AmountPaid=" & Format$(amountPaid, "0.00") & ", Status='" & SqlEscape(status) & "', LastUpdated=" & SqlDate(Date) & " WHERE InvoiceNumber='" & SqlEscape(invoiceNumber) & "'"
    WriteAudit "Invoices", invoiceNumber, "Payment", "Invoice updated to status " & status
End Sub

Private Sub WriteAudit(ByVal entityName As String, ByVal entityKey As String, ByVal actionName As String, ByVal details As String)
    gConn.Execute "INSERT INTO AuditTrail (EventTime, Entity, EntityKey, Action, Details) VALUES (" & SqlDate(Now) & ", '" & SqlEscape(entityName) & "', '" & SqlEscape(entityKey) & "', '" & SqlEscape(actionName) & "', '" & SqlEscape(details) & "')"
End Sub

Public Function CountByStatus(ByVal tableName As String, ByVal statusColumn As String) As Object
    Dim rs As Object
    Dim results As Object
    Set rs = CreateObject("ADODB.Recordset")
    Set results = CreateObject("Scripting.Dictionary")

    rs.Open "SELECT " & statusColumn & " AS StatusValue, COUNT(*) AS TotalCount FROM [" & tableName & "] GROUP BY " & statusColumn, gConn, AD_OPEN_STATIC, AD_LOCK_READONLY, AD_CMD_TEXT

    Do Until rs.EOF
        Dim key As String
        key = modUtilities.NullToString(rs.Fields("StatusValue").Value)
        If results.Exists(key) Then
            results(key) = results(key) + rs.Fields("TotalCount").Value
        Else
            results.Add key, rs.Fields("TotalCount").Value
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    Set CountByStatus = results
End Function

Public Function GetTotalPremium() As Currency
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT SUM(Premium) AS TotalPremium FROM Policies", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    GetTotalPremium = modUtilities.NullToCurrency(rs.Fields("TotalPremium").Value)
    rs.Close
    Set rs = Nothing
End Function

Public Function GetOpenClaimAmount() As Currency
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT SUM(Amount) AS TotalClaims FROM Claims WHERE Status <> 'Closed'", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    GetOpenClaimAmount = modUtilities.NullToCurrency(rs.Fields("TotalClaims").Value)
    rs.Close
    Set rs = Nothing
End Function

Public Function GetPastDuePremium() As Currency
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT SUM(AmountDue - AmountPaid) AS Outstanding FROM Invoices WHERE Status <> 'Paid'", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    GetPastDuePremium = modUtilities.NullToCurrency(rs.Fields("Outstanding").Value)
    rs.Close
    Set rs = Nothing
End Function

Public Function GetExpiringPolicies(ByVal cutoffDate As Date) As Long
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) AS TotalExpiring FROM Policies WHERE ExpirationDate <= #" & cutoffDate & "#", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    GetExpiringPolicies = rs.Fields("TotalExpiring").Value
    rs.Close
    Set rs = Nothing
End Function

Public Function GetAgentsWithExpiredLicenses() As Long
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) AS Expired FROM Agents WHERE LicenseExpiration < #" & Date & "#", gConn, AD_OPEN_FORWARD_ONLY, AD_LOCK_READONLY, AD_CMD_TEXT
    GetAgentsWithExpiredLicenses = rs.Fields("Expired").Value
    rs.Close
    Set rs = Nothing
End Function

Public Sub UpdateAgent(ByVal agentCode As String, ByVal active As Boolean, ByVal licenseExpiration As Date)
    Dim activeLiteral As String
    If active Then
        activeLiteral = "TRUE"
    Else
        activeLiteral = "FALSE"
    End If

    gConn.Execute "UPDATE Agents SET Active=" & activeLiteral & ", LicenseExpiration=" & SqlDate(licenseExpiration) & " WHERE AgentCode='" & SqlEscape(agentCode) & "'"
    WriteAudit "Agents", agentCode, "Update", "Agent status/license updated"
End Sub

Public Sub UpdateUnderwritingCase(ByVal policyNumber As String, ByVal newStatus As String, ByVal riskScore As Integer)
    gConn.Execute "UPDATE UnderwritingCases SET Status='" & SqlEscape(newStatus) & "', RiskScore=" & riskScore & " WHERE PolicyNumber='" & SqlEscape(policyNumber) & "'"
    WriteAudit "Underwriting", policyNumber, "Update", "Underwriting case moved to " & newStatus
End Sub

Public Sub UpdateDocumentCompliance(ByVal policyNumber As String, ByVal documentName As String, ByVal newStatus As String)
    gConn.Execute "UPDATE Documents SET ComplianceStatus='" & SqlEscape(newStatus) & "' WHERE PolicyNumber='" & SqlEscape(policyNumber) & "' AND DocumentName='" & SqlEscape(documentName) & "'"
    WriteAudit "Documents", policyNumber & "-" & documentName, "Compliance", "Document marked " & newStatus
End Sub
