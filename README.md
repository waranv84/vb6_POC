# VB6 Insurance Operations Suite Proof of Concept

This repository hosts a Visual Basic 6 (VB6) proof-of-concept desktop application called **Insurance Operations Suite**. The sample solution showcases how an insurance carrier could coordinate policy administration, claims, billing, agency compliance, and underwriting workflows from a single VB6 executable while persisting data to a standalone Microsoft Access database.

## Project Highlights

- **Modular navigation dashboard** with live operational metrics, compliance alerts, and audit activity surfaced on the main form.
- **Policy Administration** form for viewing, creating, and updating commercial policies with status tracking and narrative summaries.
- **Claim Management** console for advancing claim status, adjusting reserves, and reviewing adjuster details.
- **Billing & Collections** workspace that applies partial payments, marks invoices paid, and highlights aging receivables.
- **Agency Management** screen to monitor licensing, territory assignments, and active/inactive state for producing agents.
- **Underwriting Workbench** to triage submissions, adjust risk scores, and escalate complex risks.
- **Document Compliance** center to keep audit trails of outstanding compliance documents across policies.
- **Analytics Dashboard** aggregating policy, claim, billing, and compliance distributions along with a narrative snapshot of key KPIs.

All modules share a common Access (`.mdb`) data store located in the `data` folder next to the project. On first launch the application creates the database, applies the schema (via ADOX), seeds representative sample data, and records audit events as updates occur.

## Project Structure

The VB6 project lives in the `EnterpriseSuite` directory and contains:

- Forms: `frmMain`, `frmPolicies`, `frmClaims`, `frmBilling`, `frmAgents`, `frmUnderwriting`, `frmDocuments`, `frmAnalytics`, `frmSettings`, `frmAbout`.
- Modules: `modDatabase` (Access setup and queries), `modUtilities` (logging and helpers), `modReporting` (dashboard rollups and narratives).
- Class modules: `clsPolicy`, `clsClaim`, `clsAgent`, `clsInvoice`, `clsUnderwritingCase`, `clsDocument` representing core insurance entities.
- Access data folder: `data/InsuranceSuite.mdb` (auto-generated on first run) plus `InsuranceSuite.log` for user activity traces.

## Opening the Project in VB6

1. Copy the entire `EnterpriseSuite` folder to a Windows machine with VB6 (Visual Studio 6) installed.
2. Launch the VB6 IDE and open `EnterpriseSuite.vbp`.
3. Add the following references (Project → References…) if they are not already resolved:
   - **Microsoft ActiveX Data Objects 2.6+ Library** (e.g., `MSADO28.TLB`).
   - **Microsoft ADO Ext. 2.x for DDL and Security** (`MSADOX.DLL`).
   - **Microsoft Scripting Runtime** (`SCRRUN.DLL`) for dictionary support.
4. Press **F5** to run the project or use **File → Make InsuranceOperationsSuite.exe** to build an executable.

The application automatically creates `data\InsuranceSuite.mdb` next to the executable, so no external database server is required. The seeded dataset can be refreshed by deleting the `.mdb` file before the next launch.

## Running in This Environment

The current Linux container does not include the VB6 runtime or compiler, so the `.vbp` cannot be executed or compiled here. Please follow the Windows setup steps above to explore the Insurance Operations Suite locally.
