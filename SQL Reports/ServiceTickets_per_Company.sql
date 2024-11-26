/*

Created by   : Sven Buijvoets
Create date  : 26-11-2024
Version      : 1.0
Purpose      : Retrieve the number of Service Tickets for each company in ConnectWise PSA in the past 12 months, including customers with 0 invoices
               In order to be able to identify companies that have created 0 Service Tickets. 

Important    : This will return all companies including Vendor, Owner and others.
               All company statuses are includen except "Inactive/Gone"               

*/

WITH FilteredCompanies AS (
    SELECT
        Company_RecID,
        Company_Name,
        Account_Nbr,
		Location,
        Company_Status_Desc
    FROM v_rpt_Company
    WHERE Company_Status_Desc NOT IN ('Inactive/Gone')
),
FilteredTickets AS (
    SELECT
        company_recid,
        COUNT(*) AS TicketCount
    FROM v_rpt_service
    WHERE date_entered >= DATEADD(YEAR, -1, GETDATE())
    GROUP BY company_recid
)
SELECT
    fc.Company_RecID,
    fc.Company_Name,
	fc.Location,
    fc.Account_Nbr AS AccountNumber,
    fc.Company_Status_Desc,
    ISNULL(ft.TicketCount, 0) AS TotalTickets
FROM FilteredCompanies fc
LEFT JOIN FilteredTickets ft
    ON fc.Company_RecID = ft.company_recid
ORDER BY TotalTickets ASC;
