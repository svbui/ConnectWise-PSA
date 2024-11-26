/*

Created by   : Sven Buijvoets
Create date  : 26-11-2024
Version      : 1.0
Purpose      : Retrieve the number of not Expired agreements for each company in ConnectWise PSA, including customers with 0 active agreements
               In order to be able to identify companies who have no Active or Inactive (future start date) agreements. 

Important    : This will return all companies including Vendor, Owner and others.
               All company statuses are included except "Inactive/Gone"               

*/

WITH FilteredCompanies AS (
    SELECT
        Company_RecID,
        Company_Name,
		Location,
        Account_Nbr,
        Company_Status_Desc
    FROM v_rpt_Company
    WHERE Company_Status_Desc NOT IN ('Inactive/Gone')
),
FilteredAgreements AS (
    SELECT
        AGR_Name,
		agr_type_desc,
        company_recid,
        datestart,
        dateend,
        Agreement_Status
    FROM v_rpt_AgreementList
    WHERE Agreement_Status <> 'Expired'
),
CompanyAgreementCounts AS (
    SELECT
        fc.Company_RecID,
        COUNT(fa.AGR_Name) AS TotalAgreements
    FROM FilteredCompanies fc
    LEFT JOIN FilteredAgreements fa
        ON fc.Company_RecID = fa.company_recid
    GROUP BY fc.Company_RecID
)
SELECT
    fc.Company_RecID,
    fc.Company_Name,
	fc.Location,
    fc.Account_Nbr AS AccountNumber,
    fc.Company_Status_Desc,
    fa.AGR_Name,
	fa.agr_type_desc,
    fa.datestart,
    fa.dateend,
    fa.Agreement_Status,
    ISNULL(cac.TotalAgreements, 0) AS TotalAgreements
FROM FilteredCompanies fc
LEFT JOIN FilteredAgreements fa
    ON fc.Company_RecID = fa.company_recid
LEFT JOIN CompanyAgreementCounts cac
    ON fc.Company_RecID = cac.Company_RecID
ORDER BY TotalAgreements ASC, fc.Company_Name;
