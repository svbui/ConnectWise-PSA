/*

Created by   : Sven Buijvoets
Create date  : 26-11-2024
Version      : 1.0
Purpose      : Retrieve the number of Invoices for each company in ConnectWise PSA in the past 12 months, including customers with 0 invoices
               In order to be able to identify companies have received 0 invoices. 

Important    : This will return all companies including Vendor, Owner and others.
               All company statuses are includen except "Inactive/Gone"               

*/

WITH FilteredInvoices AS (
    SELECT 
        i.Company_RecID,
        i.Invoice_Number,
        i.Invoice_Type,
        i.Date_invoice,
        i.BusGroup,
        i.Reference,
        i.Agreement_name,
        i.AGR_Header_recID,
        COUNT(i.Invoice_Number) OVER (PARTITION BY i.Company_RecID) AS Total_Invoices,
        ROW_NUMBER() OVER (PARTITION BY i.Company_RecID ORDER BY i.Date_invoice DESC) AS rn
    FROM 
        v_rpt_Invoices i
    WHERE 
        i.Date_invoice >= DATEADD(MONTH, -12, GETDATE())
),
LastInvoices AS (
    SELECT 
        Company_RecID,
        Invoice_Number,
        Invoice_Type,
        Date_invoice,
        BusGroup,
        Reference,
        Agreement_name,
        AGR_Header_recID,
        Total_Invoices
    FROM 
        FilteredInvoices
    WHERE 
        rn = 1
)
SELECT 
    c.Company_RecID,
    c.Company_Name,
    c.Company_Type_Desc,
    c.Company_Status_Desc,
    c.Account_Nbr,
    c.Location,
    li.Invoice_Number,
    li.Invoice_Type,
    li.Date_invoice,
    li.BusGroup,
    li.Reference,
    li.Agreement_name,
    li.AGR_Header_recID,
    COALESCE(li.Total_Invoices, 0) AS Total_Invoices
FROM 
    v_rpt_Company c
LEFT JOIN 
    LastInvoices li
ON 
    c.Company_RecID = li.Company_RecID
WHERE 
    c.Company_Status_Desc <> 'Inactive/Gone'
ORDER BY 
    COALESCE(li.Total_Invoices, 0) ASC;
