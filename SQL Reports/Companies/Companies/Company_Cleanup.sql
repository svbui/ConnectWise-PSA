/*

File       :    Company_Cleanup.sql.sql
Version    :    v1.0
Author     :    https://github.com/svbui
Purpose    :    Create a report that identifies potentially dormant or obsolete companies and ranks them for cleanup, while still keeping recently added companies visible for validation. 

                Rules:
                - Look at Companies that are any other than Inactive (Company_Status_Desc <> 'Inactive/Gone'
                - It has no invoice dated on or after @ExcludeInvoicesAfter
                - 

                Results:
                - Recently_Added_Flag = 1: Date_Entered between @RecentStart and @RecentEnd | 0: otherwise
                - No_Invoice_Since_Stale_Flag = 1: No invoice OR last invoice < @StaleSince  | 0: Has invoice >= @StaleSince
                - No_Active_Agreements_Flag = 1: has 0 active agreements | 0 = Has active agreements
                - No_Ticket_Since_Stale_Flag =
                - No_Project_Since_Stale_Flag =
                - Clean up Bucket: 
                    - Green: No financial activity + no operational activity.
                    - Red: No billing structure + no revenue. 
                    - Yellow: Risk of lost revenue 
                - Cleanup Score:
                    - No invoice since stale	+40
                    - No active agreements	+25
                    - No ticket activity	+20
                    - No project activity	+15
                    - Recently added	−15


                Prerequisites:
                - Database Access
                    - CW PSA On-Premise login to your SQL Server
                    - CW PSA as SaaS, make sure you have CDA Access
                - SQL Tooling like SQl Server Management Studio

*/

DECLARE @ExcludeInvoicesAfter date = '2025-01-01';
DECLARE @StaleSince          date = '2024-01-01';

-- Recently added window (inclusive start, exclusive end)
DECLARE @RecentStart         date = '2025-01-01';
DECLARE @RecentEnd           date = '2027-01-01';

WITH CompanyBase AS (
    SELECT
          c.Company_RecID
        , c.Company_ID
        , c.Company_Name
        , c.Company_Type_Desc
        , c.Company_Status_Desc
        , c.Account_Nbr
        , c.Location        AS Company_TerritoryOrCountry
        , c.Date_Entered    AS Company_Date_Entered
        , c.Last_Update_UTC AS Company_Last_Updated_UTC
    FROM cwwebapp_dustin.dbo.v_rpt_Company AS c
    WHERE c.Company_Status_Desc <> 'Inactive/Gone'
      AND NOT EXISTS (
          SELECT 1
          FROM cwwebapp_dustin.dbo.v_rpt_Invoices i
          WHERE i.Company_RecID = c.Company_RecID
            AND i.Date_Invoice >= @ExcludeInvoicesAfter
      )
),
InvoiceAgg AS (
    SELECT
          i.Company_RecID
        , MAX(i.Date_Invoice) AS Last_Invoice_Date
    FROM cwwebapp_dustin.dbo.v_rpt_Invoices i
    GROUP BY i.Company_RecID
),
LastInvoiceDetail AS (
    SELECT
          x.Company_RecID
        , x.Date_Invoice   AS Last_Invoice_Date
        , x.Invoice_Number AS Last_Invoice_Number
    FROM (
        SELECT
              i.Company_RecID
            , i.Date_Invoice
            , i.Invoice_Number
            , ROW_NUMBER() OVER (
                PARTITION BY i.Company_RecID
                ORDER BY i.Date_Invoice DESC, i.Invoice_Number DESC
              ) AS rn
        FROM cwwebapp_dustin.dbo.v_rpt_Invoices i
    ) x
    WHERE x.rn = 1
),
ActiveAgreements AS (
    SELECT
          a.Company_RecID
        , COUNT(*) AS Active_Agreements
    FROM cwwebapp_dustin.dbo.v_rpt_AgreementList a
    WHERE a.Agreement_Status = 'Active'
      AND (a.DateEnd IS NULL OR a.DateEnd >= CAST(GETDATE() AS date))
    GROUP BY a.Company_RecID
),
LastTicketUpdate AS (
    SELECT
          s.Company_RecID
        , MAX(s.Last_Update) AS Last_Ticket_Updated
    FROM cwwebapp_dustin.dbo.v_rpt_Service s
    GROUP BY s.Company_RecID
),
LastProjectUpdate AS (
    SELECT
          s.Company_RecID
        , MAX(p.Last_Update) AS Last_Project_Updated
    FROM cwwebapp_dustin.dbo.v_rpt_Project p
    INNER JOIN cwwebapp_dustin.dbo.v_rpt_Service s
        ON s.SR_Service_RecID = p.SR_Service_RecID
    GROUP BY s.Company_RecID
),

/* One place to compute flags so we don’t repeat CASE logic everywhere */
Flags AS (
    SELECT
          cb.*
        , ia.Last_Invoice_Date
        , lid.Last_Invoice_Number
        , ltu.Last_Ticket_Updated
        , lpu.Last_Project_Updated
        , ISNULL(aa.Active_Agreements, 0) AS Active_Agreements

        , CASE
            WHEN cb.Company_Date_Entered >= @RecentStart
             AND cb.Company_Date_Entered <  @RecentEnd
            THEN 1 ELSE 0
          END AS Recently_Added_Flag

        , CASE
            WHEN (ia.Last_Invoice_Date IS NULL OR ia.Last_Invoice_Date < @StaleSince)
            THEN 1 ELSE 0
          END AS No_Invoice_Since_Stale_Flag

        , CASE
            WHEN (ltu.Last_Ticket_Updated IS NULL OR ltu.Last_Ticket_Updated < @StaleSince)
            THEN 1 ELSE 0
          END AS No_Ticket_Since_Stale_Flag

        , CASE
            WHEN (lpu.Last_Project_Updated IS NULL OR lpu.Last_Project_Updated < @StaleSince)
            THEN 1 ELSE 0
          END AS No_Project_Since_Stale_Flag

        , CASE
            WHEN ISNULL(aa.Active_Agreements, 0) = 0
            THEN 1 ELSE 0
          END AS No_Active_Agreements_Flag

    FROM CompanyBase cb
    LEFT JOIN InvoiceAgg         ia  ON ia.Company_RecID  = cb.Company_RecID
    LEFT JOIN LastInvoiceDetail  lid ON lid.Company_RecID = cb.Company_RecID
    LEFT JOIN ActiveAgreements   aa  ON aa.Company_RecID  = cb.Company_RecID
    LEFT JOIN LastTicketUpdate   ltu ON ltu.Company_RecID = cb.Company_RecID
    LEFT JOIN LastProjectUpdate  lpu ON lpu.Company_RecID = cb.Company_RecID
)

SELECT
      f.Company_RecID
    , f.Company_ID
    , f.Company_Name
    , f.Company_Type_Desc
    , f.Company_Status_Desc            AS Company_Status
    , f.Account_Nbr                    AS Company_Account_Number
    , f.Company_TerritoryOrCountry     AS Company_Territory_Country
    , f.Company_Date_Entered

    , CASE WHEN f.Recently_Added_Flag = 1 THEN 'Recently Added' ELSE 'Not Recent' END AS Recently_Added
    , f.Recently_Added_Flag

    , f.Company_Last_Updated_UTC

    , f.Last_Invoice_Date              AS Last_Invoice_Sent_Date
    , f.Last_Invoice_Number            AS Last_Invoice_Number

    , f.Active_Agreements
    , f.Last_Ticket_Updated
    , f.Last_Project_Updated

    /* ---- Bucket assignment (priority: GREEN > RED > YELLOW > OK) ---- */
    , CASE
        WHEN f.No_Invoice_Since_Stale_Flag = 1
         AND f.No_Ticket_Since_Stale_Flag = 1
         AND f.No_Project_Since_Stale_Flag = 1
        THEN 'GREEN - Zombie (no invoice/ticket/project since ' + CONVERT(varchar(10), @StaleSince, 120) + ')'

        WHEN f.No_Active_Agreements_Flag = 1
         AND f.No_Invoice_Since_Stale_Flag = 1
        THEN 'RED - 0 active agreements + no invoice since ' + CONVERT(varchar(10), @StaleSince, 120)

        WHEN f.No_Active_Agreements_Flag = 0
         AND f.No_Invoice_Since_Stale_Flag = 1
        THEN 'YELLOW - Active agreements but no invoice since ' + CONVERT(varchar(10), @StaleSince, 120)

        ELSE 'OK / Needs review'
      END AS Cleanup_Bucket

    /* ---- Cleanup score (0-100), higher = more likely cleanup candidate ---- */
    , (
          (CASE WHEN f.No_Invoice_Since_Stale_Flag = 1 THEN 40 ELSE 0 END)
        + (CASE WHEN f.No_Active_Agreements_Flag  = 1 THEN 25 ELSE 0 END)
        + (CASE WHEN f.No_Ticket_Since_Stale_Flag = 1 THEN 20 ELSE 0 END)
        + (CASE WHEN f.No_Project_Since_Stale_Flag= 1 THEN 15 ELSE 0 END)
        - (CASE WHEN f.Recently_Added_Flag        = 1 THEN 15 ELSE 0 END)
      ) AS Cleanup_Score

    , CASE
        WHEN (
              (CASE WHEN f.No_Invoice_Since_Stale_Flag = 1 THEN 40 ELSE 0 END)
            + (CASE WHEN f.No_Active_Agreements_Flag  = 1 THEN 25 ELSE 0 END)
            + (CASE WHEN f.No_Ticket_Since_Stale_Flag = 1 THEN 20 ELSE 0 END)
            + (CASE WHEN f.No_Project_Since_Stale_Flag= 1 THEN 15 ELSE 0 END)
            - (CASE WHEN f.Recently_Added_Flag        = 1 THEN 15 ELSE 0 END)
          ) >= 70
        THEN 'Strong candidate: set Inactive/Gone (after quick check)'

        WHEN (
              (CASE WHEN f.No_Invoice_Since_Stale_Flag = 1 THEN 40 ELSE 0 END)
            + (CASE WHEN f.No_Active_Agreements_Flag  = 1 THEN 25 ELSE 0 END)
            + (CASE WHEN f.No_Ticket_Since_Stale_Flag = 1 THEN 20 ELSE 0 END)
            + (CASE WHEN f.No_Project_Since_Stale_Flag= 1 THEN 15 ELSE 0 END)
            - (CASE WHEN f.Recently_Added_Flag        = 1 THEN 15 ELSE 0 END)
          ) BETWEEN 45 AND 69
        THEN 'Review: likely cleanup / confirm ownership + billing'

        WHEN (
              (CASE WHEN f.No_Invoice_Since_Stale_Flag = 1 THEN 40 ELSE 0 END)
            + (CASE WHEN f.No_Active_Agreements_Flag  = 1 THEN 25 ELSE 0 END)
            + (CASE WHEN f.No_Ticket_Since_Stale_Flag = 1 THEN 20 ELSE 0 END)
            + (CASE WHEN f.No_Project_Since_Stale_Flag= 1 THEN 15 ELSE 0 END)
            - (CASE WHEN f.Recently_Added_Flag        = 1 THEN 15 ELSE 0 END)
          ) BETWEEN 25 AND 44
        THEN 'Low priority: keep, revisit later'

        ELSE 'Keep'
      END AS Cleanup_Recommendation

    /* Helpful raw flags for filtering in BI */
    , f.No_Invoice_Since_Stale_Flag
    , f.No_Active_Agreements_Flag
    , f.No_Ticket_Since_Stale_Flag
    , f.No_Project_Since_Stale_Flag

FROM Flags f
ORDER BY
    Cleanup_Score DESC,
    Cleanup_Bucket,
    f.Company_Name;
