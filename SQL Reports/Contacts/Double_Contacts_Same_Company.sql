/*

File       :    Double_Contacts_Same_Company.sql
Version    :    v1.0
Author     :    https://github.com/svbui
Purpose    :    Create a report that extracts Contacts with the same e-mail address in the in the same Company. It looks at all e-mail addresses of the contacts, not just the primary. 

                Rules:
                - merge_with must be an Active contact (Inactive_Flag = 0)
                - if multiple active → pick the Primary_Contact (Default_Flag = 1)
                - if none of the actives are primary → pick lowest Contact_RecID
                - if no active contacts exist for that email within the company → merge_with = NULL (CW PSA does not allow to merge with an Inactive Contact)

                Results:
                - PRIMARY = this row is the winner (target to keep) → merge_with = NULL
                - MERGE = this row should be merged into the winner → merge_with = <winner Contact_RecID>
                - NO_ACTIVE_TARGET = duplicates exist but none of the contacts are active → merge_with = NULL

                The output of this can be exported to CSV to be used later with the API of ConnectWise PSA ( $CWM_URL/company/contacts/$Contact_RecID?transferContactId=$merge_with )

                IMPORTANT!
                - Please keep in mind that you can not merge a contact with a contact in another Company, this is only to find duplicates that are present within the same Company.
                - This works the same as merging Contacts in the UI, please understand what this does to the information on a Contact that is to be merged

                Prerequisites:
                - Database Access
                    - CW PSA On-Premise login to your SQL Server
                    - CW PSA as SaaS, make sure you have CDA Access
                - SQL Tooling like SQl Server Management Studio
*/



WITH EmailData AS (
    SELECT
        c.Company_RecID,
        co.Company_Name,
        LOWER(LTRIM(RTRIM(cc.Contact_Communication_Desc))) AS EmailAddress,

        c.Contact_RecID,
        c.First_Name,
        c.Last_Name,
        ISNULL(c.Inactive_Flag, 0) AS Contact_InactiveFlag,
        ISNULL(c.Default_Flag, 0)  AS Primary_Contact,

        cc.Contact_Communication_RecID
    FROM cwwebapp_dustin.dbo.v_rpt_ContactCommunication AS cc
    INNER JOIN cwwebapp_dustin.dbo.v_rpt_Contact AS c
        ON c.Contact_RecID = cc.Contact_RecID
    INNER JOIN cwwebapp_dustin.dbo.v_rpt_Company AS co
        ON co.Company_RecID = c.Company_RecID
    WHERE
        cc.Communication_Name = 'Email'
        AND cc.Contact_Communication_Desc IS NOT NULL
),
DupKeys AS (
    SELECT
        Company_RecID,
        EmailAddress
    FROM EmailData
    GROUP BY
        Company_RecID,
        EmailAddress
    HAVING COUNT(*) > 1
),
Winners AS (
    SELECT DISTINCT
        ed.Company_RecID,
        ed.EmailAddress,
        CASE
            WHEN MAX(CASE WHEN ed.Contact_InactiveFlag = 0 THEN 1 ELSE 0 END)
                 OVER (PARTITION BY ed.Company_RecID, ed.EmailAddress) = 1
            THEN
                FIRST_VALUE(ed.Contact_RecID) OVER (
                    PARTITION BY ed.Company_RecID, ed.EmailAddress
                    ORDER BY
                        CASE WHEN ed.Contact_InactiveFlag = 0 THEN 0 ELSE 1 END,  -- Active first
                        CASE WHEN ed.Primary_Contact = 1 THEN 0 ELSE 1 END,       -- Primary next
                        ed.Contact_RecID ASC                                      -- Lowest ID
                )
            ELSE NULL
        END AS merge_with
    FROM EmailData ed
)

SELECT
    ed.Company_RecID,
    ed.Company_Name,
    ed.EmailAddress,
    ed.Contact_RecID,
    ed.First_Name,
    ed.Last_Name,
    ed.Contact_InactiveFlag,
    ed.Primary_Contact,
    ed.Contact_Communication_RecID,
    CASE
        WHEN w.merge_with IS NULL THEN NULL
        WHEN ed.Contact_RecID = w.merge_with THEN 'PRIMARY'
        ELSE CAST(w.merge_with AS varchar(20))
    END AS merge_with
FROM EmailData ed
INNER JOIN DupKeys d
    ON d.Company_RecID = ed.Company_RecID
   AND d.EmailAddress  = ed.EmailAddress
INNER JOIN Winners w
    ON w.Company_RecID = ed.Company_RecID
   AND w.EmailAddress  = ed.EmailAddress
ORDER BY
    ed.Company_Name,
    ed.EmailAddress,
    ed.Contact_InactiveFlag,
    ed.Primary_Contact DESC,
    ed.Contact_RecID;
