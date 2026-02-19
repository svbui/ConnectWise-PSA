/*

File       :    Double_Contact_Same_Company_incl_InActiveMerge.sql
Version    :    v1.0
Author     :    https://github.com/svbui
Purpose    :    Create a report that extracts Contacts with the same e-mail address in the in the same Company. It looks at all e-mail addresses of the contacts, not just the primary. 

                Rules:
                - merge_with does NOT have to be an Active contact (Inactive_Flag = 0)
                - if multiple active → pick the Primary_Contact (Default_Flag = 1)
                - if none of the actives are primary → pick lowest Contact_RecID that is active
                - if no active contacts exist for that email within the company → pick lowest Contact_RecID

                Results:
                - PRIMARY = this row is the winner (target to keep) → merge_with = NULL
                - MERGE = this row should be merged into the winner → merge_with = <winner Contact_RecID>
                - NO_ACTIVE_TARGET = duplicates exist but none of the contacts are active → merge_with = NULL

                The output of this can be exported to CSV to be used later with the API of ConnectWise PSA ( $CWM_URL/company/contacts/$Contact_RecID?transferContactId=$merge_with )

                IMPORTANT!
                - Please keep in mind that you can not merge a contact with a contact in another Company, this is only to find duplicates that are present within the same Company.
                - This works the same as merging Contacts in the UI, please understand what this does to the information on a Contact that is to be merged
                - For this to work there are multiple API calls needed, first you need to activate the Contact_RecID marked as Primary, than you can merge and than set the Primary back to Inactive, so there are 2 calls needed for the PRIMARY with in between 1 call for each MERGE.

                Prerequisites:
                - Database Access
                    - CW PSA On-Premise login to your SQL Server
                    - CW PSA as SaaS, make sure you have CDA Access
                - SQL Tooling like SQl Server Management Studio
                - API Tooling like Postman or use any scripting tool
*/



WITH EmailData AS (
    SELECT
        c.Company_RecID,
        co.Company_Name,
        -- Normalize to avoid case/space duplicates
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
        FIRST_VALUE(ed.Contact_RecID) OVER (
            PARTITION BY ed.Company_RecID, ed.EmailAddress
            ORDER BY
                CASE WHEN ed.Contact_InactiveFlag = 0 THEN 0 ELSE 1 END,  -- Prefer Active, else Inactive
                CASE WHEN ed.Primary_Contact = 1 THEN 0 ELSE 1 END,       -- Prefer Primary
                ed.Contact_RecID ASC                                      -- Lowest ID wins
        ) AS winner_contact_recid
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

    -- merge_with: NULL for winner, winner id for rows to merge
    CASE
        WHEN ed.Contact_RecID = w.winner_contact_recid THEN NULL
        ELSE w.winner_contact_recid
    END AS merge_with,

    CASE
        WHEN ed.Contact_RecID = w.winner_contact_recid THEN 'PRIMARY'
        ELSE 'MERGE'
    END AS merge_action

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
