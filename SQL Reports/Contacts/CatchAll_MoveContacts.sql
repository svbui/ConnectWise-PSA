USE cwwebapp_contoso;   -- TODO: replace with your Database name
GO

DECLARE @CatchAllCompanyRecID INT = 20130;   -- TODO: replace with your CatchAllCompany Company_RecID
DECLARE @OwnCompanyRecID      INT = 99999;   -- TODO: replace with the Company_RecID to use when Own Company domain matches




/*
File       :    CatchAll_MoveContacts.sql
Version    :    v1.0
Author     :    https://github.com/svbui
Purpose    :    Create a report that identifies Contacts currently linked to the Catch All Company
                and suggests which Company they most likely belong to, based on the domain part of
                their email address.

                The output of this can be reviewed and later exported to CSV to be used with the
                ConnectWise PSA API for moving Contacts to another Company.

                IMPORTANT!
                - This query does NOT move Contacts, it only prepares the data for review and export.
                - The API requires:
                    - Contact_RecID
                    - MoveToCompanyRecID
                    - MoveToCompanyAddressRecID
                - The selected Company Address is determined from dbo.v_rpt_CompanyAddress:
                    - first prefer Default_Flag = 1
                    - if no default exists, prefer active addresses
                    - if still multiple exist, use the lowest Company_Address_RecID
                - Normal domain matching supports historical '_old' domain cleanup:
                    - example: domain.no_old is normalized to domain.no
                    - this normalization is only used for normal matching
                    - Vendor / Own Company / Public / Do Not Move matching uses the raw EmailDomain

                Rules:
                - Own Company domains are force-mapped to @OwnCompanyRecID
                - Do Not Move domains are never moved automatically
                - Public domains are never moved automatically
                - Vendor / 3rd Party domains only match Companies where:
                    Company_Type_Desc LIKE '%Vendor%'
                - Normal domains match to Companies using NormalizedEmailDomain
                - If the best normal match has fewer than 3 matching Contacts on the candidate Company,
                  MatchResult becomes:
                    Potential Vendor / 3rd Party
                  and ReadyForApiUpdateFlag is set to 0
                - For multiple matches, the best suggestion is based on:
                    1. highest CandidateContactCountOnCompany
                    2. Active company preferred
                    3. lowest Company_RecID

                Results:
                - Own Company Match
                    Domain belongs to your own organisation and is force-mapped to @OwnCompanyRecID
                - Do Not Move Domain
                    Domain is protected and should not be moved
                - Public Domain
                    Domain belongs to a shared mailbox provider and should not be moved automatically
                - Vendor Match - Active / Not Active
                    Vendor domain matched to exactly one Vendor company
                - Vendor Multiple Match
                    Vendor domain matched multiple Vendor companies
                - Single Match - Active / Not Active
                    Normal domain matched exactly one company
                - Multiple Match
                    Normal domain matched multiple companies
                - No Match
                    No candidate company found
                - Potential Vendor / 3rd Party
                    Normal domain found a weak candidate match (< 3 matching contacts on best candidate)

                Most important fields for API export:
                - Contact_RecID
                - MoveToCompanyRecID
                - MoveToCompanyAddressRecID

                Variables to set:
                - @CatchAllCompanyRecID
                    The Company_RecID of the Catch All Company
                - @OwnCompanyRecID
                    The Company_RecID to use for Own Company domain matches

                Domain lists to maintain inside the query:
                - Own Company domains
                - Do Not Move domains
                - Public domains
                - Vendor / 3rd Party domains

                Review guidance:
                - ReadyForApiUpdateFlag = 1 means the row is considered suitable for API move prep
                - SuggestedMoveToCompanyRecID is the best review suggestion when a row is not API-ready
                - MoveToCompanyRecID is the intended move target for approved rows
                - Rows marked Potential Vendor / 3rd Party should be manually reviewed first

                Prerequisites:
                - Database Access
                    - CW PSA On-Premise login to your SQL Server
                    - CW PSA as SaaS, make sure you have CDA Access
                - SQL Tooling like SQL Server Management Studio
                - API Tooling like Postman or any scripting tool

                Typical workflow:
                1. Set @CatchAllCompanyRecID
                2. Set @OwnCompanyRecID
                3. Maintain the domain lists
                4. Run the query
                5. Review MatchType, MatchResult, ConfidenceScore, CompanyType, and CompanyStatus
                6. Export approved rows for API use with:
                    - Contact_RecID
                    - MoveToCompanyRecID
                    - MoveToCompanyAddressRecID
*/





SET NOCOUNT ON;

IF OBJECT_ID('tempdb..#CatchAllContacts')       IS NOT NULL DROP TABLE #CatchAllContacts;
IF OBJECT_ID('tempdb..#OtherCompanyDomains')    IS NOT NULL DROP TABLE #OtherCompanyDomains;
IF OBJECT_ID('tempdb..#CompanyAddressChoice')   IS NOT NULL DROP TABLE #CompanyAddressChoice;
IF OBJECT_ID('tempdb..#DomainOverallCounts')    IS NOT NULL DROP TABLE #DomainOverallCounts;
IF OBJECT_ID('tempdb..#DomainCompanyCounts')    IS NOT NULL DROP TABLE #DomainCompanyCounts;
IF OBJECT_ID('tempdb..#BestMultipleMatch')      IS NOT NULL DROP TABLE #BestMultipleMatch;

------------------------------------------------------------
-- Step 1: CatchAll contacts with parsed domain + flags
------------------------------------------------------------
SELECT
    c.Contact_RecID,
    c.Company_RecID AS CurrentCompanyRecID,
    co.Company_Name AS CurrentCompanyName,
    co.Location AS CurrentCompanyCountry,
    c.First_Name,
    c.Last_Name,
    c.Inactive_Flag AS ContactInactiveFlag,
    cc.Contact_Communication_RecID,
    cc.Contact_Communication_Desc AS EmailAddress,
    d.EmailDomain,
    d.NormalizedEmailDomain,

    IsOwnCompanyDomain =
        CASE
            WHEN d.EmailDomain IN (
                -- TODO: maintain this list of your own/internal company domains
            'contoso.com',
            'contoso.eu'
            ) THEN 1 ELSE 0
        END,

    IsDoNotMoveDomain =
        CASE
            WHEN d.EmailDomain IN (
                -- TODO: maintain this list of domains that should never be moved
                'haveibeenpwned.com'
            ) THEN 1 ELSE 0
        END,

    IsPublicDomain =
        CASE
            WHEN d.EmailDomain IN (
                'aol.com',
                'bredband.se',
                'bredband2.se',
                'caiway.nl',
                'gmail.com',
                'foxmail.com',
                'googlemail.com',
                'hetnet.nl',
                'home.nl',
                'hotmail.com',
                'hotmail.dk',
                'hotmail.fi',
                'hotmail.fr',
                'hotmail.nl',
                'hotmail.no',
                'hotmail.se',
                'icloud.com',
                'kpnmail.nl',
                'kpnplanet.nl',
                'live.co.uk',
                'live.com',
                'live.dk',
                'live.fi',
                'live.nl',
                'live.no',
                'live.se',
                'mac.com',
                'mail.dk',
                'mail.ru',
                'me.com',
                'msn.com',
                'online.no',
                'orange.com',
                'orange.fr',
                'outlook.com',
                'outlook.dk',
                'outlook.fi',
                'outlook.nl',
                'outlook.no',
                'outlook.se',
                'planet.nl',
                'proton.me',
                'protonmail.com',
                'quikcnet.nl',
                'tele2.com',
                'tele2.dk',
                'tele2.fi',
                'tele2.nl',
                'tele2.no',
                'tele2.se',
                'upcmail.nl',
                'xs4all.nl',
                'yahoo.com',
                'yahoo.dk',
                'yahoo.fi',
                'yahoo.nl',
                'yahoo.no',
                'yahoo.se'
            ) THEN 1 ELSE 0
        END,

    IsVendorDomain =
        CASE
            WHEN d.EmailDomain IN (
                -- TODO: maintain this list of vendor / distributor / OEM / partner domains
                'brightgauge.com',
                'connectwise.com'               
            ) THEN 1 ELSE 0
        END
INTO #CatchAllContacts
FROM dbo.v_rpt_ContactCommunication AS cc
INNER JOIN dbo.v_rpt_Contact AS c
    ON c.Contact_RecID = cc.Contact_RecID
INNER JOIN dbo.v_rpt_Company AS co
    ON co.Company_RecID = c.Company_RecID
CROSS APPLY (
    SELECT
        EmailDomain = LOWER(LTRIM(RTRIM(
            SUBSTRING(
                cc.Contact_Communication_Desc,
                CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                LEN(cc.Contact_Communication_Desc)
            )
        ))),
        NormalizedEmailDomain =
            CASE
                WHEN LOWER(LTRIM(RTRIM(
                    SUBSTRING(
                        cc.Contact_Communication_Desc,
                        CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                        LEN(cc.Contact_Communication_Desc)
                    )
                ))) LIKE '%_old'
                THEN LEFT(
                    LOWER(LTRIM(RTRIM(
                        SUBSTRING(
                            cc.Contact_Communication_Desc,
                            CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                            LEN(cc.Contact_Communication_Desc)
                        )
                    ))),
                    LEN(LOWER(LTRIM(RTRIM(
                        SUBSTRING(
                            cc.Contact_Communication_Desc,
                            CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                            LEN(cc.Contact_Communication_Desc)
                        )
                    )))) - 4
                )
                ELSE LOWER(LTRIM(RTRIM(
                    SUBSTRING(
                        cc.Contact_Communication_Desc,
                        CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                        LEN(cc.Contact_Communication_Desc)
                    )
                )))
            END
) AS d
WHERE
    c.Company_RecID = @CatchAllCompanyRecID
    AND cc.Contact_Communication_Desc IS NOT NULL
    AND cc.Contact_Communication_Desc LIKE '%@%'
    AND LOWER(cc.Contact_Communication_Desc) NOT LIKE '%_old@%';

CREATE INDEX IX_CatchAllContacts_ContactRecID
    ON #CatchAllContacts (Contact_RecID);

CREATE INDEX IX_CatchAllContacts_EmailDomain
    ON #CatchAllContacts (EmailDomain);

CREATE INDEX IX_CatchAllContacts_NormalizedEmailDomain
    ON #CatchAllContacts (NormalizedEmailDomain);

------------------------------------------------------------
-- Step 2: Aggregate other companies by domain
------------------------------------------------------------
SELECT
    c.Company_RecID,
    co.Company_Name,
    co.Company_Status_Desc,
    co.Company_Type_Desc,
    co.Location AS CompanyCountry,
    d.EmailDomain,
    d.NormalizedEmailDomain,
    COUNT(*) AS CandidateContactCountOnCompany
INTO #OtherCompanyDomains
FROM dbo.v_rpt_ContactCommunication AS cc
INNER JOIN dbo.v_rpt_Contact AS c
    ON c.Contact_RecID = cc.Contact_RecID
INNER JOIN dbo.v_rpt_Company AS co
    ON co.Company_RecID = c.Company_RecID
CROSS APPLY (
    SELECT
        EmailDomain = LOWER(LTRIM(RTRIM(
            SUBSTRING(
                cc.Contact_Communication_Desc,
                CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                LEN(cc.Contact_Communication_Desc)
            )
        ))),
        NormalizedEmailDomain =
            CASE
                WHEN LOWER(LTRIM(RTRIM(
                    SUBSTRING(
                        cc.Contact_Communication_Desc,
                        CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                        LEN(cc.Contact_Communication_Desc)
                    )
                ))) LIKE '%_old'
                THEN LEFT(
                    LOWER(LTRIM(RTRIM(
                        SUBSTRING(
                            cc.Contact_Communication_Desc,
                            CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                            LEN(cc.Contact_Communication_Desc)
                        )
                    ))),
                    LEN(LOWER(LTRIM(RTRIM(
                        SUBSTRING(
                            cc.Contact_Communication_Desc,
                            CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                            LEN(cc.Contact_Communication_Desc)
                        )
                    )))) - 4
                )
                ELSE LOWER(LTRIM(RTRIM(
                    SUBSTRING(
                        cc.Contact_Communication_Desc,
                        CHARINDEX('@', cc.Contact_Communication_Desc) + 1,
                        LEN(cc.Contact_Communication_Desc)
                    )
                )))
            END
) AS d
WHERE
    c.Company_RecID <> @CatchAllCompanyRecID
    AND cc.Contact_Communication_Desc IS NOT NULL
    AND cc.Contact_Communication_Desc LIKE '%@%'
    AND LOWER(cc.Contact_Communication_Desc) NOT LIKE '%_old@%'
GROUP BY
    c.Company_RecID,
    co.Company_Name,
    co.Company_Status_Desc,
    co.Company_Type_Desc,
    co.Location,
    d.EmailDomain,
    d.NormalizedEmailDomain;

CREATE INDEX IX_OtherCompanyDomains_EmailDomain
    ON #OtherCompanyDomains (EmailDomain);

CREATE INDEX IX_OtherCompanyDomains_NormalizedEmailDomain
    ON #OtherCompanyDomains (NormalizedEmailDomain);

CREATE INDEX IX_OtherCompanyDomains_CompanyRecID
    ON #OtherCompanyDomains (Company_RecID);

------------------------------------------------------------
-- Step 3: Choose one address per company
-- Prefer Default_Flag = 1, then active, then lowest RecID
------------------------------------------------------------
SELECT
    ranked.Company_RecID,
    ranked.Company_Address_RecID,
    ranked.Site_Name,
    ranked.Default_Flag,
    ranked.Inactive_Flag
INTO #CompanyAddressChoice
FROM (
    SELECT
        ca.Company_RecID,
        ca.Company_Address_RecID,
        ca.Site_Name,
        ca.Default_Flag,
        ca.Inactive_Flag,
        ROW_NUMBER() OVER (
            PARTITION BY ca.Company_RecID
            ORDER BY
                CASE WHEN ca.Default_Flag = 1 THEN 0 ELSE 1 END,
                CASE WHEN ca.Inactive_Flag = 0 THEN 0 ELSE 1 END,
                ca.Company_Address_RecID ASC
        ) AS rn
    FROM dbo.v_rpt_CompanyAddress AS ca
) AS ranked
WHERE ranked.rn = 1;

CREATE INDEX IX_CompanyAddressChoice_CompanyRecID
    ON #CompanyAddressChoice (Company_RecID);

------------------------------------------------------------
-- Step 4: Overall domain counts
------------------------------------------------------------
SELECT
    MatchDomain,
    COUNT(DISTINCT Company_RecID) AS DomainCompanyCountAcrossAllCompanies,
    SUM(CandidateContactCountOnCompany) AS DomainContactCountAcrossAllCompanies
INTO #DomainOverallCounts
FROM (
    SELECT
        Company_RecID,
        CandidateContactCountOnCompany,
        EmailDomain AS MatchDomain
    FROM #OtherCompanyDomains
    WHERE ISNULL(Company_Type_Desc, '') LIKE '%Vendor%'

    UNION ALL

    SELECT
        Company_RecID,
        CandidateContactCountOnCompany,
        NormalizedEmailDomain AS MatchDomain
    FROM #OtherCompanyDomains
) AS x
GROUP BY MatchDomain;

CREATE INDEX IX_DomainOverallCounts_MatchDomain
    ON #DomainOverallCounts (MatchDomain);

------------------------------------------------------------
-- Step 5: Match CatchAll contacts to candidate companies
-- Vendor domains match exact EmailDomain and only Vendor companies
-- Normal domains match NormalizedEmailDomain
------------------------------------------------------------
SELECT
    cac.Contact_RecID,
    cac.EmailDomain,
    cac.NormalizedEmailDomain,
    ocd.Company_RecID AS CandidateCompanyRecID,
    ocd.Company_Name AS CandidateCompanyName,
    ocd.Company_Status_Desc AS CandidateCompanyStatus,
    ocd.Company_Type_Desc AS CandidateCompanyType,
    ocd.CompanyCountry AS CandidateCompanyCountry,
    ocd.CandidateContactCountOnCompany,
    addr.Company_Address_RecID AS CandidateCompanyAddressRecID
INTO #DomainCompanyCounts
FROM #CatchAllContacts AS cac
INNER JOIN #OtherCompanyDomains AS ocd
    ON (
        (cac.IsVendorDomain = 1 AND ocd.EmailDomain = cac.EmailDomain)
        OR
        (cac.IsVendorDomain = 0 AND ocd.NormalizedEmailDomain = cac.NormalizedEmailDomain)
    )
LEFT JOIN #CompanyAddressChoice AS addr
    ON addr.Company_RecID = ocd.Company_RecID
WHERE
    cac.IsOwnCompanyDomain = 0
    AND cac.IsDoNotMoveDomain = 0
    AND cac.IsPublicDomain = 0
    AND (
        (cac.IsVendorDomain = 1 AND ISNULL(ocd.Company_Type_Desc, '') LIKE '%Vendor%')
        OR
        (cac.IsVendorDomain = 0)
    );

CREATE INDEX IX_DomainCompanyCounts_ContactRecID
    ON #DomainCompanyCounts (Contact_RecID);

CREATE INDEX IX_DomainCompanyCounts_ContactRecID_CompanyRecID
    ON #DomainCompanyCounts (Contact_RecID, CandidateCompanyRecID);

------------------------------------------------------------
-- Step 6: Best suggestion for multiple matches
------------------------------------------------------------
SELECT
    ranked.Contact_RecID,
    ranked.CandidateCompanyRecID AS SuggestedMoveToCompanyRecID,
    ranked.CandidateCompanyName AS SuggestedCompanyName,
    ranked.CandidateCompanyStatus AS SuggestedCompanyStatus,
    ranked.CandidateCompanyType AS SuggestedCompanyType,
    ranked.CandidateCompanyCountry AS SuggestedCompanyCountry,
    ranked.CandidateCompanyAddressRecID AS SuggestedMoveToCompanyAddressRecID,
    ranked.CandidateContactCountOnCompany AS SuggestedCompanyHitCount
INTO #BestMultipleMatch
FROM (
    SELECT
        dcc.Contact_RecID,
        dcc.CandidateCompanyRecID,
        dcc.CandidateCompanyName,
        dcc.CandidateCompanyStatus,
        dcc.CandidateCompanyType,
        dcc.CandidateCompanyCountry,
        dcc.CandidateCompanyAddressRecID,
        dcc.CandidateContactCountOnCompany,
        ROW_NUMBER() OVER (
            PARTITION BY dcc.Contact_RecID
            ORDER BY
                dcc.CandidateContactCountOnCompany DESC,
                CASE WHEN dcc.CandidateCompanyStatus IN ('Active', 'ACTIVE') THEN 0 ELSE 1 END,
                dcc.CandidateCompanyRecID ASC
        ) AS rn
    FROM #DomainCompanyCounts AS dcc
) AS ranked
WHERE ranked.rn = 1;

CREATE INDEX IX_BestMultipleMatch_ContactRecID
    ON #BestMultipleMatch (Contact_RecID);

------------------------------------------------------------
-- Step 7: Final result
------------------------------------------------------------
WITH MatchSummary AS (
    SELECT
        cac.Contact_RecID,
        cac.CurrentCompanyRecID,
        cac.CurrentCompanyName,
        cac.CurrentCompanyCountry,
        cac.First_Name,
        cac.Last_Name,
        cac.ContactInactiveFlag,
        cac.Contact_Communication_RecID,
        cac.EmailAddress,
        cac.EmailDomain,
        cac.NormalizedEmailDomain,
        cac.IsOwnCompanyDomain,
        cac.IsDoNotMoveDomain,
        cac.IsPublicDomain,
        cac.IsVendorDomain,
        COUNT(DISTINCT dcc.CandidateCompanyRecID) AS UniqueMatchingCompanyCount,
        MIN(dcc.CandidateCompanyRecID) AS SingleMatchCompanyRecID
    FROM #CatchAllContacts AS cac
    LEFT JOIN #DomainCompanyCounts AS dcc
        ON dcc.Contact_RecID = cac.Contact_RecID
    GROUP BY
        cac.Contact_RecID,
        cac.CurrentCompanyRecID,
        cac.CurrentCompanyName,
        cac.CurrentCompanyCountry,
        cac.First_Name,
        cac.Last_Name,
        cac.ContactInactiveFlag,
        cac.Contact_Communication_RecID,
        cac.EmailAddress,
        cac.EmailDomain,
        cac.NormalizedEmailDomain,
        cac.IsOwnCompanyDomain,
        cac.IsDoNotMoveDomain,
        cac.IsPublicDomain,
        cac.IsVendorDomain
),
SingleMatchDetails AS (
    SELECT
        ms.Contact_RecID,
        dcc.CandidateCompanyRecID,
        dcc.CandidateCompanyName,
        dcc.CandidateCompanyStatus,
        dcc.CandidateCompanyType,
        dcc.CandidateCompanyCountry,
        dcc.CandidateCompanyAddressRecID,
        dcc.CandidateContactCountOnCompany
    FROM MatchSummary AS ms
    INNER JOIN #DomainCompanyCounts AS dcc
        ON dcc.Contact_RecID = ms.Contact_RecID
       AND dcc.CandidateCompanyRecID = ms.SingleMatchCompanyRecID
    WHERE ms.UniqueMatchingCompanyCount = 1
),
CandidateArrays AS (
    SELECT
        dcc.Contact_RecID,
        '[' + STRING_AGG(CAST(dcc.CandidateCompanyRecID AS VARCHAR(20)), ',')
            WITHIN GROUP (ORDER BY dcc.CandidateCompanyRecID) + ']' AS MoveToCompanyRecIDArray
    FROM (
        SELECT DISTINCT
            Contact_RecID,
            CandidateCompanyRecID
        FROM #DomainCompanyCounts
    ) AS dcc
    GROUP BY dcc.Contact_RecID
),
VendorDomainCounts AS (
    SELECT
        cac.EmailDomain,
        COUNT(*) AS VendorDomainContactCountInCatchAll
    FROM #CatchAllContacts AS cac
    WHERE cac.IsVendorDomain = 1
    GROUP BY cac.EmailDomain
),
OwnCompanyAddress AS (
    SELECT
        addr.Company_RecID,
        addr.Company_Address_RecID
    FROM #CompanyAddressChoice AS addr
    WHERE addr.Company_RecID = @OwnCompanyRecID
)

SELECT
    ms.Contact_RecID,
    ms.CurrentCompanyRecID,
    ms.CurrentCompanyName,
    ms.CurrentCompanyCountry,
    ms.First_Name,
    ms.Last_Name,
    ms.ContactInactiveFlag,
    ms.Contact_Communication_RecID,
    ms.EmailAddress,
    ms.EmailDomain,
    ms.NormalizedEmailDomain,
    ms.IsOwnCompanyDomain,
    ms.IsDoNotMoveDomain,
    ms.IsPublicDomain,
    ms.IsVendorDomain,
    ms.UniqueMatchingCompanyCount,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 'Own Company Match'
        WHEN ms.IsDoNotMoveDomain = 1 THEN 'Do Not Move Domain'
        WHEN ms.IsPublicDomain = 1 THEN 'Public Domain'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 0 THEN 'Vendor / 3rd Party Domain'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1
             AND smd.CandidateCompanyStatus IN ('Active', 'ACTIVE') THEN 'Vendor Match - Active'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN 'Vendor Match - Not Active'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount > 1 THEN 'Vendor Multiple Match'
        WHEN ms.UniqueMatchingCompanyCount = 0 THEN 'No Match'
        WHEN ms.UniqueMatchingCompanyCount = 1
             AND smd.CandidateCompanyStatus IN ('Active', 'ACTIVE') THEN 'Single Match - Active'
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN 'Single Match - Not Active'
        ELSE 'Multiple Match'
    END AS MatchType,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN @OwnCompanyRecID
        WHEN ms.IsDoNotMoveDomain = 1 THEN NULL
        WHEN ms.IsPublicDomain = 1 THEN NULL
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN ms.SingleMatchCompanyRecID
        WHEN ms.IsVendorDomain = 0 AND ms.UniqueMatchingCompanyCount = 1 THEN ms.SingleMatchCompanyRecID
        ELSE NULL
    END AS MoveToCompanyRecID,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN oca.Company_Address_RecID
        WHEN ms.IsDoNotMoveDomain = 1 THEN NULL
        WHEN ms.IsPublicDomain = 1 THEN NULL
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyAddressRecID
        WHEN ms.IsVendorDomain = 0 AND ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyAddressRecID
        ELSE NULL
    END AS MoveToCompanyAddressRecID,

    CASE
        WHEN ms.IsOwnCompanyDomain = 0
         AND ms.IsDoNotMoveDomain = 0
         AND ms.IsPublicDomain = 0
         AND ms.UniqueMatchingCompanyCount > 1
        THEN ca.MoveToCompanyRecIDArray
        ELSE NULL
    END AS MoveToCompanyRecIDArray,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 'Own Company Domain Override'
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyName
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedCompanyName
        ELSE NULL
    END AS SuggestedCompanyName,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN @OwnCompanyRecID
        WHEN ms.IsDoNotMoveDomain = 1 THEN NULL
        WHEN ms.IsPublicDomain = 1 THEN NULL
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN ms.SingleMatchCompanyRecID
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedMoveToCompanyRecID
        ELSE NULL
    END AS SuggestedMoveToCompanyRecID,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN oca.Company_Address_RecID
        WHEN ms.IsDoNotMoveDomain = 1 THEN NULL
        WHEN ms.IsPublicDomain = 1 THEN NULL
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyAddressRecID
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedMoveToCompanyAddressRecID
        ELSE NULL
    END AS SuggestedMoveToCompanyAddressRecID,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN ms.CurrentCompanyCountry
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyCountry
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedCompanyCountry
        ELSE NULL
    END AS CompanyCountry,

    CASE
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyStatus
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedCompanyStatus
        ELSE NULL
    END AS CompanyStatus,

    CASE
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateCompanyType
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedCompanyType
        ELSE NULL
    END AS CompanyType,

    CASE
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateContactCountOnCompany
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN bmm.SuggestedCompanyHitCount
        ELSE NULL
    END AS CandidateContactCountOnChosenCompany,

    doc.DomainCompanyCountAcrossAllCompanies,
    doc.DomainContactCountAcrossAllCompanies,
    vdc.VendorDomainContactCountInCatchAll,

    CASE
        WHEN ms.IsVendorDomain = 1
         AND ISNULL(vdc.VendorDomainContactCountInCatchAll, 0) >= 5
            THEN 'Create Vendor company candidate'
        WHEN ms.IsVendorDomain = 1
         AND ISNULL(vdc.VendorDomainContactCountInCatchAll, 0) >= 2
            THEN 'Review vendor domain usage'
        WHEN ms.IsVendorDomain = 1
            THEN 'Ignore or review'
        ELSE NULL
    END AS VendorHandlingSuggestion,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 100
        WHEN ms.IsDoNotMoveDomain = 1 THEN 0
        WHEN ms.IsPublicDomain = 1 THEN 0

        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 0 THEN
            CASE
                WHEN ISNULL(vdc.VendorDomainContactCountInCatchAll, 0) >= 5 THEN 35
                WHEN ISNULL(vdc.VendorDomainContactCountInCatchAll, 0) >= 2 THEN 20
                ELSE 10
            END

        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN
            CASE
                WHEN smd.CandidateCompanyStatus IN ('Active', 'ACTIVE') THEN 60
                ELSE 40
            END
            + CASE
                WHEN ISNULL(smd.CandidateContactCountOnCompany, 0) >= 5 THEN 20
                WHEN ISNULL(smd.CandidateContactCountOnCompany, 0) >= 3 THEN 10
                ELSE 0
              END

        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount > 1 THEN
            25
            + CASE
                WHEN bmm.SuggestedCompanyStatus IN ('Active', 'ACTIVE') THEN 10
                ELSE 0
              END
            + CASE
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 5 THEN 25
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 3 THEN 15
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 2 THEN 10
                ELSE 0
              END

        WHEN ms.UniqueMatchingCompanyCount = 0 THEN 0

        WHEN ms.UniqueMatchingCompanyCount = 1 THEN
            CASE
                WHEN smd.CandidateCompanyStatus IN ('Active', 'ACTIVE') THEN 60
                ELSE 40
            END
            + CASE
                WHEN ISNULL(smd.CandidateContactCountOnCompany, 0) >= 5 THEN 20
                WHEN ISNULL(smd.CandidateContactCountOnCompany, 0) >= 3 THEN 10
                ELSE 0
              END
            - CASE
                WHEN ISNULL(smd.CandidateContactCountOnCompany, 0) < 3 THEN 25
                ELSE 0
              END

        WHEN ms.UniqueMatchingCompanyCount > 1 THEN
            25
            + CASE
                WHEN bmm.SuggestedCompanyStatus IN ('Active', 'ACTIVE') THEN 10
                ELSE 0
              END
            + CASE
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 5 THEN 25
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 3 THEN 15
                WHEN ISNULL(bmm.SuggestedCompanyHitCount, 0) >= 2 THEN 10
                ELSE 0
              END
        ELSE 0
    END AS ConfidenceScore,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 1
        WHEN ms.IsDoNotMoveDomain = 1 THEN 0
        WHEN ms.IsPublicDomain = 1 THEN 0
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN 1
        WHEN ms.IsVendorDomain = 1 THEN 0
        WHEN ms.UniqueMatchingCompanyCount = 1
             AND ISNULL(smd.CandidateContactCountOnCompany, 0) >= 3 THEN 1
        ELSE 0
    END AS ReadyForApiUpdateFlag,

    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 'Own company domain match - auto-candidate'
        WHEN ms.IsDoNotMoveDomain = 1 THEN 'Do not move - protected domain rule'
        WHEN ms.IsPublicDomain = 1 THEN 'Public domain - manual review'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 0 THEN 'Manual review - vendor domain with no Vendor company match'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN 'Auto-candidate'
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount > 1 THEN 'Manual review - multiple Vendor company matches'
        WHEN ms.IsVendorDomain = 0
             AND ms.UniqueMatchingCompanyCount > 0
             AND ISNULL(
                 CASE
                     WHEN ms.UniqueMatchingCompanyCount = 1 THEN smd.CandidateContactCountOnCompany
                     ELSE bmm.SuggestedCompanyHitCount
                 END, 0
             ) < 3
             THEN 'Potential Vendor / 3rd Party'
        WHEN ms.UniqueMatchingCompanyCount = 0 THEN 'No domain match found'
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN 'Auto-candidate'
        ELSE 'Manual review - suggested company based on highest hit count'
    END AS MatchResult

FROM MatchSummary AS ms
LEFT JOIN SingleMatchDetails AS smd
    ON smd.Contact_RecID = ms.Contact_RecID
LEFT JOIN CandidateArrays AS ca
    ON ca.Contact_RecID = ms.Contact_RecID
LEFT JOIN #BestMultipleMatch AS bmm
    ON bmm.Contact_RecID = ms.Contact_RecID
LEFT JOIN VendorDomainCounts AS vdc
    ON vdc.EmailDomain = ms.EmailDomain
LEFT JOIN #DomainOverallCounts AS doc
    ON doc.MatchDomain = CASE
        WHEN ms.IsVendorDomain = 1 THEN ms.EmailDomain
        ELSE ms.NormalizedEmailDomain
    END
LEFT JOIN OwnCompanyAddress AS oca
    ON 1 = 1
ORDER BY
    CASE
        WHEN ms.IsOwnCompanyDomain = 1 THEN 1
        WHEN ms.IsDoNotMoveDomain = 1 THEN 2
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount = 1 THEN 3
        WHEN ms.UniqueMatchingCompanyCount = 1 THEN 4
        WHEN ms.IsVendorDomain = 1 AND ms.UniqueMatchingCompanyCount > 1 THEN 5
        WHEN ms.UniqueMatchingCompanyCount > 1 THEN 6
        WHEN ms.IsVendorDomain = 1 THEN 7
        WHEN ms.IsPublicDomain = 1 THEN 8
        ELSE 9
    END,
    ConfidenceScore DESC,
    ms.NormalizedEmailDomain,
    ms.Last_Name,
    ms.First_Name;
