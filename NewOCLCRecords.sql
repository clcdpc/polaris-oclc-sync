-- Declare parameters
DECLARE @TargetParentOrganizationID INT = {0}; -- Example: Set to the OrganizationID of the library system.
DECLARE @CreationStartDate DATETIME = '{1}'; -- Example: Set the desired start date for NEW items and the floor for the lost, missing, withdrawn history.

-- CTE to find the first item created for each title within the target organization
WITH FirstItemCreationDateInTargetParentOrg AS (
    SELECT
        cir_cte.AssociatedBibRecordID,
        MIN(cir_cte.FirstAvailableDate) AS MinCreationDateForBib
    FROM
        polaris.polaris.CircItemRecords cir_cte
    JOIN
        polaris.polaris.Organizations org_cte ON cir_cte.AssignedBranchID = org_cte.OrganizationID
    WHERE
        org_cte.ParentOrganizationID = @TargetParentOrganizationID
    GROUP BY
        cir_cte.AssociatedBibRecordID
),	
-- CTE to pre-calculate history information for all items
ItemHistorySummary AS (
    SELECT
        irh.ItemRecordID,
        COUNT(*) AS HistoryCount,
        -- This flag will now be 1 if the OldItemStatusID was Lost, Missing, or Withdrawn ON OR AFTER the @CreationStartDate
        MAX(CASE
                WHEN irh.OldItemStatusID IN (7, 10, 11) -- Check for Lost (7), Missing (10), or Withdrawn (11) statuses
                     AND irh.TransactionDate >= @CreationStartDate
                THEN 1
                ELSE 0
            END) AS HasRecentLostHistory
    FROM
        polaris.polaris.ItemRecordHistory irh
    GROUP BY
        irh.ItemRecordID
)
-- Main query
SELECT DISTINCT
    '' AS LSN,
	-- Updated line to format the OCLC Number
    '(OCoLC)' + CAST(MIN(CAST(SUBSTRING(o035.SystemControlNumber, 8, LEN(o035.SystemControlNumber) - 7) AS BIGINT)) AS VARCHAR(40)) AS OCLC_Number
FROM
    polaris.polaris.Organizations org
JOIN
    polaris.polaris.CircItemRecords cir ON org.OrganizationID = cir.AssignedBranchID
JOIN
    FirstItemCreationDateInTargetParentOrg fic ON cir.AssociatedBibRecordID = fic.AssociatedBibRecordID
LEFT JOIN
    ItemHistorySummary ihs ON cir.ItemRecordID = ihs.ItemRecordID -- Join the pre-calculated history
LEFT JOIN
    polaris.polaris.ItemStatuses istat ON cir.ItemStatusID = istat.ItemStatusID
JOIN
    polaris.polaris.BibliographicRecords br ON cir.AssociatedBibRecordID = br.BibliographicRecordID
-- Changed to INNER JOIN to ensure an OCLC number exists
JOIN
    polaris.polaris.BibliographicTag035Index o035 ON cir.AssociatedBibRecordID = o035.BibliographicRecordID
WHERE
    -- General filters that apply to ALL results
    org.ParentOrganizationID = @TargetParentOrganizationID
    AND cir.ILLFlag = 0
    AND cir.RecordStatusID = 1 -- Only final item records
    -- Added filters for valid OCLC numbers
    AND o035.SystemControlNumber LIKE '(OCoLC)%'
    AND ISNUMERIC(SUBSTRING(o035.SystemControlNumber, 8, LEN(o035.SystemControlNumber) - 7)) = 1
    AND
    (
        -- Scenario 1: Item is genuinely NEW
        (
            cir.FirstAvailableDate >= @CreationStartDate
            AND cir.FirstAvailableDate = fic.MinCreationDateForBib -- It is the first item created for this title
        )
        OR
        -- Scenario 2: Item has been "Resurrected"
        (
            cir.ItemStatusID IN (1, 2, 3, 4, 5, 6) -- It is currently in a normal circulating status
            AND ihs.HasRecentLostHistory = 1 -- And it has a "Lost", "Missing", or "Withdrawn" entry (based on OldItemStatusID) on or after the start date
        )
    )
-- Group by AssociatedBibRecordID to ensure the MIN aggregate function works correctly per bib record
GROUP BY
    cir.AssociatedBibRecordID
-- If you only want unique OCLC numbers, ordering by the OCLC_Number itself makes the most sense.
ORDER BY
    OCLC_Number;
