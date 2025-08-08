-- Declare and initialize parameters
DECLARE @TargetParentOrganizationID INT = {0}; -- Example: The Organization ID for your library system.
DECLARE @DeletionStartDate DATETIME = '{1}'; -- Example: The start date for the deletion period.
DECLARE @DeletedStatusID INT = (SELECT RecordStatusID FROM polaris.polaris.RecordStatuses WHERE RecordStatusName = 'Deleted');

-- A table variable to hold the list of item statuses that are considered terminal/inactive.
-- As per the image provided: 7=Lost, 8=Claim Returned, 9=Claim Never Had, 10=Missing, 11=Withdrawn, 16=Unavailable, 20=Does Not Circulate, 21=Claim Missing Parts
DECLARE @TerminalStatuses TABLE (ItemStatusID INT PRIMARY KEY);
INSERT INTO @TerminalStatuses (ItemStatusID) VALUES (7), (8), (9), (10), (11), (16), (20), (21);

-- CTE to gather and summarize item information for the target organization
WITH ItemInfo AS (
    SELECT
        cir.AssociatedBibRecordID,
        -- Count of "active" items (i.e., not deleted and not in a terminal status)
        SUM(CASE
            WHEN cir.RecordStatusID <> @DeletedStatusID AND ts.ItemStatusID IS NULL THEN 1
            ELSE 0
        END) AS ActiveItemCount,

        -- Count of items deleted within the specified period
        SUM(CASE
            WHEN cir.RecordStatusID = @DeletedStatusID AND cir.RecordStatusDate >= @DeletionStartDate THEN 1
            ELSE 0
        END) AS DeletedInPeriodCount,

        -- Count of items that have one of the terminal statuses
        SUM(CASE
            WHEN ts.ItemStatusID IS NOT NULL THEN 1
            ELSE 0
        END) AS TerminalStatusCount,

        -- Total number of items attached to the bib
        COUNT(cir.ItemRecordID) AS TotalItemCount
    FROM
        polaris.polaris.CircItemRecords AS cir
    JOIN
        polaris.polaris.Organizations AS org ON cir.AssignedBranchID = org.OrganizationID
    LEFT JOIN
        @TerminalStatuses AS ts ON cir.ItemStatusID = ts.ItemStatusID
    WHERE
        org.ParentOrganizationID = @TargetParentOrganizationID
    GROUP BY
        cir.AssociatedBibRecordID
)
-- Select OCLC numbers for bibs that meet the deletion or terminal status criteria
SELECT DISTINCT
    '' AS LSN,
    '(OCoLC)' + CAST(CAST(SUBSTRING(o035.SystemControlNumber, 8, LEN(o035.SystemControlNumber) - 7) AS BIGINT) AS VARCHAR(40)) AS OCLC_Number
FROM
    ItemInfo ii
JOIN
    polaris.polaris.BibliographicTag035Index o035 ON ii.AssociatedBibRecordID = o035.BibliographicRecordID
WHERE
    -- Condition 1: No "active" items must remain on the bib. All items must be either deleted or in a terminal status.
    ii.ActiveItemCount = 0
    -- Condition 2: The bib must be included if EITHER at least one item was deleted in the period OR if ALL items on the bib have a terminal status.
    AND (ii.DeletedInPeriodCount > 0 OR ii.TerminalStatusCount = ii.TotalItemCount)
    AND UPPER(LEFT(o035.SystemControlNumber, 7)) = '(OCOLC)'
    AND LEN(o035.SystemControlNumber) > 7
    AND ISNUMERIC(SUBSTRING(o035.SystemControlNumber, 8, LEN(o035.SystemControlNumber) - 7)) = 1;
