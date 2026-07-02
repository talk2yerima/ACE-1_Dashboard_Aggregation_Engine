-- Deduplicate each (DataElement × Facility × Sex × AgeGroup × CategoryComboId × ReportingDate)
-- by keeping only the latest submitted row, then SUM across CategoryCombos.
-- This prevents inflated totals caused by re-submissions of the same category combo.
WITH deduped AS (
    SELECT DISTINCT ON (
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "CategoryComboId",
        "ReportingDate"
    )
        "DataElementId",
        "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "CategoryComboId",
        "ReportingDate",
        "Value"
    FROM public."Addons"
    WHERE "DataElementId" IN (
        'y1C9SlSiZhQ',  -- DR POST-RESP - PHYSICAL and/or EMOTIONAL Violence_v22
        'uXsX7BTOuPa',  -- DR POST-RESP - Sexual Violence_v22
        'U3buuLpx4pV',  -- TB (separate)
        'W47dKBCOu3Q',  -- TB (separate)
        'U1u8KRnRN3o',  -- New OPD Attendance
        'unXz0P5Fyad',  -- HIV Risk Screened (Facility)
        'wQAwmcMTaBw',  -- HIV Risk Screened (Community)
        'sGypnO741EM',  -- HIV Test Eligible (Facility)
        'x6yYzcbQrSu',  -- HIV Test Eligible (Community)
        'mC5eo9Ouku0'   -- ANC1 Attendees (PMTCT)
    )
    AND "ReportingDate" >= $1
    AND "ReportingDate" <= $2
    ORDER BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "CategoryComboId",
        "ReportingDate",
        "UpdatedAt" DESC
),
violence_combined AS (
    SELECT
        CASE "DataElementId"
            WHEN 'y1C9SlSiZhQ' THEN 'DR POST-RESP - PHYSICAL and/or EMOTIONAL Violence_v22'
            WHEN 'uXsX7BTOuPa' THEN 'DR POST-RESP - Sexual Violence_v22'
        END      AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" IN ('y1C9SlSiZhQ', 'uXsX7BTOuPa')
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
tb_separate AS (
    SELECT
        "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" IN ('U3buuLpx4pV', 'W47dKBCOu3Q')
    GROUP BY
        "Section",
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
new_opd_attendance AS (
    SELECT
        'New OPD Attendance' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'U1u8KRnRN3o'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
hiv_risk_screened_facility AS (
    SELECT
        'HIV Risk Screened (Facility)' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'unXz0P5Fyad'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
hiv_risk_screened_community AS (
    SELECT
        'HIV Risk Screened (Community)' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'wQAwmcMTaBw'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
hiv_test_eligible_facility AS (
    SELECT
        'HIV Test Eligible (Facility)' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'sGypnO741EM'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
hiv_test_eligible_community AS (
    SELECT
        'HIV Test Eligible (Community)' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'x6yYzcbQrSu'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
),
anc1_attendees AS (
    SELECT
        'ANC1 Attendees (PMTCT)' AS "Section",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        SUM("Value") AS "Value",
        "DataElementId",
        "ReportingDate"
    FROM deduped
    WHERE "DataElementId" = 'mC5eo9Ouku0'
    GROUP BY
        "DataElementId",
        "State",
        "Facility",
        "Datim",
        "Sex",
        "AgeGroup",
        "ReportingDate"
)
SELECT * FROM violence_combined
UNION ALL
SELECT * FROM tb_separate
UNION ALL
SELECT * FROM new_opd_attendance
UNION ALL
SELECT * FROM hiv_risk_screened_facility
UNION ALL
SELECT * FROM hiv_risk_screened_community
UNION ALL
SELECT * FROM hiv_test_eligible_facility
UNION ALL
SELECT * FROM hiv_test_eligible_community
UNION ALL
SELECT * FROM anc1_attendees
ORDER BY
    "ReportingDate" DESC,
    "State",
    "Facility",
    "Datim",
    "Section",
    "Sex",
    "AgeGroup";
