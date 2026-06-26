WITH violence_combined AS (
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
    FROM public."Addons"
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
