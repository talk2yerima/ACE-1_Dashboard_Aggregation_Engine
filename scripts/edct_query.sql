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
)
SELECT * FROM violence_combined
UNION ALL
SELECT * FROM tb_separate
ORDER BY
    "ReportingDate" DESC,
    "State",
    "Facility",
    "Datim",
    "Section",
    "Sex",
    "AgeGroup";
