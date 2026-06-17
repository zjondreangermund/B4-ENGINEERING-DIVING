-- Enrich imported Job Cards after importing the new Job Card Conversion workbooks.
-- The uploaded files show the Job Card Info sheet carries these fields in the same row:
-- Date, Month, Job of Day, Client, Job Nr, OPS manager, Description, staff columns,
-- vessel/equipment columns, Cost Labour / Labour Normal Time, Labour Overtime,
-- Cost Equipment and Workshop Cost.
--
-- Run after importing the workbook data into job_card_entries / vessel_entries.
-- This backfills imported cards that show blank Client, Vessel, Staff or Equipment.

WITH vessel_summary AS (
  SELECT
    job_number,
    MAX(NULLIF(client_name,'')) AS client_name,
    MAX(NULLIF(ops_manager,'')) AS ops_manager,
    MAX(NULLIF(division,'')) AS division,
    MAX(NULLIF(description,'')) AS description,
    STRING_AGG(DISTINCT NULLIF(vessel_name,''), ', ' ORDER BY NULLIF(vessel_name,'')) FILTER (WHERE NULLIF(vessel_name,'') IS NOT NULL) AS vessel_name,
    STRING_AGG(DISTINCT NULLIF(equipment_used,''), ', ' ORDER BY NULLIF(equipment_used,'')) FILTER (WHERE NULLIF(equipment_used,'') IS NOT NULL) AS equipment_used
  FROM vessel_entries
  GROUP BY job_number
),
register_summary AS (
  SELECT
    job_number,
    MAX(NULLIF(client_name,'')) AS client_name,
    MAX(NULLIF(ops_manager,'')) AS ops_manager,
    MAX(NULLIF(division,'')) AS division,
    MAX(NULLIF(description,'')) AS description
  FROM job_register_entries
  GROUP BY job_number
)
UPDATE job_card_entries c
SET
  client_name = COALESCE(NULLIF(c.client_name,''), v.client_name, r.client_name),
  vessel_name = COALESCE(NULLIF(c.vessel_name,''), v.vessel_name),
  equipment_used = COALESCE(NULLIF(c.equipment_used,''), v.equipment_used),
  ops_manager = COALESCE(NULLIF(c.ops_manager,''), v.ops_manager, r.ops_manager),
  division = COALESCE(NULLIF(c.division,''), v.division, r.division),
  description = COALESCE(NULLIF(c.description,''), v.description, r.description)
FROM register_summary r
FULL JOIN vessel_summary v ON v.job_number = r.job_number
WHERE c.job_number = COALESCE(v.job_number, r.job_number)
  AND COALESCE(c.source,'import') = 'import';

-- Show imported cards still missing important fields after enrichment.
SELECT
  job_number,
  job_date,
  client_name,
  vessel_name,
  equipment_used,
  staff_names,
  ops_manager,
  division
FROM job_card_entries
WHERE COALESCE(source,'import')='import'
  AND (
    COALESCE(client_name,'')=''
    OR COALESCE(vessel_name,'')=''
    OR COALESCE(ops_manager,'')=''
  )
ORDER BY job_date DESC NULLS LAST, job_number
LIMIT 100;
