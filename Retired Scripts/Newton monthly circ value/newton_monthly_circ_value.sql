--changed from weekly to monthly 4/28/25

SELECT
CAST(SUM(i.price) AS MONEY) AS value,
COUNT(c.id) AS circ_count,
(CAST(SUM(i.price) AS MONEY) / COUNT(c.id)) as value_per_circ,
localtimestamp - interval '1 month' AS start_time,
localtimestamp AS end_time
FROM
sierra_view.circ_trans c
JOIN
sierra_view.item_record i
ON
c.item_record_id = i.id
WHERE
c.op_code IN ('o', 'r')
AND
c.transaction_gmt >= (localtimestamp - interval '1 month')
AND
i.location_code LIKE 'ntn%';
