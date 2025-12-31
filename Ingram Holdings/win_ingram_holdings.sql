--print items and orders for Winchester

DROP TABLE IF EXISTS win_bibs;

CREATE TEMP TABLE win_bibs AS

SELECT
b.bib_record_id

FROM
sierra_view.bib_record_property b
LEFT JOIN
sierra_view.bib_record_item_record_link l
ON
b.bib_record_id = l.bib_record_id
LEFT JOIN
sierra_view.item_record i
ON
l.item_record_id = i.id  AND i.location_code ~ '^wi' AND i.item_status_code NOT IN ('w','$','n')
LEFT JOIN
sierra_view.bib_record_order_record_link ol
ON
b.bib_record_id = ol.bib_record_id
LEFT JOIN
sierra_view.order_record o
ON
ol.order_record_id = o.id AND o.accounting_unit_code_num = '36'
LEFT JOIN
sierra_view.varfield v
ON
i.id = v.record_id AND v.varfield_type_code = 'v'

WHERE i.id IS NOT NULL OR o.id IS NOT NULL
GROUP BY 1
HAVING
COUNT(i.id) FILTER(WHERE v.field_content IS NOT NULL) = 0
;

DROP TABLE IF EXISTS isbn;

CREATE TEMP TABLE isbn AS

SELECT
s.record_id,
COALESCE(STRING_AGG(DISTINCT SUBSTRING(s.content FROM '^\d{9,12}[\d|X]'),'|'),'') AS "isbns"
FROM
win_bibs b
JOIN
sierra_view.subfield s
ON
b.bib_record_id = s.record_id AND s.marc_tag = '020' AND s.tag = 'a'
GROUP BY 1
;

SELECT
i.record_id,
--trying to use phase entry instead of subfield for 001
COALESCE(SUBSTRING(o.content FROM '[0-9]+'),'') AS "001",
i.isbns AS "020"--,
--COALESCE(STRING_AGG(DISTINCT SUBSTRING(s3.content FROM '^\d+'),'|'),'') AS "024"

FROM
isbn i
JOIN
sierra_view.bib_record_property b
ON
i.record_id = b.bib_record_id AND b.material_code NOT IN ('3','6','b','g','h','l','s','w','y')
LEFT JOIN
sierra_view.subfield o
ON
i.record_id = o.record_id AND o.marc_tag = '001'
/*LEFT JOIN
sierra_view.subfield s3
ON
bi.bib_record_id = s3.record_id AND s3.marc_tag = '024' AND s3.tag = 'a'

WHERE 
s3.content IS NOT NULL*/
