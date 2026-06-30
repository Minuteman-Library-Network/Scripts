SELECT
  rm.record_num AS patronid,
  b.index_entry AS barcode,
  f.first_name AS firstname,
  f.last_name AS lastname,
  p.birth_date_gmt AS birthdate,
  p.home_library_code AS homelibr,
  a.addr1 AS address1,
  a.addr2 AS address2,
  a.city AS city,
  a.region AS STATE,
  a.postal_code AS zip,
  rm.creation_date_gmt AS creationdate,
  p.expiration_date_gmt AS expirationdate,
  p.activity_gmt AS circactive,
  p.ptype_code AS ptype,
  e.field_content AS email,
  p.owed_amt AS moneyowed,
  rm.record_type_code||rm.record_num AS pnum,
  DATE_PART('year', p.birth_date_gmt) AS birthyear

FROM sierra_view.patron_record p
JOIN sierra_view.record_metadata rm
  ON p.id = rm.id
JOIN sierra_view.phrase_entry b
  ON p.id = b.record_id
  AND b.varfield_type_code = 'b'
  AND b.occurrence = 0
JOIN sierra_view.patron_record_fullname f
  ON p.id = f.patron_record_id
JOIN sierra_view.patron_record_address a
  ON p.id = a.patron_record_id
LEFT JOIN sierra_view.varfield e
  ON p.id = e.record_id
  AND e.varfield_type_code = 'z'
  AND e.occ_num = 0
  
WHERE p.ptype_code IN ('26','126','326')