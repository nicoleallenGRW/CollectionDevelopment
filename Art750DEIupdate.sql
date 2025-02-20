--Create shelf list of all items in a particular collection by call number. Also contains key DEI data tags for staff.
SELECT 
itemcall.call_number_norm as "call no.",
title.best_title AS "Title",
title.best_author AS "Author",
CASE
 WHEN title.publish_year IS NULL THEN 0
 ELSE title.publish_year END AS "Pub. Year",
item.record_creation_date_gmt AS "Item Created",
CASE
 WHEN to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM') IS NULL THEN 'No data'
 ELSE to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM')
 END AS "Last Checkin",
item.checkout_total AS "Total Checkouts",
item.renewal_total AS "Total Renewals",
item.year_to_date_checkout_total AS "YTD",
item.last_year_to_date_checkout_total AS "LYR",
CASE
 WHEN checkout.loanrule_code_num > 0 THEN 'YES'
 ELSE 'NO' END AS "Checked out?",
item.barcode AS "Barcode",
item.item_status_code AS "Status",
COALESCE(m590.field_content) as "590",
COALESCE(string_agg(m695.field_content,',')) AS "695"

FROM
sierra_view.phrase_entry AS call
JOIN sierra_view.bib_record_property AS title ON title.bib_record_id = call.record_id
JOIN sierra_view.bib_record_item_record_link AS link ON call.record_id = link.bib_record_id
JOIN sierra_view.item_view AS item ON item.id = link.item_record_id 
LEFT JOIN sierra_view.item_record_property AS itemcall ON itemcall.item_record_id = link.item_record_id
FULL OUTER JOIN sierra_view.checkout AS checkout ON checkout.item_record_id = link.item_record_id
LEFT JOIN sierra_view.varfield AS m590 ON title.bib_record_id = m590.record_id AND m590.marc_tag = '590'
LEFT JOIN sierra_view.varfield AS m695 ON title.bib_record_id = m695.record_id AND m695.marc_tag = '695'

WHERE

call.varfield_type_code = 'c' AND
itemcall.call_number_norm BETWEEN '750' AND '769.999' AND
item.location_code = 'gmad2' AND
itemcall.call_number <> 'pbk nonfiction' 

group by
itemcall.call_number_norm,
title.best_title,
title.best_author,
CASE
 WHEN title.publish_year IS NULL THEN 0
 ELSE title.publish_year END,
item.record_creation_date_gmt,
CASE
 WHEN to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM') IS NULL THEN 'No data'
 ELSE to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM')
 END,
item.checkout_total,
item.renewal_total,
item.year_to_date_checkout_total,
item.last_year_to_date_checkout_total,
CASE
 WHEN checkout.loanrule_code_num > 0 THEN 'YES'
 ELSE 'NO' END,
item.barcode,
item.item_status_code,
COALESCE(m590.field_content)

ORDER BY

itemcall.call_number_norm;
