SELECT 
  call.index_entry,
  item.barcode, 
  title.best_title,
  title.best_author,
  item.record_creation_date_gmt,
  concat('https://greenwichlibrary.bibliocommons.com/item/show/',id2reckey(link.bib_record_id),'086')
  

FROM 
sierra_view.phrase_entry as call 
Join sierra_view.bib_record_property as title on title.bib_record_id = call.record_id
join sierra_view.bib_record_item_record_link as link on call.record_id = link.bib_record_id
join sierra_view.item_view as item on item.id = Link.item_record_id 
left join sierra_view.item_record_property as itemcall on itemcall.item_record_id = link.item_record_id
full outer join sierra_view.checkout as Cout on Cout.item_record_id = item.id



WHERE 
item.record_creation_date_gmt > (now() - interval '7 day') and 
--item.item_status_code = 'p' and 
call.varfield_type_code = 'c' and item.location_code like 'gm%'
and call.index_entry between '800' and '899.9999'

order by

 call.index_entry   ASC;
 