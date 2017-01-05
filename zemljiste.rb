# encoding: utf-8
require 'rexml/document' # The require statement loads the REXML library.
include REXML # so that we don't have to prefix everything with REXML::...
require "uuidtools"

require 'spreadsheet' # za citanje excell-a
require 'active_support'
require 'active_support/core_ext' # za tj i ch i sh....
def check(x)
    if x.class == Float
        return x.round
    else
        return x
    end
 end

def printrows_to_array()
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open('zemljiste.xls')
  sheet = book.worksheet('zemlja')
  x = []
  
  sheet.each do |row|
    
  x << [row[0].to_s, row[1].to_s, row[2].round, row[3].to_s, row[4].to_s, row[5].to_s, row[6].to_s, row[7].to_s, check(row[8]), row[9].round, row[10].round, row[11].round, row[12].to_s, row[13].to_s, row[14].to_s, row[15].round(2), row[16].round(2) ]
  puts check(row[8])
  end
  return x
end

def AsadauXML(x)
    stringXML = ""
    x.each do |x1Napomena, x2AnalitSifra, x3BrojLista, x4KO, x5Lokacija, x6Naziv, x7Adresa, x8Vrsta, x9BrojParcele, x10Hektara, x11Ara, x12Metara, x13VrstaPrava, x14OblikSvojine, x15ObimUdela, x16knjigoVrednost, x17LikvidacionVrednost|
        @RandomBroj = UUIDTools::UUID.random_create #bez @ dobijam "dynamic constant assignment error"
        string = <<EOF
<asset>
    <ID>#{@RandomBroj}</ID>
    <GUID>00000000-0000-0000-0000-000000000000</GUID>
    <asset_type>Land</asset_type>
    <category />
    <deleted>0</deleted>
    <date_entered>2015-10-29T00:00:00+02:00</date_entered>
    <include_in_EFR>1</include_in_EFR>
    <inventory_display_timestamp>2015-10-29T00:00:00+02:00</inventory_display_timestamp>
    <inventory_timestamp>2015-09-30T00:00:00+02:00</inventory_timestamp>
    <liquidation_value_change_date>2014-07-08T00:00:00+02:00</liquidation_value_change_date>
    <write_off_date>2014-07-08T00:00:00+02:00</write_off_date>
    <case_id>0</case_id>
    <book_value>#{x16knjigoVrednost}</book_value>
    <sale_value>0</sale_value>
    <needs_insurance>0</needs_insurance>
    <account />
    <write_off_grounds />
    <search_field />
    <write_off_description />
    <liquidation_value>#{x17LikvidacionVrednost}</liquidation_value>
    <sale_value_rsd>0</sale_value_rsd>
    <income_value_rsd>0</income_value_rsd>
    <distributed_value_rsd>0</distributed_value_rsd>
    <third_party />
    <code>#{x2AnalitSifra}</code>
    <name>#{x6Naziv}</name>
    <notes>#{x1Napomena}</notes>
    <state />
    <currency_code>RSD</currency_code>
    <estimated_sale_value>0</estimated_sale_value>
    <additional_expenses_rsd>0</additional_expenses_rsd>
    <administrators_reward_rsd>0</administrators_reward_rsd>
    <ownership_type />
    <buyer_id>0</buyer_id>
    <fully_paid>0</fully_paid>
    <city />
    <town />
    <municipality />
    <address>#{x7Adresa}</address>
    <postal_code />
    <country />
    <building_net_area>0</building_net_area>
    <building_gross_area>0</building_gross_area>
    <building_purpose />
    <building_floor_count>0</building_floor_count>
    <building_cadastre_parcel />
    <building_built_date>2014-07-08T00:00:00+02:00</building_built_date>
    <building_number />
    <building_apartments_number />
    <building_room_count />
    <building_possesing_paper />
    <building_landed_bookmark />
    <building_paper_number />
    <building_geoletry />
    <building_cadastre_municipality />
    <building_registered_area>0</building_registered_area>
    <building_nonregistered_area>0</building_nonregistered_area>
    <building_legal_category />
    <building_ownership_status />
    <building_stock_latitude />
    <building_nonregistered_ownership />
    <building_privileged_buyers />
    <building_quantity_unit />
    <building_quantity>0</building_quantity>
    <building_location />
    <equipment_type />
    <equipment_quantity_unit />
    <equipment_manufacturer />
    <equipment_location />
    <equipment_eq_state />
    <equipment_quantity>0</equipment_quantity>
    <land_area_are>#{x11Ara}</land_area_are>
    <land_area_sqmetre>#{x12Metara}</land_area_sqmetre>
    <land_location>#{x5Lokacija}</land_location>
    <land_unit />
    <land_area_hectare>#{x10Hektara}</land_area_hectare>
    <land_type>#{x8Vrsta}</land_type>
    <land_cadastre_parcel>#{x9BrojParcele}</land_cadastre_parcel>
    <land_delimited>0</land_delimited>
    <land_cadastre_municipality>#{x4KO}</land_cadastre_municipality>
    <land_building_paper_number>#{x3BrojLista}</land_building_paper_number>
    <land_legal_category>#{x13VrstaPrava}</land_legal_category>
    <land_ownership_status>#{x14OblikSvojine}</land_ownership_status>
    <land_stock_latitude>#{x15ObimUdela}</land_stock_latitude>
    <land_privileged_buyers />
    <land_nonregistered_ownership />
    <stock_shares_brocker />
    <stock_shares_amount>0</stock_shares_amount>
    <stock_shares_share_percent>0</stock_shares_share_percent>
    <stock_shares_market />
    <stock_shares_company />
    <stock_shares_company_id />
    <stock_shares_privileged_buyers />
    <stock_shares_stock_symbol />
    <supplies_goods_type />
    <supplies_storage_type />
    <supplies_unit />
    <supplies_storage_location />
    <supplies_amount>0</supplies_amount>
    <supplies_expiration_date>2014-07-08T00:00:00+02:00</supplies_expiration_date>
    <supplies_storage_date>2014-07-08T00:00:00+02:00</supplies_storage_date>
    <supplies_rottenly_goods>0</supplies_rottenly_goods>
    <vehicle_last_service>2014-07-08T00:00:00+02:00</vehicle_last_service>
    <vehicle_mileage>0</vehicle_mileage>
    <vehicle_brand />
    <vehicle_registrated_until>2014-07-08T00:00:00+02:00</vehicle_registrated_until>
    <vehicle_plate_number />
    <vehicle_insurance />
    <vehicle_purchase_date>2014-07-08T00:00:00+02:00</vehicle_purchase_date>
    <vehicle_equipment />
    <vehicle_chasis_code />
    <vehicle_quantity>0</vehicle_quantity>
    <vehicle_other />
    <vehicle_color />
    <vehicle_production_date>2014-07-08T00:00:00+02:00</vehicle_production_date>
    <vehicle_model />
    <vehicle_type />
    <vehicle_manufacturer />
    <vehicle_status />
    <vehicle_engine_cubage>0</vehicle_engine_cubage>
    <vehicle_engine_power>0</vehicle_engine_power>
    <vehicle_licence_number />
    <vehicle_engine_number />
    <created_by_id>0</created_by_id>
    <license_and_patent_full_name />
    <license_and_patent_number />
    <license_and_patent_issuing_date>2014-07-08T00:00:00+02:00</license_and_patent_issuing_date>
    <license_and_patent_termination_date>2014-07-08T00:00:00+02:00</license_and_patent_termination_date>
    <license_and_patent_category />
    <license_and_patent_owner />
    <biological_asset_type />
    <biological_asset_sort />
    <biological_asset_strain />
    <biological_asset_quantity_unit />
    <biological_asset_quantity>0</biological_asset_quantity>
    <biological_asset_purpose />
</asset> 
EOF
    stringXML.concat(string)
    end
    #ovde ga samo ubaciti u fajl
    UbaciUFajl(stringXML)
end

def UbaciUFajl(string)
#doc = Document.new(string)
file = File.open("zemljiste.xml", "w")
headerString = <<EOF
<?xml version="1.0" standalone="yes"?>
<AssetDataSet xmlns="http://tempuri.org/AssetDataSet.xsd">
 #{string}
</AssetDataSet> 
EOF
file.write(headerString)
file.close 

end


AsadauXML(printrows_to_array())

