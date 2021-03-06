$:.unshift File.dirname(__FILE__)

require "byebug"
require "rubyXL"

domande_workbook = RubyXL::Parser.parse("04_RPS_ALI_103_dati_domande.xlsx")
domande = domande_workbook[0]
spid_workbook = RubyXL::Parser.parse("05_RPS_ALI_103_dati_spid.xlsx")
spid = spid_workbook[0]

domande.each_with_index do |d, i|
  break if i == 48674 
  puts "#{i} - #{d[3].value}"
  spid.each_with_index do |s, y|
    break if y == 15250
    domande_email = d[3].value
    spid_email = s[12].value
    if domande_email && spid_email && domande_email.strip.downcase == spid_email.strip.downcase
      (0..14).each_with_index do |c, z|
        domande.add_cell(i, 33 + z, s[z].value)
      end
      break
    end
  end
end

domande_workbook.write("results.xlsx")

