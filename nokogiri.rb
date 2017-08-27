require 'nokogiri' # comment added
require 'xlsxtream'
array1 = []
array2 = []
array3 = []
array4 = []
report={}
report1={}

doc = Nokogiri::XML(File.open("show.xml"))    
doc_pass = doc.xpath("//CELL//CELL_SPEC")    
doc_pass.xpath("//REAL_INDEX").each do |pass|
 array1 << pass['VALUE']
end
doc_pass.xpath("//STRING_INDEX").each do |si|
  array2 << si['VALUE']
end
doc_pass.xpath("//REAL_RANGE_INDEX").each do |hv|
 array3 << hv['HIGH_VALUE']
end
doc_pass1 = doc.xpath("//CELL//CELL_VALUE")
doc_pass1.each do |dv|
 array4 << dv['DECIMAL_VALUE']
end    
report['ri'] = array1
report['si'] = array2
report['hv'] = array3
report['dv'] = array4

puts array1.length
puts array2.length
puts array3.length
puts array4.length

count=0
array1.length.times do |i|
	# puts "#{array1[i]};#{array3[i]};#{array4[i]}"
	File.open("dup1.txt", 'a+')  { |file| file.write("#{array1[i]};#{array3[i]};#{array4[i]};#{array2[count]}; #{array2[count+1]}") }
    File.open("dup1.txt", 'a+') { |file| file.write("\n") }
    count += 3
end

Xlsxtream::Workbook.open("foo1.xlsx") do |xlsx|
  xlsx.write_worksheet "Sheet1" do |sheet|
  	count=0
  	array1.length.times do |i|
    # Date, Time, DateTime and Numeric are properly mapped
      sheet << [array1[i], array3[i], array4[i], array2[count], array2[count+1]]
      # puts "#{array2[count]}; #{array2[count+1]}"
      count += 3
    end
  end
end