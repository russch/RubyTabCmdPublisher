# Modify the name of a column in the data source. Column in question must 
# already have been renamed in Desktop to add the caption attribute. 
require 'rexml/document'
include REXML

 File.open('C://Users//russch.SIMPLE//Desktop//test.twb') do |config|
    twb = Document.new(config)
	twb.elements.each("*/datasources/datasource") do |element|
      element.each_element_with_attribute('caption', 'Tokenized City') do |childelement| 
	   # Fix up Caption Name
	   childelement.attributes['caption']='Friendly City Name'
     end
	end

	#Write to Outgoing folder
    @fqdnFileName = 'C://Users//russch.SIMPLE//Desktop//1test.twb'
 
    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName,"w") do |result|
      formatter.write(twb, result)
    end
 end