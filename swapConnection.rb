require 'rexml/document'
require 'csv'
include REXML

# Globals
@zip = "c:\\Program Files (x86)\\7-Zip\\7z.exe e -y -oc:\\tableauScripts\\Extract"
# tabcmd must be in PATH
@tabCmd = "tabcmd.exe"
@tabUser = "russch"
@tabPassword = "e9hIzZ9a"
@tabServer = "http://localhost"
@tabLogin = @tabCmd + " login -s " + @tabServer + " -u " + @tabUser + " -p " + @tabPassword
@customerListFile = "C://Users//russch.SIMPLE//SkyDrive//Documents//My Documents//GitHub//RubyTabCmdPublisher//config.csv"
@dbUser = "sa"
@dbPassword = "P@ssw0rd123"
@tabLogoff = @tabCmd + " logout"


def Login()
  # Login
  system(@tabLogin)
  Process.waitall
end

def ChangeExtract(tbWorkBook, existingExt, newExt, dstFileName, targetSite)
  put "ChangeExtract"
  @tbWorkBookPath = "c:\\tableauScripts\\" + tbWorkBook + ".twb"

  if !FileTest.exists?(@tbWorkBookPath)
    GetTBWorkBook(tbWorkBook)
  end
  
  twbx = "c:\\tableauScripts\\" + tbWorkBook + ".twb" 
  @zipcmd = @zip + " " + twbx + " " + tbWorkBook + ".twb"
  system(@zipcmd)
  Process.waitall
  
  File.open(@tbWorkBookPath) do |config|
    twb = Document.new(config)
    twb.elements.each("*/datasources/datasource/connection") do |element| 
      element.attributes["filename"]="c:\\tableauScripts\\#{newExt}" 
    end

    @fqdnFileName = "c:\\tableauScripts\\Outgoing\\" + dstFileName

    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName, 'w') do |result|
      formatter.write(twb, result)
    end
  end
  
  Publish(@fqdnFileName, targetSite)
  Process.waitall
  File.delete(@fqdnFileName)
end

def ChangeDB(tbWorkBook, tbWorkBookDest, sourceSite, targetSite, newDatabase )

 #See if file already exists
  @tbWorkBookPath = "c:\\tableauScripts\\Incoming\\" + tbWorkBook + ".twb"

  if !FileTest.exists?(@tbWorkBookPath)
    # If not, tabcmd GET it
	GetTBWorkBook(tbWorkBook, sourceSite)
  end

  File.open(@tbWorkBookPath) do |config|
    twb = Document.new(config)
    twb.elements.each("*/datasources/datasource/connection") do |element| 
	  # Change the server name to newDatabase	
      element.attributes["server"]="#{newDatabase}" 
    end
	
	#Write to Outgoing folder
    @fqdnFileName = "c:\\tableauScripts\\Outgoing\\" + tbWorkBookDest + ".twb"
 
    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName,"w") do |result|
      formatter.write(twb, result)
    end
  end
  # TabCmd Publish the report
  puts "Publishing " + tbWorkBookDest + " to " + targetSite
  PublishDB(@fqdnFileName,tbWorkBookDest, targetSite)
  Process.waitall
  #Clean up
  puts "Deleting temp files..."
  File.delete(@fqdnFileName)
  #File.delete(@tbWorkBookPath)
end

def Publish(fqdnFileName,dstFileName, site)
 
  cmd = @tabCmd + " publish --overwrite -t " + site + " -n " + dstFileName + " \"" + fqdnFileName + "\" -u " + @tabUser + " -p " + @tabPassword
  system(cmd)
  p Process.waitall
end

def PublishDB(fqdnFileName,dstFileName, site)
  # When publishing to Default site, site should be a zero length string
  # Also, wrap workbook name in quotes.
  cmd = @tabCmd + " publish --overwrite -t \"" + (site == "Default" ? "" : site) + "\" -n \"" + dstFileName + "\" \"" + fqdnFileName + "\" --db-username " + @dbUser + " --db-password " + @dbPassword + " -u " + @tabUser + " -p " + @tabPassword + " --save-db-password"
  puts cmd
  system(cmd)
  Process.waitall
end

def GetTBWorkBook(tbWorkBookName, sourceSite)
  Login()
  # Define site (or none, if default)
  siteGet = (sourceSite == "Default" ? " get /workbooks/" : " get /t/" + sourceSite + "/workbooks/")
  # Strip white spaces from workbook name for the get
  cmd = @tabCmd + siteGet + tbWorkBookName.delete(' ') + ".twb -f \"C:\\tableauScripts\\Incoming\\" + tbWorkBookName + ".twb\""
  puts "Getting " + tbWorkBookName + "..."
  puts cmd
  system(cmd)
  Process.waitall
end

def Main()
	puts "Processing..."
	csv_contents = CSV.read(@customerListFile)
	csv_contents.shift
	csv_contents.each do |row|
	  # row[0] defines if data source is an extract or database
	  # Source Tableau Server
	  @sourceServer = row[1].strip! || row[1] if row[1]	  
	  # Target Tableau Server
	  @targetServer = row[2].strip! || row[2] if row[2]	    
	  # Tableau source site
	  @sourceSite = row[3].strip! || row[3] if row[3]	  
	  # Tableau Site to be published to
	  @targetSite = row[4].strip! || row[4] if row[4]
	  # Orignal Workbook Name
	  @sourceWorkbook = row[5].strip! || row[5] if row[5]
	  # Name of workbook as published on new server
	  @targetWorkbook = row[6].strip! || row[6] if row[6]
	  # Original Data Source Name   
	  @sourceDataSource = row[7].strip! || row[7] if row[7]
	  # New Data Source Name (Or Server Name or database)
	  @targetDataSource = row[8].strip! || row[8] if row[8]

	  
	  if row[0] == "extract"
		ChangeExtract(@row1, @row2, @row3, @row4, @row5)
	  elsif row[0] == "database"
		ChangeDB(@sourceWorkbook, @targetWorkbook, @sourceSite, @targetSite, @targetDataSource )
	  end
	end

	system(@tabLogoff)
end

Main()
