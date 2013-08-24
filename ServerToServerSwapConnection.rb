require 'rexml/document'
require 'csv'
include REXML

# Globals
@zip = "\"c:\\Program Files (x86)\\7-Zip\\7z.exe\" e -y -oc:\\tableauScripts\\Extract"
# tabcmd must be in PATH
@tabCmd = "tabcmd.exe"
@tabUser = "simple\\flinstone"
@tabPassword = "1LikeBetty!"
@tabServer = "http://localhost:8000"
@tabLogin = @tabCmd + " login -s " + @tabServer + " -u " + @tabUser + " -p " + @tabPassword
@customerListFile = "C://tableauscripts//configServerToServer.csv"
@dbUser = "sa"
@dbPassword = "foo"
@tabLogoff = @tabCmd + " logout"

=begin
	To get this script to run, you need:
	
	- tabcmd installed locally and added to your PATH statement
	- 7-zip installed with the @zip variable (above) pointing to it ("\"c:\\Program Files (x86)\\7-Zip\\7z.exe\"
	- 3 Folders on the file system: c:\tableauscripts\incoming, c:\tableauscripts\outgoing, c:\tableauscripts\extract
	- config.csv filled out and accessible via @customerListFile
	
	What do the folders do?
	- \incoming is used to temporarily host files which have been grabbed from Tableau Server via tabcmd get
	- \outgoing is used to temprarily host files that have been modified and about to be published back to a Server
	- \extracts holds NEW extracts you have been created which will be swapped in for old extracts in a twbx you pulled down
	- \extracts is also a temporary holding place for the .twb which will get unzipped out of a twbx before it is modified
	   and then placed in \outgoing

=end



def Login()
  # Login
  system(@tabLogin)
  Process.waitall
end

def ChangeExtract(tbWorkBook, tbWorkBookDest, sourceSite, targetSite, newExt)#

  @tbWorkBookPath = "c:\\tableauScripts\\incoming\\" + tbWorkBook + ".twb"
  @tbUnZippedWorkBookPath = "c:\\tableauScripts\\Extract\\" + tbWorkBook + ".twb"
  #See if file already exists
  if !FileTest.exists?(@tbWorkBookPath)
    # If not, tabcmd GET it
	GetTBWorkBook(tbWorkBook, sourceSite)
  end
  
  #unzip 
  twbx = "\"c:\\tableauScripts\\incoming\\" + tbWorkBook + ".twb\"" 
  @zipcmd = @zip + " " + twbx + " \"" + tbWorkBook + ".twb\""
  system(@zipcmd)
  Process.waitall
  
  File.open(@tbUnZippedWorkBookPath) do |config|
    twb = Document.new(config)
    
	twb.elements.each("*/datasources/datasource/connection") do |element| 
	  # Change the extract file name in the original connectionto new one	
      element.attributes["dbname"]="c:\\tableauScripts\\Extract\\#{newExt}" 
    end
	

	#Write to Outgoing folder
    @fqdnFileName = "c:\\tableauScripts\\Outgoing\\" + tbWorkBookDest + ".twb"

    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName, 'w') do |result|
      formatter.write(twb, result)
    end
  end
  
  Publish(@fqdnFileName, tbWorkBookDest, targetSite)
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
  File.delete(@tbWorkBookPath)
end

def Publish(fqdnFileName,dstFileName, site)
  # When publishing to Default site, site should be a zero length string
  # Also, wrap workbook name in quotes. 
  cmd = @tabCmd + " publish --overwrite -t \"" + (site == "Default" ? "" : site) + "\" -n \"" + dstFileName + "\" \"" + fqdnFileName + "\" -u " + @tabUser + " -p " + @tabPassword
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
	  # Original Data Source Name - not really used for anything, just helps to make the config file more logical.
	  @sourceDataSource = row[7].strip! || row[7] if row[7]
	  # New Data Source Name (Or Server Name or database)
	  @targetDataSource = row[8].strip! || row[8] if row[8]
	  # New extract file name to be added to twbx as substitute for current file (type "extract" only)
	  @targetExtractName = row[9].strip! || row[9] if row[9]	  

	  
	  if row[0] == "extract"
		ChangeExtract(@sourceWorkbook, @targetWorkbook, @sourceSite, @targetSite,  @targetExtractName)
	  elsif row[0] == "database"
		ChangeDB(@sourceWorkbook, @targetWorkbook, @sourceSite, @targetSite, @targetDataSource )
	  end
	end

	system(@tabLogoff)
end

Main()
