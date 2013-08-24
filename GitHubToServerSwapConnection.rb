require 'rexml/document'
require 'csv'
require 'octokit'
include REXML

# Globals
# tabcmd must be in PATH
@tabCmd = "tabcmd.exe"
@tabUser = "simple\\flinstone"
@tabPassword = "1LikeBetty!"
@tabServer = "http://localhost:8000"
@tabLogin = @tabCmd + " login -s " + @tabServer + " -u " + @tabUser + " -p " + @tabPassword
@customerListFile = "C://tableauscripts//configGitToServer.csv"
@dbUser = "sa"
@dbPassword = "P@ssw0rd123"
@tabLogoff = @tabCmd + " logout"

=begin
	To get this script to run, you need:
	
	- tabcmd installed locally and added to your PATH statement
	- 2 Folders on the file system: c:\tableauscripts\incoming, c:\tableauscripts\outgoing
	- configGitToServer.csv filled out and accessible via @customerListFile
	
	What do the folders do?
	- \incoming is used to temporarily host files which have been grabbed from Tableau Server via tabcmd get
	- \outgoing is used to temprarily host files that have been modified and about to be published back to a Server


=end



def Login()
  # Login
  system(@tabLogin)
  Process.waitall
end


def ChangeDB(tbWorkBook, tbWorkBookDest, sourceTag, targetSite, newDatabase )

 #See if file already exists
  @tbWorkBookPath = "c:\\tableauScripts\\Incoming\\" + tbWorkBook + ".twb"

  if !FileTest.exists?(@tbWorkBookPath)
    # If not, tabcmd GET it
	GetTBWorkBook(tbWorkBook, sourceTag)
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
  p "Modifing database connection string"
  # TabCmd Publish the report
  puts "Publishing " + tbWorkBookDest + " to " + targetSite
  PublishDB(@fqdnFileName,tbWorkBookDest, targetSite)
  Process.waitall
  #Clean up
  puts "Deleting temp files..."
  File.delete(@fqdnFileName)
  File.delete(@tbWorkBookPath)
end



def PublishDB(fqdnFileName,dstFileName, site)
  # When publishing to Default site, site should be a zero length string
  # Also, wrap workbook name in quotes.
  cmd = @tabCmd + " publish --overwrite -t \"" + (site == "Default" ? "" : site) + "\" -n \"" + dstFileName + "\" \"" + fqdnFileName + "\" --db-username " + @dbUser + " --db-password " + @dbPassword + " -u " + @tabUser + " -p " + @tabPassword + " --save-db-password"
  system(cmd)
  p Process.waitall
end

def GetTBWorkBook(tbWorkBookName, sourceTag)
  
  p 'Pulling workbook from GitHub'	
  tbWorkBookName = tbWorkBookName + '.twb'
 
  # Get workbook from gitHub
  api_response = Octokit.contents @gitHubRepo, :ref => sourceTag, :path => tbWorkBookName

  text_contents = Base64.decode64( api_response.content )

  # Save File
  File.open('c:\\tableauscripts\\incoming\\' + tbWorkBookName , 'w') do |file| 
 	file.write(text_contents)
  end

end

def Main()
	p "Processing..."
	
	csv_contents = CSV.read(@customerListFile)
	csv_contents.shift
	csv_contents.each do |row|

	  # The GitHub User/Repo which contains the report we want
	  @gitHubRepo = row[0].strip! || row[0] if row[0]	  
	  # Target Tableau Server
	  @targetServer = row[1].strip! || row[1] if row[1]	    
	  # The github TAG used to give us the version of the workbook we're about to grab
	  @sourceTag = row[2].strip! || row[2] if row[2]	  
	  # Tableau Site to be published to
	  @targetSite = row[3].strip! || row[3] if row[3]
	  # Orignal Workbook Name
	  @sourceWorkbook = row[4].strip! || row[4] if row[4]
	  # Name of workbook as published on new server
	  @targetWorkbook = row[5].strip! || row[5] if row[5]
	  # Original Data Source Name - not really used for anything, just helps to make the config file more logical.
	  @sourceDataSource = row[6].strip! || row[6] if row[6]
	  # New Data Source Name (Or Server Name or database)
	  @targetDataSource = row[7].strip! || row[7] if row[7]
	  # New extract file name to be added to twbx as substitute for current file (type "extract" only)
	  @targetExtractName = row[8].strip! || row[8] if row[8]	  


	  ChangeDB(@sourceWorkbook, @targetWorkbook, @sourceTag, @targetSite, @targetDataSource )
	end

	  system(@tabLogoff)
end

Main()
