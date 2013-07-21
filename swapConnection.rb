require 'rexml/document'
require 'csv'
include REXML

# Globals
@zip = "c:\\Program Files (x86)\\7-Zip\\7z.exe e -y -oc:\\tableauScripts\\Extract"
@tabCmd = "C:\\Program Files (x86)\\Tableau\\Tableau Server\\8.0\\bin\\tabcmd.exe"
@tabUser = "russch"
@tabPassword = "tableau"
@tabServer = "http://localhost"
@tabLogin = @tabCmd + " login -s " + @tabServer + " -u " + @tabUser + " -p " + @tabPassword
@customerListFile = "C:\\tableauScripts\\swapConnection\\config.csv"
@dbUser = "russch"
@dbPassword = "tableau"
@tabLogoff = @tabCmd + " logout"

# Login
system(@tabLogin)
Process.waitall

def ChangeExt(tbWorkBook, existingExt, newExt, dstFileName, targetSite)
  @tbWorkBookPath = "c:\\tableauScripts\\Extract\\" + tbWorkBook + ".twb"

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

    @fqdnFileName = "c:\\tableauScripts\\" + dstFileName

    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName, 'w') do |result|
      formatter.write(twb, result)
    end
  end
  
  Publish(@fqdnFileName, targetSite)
  Process.waitall
  File.delete(@fqdnFileName)
end

def ChangeDB(tbWorkBook, newDatabase, dstFileName, targetSite)
  @tbWorkBookPath = "c:\\tableauScripts\\" + tbWorkBook + ".twb"

  if !FileTest.exists?(@tbWorkBookPath)
    GetTBWorkBook(tbWorkBook)
  end

  File.open(@tbWorkBookPath) do |config|
    twb = Document.new(config)
    twb.elements.each("*/datasources/datasource/connection") do |element| 
      element.attributes["dbname"]="#{newDatabase}" 
    end
 
    @fqdnFileName = "c:\\tableauScripts\\" + dstFileName + ".twb"
 
    formatter = REXML::Formatters::Default.new
    File.open(@fqdnFileName, 'w') do |result|
      formatter.write(twb, result)
    end
  end
  PublishDB(@fqdnFileName, targetSite)
  Process.waitall
  File.delete(@fqdnFileName)
end

def Publish(dstFileName, site)
  cmd = @tabCmd + " publish --overwrite -t " + site + " -n " + dstFileName + " " + dstFileName + " -u " + @tabUser + " -p " + @tabPassword
  system(cmd)
  Process.waitall
end

def PublishDB(dstFileName, site)
  cmd = @tabCmd + " publish --overwrite -t " + site + " -n " + dstFileName + " " + dstFileName + " --db-username " + @dbUser + " --db-password " + @dbPassword + " -u " + @tabUser + " -p " + @tabPassword + " --save-db-password"
  system(cmd)
  Process.waitall
end

def GetTBWorkBook(tbWorkBookName)
  cmd = @tabCmd + " get /workbooks/" + tbWorkBookName + ".twb -f C:\\tableauScripts\\" + tbWorkBookName + ".twb"
  system(cmd)
  Process.waitall
end

csv_contents = CSV.read(@customerListFile)
csv_contents.shift
csv_contents.each do |row|
  @row1 = row[1].strip! || row[1] if row[1]
  @row2 = row[2].strip! || row[2] if row[2]
  @row3 = row[3].strip! || row[3] if row[3]
  @row4 = row[4].strip! || row[4] if row[4]
  @row5 = row[5].strip! || row[5] if row[5]
  
  if row[0] == "extract"
    ChangeExt(@row1, @row2, @row3, @row4, @row5)
  elsif row[0] == "database"
    ChangeDB(@row1, @row3, @row4, @row5)
  end
end

system(@tabLogoff)
File.delete(@tbWorkBookPath)