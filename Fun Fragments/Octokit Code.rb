require 'octokit'

api_response = Octokit.contents 'russch/ReportsToDeploy', 
  :ref => 'Tableau_Rocks_v1.1', :path => 'Tableau Rocks.twb'
text_contents = Base64.decode64( api_response.content )

File.open('Tableau Rocks.twb', 'w') do |file| 
	file.write(text_contents)
end





  