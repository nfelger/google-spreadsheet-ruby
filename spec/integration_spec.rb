require File.expand_path(File.dirname(__FILE__) + '/../lib/google_spreadsheet')

describe "the whole thing" do
  it "should not fuck up" do
    session = GoogleSpreadsheet.login("niko.felger@gmail.com", "coatoug7")
    session.spreadsheet_by_key("tk5zK8MY3hbNWtuxhu5IPww").worksheets[0]
  end
end
