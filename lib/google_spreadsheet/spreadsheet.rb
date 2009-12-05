module GoogleSpreadsheet
  # Use methods in GoogleSpreadsheet::Session to get GoogleSpreadsheet::Spreadsheet object.
  class Spreadsheet

      include(Util)

      def initialize(session, worksheets_feed_url, title = nil) #:nodoc:
        @session = session
        @worksheets_feed_url = worksheets_feed_url
        @title = title
      end

      # URL of worksheet-based feed of the spreadsheet.
      attr_reader(:worksheets_feed_url)

      # Title of the spreadsheet. So far only available if you get this object by
      # GoogleSpreadsheet::Session#spreadsheets.
      attr_reader(:title)

      # Key of the spreadsheet.
      def key
        if !(@worksheets_feed_url =~
            %r{http://spreadsheets.google.com/feeds/worksheets/(.*)/private/full})
          raise(GoogleSpreadsheet::Error,
            "worksheets feed URL is in unknown format: #{@worksheets_feed_url}")
        end
        return $1
      end

      # Tables feed URL of the spreadsheet.
      def tables_feed_url
        return "http://spreadsheets.google.com/feeds/#{self.key}/tables"
      end

      # Returns worksheets of the spreadsheet as array of GoogleSpreadsheet::Worksheet.
      def worksheets
        doc = @session.get(@worksheets_feed_url)
        result = []
        for entry in doc.search("entry")
          title = as_utf8(entry.search("title").text)
          url = as_utf8(entry.search(
            "link[@rel='http://schemas.google.com/spreadsheets/2006#cellsfeed']")[0]["href"])
          result.push(Worksheet.new(@session, self, url, title))
        end
        return result.freeze()
      end

      # Adds a new worksheet to the spreadsheet. Returns added GoogleSpreadsheet::Worksheet.
      def add_worksheet(title, max_rows = 100, max_cols = 20)
        xml = <<-"EOS"
          <entry xmlns='http://www.w3.org/2005/Atom'
                 xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
            <title>#{h(title)}</title>
            <gs:rowCount>#{h(max_rows)}</gs:rowCount>
            <gs:colCount>#{h(max_cols)}</gs:colCount>
          </entry>
        EOS
        doc = @session.post(@worksheets_feed_url, xml)
        url = as_utf8(doc.search(
          "link[@rel='http://schemas.google.com/spreadsheets/2006#cellsfeed']")[0]["href"])
        return Worksheet.new(@session, self, url, title)
      end

      # Returns list of tables in the spreadsheet.
      def tables
        doc = @session.get(self.tables_feed_url)
        return doc.search("entry").map(){ |e| Table.new(@session, e) }.freeze()
      end

  end
end