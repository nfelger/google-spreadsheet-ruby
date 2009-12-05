module GoogleSpreadsheet
# Use GoogleSpreadsheet::Worksheet#add_table to create table.
# Use GoogleSpreadsheet::Worksheet#tables to get GoogleSpreadsheet::Table objects.
  class Table
    include(Util)

    def initialize(session, entry) #:nodoc:
      @columns = {}
      @worksheet_title = as_utf8(entry.search("//gs:worksheet")[0]["name"])
      @records_url = as_utf8(entry.search("content")[0]["src"])
      @session = session
    end

    # Title of the worksheet the table belongs to.
    attr_reader(:worksheet_title)

    # Adds a record.
    def add_record(values)
      fields = ""
      values.each do |name, value|
        fields += "<gs:field name='#{h(name)}'>#{h(value)}</gs:field>"
      end
      xml =<<-EOS
        <entry
            xmlns="http://www.w3.org/2005/Atom"
            xmlns:gs="http://schemas.google.com/spreadsheets/2006">
          #{fields}
        </entry>
      EOS
      @session.post(@records_url, xml)
    end

    # Returns records in the table.
    def records
      doc = @session.get(@records_url)
      return doc.search("entry").map(){ |e| Record.new(@session, e) }
    end

  end
end
