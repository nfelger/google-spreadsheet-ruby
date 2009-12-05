# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "set"
require "net/https"
require "open-uri"
require "cgi"
require "rubygems"
require "nokogiri"

Net::HTTP.version_1_2
Nokogiri::XML::Node::Encoding = "UTF8"

if RUBY_VERSION < "1.9.0"
  class String
    def force_encoding(encoding)
      return self
    end
  end
end

$:.push(File.expand_path(File.dirname(__FILE__)))

require 'google_spreadsheet/util'
require 'google_spreadsheet/session'
require 'google_spreadsheet/spreadsheet'
require 'google_spreadsheet/table'
require 'google_spreadsheet/record'

module GoogleSpreadsheet
    
    # Authenticates with given +mail+ and +password+, and returns GoogleSpreadsheet::Session
    # if succeeds. Raises GoogleSpreadsheet::AuthenticationError if fails.
    # Google Apps account is supported.
    def self.login(mail, password)
      return Session.login(mail, password)
    end
    
    # Restores GoogleSpreadsheet::Session from +path+ and returns it.
    # If +path+ doesn't exist or authentication has failed, prompts mail and password on console,
    # authenticates with them, stores the session to +path+ and returns it.
    #
    # This method requires Ruby/Password library: http://www.caliban.org/ruby/ruby-password.shtml
    def self.saved_session(path = ENV["HOME"] + "/.ruby_google_spreadsheet.token")
      session = Session.new(File.exist?(path) ? File.read(path) : nil)
      session.on_auth_fail = proc() do
        begin
          require "highline2"
        rescue LoadError
          raise(LoadError,
            "GoogleSpreadsheet.saved_session requires Highline library.\n" +
            "Run\n" +
            "  \$ sudo gem install highline\n" +
            "to install it.")
        end
        highline = HighLine.new()
        mail = highline.ask("Mail: ")
        password = highline.ask("Password: "){ |q| q.echo = false }
        session.login(mail, password)
        open(path, "w", 0600){ |f| f.write(session.auth_token) }
        true
      end
      if !session.auth_token
        session.on_auth_fail.call()
      end
      return session
    end
    
    # Raised when spreadsheets.google.com has returned error.
    class Error < RuntimeError;end

    # Raised when GoogleSpreadsheet.login has failed.
    class AuthenticationError < GoogleSpreadsheet::Error;end
    
    # Use GoogleSpreadsheet::Spreadsheet#worksheets to get GoogleSpreadsheet::Worksheet object.
    class Worksheet
        
        include(Util)
        
        def initialize(session, spreadsheet, cells_feed_url, title = nil) #:nodoc:
          @session = session
          @spreadsheet = spreadsheet
          @cells_feed_url = cells_feed_url
          @title = title

          @cells = nil
          @input_values = nil
          @modified = Set.new()
        end

        # URL of cell-based feed of the worksheet.
        attr_reader(:cells_feed_url)
        
        # URL of worksheet feed URL of the worksheet.
        def worksheet_feed_url
          # I don't know good way to get worksheet feed URL from cells feed URL.
          # Probably it would be cleaner to keep worksheet feed URL and get cells feed URL
          # from it.
          if !(@cells_feed_url =~
              %r{^http://spreadsheets.google.com/feeds/cells/(.*)/(.*)/private/full$})
            raise(GoogleSpreadsheet::Error,
              "cells feed URL is in unknown format: #{@cells_feed_url}")
          end
          return "http://spreadsheets.google.com/feeds/worksheets/#{$1}/private/full/#{$2}"
        end
        
        # GoogleSpreadsheet::Spreadsheet which this worksheet belongs to.
        def spreadsheet
          if !@spreadsheet
            if !(@cells_feed_url =~
                %r{^http://spreadsheets.google.com/feeds/cells/(.*)/(.*)/private/full$})
              raise(GoogleSpreadsheet::Error,
                "cells feed URL is in unknown format: #{@cells_feed_url}")
            end
            @spreadsheet = @session.spreadsheet_by_key($1)
          end
          return @spreadsheet
        end
        
        # Returns content of the cell as String. Top-left cell is [1, 1].
        def [](string_or_row, col=nil)
          if string_or_row.instance_of?(String)
            row, col =  string_to_position(string_or_row)
          else
            row = string_or_row
          end

          return self.cells[[row, col]] || ""
        end
        
        # Updates content of the cell.
        # Note that update is not sent to the server until you call save().
        # Top-left cell is [1, 1].
        #
        # e.g.
        #   worksheet[2, 1] = "hoge"
        #   worksheet[1, 3] = "=A1+B1"
        def []=(string_or_row, value_or_col=nil, value=nil)
          if string_or_row.instance_of?(String)
            row, col =  string_to_position(string_or_row)
            value = value_or_col
          else
            row = string_or_row
            col = value_or_col
          end

          reload() if !@cells
          @cells[[row, col]] = value
          @input_values[[row, col]] = value
          @modified.add([row, col])
          self.max_rows = row if row > @max_rows
          self.max_cols = col if col > @max_cols
        end
        
        # Returns the value or the formula of the cell. Top-left cell is [1, 1].
        #
        # If user input "=A1+B1" to cell [1, 3], worksheet[1, 3] is "3" for example and
        # worksheet.input_value(1, 3) is "=RC[-2]+RC[-1]".
        def input_value(string_or_row, col=nil)
          if string_or_row.instance_of?(String)
            row, col =  string_to_position(string_or_row)
          else
            row = string_or_row
          end

          reload() if !@cells
          return @input_values[[row, col]] || ""
        end

        # Returns the numeric value of the cell. Top-left cell is [1, 1].
        #
        # If user input "0.1" to cell [1, 3], worksheet[1, 3] is "R$ 1,23" for example.
        def numeric_value(string_or_row, col=nil)
          if string_or_row.instance_of?(String)
            row, col =  string_to_position(string_or_row)
          else
            row = string_or_row
          end

          reload() if !@cells
          return @numeric_values[[row, col]] || ""
        end
        
        # Row number of the bottom-most non-empty row.
        def num_rows
          reload() if !@cells
          return @cells.keys.map(){ |r, _| r }.max || 0
        end
        
        # Column number of the right-most non-empty column.
        def num_cols
          reload() if !@cells
          return @cells.keys.map(){ |_, c| c }.max || 0
        end
        
        # Number of rows including empty rows.
        def max_rows
          reload() if !@cells
          return @max_rows
        end
        
        # Updates number of rows.
        # Note that update is not sent to the server until you call save().
        def max_rows=(rows)
          @max_rows = rows
          @meta_modified = true
        end
        
        # Number of columns including empty columns.
        def max_cols
          reload() if !@cells
          return @max_cols
        end
        
        # Updates number of columns.
        # Note that update is not sent to the server until you call save().
        def max_cols=(cols)
          @max_cols = cols
          @meta_modified = true
        end
        
        # Title of the worksheet (shown as tab label in Web interface).
        def title
          reload() if !@title
          return @title
        end
        
        # Updates title of the worksheet.
        # Note that update is not sent to the server until you call save().
        def title=(title)
          @title = title
          @meta_modified = true
        end
        
        def cells #:nodoc:
          reload() if !@cells
          return @cells
        end
        
        # An array of spreadsheet rows. Each row contains an array of
        # columns. Note that resulting array is 0-origin so
        # worksheet.rows[0][0] == worksheet[1, 1].
        def rows(skip = 0)
          nc = self.num_cols
          result = ((1 + skip)..self.num_rows).map() do |row|
            (1..nc).map(){ |col| self[row, col] }.freeze()
          end
          return result.freeze()
        end
        
        # Reloads content of the worksheets from the server.
        # Note that changes you made by []= is discarded if you haven't called save().
        def reload()
          doc = @session.get(@cells_feed_url)
          @max_rows = doc.search("//gs:rowCount")[0].text.to_i()
          @max_cols = doc.search("//gs:colCount")[0].text.to_i()
          @title = as_utf8(doc.search("title")[0].text)
          
          @cells = {}
          @input_values = {}
          @numeric_values = {}
          for cell in doc.search("//gs:cell")
            row = cell["row"].to_i()
            col = cell["col"].to_i()
            @cells[[row, col]] = as_utf8(cell.inner_text)
            @input_values[[row, col]] = as_utf8(cell["inputValue"])
            @numeric_values[[row, col]] = cell["numericValue"].to_f if cell["numericValue"]
          end
          @modified.clear()
          @meta_modified = false
          return true
        end
        
        # Saves your changes made by []=, etc. to the server.
        def save()
          sent = false
          
          if @meta_modified
            
            ws_doc = @session.get(self.worksheet_feed_url)
            edit_url = ws_doc.search("link[@rel='edit']")[0]["href"]
            xml = <<-"EOS"
              <entry xmlns='http://www.w3.org/2005/Atom'
                     xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
                <title>#{h(self.title)}</title>
                <gs:rowCount>#{h(self.max_rows)}</gs:rowCount>
                <gs:colCount>#{h(self.max_cols)}</gs:colCount>
              </entry>
            EOS
       
            @session.put(edit_url, xml)
            
            @meta_modified = false
            sent = true
            
          end
          
          if !@modified.empty?
            
            # Gets id and edit URL for each cell.
            # Note that return-empty=true is required to get those info for empty cells.
            cell_entries = {}
            rows = @modified.map(){ |r, _| r }
            cols = @modified.map(){ |_, c| c }
            url = "#{@cells_feed_url}?return-empty=true&min-row=#{rows.min}&max-row=#{rows.max}" +
              "&min-col=#{cols.min}&max-col=#{cols.max}"
            doc = @session.get(url)
            for cell in doc.search("//gs:cell")
              row = cell["row"].to_i()
              col = cell["col"].to_i()
              cell_entries[[row, col]] = cell.parent
            end
            
            # Updates cell values using batch operation.
            xml = <<-"EOS"
              <feed xmlns="http://www.w3.org/2005/Atom"
                    xmlns:batch="http://schemas.google.com/gdata/batch"
                    xmlns:gs="http://schemas.google.com/spreadsheets/2006">
                <id>#{h(@cells_feed_url)}</id>
            EOS
            for row, col in @modified
              value = @cells[[row, col]]
              entry = cell_entries[[row, col]]
              id = entry.search("id").text
              edit_url = entry.search("link[@rel='edit']")[0]["href"]
              xml << <<-"EOS"
                <entry>
                  <batch:id>#{h(row)},#{h(col)}</batch:id>
                  <batch:operation type="update"/>
                  <id>#{h(id)}</id>
                  <link rel="edit" type="application/atom+xml"
                    href="#{h(edit_url)}"/>
                  <gs:cell row="#{h(row)}" col="#{h(col)}" inputValue="#{h(value)}"/>
                </entry>
              EOS
            end
            xml << <<-"EOS"
              </feed>
            EOS
            
            result = @session.post("#{@cells_feed_url}/batch", xml)
            for entry in result.search("atom:entry")
              interrupted = entry.search("batch:interrupted")[0]
              if interrupted
                raise(GoogleSpreadsheet::Error, "Update has failed: %s" %
                  interrupted["reason"])
              end
              if !(entry.search("batch:status")[0]["code"] =~ /^2/)
                raise(GoogleSpreadsheet::Error, "Updating cell %s has failed: %s" %
                  [entry.search("atom:id").text, entry.search("batch:status")[0]["reason"]])
              end
            end
            
            @modified.clear()
            sent = true
            
          end
          return sent
        end
        
        # Calls save() and reload().
        def synchronize()
          save()
          reload()
        end
        
        # Deletes this worksheet. Deletion takes effect right away without calling save().
        def delete
          ws_doc = @session.get(self.worksheet_feed_url)
          edit_url = ws_doc.search("link[@rel='edit']")[0]["href"]
          @session.delete(edit_url)
        end
        
        # Returns true if you have changes made by []= which haven't been saved.
        def dirty?
          return !@modified.empty?
        end
        
        # Creates table for the worksheet and returns GoogleSpreadsheet::Table.
        # See this document for details:
        # http://code.google.com/intl/en/apis/spreadsheets/docs/3.0/developers_guide_protocol.html#TableFeeds
        def add_table(table_title, summary, columns)
          column_xml = ""
          columns.each do |index, name|
            column_xml += "<gs:column index='#{h(index)}' name='#{h(name)}'/>\n"
          end

          xml = <<-"EOS"
            <entry xmlns="http://www.w3.org/2005/Atom"
              xmlns:gs="http://schemas.google.com/spreadsheets/2006">
              <title type='text'>#{h(table_title)}</title>
              <summary type='text'>#{h(summary)}</summary>
              <gs:worksheet name='#{h(self.title)}' />
              <gs:header row='1' />
              <gs:data numRows='0' startRow='2'>
                #{column_xml}
              </gs:data>
            </entry>
          EOS

          result = @session.post(self.spreadsheet.tables_feed_url, xml)
          return Table.new(@session, result)
        end
        
        # Returns list of tables for the worksheet.
        def tables
          return self.spreadsheet.tables.select(){ |t| t.worksheet_title == self.title }
        end

    end
    
    
end
