module GoogleSpreadsheet
  # Use GoogleSpreadsheet.login or GoogleSpreadsheet.saved_session to get
  # GoogleSpreadsheet::Session object.
  class Session

    include(Util)
    extend(Util)

    # The same as GoogleSpreadsheet.login.
    def self.login(mail, password)
      session = Session.new()
      session.login(mail, password)
      return session
    end

    # Creates session object with given authentication token.
    def initialize(auth_token = nil)
      @auth_token = auth_token
    end

    # Authenticates with given +mail+ and +password+, and updates current session object
    # if succeeds. Raises GoogleSpreadsheet::AuthenticationError if fails.
    # Google Apps account is supported.
    def login(mail, password)
      begin
        @auth_token = nil
        params = {
          "accountType" => "HOSTED_OR_GOOGLE",
          "Email" => mail,
          "Passwd" => password,
          "service" => "wise",
          "source" => "Gimite-RubyGoogleSpreadsheet-1.00",
        }
        response = http_request(:post,
          "https://www.google.com/accounts/ClientLogin", encode_query(params))
        @auth_token = response.slice(/^Auth=(.*)$/, 1)
      rescue GoogleSpreadsheet::Error => ex
        return true if @on_auth_fail && @on_auth_fail.call()
        raise(AuthenticationError, "authentication failed for #{mail}: #{ex.message}")
      end
    end

    # Authentication token.
    attr_accessor(:auth_token)

    # Proc or Method called when authentication has failed.
    # When this function returns +true+, it tries again.
    attr_accessor(:on_auth_fail)

    def get(url) #:nodoc:
      while true
        begin
          response = open(url, self.http_header){ |f| f.read() }
        rescue OpenURI::HTTPError => ex
          if ex.message =~ /^401/ && @on_auth_fail && @on_auth_fail.call()
            next
          end
          raise(ex.message =~ /^401/ ? AuthenticationError : GoogleSpreadsheet::Error,
            "Error #{ex.message} for GET #{url}: " + ex.io.read())
        end
        return Nokogiri.XML(response)
      end
    end

    def post(url, data) #:nodoc:
      header = self.http_header.merge({"Content-Type" => "application/atom+xml"})
      response = http_request(:post, url, data, header)
      return Nokogiri.XML(response)
    end

    def put(url, data) #:nodoc:
      header = self.http_header.merge({"Content-Type" => "application/atom+xml"})
      response = http_request(:put, url, data, header)
      return Nokogiri.XML(response)
    end

    def delete(url) #:nodoc:
      header = self.http_header.merge({"Content-Type" => "application/atom+xml"})
      response = http_request(:delete, url, nil, header)
      return Nokogiri.XML(response)
    end

    def http_header #:nodoc:
      return {"Authorization" => "GoogleLogin auth=#{@auth_token}"}
    end

    # Returns list of spreadsheets for the user as array of GoogleSpreadsheet::Spreadsheet.
    # You can specify query parameters described at
    # http://code.google.com/apis/spreadsheets/docs/2.0/reference.html#Parameters
    #
    # e.g.
    #   session.spreadsheets
    #   session.spreadsheets("title" => "hoge")
    def spreadsheets(params = {})
      query = encode_query(params)
      doc = get("http://spreadsheets.google.com/feeds/spreadsheets/private/full?#{query}")
      result = []
      for entry in doc.search("entry")
        title = as_utf8(entry.search("title").text)
        url = as_utf8(entry.search(
          "link[@rel='http://schemas.google.com/spreadsheets/2006#worksheetsfeed']")[0]["href"])
        result.push(Spreadsheet.new(self, url, title))
      end
      return result
    end

    # Returns GoogleSpreadsheet::Spreadsheet with given +key+.
    #
    # e.g.
    #   # http://spreadsheets.google.com/ccc?key=pz7XtlQC-PYx-jrVMJErTcg&hl=ja
    #   session.spreadsheet_by_key("pz7XtlQC-PYx-jrVMJErTcg")
    def spreadsheet_by_key(key)
      url = "http://spreadsheets.google.com/feeds/worksheets/#{key}/private/full"
      return Spreadsheet.new(self, url)
    end

    # Returns GoogleSpreadsheet::Spreadsheet with given +url+. You must specify either of:
    # - URL of the page you open to access the spreadsheet in your browser
    # - URL of worksheet-based feed of the spreadsheet
    #
    # e.g.
    #   session.spreadsheet_by_url(
    #     "http://spreadsheets.google.com/ccc?key=pz7XtlQC-PYx-jrVMJErTcg&hl=en")
    #   session.spreadsheet_by_url(
    #     "http://spreadsheets.google.com/feeds/worksheets/pz7XtlQC-PYx-jrVMJErTcg/private/full")
    def spreadsheet_by_url(url)
      # Tries to parse it as URL of human-readable spreadsheet.
      uri = URI.parse(url)
      if uri.host == "spreadsheets.google.com" && uri.path =~ /\/ccc$/
        if (uri.query || "").split(/&/).find(){ |s| s=~ /^key=(.*)$/ }
          return spreadsheet_by_key($1)
        end
      end
      # Assumes the URL is worksheets feed URL.
      return Spreadsheet.new(self, url)
    end

    # Returns GoogleSpreadsheet::Worksheet with given +url+.
    # You must specify URL of cell-based feed of the worksheet.
    #
    # e.g.
    #   session.worksheet_by_url(
    #     "http://spreadsheets.google.com/feeds/cells/pz7XtlQC-PYxNmbBVgyiNWg/od6/private/full")
    def worksheet_by_url(url)
      return Worksheet.new(self, nil, url)
    end

  end
end