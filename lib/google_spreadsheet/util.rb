module GoogleSpreadsheet

  module Util #:nodoc:

    module_function

    def http_request(method, url, data, header = {})
      uri = URI.parse(url)
      http = Net::HTTP.new(uri.host, uri.port)
      http.use_ssl = uri.scheme == "https"
      http.verify_mode = OpenSSL::SSL::VERIFY_NONE
      http.start() do
        path = uri.path + (uri.query ? "?#{uri.query}" : "")
        if method == :delete
          response = http.__send__(method, path, header)
        else
          response = http.__send__(method, path, data, header)
        end
        if !(response.code =~ /^2/)
          raise(GoogleSpreadsheet::Error, "Response code #{response.code} for POST #{url}: " +
            CGI.unescapeHTML(response.body))
        end
        return response.body
      end
    end

    def encode_query(params)
      return params.map(){ |k, v| uri_encode(k) + "=" + uri_encode(v) }.join("&")
    end

    def uri_encode(str)
      return URI.encode(str, /#{URI::UNSAFE}|&/n)
    end

    def h(str)
      return CGI.escapeHTML(str.to_s())
    end

    def as_utf8(str)
      str.force_encoding("UTF-8")
    end

    # Converts a string that indicates a cell in its position.
    # Ex: A1 => [1,1]; B1 => [1,2]; => z32 => [32,26]
    def string_to_position(string)
      rx = string.upcase.match(/([A-Z]+)(\d+)/)
      letter, number = rx[1..2]

      row = number.to_i
      col = 0
      letter.each_byte { |b| col *= 26; col += b-64 }

      return [row, col]
    end

  end
end