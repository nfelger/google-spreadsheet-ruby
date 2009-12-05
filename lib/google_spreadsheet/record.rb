module GoogleSpreadsheet
  # Use GoogleSpreadsheet::Table#records to get GoogleSpreadsheet::Record objects.
  class Record < Hash

      include(Util)

      def initialize(session, entry) #:nodoc:
        @session = session
        for field in entry.children.collect.delete_if { |x| x.name != "field" }
          self[as_utf8(field["name"])] = as_utf8(field.inner_text)
        end
      end

      def inspect #:nodoc:
        content = self.map(){ |k, v| "%p => %p" % [k, v] }.join(", ")
        return "\#<%p:{%s}>" % [self.class, content]
      end

  end
end