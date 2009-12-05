Gem::Specification.new do |s|
  s.name = %q{google-spreadsheet-ruby}
  s.version = "0.0.6"
  s.authors = ["Hiroshi Ichikawa"]
  s.date = %q{2009-09-26}
  s.description = %q{This is a library to read/write Google Spreadsheet.}
  s.email = ["gimite+github@gmail.com"]
  s.extra_rdoc_files = ["README.rdoc"]
  s.files = ["README.rdoc", "lib/google_spreadsheet.rb"]
  s.has_rdoc = true
  s.homepage = %q{http://github.com/nfelger/google-spreadsheet-ruby/tree/master}
  s.rdoc_options = ["--main", "README.rdoc"]
  s.require_paths = ["lib"]
  s.rubygems_version = %q{1.2.0}
  s.summary = %q{This is a library to read/write Google Spreadsheet.}

  if s.respond_to? :specification_version
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 2
  end
end
