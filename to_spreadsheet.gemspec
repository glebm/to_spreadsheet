$LOAD_PATH << File.join(File.dirname(__FILE__), 'lib')
require 'to_spreadsheet/version'

include_files = ["README*", "LICENSE", "Rakefile", "init.rb",
                 "{lib,tasks,test,rails,generators,shoulda_macros}/**/*"].map do |glob|
  Dir[glob]
end.flatten

exclude_files = ["**/*.rbc", "test/s3.yml", "test/debug.log", "test/to_spreadsheet.db", "test/doc", "test/doc/*",
                 "test/pkg", "test/pkg/*", "test/tmp", "test/tmp/*"].map do |glob|
  Dir[glob]
end.flatten


Gem::Specification.new do |s|
  s.name              = "to_spreadsheet"
  s.email             = "glex.spb@gmail.com"
  s.author            = "Gleb Mazovetskiy"
  s.homepage          = "https://github.com/glebm/to_spreadsheet"
  s.summary           = "Render existing views as Excel documents with style!"
  s.description       = "Render XLSX from Rails using existing views ( .*.html => .xlsx )"
  s.files             = Dir["lib/**/*"] + ["MIT-LICENSE", "Rakefile", "README.rdoc"]
  s.version           = ToSpreadsheet::VERSION
  s.platform          = Gem::Platform::RUBY
  s.files             = include_files - exclude_files
  s.require_path      = "lib"
  s.test_files        = Dir["spec/**/*_spec.rb"]
  s.has_rdoc          = true
  s.extra_rdoc_files  = Dir["README*"]
  s.add_dependency 'rails'
  s.add_dependency 'nokogiri'
  s.add_dependency 'axlsx'
  s.add_dependency 'chronic'
  s.add_development_dependency 'haml-rails'
  s.add_development_dependency 'rspec-rails'
end
