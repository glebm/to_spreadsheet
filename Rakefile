# encoding: UTF-8
require 'rubygems'
require 'bundler/setup'

require 'rake'
require 'rdoc/task'
require 'rspec/core/rake_task'
RSpec::Core::RakeTask.new(:spec)

task :generate_pdf do
  Haml::Engine.new(File.read('spec/support/table.html.haml').render
  xls_io = ToSpreadsheet::XLS.to_io(html)
  f = File.open('/tmp/spreadsheet.xls', 'wb')
  Spreadsheet.writer(f).write_workbook(spreadsheet, f)
  f.close
end