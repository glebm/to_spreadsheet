# encoding: UTF-8
require 'rubygems'
begin
  require 'bundler/setup'
rescue LoadError
  puts 'You must `gem install bundler` and `bundle install` to run rake tasks'
end

require 'rake'
require 'rake/rdoctask'

require 'rake/testtask'

Rake::TestTask.new(:test) do |t|
  Dir.chdir('test/test_app')
  Kernel.exec('rake db:test:prepare')
  Kernel.exec('rake test')
  Dir.chdir('../..')
end

task :default => :test