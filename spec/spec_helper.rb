$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
$LOAD_PATH.unshift(File.dirname(__FILE__))

require "rubygems"
require 'officex2str'
require "rspec"
require 'nokogiri'
require 'zipruby'
require 'kconv'

# Requires supporting files with custom matchers and macros, etc,
# in ./support/ and its subdirectories.
#Dir["#{File.dirname(__FILE__)}/support/**/*.rb"].each {|f| require f}

RSpec.configure do |config|

end
