# -*- encoding: utf-8 -*-
require File.expand_path('../lib/officex2str/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["interu"]
  gem.email         = ["interu@sonicgarden.jp"]
  gem.description   = %q{convert office 2010 files to str}
  gem.summary       = %q{convert office 2010 files(docx,xlsx,pptx) to str}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($\)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.name          = "officex2str"
  gem.require_paths = ["lib"]
  gem.version       = Officex2str::VERSION

  gem.add_development_dependency "rake", ["= 0.9.2"]
  gem.add_development_dependency "nokogiri", [">= 1.4.7"]
  gem.add_development_dependency "zipruby", ["= 0.3.6"]

end
