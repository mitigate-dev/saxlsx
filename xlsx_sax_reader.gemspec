# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xlsx_sax_reader/version'

Gem::Specification.new do |spec|
  spec.name          = "xlsx_sax_reader"
  spec.version       = XlsxSaxReader::VERSION
  spec.authors       = ["Gabriel Naiman"]
  spec.email         = ["gabynaiman@gmail.com"]
  spec.description   = 'Fast xlsx reader on top of Ox SAX parser'
  spec.summary       = 'Fast xlsx reader on top of Ox SAX parser'
  spec.homepage      = "https://github.com/gabynaiman/xlsx_sax_reader"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency 'rubyzip', '~> 1.0'
  spec.add_dependency "ox"
  spec.add_dependency "parallel_enumerable"

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rspec"
  spec.add_development_dependency 'simplecov'
  spec.add_development_dependency 'ruby-prof'

end
