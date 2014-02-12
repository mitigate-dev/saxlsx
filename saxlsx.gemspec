# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'saxlsx/version'

Gem::Specification.new do |spec|
  spec.name          = "saxlsx"
  spec.version       = Saxlsx::VERSION
  spec.authors       = ["Edgars Beigarts"]
  spec.email         = ["edgars.beigarts@makit.lv"]
  spec.description   = 'Fast xlsx reader on top of Ox SAX parser'
  spec.summary       = 'Fast xlsx reader on top of Ox SAX parser'
  spec.homepage      = "https://github.com/mak-it/saxlsx"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^spec/})
  spec.require_paths = ["lib"]

  spec.add_dependency 'rubyzip', '~> 1.0'
  spec.add_dependency 'ox'

  spec.add_development_dependency 'bundler', "~> 1.5"
  spec.add_development_dependency 'rake'
  spec.add_development_dependency 'rspec'
  spec.add_development_dependency 'simplecov'
end
