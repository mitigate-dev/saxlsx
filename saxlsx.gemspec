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

  spec.required_ruby_version = '>= 2.1.0'

  spec.add_dependency 'rubyzip', '~> 1.0'
  spec.add_dependency 'ox', '~> 2.1'

  spec.add_development_dependency 'bundler', ">= 1.5"
  spec.add_development_dependency 'rake', '~> 10.1'
  spec.add_development_dependency 'rspec', '~> 2.14'
  spec.add_development_dependency 'simplecov', '~> 0.8'
end
