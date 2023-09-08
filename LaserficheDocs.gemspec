# coding: utf-8

Gem::Specification.new do |spec|
  spec.name          = "LaserficheDocs"
  spec.version       = "0.0.1"
  spec.authors       = ["Robert Fulton"]
  spec.email         = ["robert.fulton@laserfiche.com"]

  spec.summary       = %q{Jekyll theme for Laserfiche Technical Documentation based on Just-The-Docs}
  spec.homepage      = "https://github.com/Laserfiche/laserfiche-sharepoint-integration"
  spec.license       = "MIT"
  spec.metadata      = {
    "bug_tracker_uri"   => "https://github.com/Laserfiche/laserfiche-sharepoint-integration/issues",
    "changelog_uri"     => "https://github.com/Laserfiche/laserfiche-sharepoint-integration/CHANGELOG.md",
    "documentation_uri" => "https://laserfiche.github.io/laserfiche-sharepoint-integration/"
    "source_code_uri"   => "https://github.com/Laserfiche/laserfiche-sharepoint-integration",
  }

  # spec.files         = `git ls-files -z ':!:*.jpg' ':!:*.png'`.split("\x0").select { |f| f.match(%r{^(assets|bin|_layouts|_includes|lib|Rakefile|_sass|LICENSE|README|CHANGELOG|favicon)}i) }
  # spec.executables   << 'just-the-docs'

  spec.add_development_dependency "bundler", ">= 2.3.5"
  spec.add_runtime_dependency "jekyll", ">= 3.8.5"
  spec.add_runtime_dependency "jekyll-seo-tag", ">= 2.0"
  spec.add_runtime_dependency "jekyll-include-cache"
  spec.add_runtime_dependency "rake", ">= 12.3.1"
end
