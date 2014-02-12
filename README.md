# Saxlsx

[![Build Status](https://travis-ci.org/mak-it/saxlsx.png?branch=master)](https://travis-ci.org/mak-it/saxlsx)

Fast XLSX reader on top of Ox SAX parser.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'saxlsx
```

And then execute:

```bash
$ bundle
```

Or install it yourself as:

```bash
$ gem install saxlsx
```

## Usage

```ruby
Saxslsx::Workbook.open filename do |w|
  w.sheets.each do |s|
    s.rows.each do |r|
      puts r.inspect
    end
  end
end
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
