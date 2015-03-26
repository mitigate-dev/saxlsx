# Saxlsx

[![Build Status](https://travis-ci.org/mak-it/saxlsx.png?branch=master)](https://travis-ci.org/mak-it/saxlsx)

**Fast** and memory efficient XLSX reader on top of Ox SAX parser.

It reads row by row and doesn't store the whole sheet in memory, so this
approach is more suitable when parsing big files. This also means that functions
and references will not work, as this style of parsing doesn't know
anything about other rows.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'saxlsx'
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
Saxlsx::Workbook.open filename do |w|
  w.sheets.each do |s|
    s.rows.each do |r|
      puts r.inspect
    end
  end
end
```

## How fast is it?

```bash
$ rake bench
```

```
Shared Strings

creek                  2.970000   0.070000   3.040000 (  3.254889)
oxcelix                1.290000   0.050000   1.340000 (  1.393971)
rubyXL                 3.910000   0.120000   4.030000 (  4.241886)
saxlsx                 0.850000   0.020000   0.870000 (  0.910578)
simple_xlsx_reader     2.150000   0.040000   2.190000 (  2.296939)

Inline Strings

creek                  4.040000   0.300000   4.340000 (  4.406562)
oxcelix                ERROR
rubyXL                 5.100000   0.280000   5.380000 (  5.449917)
saxlsx                 1.530000   0.080000   1.610000 (  1.628584)
simple_xlsx_reader     1.840000   0.030000   1.870000 (  1.904601)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
