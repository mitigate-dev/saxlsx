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

creek                  2.950000   0.060000   3.010000 (  3.243032)
rubyXL                 4.050000   0.130000   4.180000 (  4.351446)
saxlsx                 0.850000   0.020000   0.870000 (  0.918124)
simple_xlsx_reader     2.100000   0.040000   2.140000 (  2.255457)

Inline Strings

creek                  3.570000   0.300000   3.870000 (  4.000918)
rubyXL                 4.520000   0.270000   4.790000 (  4.827204)
saxlsx                 1.360000   0.030000   1.390000 (  1.388722)
simple_xlsx_reader     2.220000   0.030000   2.250000 (  2.290375)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
