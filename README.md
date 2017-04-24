# Saxlsx

[![Join the chat at https://gitter.im/mak-it/saxlsx](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/mak-it/saxlsx?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

[![Build Status](https://travis-ci.org/mak-it/saxlsx.svg?branch=master)](https://travis-ci.org/mak-it/saxlsx)

**Fast** and memory efficient XLSX reader on top of Ox SAX parser.

It reads row by row and doesn't store the whole sheet in memory, so this
approach is more suitable when parsing big files.

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
Saxlsx::Workbook.open filename, auto_format: true do |w|
  w.sheets.each do |s|
    puts s.rows.count
    s.rows.each do |r|
      puts r.inspect
    end
  end
end
```

By default `saxlsx` will try to convert `General` type cells that look like
numbers to ruby floats or integers. You can disable this feature
using `auto_format: false`.

## How fast is it?

```bash
$ rake bench
```

ruby 2.2.1 on OS X 10.10

```
Shared Strings

                           user     system      total        real
creek                  3.000000   0.070000   3.070000 (  3.254653)
dullard                2.070000   0.030000   2.100000 (  2.109380)
oxcelix                1.100000   0.020000   1.120000 (  1.128109)
roo                   11.980000   0.350000  12.330000 ( 12.473732)
rubyXL                 3.560000   0.070000   3.630000 (  3.641981)
saxlsx                 0.660000   0.040000   0.700000 (  0.693199)
simple_xlsx_reader     1.650000   0.030000   1.680000 (  1.686309)

Inline Strings

                           user     system      total        real
creek                  2.960000   0.260000   3.220000 (  3.304050)
dullard                2.380000   0.220000   2.600000 (  2.639881)
oxcelix                   ERROR
roo                   11.620000   0.260000  11.880000 ( 12.136643)
rubyXL                 4.060000   0.480000   4.540000 (  4.612416)
saxlsx                 0.870000   0.210000   1.080000 (  1.091133)
simple_xlsx_reader     1.910000   0.040000   1.950000 (  1.961555)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
