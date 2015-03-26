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
Saxlsx::Workbook.open filename do |w|
  w.sheets.each do |s|
    puts s.rows.count
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

creek                  3.200000   0.060000   3.260000 (  3.392446)
dullard                3.020000   0.070000   3.090000 (  3.154582)
oxcelix                1.720000   0.060000   1.780000 (  1.824003)
roo                   14.580000   0.490000  15.070000 ( 15.311664)
rubyXL                 5.180000   0.160000   5.340000 (  5.432973)
saxlsx                 1.020000   0.010000   1.030000 (  1.044249)
simple_xlsx_reader     2.580000   0.050000   2.630000 (  2.660065)

Inline Strings

creek                  3.970000   0.330000   4.300000 (  4.395139)
dullard                3.700000   0.360000   4.060000 (  4.174130)
oxcelix                ERROR
roo                   14.700000   0.550000  15.250000 ( 15.644152)
rubyXL                 5.830000   0.440000   6.270000 (  6.365362)
saxlsx                 1.500000   0.040000   1.540000 (  1.679372)
simple_xlsx_reader     2.790000   0.070000   2.860000 (  2.916020)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
