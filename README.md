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

ruby 2.4.1 on OS X 10.12

```
Shared Strings

                           user     system      total        real
creek                  2.310000   0.040000   2.350000 (  2.414795)
dullard                1.660000   0.050000   1.710000 (  1.753160)
oxcelix                1.100000   0.060000   1.160000 (  1.169371)
roo                    2.770000   0.120000   2.890000 (  2.943509)
rubyXL                 3.570000   0.090000   3.660000 (  3.716661)
saxlsx                 0.620000   0.040000   0.660000 (  0.665471)
simple_xlsx_reader     1.590000   0.030000   1.620000 (  1.643398)

Inline Strings

                           user     system      total        real
creek                  2.840000   0.200000   3.040000 (  3.335055)
dullard                2.200000   0.250000   2.450000 (  2.613793)
oxcelix                   ERROR
roo                    3.550000   0.160000   3.710000 (  3.749528)
rubyXL                 3.890000   0.390000   4.280000 (  4.340749)
saxlsx                 0.930000   0.240000   1.170000 (  1.173783)
simple_xlsx_reader     1.760000   0.040000   1.800000 (  1.818311)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
