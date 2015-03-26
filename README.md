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
creek                  2.610000   0.060000   2.670000 (  2.704594)
rubyXL                 3.830000   0.130000   3.960000 (  3.985651)
saxlsx                 0.750000   0.010000   0.760000 (  0.785445)
simple_xlsx_reader     1.870000   0.040000   1.910000 (  1.940999)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
