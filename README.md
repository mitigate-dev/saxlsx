# Saxlsx

[![Join the chat at https://gitter.im/mak-it/saxlsx](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/mak-it/saxlsx?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

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

ruby 2.7 on OS X

```
Shared Strings

                           user     system      total        real
creek                  1.296539   0.029374   1.325913 (  1.340820)
dullard                1.178981   0.025073   1.204054 (  1.221381)
oxcelix                0.985258   0.025028   1.010286 (  1.023730)
roo                    0.971155   0.029964   1.001119 (  1.016452)
rubyXL                 2.979334   0.055708   3.035042 (  3.079301)
saxlsx                 0.473398   0.011342   0.484740 (  0.490247)
simple_xlsx_reader     1.209074   0.024579   1.233653 (  1.249957)

Inline Strings

                           user     system      total        real
creek                  1.471115   0.075182   1.546297 (  1.567045)
dullard                1.338499   0.085116   1.423615 (  1.443386)
oxcelix                ERROR
roo                    1.133878   0.052834   1.186712 (  1.208369)
rubyXL                 3.213630   0.070255   3.283885 (  3.324428)
saxlsx                 0.667601   0.024265   0.691866 (  0.696603)
simple_xlsx_reader     1.350298   0.028411   1.378709 (  1.396583)
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
