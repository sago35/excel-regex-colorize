# excel-regex-colorize

Colorize string by regex

## Usage

See demo.

![](https://raw.githubusercontent.com/wiki/sago35/excel-regex-colorize/demo.gif)

## Install

Use `bin\excel-regex-colorize-installer.xls` to nstall and uninstall.

## Development

    # Decombine macro files to src/*
    $ cscript //nologo vbac.wsf decombine

    # Edit macro files
    $ vim src\excel-regex-colorize.xla\regex_colorize.bas

    # Combine macro files
    $ cscript //nologo vbac.wsf combine

    # Commit your changes
    $ git --all .

### Requirement

  * [vbac.wsf](https://github.com/vbaidiot/Ariawase)

## Licence

[MIT](http://opensource.org/licenses/mit-license.php)

## Author

[sago35](https://github.com/sago35)
