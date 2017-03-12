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
