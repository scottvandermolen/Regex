# Regular Expressions Library for ASP

Perl-Compatible Regular Expressions are part of the PHP core. ASP also supports regular expressions through the RegExp object. This function library brings the various regular expressions functions from PHP to ASP.

## Project Status

No further development is currently planned, as this is considered complete.

## Installation

Place Regex.lib.asp in any location on your web server, or on another machine on the same network. For additional security, you may wish to place it in a location that isn't directly accessible by users.

The file Regex.test.asp is not required in order to use the library and does not need to be placed on the web server unless you want to run unit tests.

## Usage

```vbscript
<!--#include file="Regex.lib.asp"-->
<%
dim sentence
sentence = "The quick brown fox jumps over the lazy dog."

' Displays the number 1
Response.Write preg_match("/dog/", sentence)

' Displays the number 2
Response.Write preg_match_all("/the/i", sentence)

' Changes the word brown to red.
sentence = preg_replace("/brown/", "red", sentence)
%>
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

See Regex.test.asp for unit tests.

## Authors

Version 1.0 written May 2009 by Scott Vander Molen (based on POSIX Regular Expression functions, deprecated in PHP 5.3.0 and removed in PHP 7.0.0)

Version 2.0 written January 2024 by Scott Vander Molen (based on Perl-Compatible Regular Expression functions)

## License
This work is licensed under a [Creative Commons Attribution 4.0 International License](https://creativecommons.org/licenses/by/4.0/).