# MagicSymbolsVBA
I needed a simple yet very robust function that would replace placeholder symbols with dynamic data.

Example: Input: `"Username %u, date is %d."`
The function would replace `%d` with the current date and `%u` with the current username.

Sounds simple to implement, but here’s a twist: **Regexps** and **Replace()** are not safe. In the example above, if `%u` after replacement would contain another token (let’s say `“Andy%d“`), there will be a recursive replacement and the output will mangled: `“Username Andy20170820, date is 20170820“`.

I could write this efficiently in C++ or some other *“proper”* language but I’m relegated to the VBA. Asked on the StackOverflow but nobody was able to provide any proper answer.

Hence, decided to deal with it myself and write the function.

Enjoy!

StackOverflow link for reference:
https://stackoverflow.com/questions/48460409/safely-and-efficiently-parse-and-replace-tokens-in-a-string

Usage:

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Replace magic symbols (placeholders) with dynamic data.
''
'' Arguments: a string full of magic.
''
'' Placeholders consist of one symbol prepended with a %:
''    %d - current date
''    %t - current time
''    %u - username (user ID)
''    %n - full user name (usually name and surname)
''    %% - literal % (placeholder escape)
''    Using an unsupported magic symbol will treat the % literally, as if it had been escaped.
''    A single placeholder terminating the string will also be treated literally.
''    Magic symbols are case-sensitive.
''
'' Returns:   A string with no magic but with lots of beauty.
''
'' Examples:
'' "Today is %d" becomes "Today is 2018-01-26"
'' "Beautiful time: %%%t%%" yields "Beautiful time: %16:10:51%"
'' "There are %zero% magic symbols %here%.", true to its message, outputs "There are %zero% magic symbols %here%."
'' "%%% looks lovely %%%" would show "%% looks lovely %%" - one % for the escaped "%%" and the second one for the unused "%"!
''
'' Alexander Ivashkin, 26 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
