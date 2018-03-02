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
