# ?? Excel MID function to extract text from the middle of a string

## Excel MID function to extract text from the middle of a string



MID is one of the Text functions that Microsoft Excel provides for manipulating text strings. At the most basic level, it is used to extract a substring from the middle of the text string. In this tutorial, we will discuss the syntax and specificities of the Excel MID function, and then you will learn a few creative uses to accomplish challenging tasks.

* [Excel MID function syntax](https://www.ablebits.com/office-addins-blog/excel-mid-function/#mid-function-syntax)
* [How to use MID in Excel - formula examples](https://www.ablebits.com/office-addins-blog/excel-mid-function/#mid-formula-examples)
  * [Extract first and last name](https://www.ablebits.com/office-addins-blog/excel-mid-function/#extract-first-last-name)
  * [Get substring between 2 delimiters](https://www.ablebits.com/office-addins-blog/excel-mid-function/#substring-between-2-delimiters)
  * [Extract Nth word from text string](https://www.ablebits.com/office-addins-blog/excel-mid-function/#extract-nth-word)
  * [Pull a word containing a specific character(s)](https://www.ablebits.com/office-addins-blog/excel-mid-function/#word-containing-specific-character)
  * [How to force Excel Mid to return a number](https://www.ablebits.com/office-addins-blog/excel-mid-function/#force-mid-return-number)

### Excel MID function - syntax and basic uses <a href="#mid-function-syntax" id="mid-function-syntax"></a>

Generally speaking, the MID function in Excel is designed to pull a substring from the middle of the original text string. Technically speaking, the MID function returns the specified number of characters starting at the position you specify.

The Excel MID function has the following arguments:

MID(text, start\_num, num\_chars)

Where:

* Text is the original text string.
* Start\_num is the position of the first character that you want to extract.
* Num\_chars is the number of characters to extract.

All 3 arguments are required.

For example, to pull 7 characters from the text string in A2, starting with the 8th character, use this formula:

`=MID(A2,8, 7)`

The result might look something similar to this:\
![Using the MID function in Excel](https://cdn.ablebits.com/\_img-blog/excel-mid/excel-mid-function.png)

#### 5 things you should know about Excel MID function

As you have just seen, there's no rocket science in using the MID function in Excel. And remembering the following simple facts will keep you safe from most common errors.

1. The MID function always returns a text string, even if the extracted substring contains only digits. This may be critical if you wish to use the result of your Mid formula within other calculations. To convert an output into a number, use MID in combination with the VALUE function as shown in [this example](https://www.ablebits.com/office-addins-blog/excel-mid-function/#force-mid-return-number).
2. If _start\_num_ is greater than the overall length of the original text, an Excel Mid formula returns an empty string ("").
3. If _start\_num_ is less than 1, a Mid formula returns the [#VALUE! error](https://www.ablebits.com/office-addins-blog/value-error-excel/).
4. If num\_chars is less than 0 (negative number), a Mid formula returns the #VALUE! error. If num\_chars is equal to 0, it outputs an empty string (blank cell).
5. If the sum of _start\_num_ and _num\_chars_ exceeds the total length of the original string, the Excel MID function returns a substring starting from _start\_num_ and up to the last character.

### Excel MID function - formula examples <a href="#mid-formula-examples" id="mid-formula-examples"></a>

When dealing with real-life tasks in Excel, you will most often need to use MID in combination with other functions as demonstrated in the following examples.

#### How to extract first and last name <a href="#extract-first-last-name" id="extract-first-last-name"></a>

If you've had a chance to read our recent tutorials, you already know how to pull the first name using the [LEFT function](https://www.ablebits.com/office-addins-blog/excel-left-function/#substring-before-character) and get the last name with the [RIGHT function](https://www.ablebits.com/office-addins-blog/excel-right-function/). But as is often the case in Excel, the same thing can be done in a variety of ways.

**MID formula to get the first name**

Assuming the full name is in cell A2, first and last names separated with a space character, you can pull the first name using this formula:

`=MID(A2,1,SEARCH(" ",A2)-1)`

The [SEARCH](https://www.ablebits.com/office-addins-blog/excel-find-search-functions/#Excel-SEARCH-function) function is used to scan the original string for the space character (" ") and return its position, from which you subtract 1 to avoid trailing spaces. And then, you use the MID function to return a substring beginning with the fist character and up to the character preceding the space, thus fetching the first name.\


**MID formula to get the last name**

To extract the last name from A2, use this formula:

`=TRIM(MID(A2,SEARCH(" ",A2),LEN(A2)))`

Again, you use the SEARCH function to determine the starting position (a space). There is no need for us to calculate the end position exactly (as you remember, if _start\_num_ and _num\_chars_ combined is bigger than the total string length, all remaining characters are returned). So, in the _num\_chars_ argument, you simply supply the total length of the original string returned by the [LEN function](https://www.ablebits.com/office-addins-blog/excel-len-functions-count-characters-cell/#Excel-LEN-function). Instead of LEN, you can put a number that represents the longest surname you expect to find, for example 100. Finally, the [TRIM function](https://www.ablebits.com/office-addins-blog/excel-trim-function/) removes extra spaces, and you get the following result:\
![Excel MID formulas to extract the first and last names](https://cdn.ablebits.com/\_img-blog/excel-mid/excel-mid-formulas.png)\


#### How to get substring between 2 delimiters <a href="#substring-between-2-delimiters" id="substring-between-2-delimiters"></a>

Taking the previous example further, if besides first and last names cell A2 also contains a middle name, how do you extract it?

Technically, the task boils down to working out the positions of two spaces in the original string, and you can have it done in this way:

* Like in the [previous example](https://www.ablebits.com/office-addins-blog/excel-mid-function/#extract-first-last-name), use the SEARCH function to determine the position of the first space (" "), to which you add 1 because you want to start with the character that follows the space. Thus, you get the start\_num argument of your Mid formula: SEARCH(" ",A2)+1
*   Next, get the position of the 2nd space character by using nested Search functions that instruct Excel to start searching from the 2nd occurrence of the space character: SEARCH(" ",A2,SEARCH(" ",A2)+1)

    To find out the number of characters to return, subtract the position of the 1st space from the position of the 2nd space, and then subtract 1 from the result since you don't want any extra spaces in the resulting substring. Thus, you have the num\_chars argument: SEARCH (" ", A2, SEARCH (" ",A2)+1) - SEARCH (" ",A2)

With all the arguments put together, here comes the Excel Mid formula to extract a substring between 2 space characters:

\=MID(A2, SEARCH(" ",A2)+1, SEARCH (" ", A2, SEARCH (" ",A2)+1) - SEARCH (" ",A2)-1)

The following screenshot shows the result:\
![Mid formula to get substring between 2 spaces](https://cdn.ablebits.com/\_img-blog/excel-mid/mid-formula-between-chars.png)

In a similar manner, you can extract a substring between any other delimiters:

MID(_string_, SEARCH(_delimiter_, _string_)+1, SEARCH (_delimiter_, _string_, SEARCH (_delimiter_, _string_)+1) - SEARCH (_delimiter_, _string_)-1)

For example, to pull a substring that is separated by a comma and a space, use this formula:

\=MID(A2,SEARCH(", ",A2)+1,SEARCH(", ",A2,SEARCH(", ",A2)+1)-SEARCH(", ",A2)-1)

In the following screenshot, this formula is used to extract the state, and it does the job perfectly:\
![Mid formula to extract a substring separated by a comma and a space.](https://cdn.ablebits.com/\_img-blog/excel-mid/mid-formula-between-chars2.png)\


#### How to extract Nth word from a text string <a href="#extract-nth-word" id="extract-nth-word"></a>

This example demonstrates an inventive use of a complex Mid formula in Excel, which includes 5 different functions:

* LEN - to get the total string length.
* REPT - repeat a specific character a given number of times.
* SUBSTITUTE - replace one character with another.
* MID - extract a substring.
* TRIM - remove extra spaces.

The generic formula is as follows:

TRIM(MID(SUBSTITUTE(_string_," ",REPT(" ",LEN(_string_))), (_N_-1)\*LEN(_string_)+1, LEN(_string_)))

Where:

* _String_ is the original text string from which you want to extract the desired word.
* _N_ is the number of word to be extracted.

For instance, to pull the 2nd word from the string in A2, use this formula:

`=TRIM(MID(SUBSTITUTE(A2," ",REPT(" ",LEN(A2))), (2-1)*LEN(A2)+1, LEN(A2)))`

Or, you can input the number of the word to extract (N) in some cell and reference that cell in your formula, like shown in the screenshot below:\
![Excel Mid formula to extract Nth word from a text string](https://cdn.ablebits.com/\_img-blog/excel-mid/extract-word-string.png)

**How this formula works**

In essence, the formula wraps each word in the original string with many spaces, finds the desired "spaces-word-spaces" block, extracts it, and then removes extra spaces. To be more specific, the formula works with the following logic:

*   The SUBSTITUTE and REPT functions replace each space in the string with multiple spaces. The number of additional spaces is equal to the total length of the original string returned by LEN: SUBSTITUTE(A2," ",REPT(" ",LEN(A2)))

    You can think of an intermediate result as of "asteroids" of words drifting in space, like this: _spaces-word1-spaces-word2-spaces-word3-…_ This "spacious" string is supplied to the text argument of our Mid formula.
* Next, you work out the starting position of the substring of interest (start\_num argument) using the following equation: (N-1)\*LEN(A1)+1. This calculation returns either the position of the first character of the desired word or, more often, the position of some space character in the preceding space separation.
* The number of characters to extract (num\_chars argument) is the easiest part - you simply take the overall length of the original string: LEN(A2). At this point, you are left with _spaces-desired word-spaces_ substring.
* Finally, the TRIM function gets rid of leading and trailing spaces.

The above formula works fine in most situations. However, if there happen to be 2 or more consecutive spaces between words, it yields wrong results. To fix this, nest another TRIM function inside SUBSTITUTE to remove excess in-between spaces except for a single space character between words, like this:

`=TRIM(MID(SUBSTITUTE(TRIM(A2)," ",REPT(" ",LEN(A2))), (B2-1)*LEN(A2)+1, LEN(A2)))`

The following screenshot demonstrates the improved formula in action:\
![An improved Mid formula to extract Nth word from text string](https://cdn.ablebits.com/\_img-blog/excel-mid/extract-word-string2.png)

If your source strings contain multiple spaces between words as well as very big and very small words, additionally embed a TRIM function into each LEN, just to keep you on the safe side:

`=TRIM(MID(SUBSTITUTE(TRIM(A2)," ",REPT(" ",LEN(TRIM(A2)))), (B2-1)*LEN(TRIM(A2))+1, LEN(TRIM(A2))))`

I agree that this formula looks a bit cumbersome, but it impeccably handles all kinds of strings.\


#### How to extract a word containing a specific character(s) <a href="#word-containing-specific-character" id="word-containing-specific-character"></a>

This example shows another non-trivial Excel Mid formula that pulls a word containing a specific character(s) from anywhere in the original text string:

TRIM(MID(SUBSTITUTE(_string_," ",REPT(" ",99)),MAX(1,FIND(_char_,SUBSTITUTE(_string_," ",REPT(" ",99)))-50),99))

Assuming the original text is in cell A2, and you are looking to get a substring containing the "$" character (the price), the formula takes the following shape:

`=TRIM(MID(SUBSTITUTE(A2," ",REPT(" ",99)),MAX(1,FIND("$",SUBSTITUTE(A2," ",REPT(" ",99)))-50),99))`\
![Mid formula to extract a word containing a specific character](https://cdn.ablebits.com/\_img-blog/excel-mid/extract-word-containing-char.png)

In a similar fashion, you can extract email addresses (based on the "@" char), web-site names (based on "www"), and so on.

**How this formula works**

Like in the [previous example](https://www.ablebits.com/office-addins-blog/excel-mid-function/#extract-nth-word), the SUBSTITUTE and REPT functions turn every single space in the original text string into multiple spaces, more precisely, 99 spaces.

The FIND function locates the position of the desired character ($ in this example), from which you subtract 50. This takes you 50 characters back and puts somewhere in the middle of the 99-spaces block that precedes the substring containing the specified character.

The MAX function is used to handle the situation when the desired substring appears in the beginning of the original text string. In this case, the result of FIND()-50 will be a negative number, and MAX(1, FIND()-50) replaces it with 1.

From that starting point, the MID function collects the next 99 characters and returns the substring of interest surrounded by lots of spaces, like this: _spaces-substring-spaces_. As usual, the TRIM function helps you eliminate extra spaces.

Tip. If the substring to be extracted is very big, replace 99 and 50 with bigger numbers, say 1000 and 500.

#### How to force an Excel Mid formula to return a number <a href="#force-mid-return-number" id="force-mid-return-number"></a>

Like other Text functions, Excel MID always returns a text string, even if it contains only digits and looks much like a number. To turn the output into a number, simply "warp" your Mid formula into the VALUE function that converts a text value representing a number to a number.

For example, to extract a 3-char substring beginning with the 7th character and convert it to a number, use this formula:

`=VALUE(MID(A2,7,3))`

The screenshot below demonstrates the result. Please notice the right-aligned numbers pulled into column B, as opposed to the original left-aligning text strings in column A:\
![Use the Excel MID function together with VALUE to return a number](https://cdn.ablebits.com/\_img-blog/excel-mid/mid-function-output-number.png)

The same approach works for more complex formulas as well. In the above example, assuming the error codes are of a variable length, you can extract them using the Mid formula that [gets a substring between 2 delimiters](https://www.ablebits.com/office-addins-blog/excel-mid-function/#substring-between-2-delimiters), nested within the VALUE function:

`=VALUE(MID(A2,SEARCH(":",A2)+1,SEARCH(":",A2,SEARCH(":",A2)+1)-SEARCH(":",A2)-1))`\
![Nest a Mid formula in the VALUE function to turn the output into a number.](https://cdn.ablebits.com/\_img-blog/excel-mid/mid-function-output-number2.png)\


This is how you use the MID function in Excel. To better understand the formulas discussed in this tutorial, you are welcome to download a sample workbook below. I thank you for reading and hope to see you on our blog next week!

### Download practice workbook

[Excel MID function - formula examples](https://cdn.ablebits.com/excel-tutorials-examples/excel-mid-function.xlsx) (.xlsx file)\


### More examples of using the MID function in Excel:

* [Extract text between two characters](https://www.ablebits.com/office-addins-blog/extract-text-between-two-characters-excel-google/) - how to find and extract text from string between two characters or words in Excel and Google Sheets.
* [How to remove only leading spaces](https://www.ablebits.com/office-addins-blog/excel-trim-function/#remove-leading-spaces-only) - how to remove only spaces before words, leaving multiple in-between spaces intact.
* [Extract N characters following a certain character](https://www.ablebits.com/office-addins-blog/excel-find-search-functions/#Excel-FIND-formula-example3) - how to extract a substring of a given length after the specified occurrence of the delimiter/character.
* [Extract text between parentheses](https://www.ablebits.com/office-addins-blog/excel-find-search-functions/#Excel-FIND-formula-example4) - self explanatory :)
* [How to extract domain names from URLs](https://www.ablebits.com/office-addins-blog/extract-domain-names-from-url-excel/) - a very clever formula to extract domain names with or without www. and with any protocol (http, https, ftp etc.).
