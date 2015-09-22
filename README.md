# Word corrupt document checker

I created this tool to fix corrupt Word documents (non-binary, just open xml files). Mainly this applies to the .docx format. It basically just checks through a list of known corruptions and then applies a fix if it comes across one of those scenarios in the file. It will create a copy of the file you provide and attempt to fix the copy.  The original file should be left unchanged.

![ScreenShot](http://i.imgur.com/43TaZtV.jpg)

There are still many other types of corrupt documents that I have not seen, so please free to submit questions or feedback in the Discussions section.

 The tool requires the .NET framework version 4.5.1 or higher.
