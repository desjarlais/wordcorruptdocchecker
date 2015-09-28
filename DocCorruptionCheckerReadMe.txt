Corruption Readme:

This readme is an explanation of each type of corruption that I've noticed when customers send me their problem Word documents.  

1. invalid oMath corruptions
- this type of corruption happens when we write out the incorrect order of oMath tags
- invalid tag = <mc:AlternateContent><mc:Choice Requires="wps"><m:oMath>
- valid tag = <m:oMath><mc:AlternateContent><mc:Choice Requires="wps">
- this happens with multiple namespaces, so we need to check for all possible namespaces (wps, wpg, wpi, wpc)

2. invalid mcChoice corruptions
- this type of corruption happens with invalid tags after the closing </mc:Choice> tag
- this is a regular expression to look for the mc:choice tag and ANY tag after it
- if the tag that follows the choice tag is not one of the valid combinations, we need to replace it with a valid combination

3. invalid fallback corruptions
- this type of corruption happens with incorrect fallback tags
- there are two possible invalid tag combinations i've seen so far
- they are as follows and include the text that can be used to replace the invalid tags with to make it valid
   1. <mc:Fallback><w:pict/> = </mc:AlternateContent></w:r>
   2. <mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r> = </mc:AlternateContent></w:r>
- i use a regular expression here also to check for <mc:Fallback><w:pict/> followed by ANY tag after it