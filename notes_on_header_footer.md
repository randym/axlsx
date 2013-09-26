## positioning
&L - code for "left section"

&C - code for "center section".

&R - code for "right section".

￼(there are three header / footer locations, "left", "center", and "right"). When two or
￼more occurrences of a section marker exist, the contents from all identical markers are concatenated, in the order of
￼appearance, and placed into the section section.
￼
## font styles

&font size - code for "text font size", where font size is a font size in points. 

&K - code for "text font color" 
    - RGB Color is specified as RRGGBB
    - Theme Color is specified as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade value, NN is the tint/shade value.

&"font name,font type" - code for "text font name" and "text font type" where font name and font type
 are ￼￼strings specifying the name and type of the font, separated by a comma. When a hyphen appears 
in font name it means none specified. Both of font name and font type can be localized values. 
￼￼Although ISO/IEC 14496-22 permits commas in font family/subfamily/full names, name and font type
, the lexically first comma in the font name is the one recognized as the separating comma. 

&"-,Bold" - code for "bold font style"

&B - also means "bold font style".

&"-,Regular" - code for "regular font style" 

&"-,Italic" - code for "italic font style"
&I - also means "italic font style"

&"-,Bold Italic" code for "bold italic font style"

&S - code for "text strikethrough" on / off 
&X - code for "text super script" on / off 
&Y - code for "text subscript" on / off
￼

## Workbook info and page numbering

&P - code for "current page #"

&N - code for "total pages"

&D - code for "date"

&T - code for "time"

&G - code for "picture as background" 

&U - code for "text single underline" 

&E - code for "double underline"

&Z - code for "this workbook's file path" 

&F - code for "this workbook's file name"

&A - code for "sheet tab name"

&+ - code for add to page #.

&- - code for subtract from page #. 

&O - code for "outline style"
&H - code for "shadow style"
