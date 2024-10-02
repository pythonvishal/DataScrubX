' ################################# Regular - No Group ####################################
' TextCamelCase
' TextSlugify : -
' TextCapitalize
' TextUpperCase
' TextLowerCase
' TextSnakeCase : _
' TextTruncate
' TextCharAt
' TextInsert - (text, what_to_be_inserted, position)
' TextRepeat - (text_to_be_repeat, times)
' TextReverse
' TextReplace - first occurance (text, word_to_be_replaced, new_word)
' TextSubstring - v.substring('beach', 1, optional) return 'each'
' TextCountCharacter - count no. of characters including space
' TextCountWords - Count no. of words
' TextSprintf - v.sprintf('{1} {2} {1} {2}, watcha gonna {3}', 'bad', 'boys', 'do')
' TextIndexOf - (text, 'string')
' TextLastIndexOf - (text, 'string')
' TextWordAtPosition
' TextClearFormating

' Function to get the word at a specified position in a text string
' Returns the word at the specified position, or a value error if the position is invalid
Function TextWordAtPosition(text As String, position As Integer) As Variant
    ' Split the text into an array of words using space as the delimiter
    Dim words() As String
    words = Split(text, " ")
    
    ' Check if the position is within the range of the array
    If position < 1 Or position > UBound(words) + 1 Then
        ' Return a value error if the position is invalid
        TextWordAtPosition = CVErr(xlErrValue)
    Else
        ' Return the word at the specified position
        TextWordAtPosition = words(position - 1)
    End If
End Function

' Function to convert text to camelCase
Function TextCamelCase(text As String) As String
    Dim words() As String
    Dim i As Integer
    Dim result As String
    
    ' Split the text into words
    words = Split(text, " ")
    
    ' Convert the first word to lowercase
    result = LCase(words(0))
    
    ' Capitalize the first letter of each subsequent word
    For i = 1 To UBound(words)
        result = result & UCase(Left(words(i), 1)) & LCase(Mid(words(i), 2))
    Next i
    
    ' Return the camelCase string
    TextCamelCase = result
End Function


' Function to convert text to a slug (lowercase words separated by hyphens)
Function TextSlugify(text As String) As String
    Dim words() As String
    Dim i As Integer
    Dim result As String
    
    ' Split the text into words
    words = Split(text, " ")
    
    ' Convert each word to lowercase and join with hyphens
    For i = 0 To UBound(words)
        If i = 0 Then
            result = LCase(words(i))
        Else
            result = result & "-" & LCase(words(i))
        End If
    Next i
    
    ' Return the slugified string
    TextSlugify = result
End Function

' Function to capitalize the first letter of each word in a string
Function TextCapitalize(text As String) As String
    Dim words() As String
    Dim i As Integer
    Dim result As String
    
    ' Split the text into words
    words = Split(text, " ")
    
    ' Capitalize the first letter of each word
    For i = 0 To UBound(words)
        result = result & UCase(Left(words(i), 1)) & LCase(Mid(words(i), 2)) & " "
    Next i
    
    ' Trim the trailing space and return the capitalized string
    TextCapitalize = Trim(result)
End Function

' Function to convert text to uppercase
Function TextUpperCase(text As String) As String
    ' Convert the text to uppercase
    TextUpperCase = UCase(text)
End Function' Function to convert text to uppercase


' Function to convert text to lowercase
Function TextLowerCase(text As String) As String
    ' Convert the text to lowercase
    TextLowerCase = LCase(text)
End Function

' Function to convert text to snake_case (lowercase words separated by underscores)
Function TextSnakeCase(text As String) As String
    Dim words() As String
    Dim i As Integer
    Dim result As String
    
    ' Split the text into words
    words = Split(text, " ")
    
    ' Convert each word to lowercase and join with underscores
    For i = 0 To UBound(words)
        If i = 0 Then
            result = LCase(words(i))
        Else
            result = result & "_" & LCase(words(i))
        End If
    Next i
    
    ' Return the snake_case string
    TextSnakeCase = result
End Function

' Function to truncate text to a specified length
Function TextTruncate(text As String, length As Integer) As String
    ' Truncate the text to the specified length
    If Len(text) > length Then
        TextTruncate = Left(text, length)
    Else
        TextTruncate = text
    End If
End Function

' Function to return the character at a specified position in a string
Function TextCharAt(text As String, position As Integer) As String
    ' Return the character at the specified position
    If position > 0 And position <= Len(text) Then
        TextCharAt = Mid(text, position, 1)
    Else
        TextCharAt = ""
    End If
End Function

' Function to reverse the given text using the built-in StrReverse function
Function TextReverse(text As String) As String
    ' Use the built-in StrReverse function to reverse the text
    TextReverse = StrReverse(text)
End Function


Function TextReplace(text As String, old_text As String, new_text As String, Optional position As Variant) As String
    ' If position is provided, replace only the occurrence at the specified position
    If Not IsMissing(position) Then
        Dim startPos As Long
        startPos = InStr(1, text, old_text)
        
        ' Loop to find the specified occurrence
        Dim i As Integer
        For i = 1 To position
            If startPos = 0 Then
                ' If the specified occurrence is not found, return the original text
                TextReplace = text
                Exit Function
            End If
            startPos = InStr(startPos + 1, text, old_text)
        Next i
        
        ' If the specified occurrence is found, replace it
        If startPos > 0 Then
            TextReplace = Left(text, startPos - 1) & new_text & Mid(text, startPos + Len(old_text))
        Else
            ' If the specified occurrence is not found, return the original text
            TextReplace = text
        End If
    Else
        ' If position is not provided, replace all occurrences
        TextReplace = Replace(text, old_text, new_text)
    End If
End Function


Function TextSubstring(text As String, start_position As Integer, Optional length As Variant) As String
    ' Check if the optional length parameter is provided
    If IsMissing(length) Then
        ' If length is not provided, return from start_position to end of string
        TextSubstring = Mid(text, start_position)
    Else
        ' If length is provided, return the substring of the specified length
        TextSubstring = Mid(text, start_position, length)
    End If
End Function

Function TextCountCharacter(text As String) As Integer
    ' Return the length of the text, which is the count of characters
    TextCountCharacter = Len(text)
End Function

Function TextCountWords(text As String) As Integer
    ' Split the text into an array of words using space as the delimiter
    Dim words() As String
    words = Split(text, " ")
    
    ' Return the number of elements in the array, which is the count of words
    TextCountWords = UBound(words) - LBound(words) + 1
End Function

' Function to format a string using placeholders with curly braces
' Example: TextSprintf("{1} {2} {1} {2}, watcha gonna {3}", "bad", "boys", "do")
Function TextSprintf(formatString As String, ParamArray args() As Variant) As String
    Dim result As String
    result = formatString
    
    ' Loop through each argument and replace the placeholders in the format string
    Dim i As Integer
    For i = LBound(args) To UBound(args)
        ' Replace placeholders like {1}, {2}, etc. with the corresponding argument
        result = Replace(result, "{" & (i + 1) & "}", args(i))
    Next i
    
    TextSprintf = result
End Function

' Function to find the first occurrence of a substring within a string
' Returns the index position (1-based) of the first occurrence, or 0 if not found
Function TextIndexOf(text As String, substring As String) As Integer
    ' Use the InStr function to find the first occurrence of the substring
    TextIndexOf = InStr(1, text, substring)
End Function

' Function to find the last occurrence of a substring within a string
' Returns the index position (1-based) of the last occurrence, or 0 if not found
Function TextLastIndexOf(text As String, substring As String) As Integer
    ' Initialize the position to 0
    Dim position As Integer
    position = 0
    
    ' Use a loop to find the last occurrence of the substring
    Do
        ' Find the next occurrence of the substring
        Dim nextPosition As Integer
        nextPosition = InStr(position + 1, text, substring)
        
        ' If found, update the position
        If nextPosition > 0 Then
            position = nextPosition
        End If
    Loop While nextPosition > 0
    
    ' Return the last found position
    TextLastIndexOf = position
End Function

' Function to insert a string into another string at a specified position
' Returns the modified string, or a value error if the position is invalid
Function TextInsert(text As String, insertString As String, position As Integer) As Variant
    ' Check if the position is valid
    If position < 1 Or position > Len(text) + 1 Then
        ' Return a value error if the position is invalid
        TextInsert = CVErr(xlErrValue)
    Else
        ' Insert the string at the specified position
        TextInsert = Left(text, position - 1) & insertString & Mid(text, position)
    End If
End Function

' Function to repeat a given text a specified number of times with spaces in between
' Returns the repeated text with spaces, or a value error if something goes wrong
Function TextRepeat(text As String, n_times As Integer) As Variant
    ' Check if n_times is a valid positive integer
    If n_times < 1 Then
        ' Return a value error if n_times is not valid
        TextRepeat = CVErr(xlErrValue)
    Else
        ' Initialize the result as an empty string
        Dim result As String
        result = ""
        
        ' Repeat the text n_times with spaces in between
        Dim i As Integer
        For i = 1 To n_times
            result = result & text
            ' Add a space after each repetition except the last one
            If i < n_times Then
                result = result & " "
            End If
        Next i
        
        ' Return the repeated text with spaces
        TextRepeat = result
    End If
End Function
' ################################## Regular End ###########################################

' ################################## Set Color Text or BackgrounColor ######################
Function TextSetColor(text As String, colorName As String) As String
    Dim colorCode As Long
    colorCode = GetColorCode(LCase(colorName))
    
    ' Apply the font color to the cell that called the function
    Application.Volatile
    On Error Resume Next
    Application.Caller.Font.Color = colorCode
    
    ' Return the text
    TextSetColor = text
End Function

Function GetColorCode(colorName As String) As Long
    Select Case LCase(colorName)
        Case "red"
            GetColorCode = RGB(255, 0, 0)
        Case "black"
            GetColorCode = RGB(0, 0, 0)
        Case "white"
            GetColorCode = RGB(255, 255, 255)
        Case "yellow"
            GetColorCode = RGB(255, 255, 0)
        Case "green"
            GetColorCode = RGB(0, 255, 0)
        Case Else
            Err.Raise vbObjectError + 1, , "Invalid color name"
    End Select
End Function
' ############################### Regex ##################################
' Needs to enable MICROSOFT VBA REGEX RUNTIME 5.5

' Helper Function for extraction and validation
Function GetPattern(identifierType As String) As String
    Select Case identifierType
        Case "GSTIN"
            GetPattern = "[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}"
        Case "PAN"
            GetPattern = "[A-Z]{5}[0-9]{4}[A-Z]{1}$"
        Case "CIN"
            GetPattern = "[LU]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}"
        Case "DIN"
            GetPattern = "[0-9]{8}$"
        Case "TAN"
            GetPattern = "[A-Z]{4}[0-9]{5}[A-Z]{1}"
        Case "EMAIL"
            GetPattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
        Case "PHONE"
            GetPattern = "\+?(\d[\d-. ]+)?(\([\d-. ]+\))?[\d-. ]+\d"
        Case "URL"
            GetPattern = "(https?|ftp)://[^\s/$.?#].[^\s]*"
        Case Else
            Err.Raise vbObjectError + 1, , "Invalid identifier type"
    End Select
End Function

' TextRegexReplace(text As String, pattern As String, replacement As String, Optional index As Variant)
' TextRegexTest()
' TextRegexSearch()
' TextRegexExtract(text, regex, optional - 0 for first occurance and so on)
' TextRemovePuncutaions
' TextRemoveNumerics
' TextRemoveExtraWhiteSpaces(text)

Function TextRegexReplace(text As String, pattern As String, replacement As String, Optional index As Variant) As Variant
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to the provided regex
    regExObj.Pattern = pattern
    
    ' Set the global property to True to find all matches
    regExObj.Global = True
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Check if any match is found
    If regExObj.Test(text) Then
        ' If index is provided, replace only the occurrence at the specified index
        If Not IsMissing(index) Then
            ' Execute the regex search
            Dim matches As Object
            Set matches = regExObj.Execute(text)
            
            ' Check if the specified index is within the range of matches
            If matches.Count > index Then
                ' Get the start position of the specified match
                Dim startPos As Integer
                startPos = matches(index).FirstIndex + 1
                
                ' Replace the specified match
                TextRegexReplace = Left(text, startPos - 1) & replacement & Mid(text, startPos + Len(matches(index).Value))
            Else
                ' Return a value error if the index is out of range
                TextRegexReplace = CVErr(xlErrValue)
            End If
            
            ' Clean up the matches object
            Set matches = Nothing
        Else
            ' If index is not provided, replace all occurrences
            TextRegexReplace = regExObj.Replace(text, replacement)
        End If
    Else
        ' Return a value error if no match is found
        TextRegexReplace = CVErr(xlErrValue)
    End If
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
End Function

Function TextRegexExtract(text As String, pattern As String, Optional index As Variant) As Variant
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to the provided regex
    regExObj.Pattern = pattern
    
    ' Set the global property to True to find all matches
    regExObj.Global = True
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Execute the regex search
    Dim matches As Object
    Set matches = regExObj.Execute(text)
    
    ' Check if occurrence is provided
    If IsMissing(index) Then
        ' If no occurrence is provided, return all matches as a comma-separated string
        Dim allMatches As String
        Dim i As Integer
        For i = 0 To matches.Count - 1
            allMatches = allMatches & matches(i).Value & ","
        Next i
        ' Remove the trailing comma
        If Len(allMatches) > 0 Then
            allMatches = Left(allMatches, Len(allMatches) - 1)
        End If
        TextRegexExtract = allMatches
    Else
        ' Check if the specified occurrence is within the range of matches
        If matches.Count > index Then
            ' Return the specified occurrence match
            TextRegexExtract = matches(index).Value
        Else
            ' Return False if no match is found or occurrence is out of range
            TextRegexExtract = CVErr(xlErrValue)
        End If
    End If
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
    Set matches = Nothing
End Function

' Some Child Functions
Function TextExtractEmail(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("EMAIL")

    If IsMissing(index) Then
        TextExtractEmail = TextRegexExtract(text, pattern)
    Else
        TextExtractEmail = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractPAN(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("PAN")

    If IsMissing(index) Then
        TextExtractPAN = TextRegexExtract(text, pattern)
    Else
        TextExtractPAN = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractTAN(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("TAN")

    If IsMissing(index) Then
        TextExtractTAN = TextRegexExtract(text, pattern)
    Else
        TextExtractTAN = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractGSTIN(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("GSTIN")

    If IsMissing(index) Then
        TextExtractGSTIN = TextRegexExtract(text, pattern)
    Else
        TextExtractGSTIN = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractURL(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("URL")

    If IsMissing(index) Then
        TextExtractURL = TextRegexExtract(text, pattern)
    Else
        TextExtractURL = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractCIN(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("CIN")

    If IsMissing(index) Then
        TextExtractCIN = TextRegexExtract(text, pattern)
    Else
        TextExtractCIN = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractDIN(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("DIN")

    If IsMissing(index) Then
        TextExtractDIN = TextRegexExtract(text, pattern)
    Else
        TextExtractDIN = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextExtractPHONE(text As String, Optional index As Variant) As Variant
    Dim pattern As String
    pattern = GetPattern("PHONE")

    If IsMissing(index) Then
        TextExtractPHONE = TextRegexExtract(text, pattern)
    Else
        TextExtractPHONE = TextRegexExtract(text, pattern, index)
    End If
End Function

Function TextRegexSearch(text As String, pattern As String) As Variant
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to the provided regex
    regExObj.Pattern = pattern
    
    ' Set the global property to False to find only the first match
    regExObj.Global = False
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Execute the regex search
    Dim matches As Object
    Set matches = regExObj.Execute(text)
    
    ' Check if any match is found
    If matches.Count > 0 Then
        ' Return the start position of the first match
        TextRegexSearch = matches(0).FirstIndex
    Else
        ' Return False if no match is found
        TextRegexSearch = False
    End If
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
    Set matches = Nothing
End Function


Function TextRegexTest(text As String, regex As String) As Boolean
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to the provided regex
    regExObj.Pattern = regex
    
    ' Set the global property to False to find only the first match
    regExObj.Global = False
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Test the text against the regex pattern
    TextRegexTest = regExObj.Test(text)
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
End Function

' ###### TextRemove with Regex

Function TextRemovePunctuations(text As String) As String
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to match any punctuation character
    regExObj.Pattern = "[^\w\s]"
    
    ' Set the global property to True to remove all punctuation characters
    regExObj.Global = True
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Perform the replacement to remove punctuation
    TextRemovePunctuations = regExObj.Replace(text, "")
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
End Function


Function TextRemoveNumerics(text As String) As String
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to match any numeric character
    regExObj.Pattern = "\d"
    
    ' Set the global property to True to remove all numeric characters
    regExObj.Global = True
    
    ' Set the ignore case property to True if case-insensitive matching is desired
    regExObj.IgnoreCase = True
    
    ' Perform the replacement to remove numeric characters
    TextRemoveNumerics = regExObj.Replace(text, "")
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
End Function

' Function to remove extra spaces between words and trim leading/trailing spaces
Function TextRemoveExtraWhiteSpaces(text As String) As String
    ' Trim leading and trailing spaces
    text = Trim(text)
    
    ' Create a new RegExp object
    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to match one or more whitespace characters
    regExObj.Pattern = "\s+"
    
    ' Set the global property to True to replace all matches
    regExObj.Global = True
    
    ' Replace multiple spaces with a single space
    TextRemoveExtraWhiteSpaces = regExObj.Replace(text, " ")
    
    ' Clean up the RegExp object
    Set regExObj = Nothing
End Function

' Function to remove common stop words from a given text
' Returns the text with stop words removed
Function TextRemoveStopWords(text As String) As String
    ' Define a list of common stop words
    Dim stopWords As Variant
    stopWords = Array("a", "an", "and", "are", "as", "at", "be", "but", "by", "for", "if", "in", "into", "is", "it", "no", "not", "of", "on", "or", "such", "that", "the", "their", "then", "there", "these", "they", "this", "to", "was", "will", "with","i", "me", "my", "myself", "we", "our", "ours", "ourselves", "you", "your", "yours", "yourself", "yourselves", "he", "him", "his", "himself", "she", "her", "hers", "herself", "it", "its", "itself", "they", "them", "their", "theirs", "themselves", "what", "which", "who", "whom", "this", "that", "these", "those", "am", "is", "are", "was", "were", "be", "been", "being", "have", "has", "had", "having", "do", "does", "did", "doing", "a", "an", "the", "and", "but", "if", "or", "because", "as", "until", "while", "of", "at", "by", "for", "with", "about", "against", "between", "into", "through", "during", "before", "after", "above", "below", "to", "from", "up", "down", "in", "out", "on", "off", "over", "under", "again", "further", "then", "once", "here", "there", "when", "where", "why", "how", "all", "any", "both", "each", "few", "more", "most", "other", "some", "such", "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very", "s", "t", "can", "will", "just", "don", "should", "now")

    ' Split the text into an array of words
    Dim words() As String
    words = Split(text, " ")
    
    ' Initialize an empty string to hold the result
    Dim result As String
    result = ""
    
    ' Loop through each word in the text
    Dim i As Integer
    For i = LBound(words) To UBound(words)
        ' Check if the word is a stop word
        If IsError(Application.Match(LCase(words(i)), stopWords, 0)) Then
            ' If the word is not a stop word, add it to the result
            result = result & words(i) & " "
        End If
    Next i
    
    ' Trim any trailing spaces from the result
    TextRemoveStopWords = Trim(result)
End Function

' ################################ Text Data Validation #############################
' TextValidateEmail
' TextValidatePhoneNumber
' TextValidateGSTIN
' TextValidatePAN
' TextValidateCIN
' TextValidateDIN
' TextValidateTAN
' TextValidateURL
' TextValidate

' helper function defined at beginning of regex
Function ValidateIdentifier(text As String, identifierType As String) As Boolean
    Dim pattern As String
    pattern = GetPattern(identifierType)
    
    ' Use TextRegexExtract to check for matches
    Dim result As Variant
    result = TextRegexTest(text, pattern)
    
    ' Return True if matches are found, otherwise False
    If result = False Then
        ValidateIdentifier = False
    Else
        ValidateIdentifier = True
    End If
End Function

' Wrapper functions for each identifier type
Function TextValidateGSTIN(text As String) As Boolean
    TextValidateGSTIN = ValidateIdentifier(text, "GSTIN")
End Function

Function TextValidatePAN(text As String) As Boolean
    TextValidatePAN = ValidateIdentifier(text, "PAN")
End Function

Function TextValidateCIN(text As String) As Boolean
    TextValidateCIN = ValidateIdentifier(text, "CIN")
End Function

Function TextValidateDIN(text As String) As Boolean
    TextValidateDIN = ValidateIdentifier(text, "DIN")
End Function

Function TextValidateTAN(text As String) As Boolean
    TextValidateTAN = ValidateIdentifier(text, "TAN")
End Function

Function TextValidateEMAIL(text As String) As Boolean
    TextValidateEMAIL = ValidateIdentifier(text, "EMAIL")
End Function

Function TextValidatePHONE(text As String) As Boolean
    TextValidatePHONE = ValidateIdentifier(text, "PHONE")
End Function

Function TextValidateURL(text As String) As Boolean
    TextValidateURL = ValidateIdentifier(text, "URL")
End Function

' ############################## Regex End ##################################


' #############################################################################
' ################################ TEST Conditions ############################
' #############################################################################
' TextIsBold
' TextIsItalic
' TextIsUnderline
' TextIsBlank
' TextIsEmpty
' TextIsNumeric
' TextIsInteger
' TextIsFloat
' TextIsDate
' TextIsLowerCase
' TextIsUpperCase
' TextIsString
' TextStartsWith
' TextEndsWith
' TextIncludes
' ##############

' Function to check if the text in a cell is bold
Function TextIsBold(cell As Range) As Boolean
    ' Check if the font of the text in the cell is bold
    TextIsBold = cell.Font.Bold
End Function

' Function to check if the text in a cell is italic
Function TextIsItalic(cell As Range) As Boolean
    ' Check if the font of the text in the cell is italic
    TextIsItalic = cell.Font.Italic
End Function

' Function to check if the text in a cell is underlined
Function TextIsUnderline(cell As Range) As Boolean
    ' Check if the font of the text in the cell is underlined
    TextIsUnderline = cell.Font.Underline <> xlUnderlineStyleNone
End Function

' Function to check if the text is blank (contains only spaces or is empty)
Function TextIsBlank(text As String) As Boolean
    ' Trim the text to remove leading and trailing spaces and check if it's empty
    TextIsBlank = Trim(text) = ""
End Function

' Function to check if the text is empty
Function TextIsEmpty(text As String) As Boolean
    ' Check if the text is an empty string
    TextIsEmpty = text = ""
End Function

' Function to check if the text is numeric
Function TextIsNumeric(text As String) As Boolean
    ' Use the IsNumeric function to check if the text is numeric
    TextIsNumeric = IsNumeric(text)
End Function

' Function to check if the text is an integer
Function TextIsInteger(text As String) As Boolean
    ' Check if the text is numeric and if it equals its integer conversion
    If IsNumeric(text) Then
        TextIsInteger = (CLng(text) = Val(text))
    Else
        TextIsInteger = False
    End If
End Function

' Function to check if the text is a float (decimal number)
Function TextIsFloat(text As String) As Boolean
    ' Check if the text is numeric and if it contains a decimal point
    If IsNumeric(text) Then
        TextIsFloat = (InStr(1, text, ".") > 0)
    Else
        TextIsFloat = False
    End If
End Function

' Function to check if the text is in lowercase
Function TextIsLowerCase(text As String) As Boolean
    ' Compare the text with its lowercase conversion
    TextIsLowerCase = (text = LCase(text))
End Function

' Function to check if the text is in uppercase
Function TextIsUpperCase(text As String) As Boolean
    ' Compare the text with its uppercase conversion
    TextIsUpperCase = (text = UCase(text))
End Function

' Function to check if the text is a string (contains alphabetic characters and spaces)
Function TextIsString(text As String) As Boolean
    Dim i As Integer
    Dim isString As Boolean
    isString = True
    
    ' Loop through each character in the text
    For i = 1 To Len(text)
        ' Check if the character is not a letter or space
        If Not (Mid(text, i, 1) Like "[A-Za-z ]") Then
            isString = False
            Exit For
        End If
    Next i
    
    ' Ensure the text is not purely numeric
    If IsNumeric(text) Then
        isString = False
    End If
    
    TextIsString = isString
End Function

' Function to check if a given text is a valid date
' Returns True if valid, False otherwise
Function TextIsDate(text As String) As Boolean
    ' Use the IsDate function to check if the text can be converted to a date
    TextIsDate = IsDate(text)
End Function

' Function to check if the text starts with a specified substring (case-insensitive)
Function TextStartsWith(text As String, startswith As String) As Boolean
    ' Convert both text and startswith to lowercase
    Dim lowerText As String
    Dim lowerStartsWith As String
    lowerText = LCase(text)
    lowerStartsWith = LCase(startswith)
    
    ' Check if the text starts with the specified substring
    TextStartsWith = (Left(lowerText, Len(lowerStartsWith)) = lowerStartsWith)
End Function

' Function to check if the text ends with a specified substring (case-insensitive)
Function TextEndsWith(text As String, endswith As String) As Boolean
    ' Convert both text and endswith to lowercase
    Dim lowerText As String
    Dim lowerEndsWith As String
    lowerText = LCase(text)
    lowerEndsWith = LCase(endswith)
    
    ' Check if the text ends with the specified substring
    TextEndsWith = (Right(lowerText, Len(lowerEndsWith)) = lowerEndsWith)
End Function

' Function to check if the text includes a specified substring (case-insensitive)
Function TextIncludes(text As String, include_string As String) As Boolean
    ' Convert both text and include_string to lowercase
    Dim lowerText As String
    Dim lowerIncludeString As String
    lowerText = LCase(text)
    lowerIncludeString = LCase(include_string)
    
    ' Check if the text includes the specified substring
    TextIncludes = (InStr(1, lowerText, lowerIncludeString, vbTextCompare) > 0)
End Function
' ################################ Test Conditions End ##############################
' ###################################################################################


' ################################ Developer And License #############################

Function TextLicense()
    TextLicense = "Visit https://vishalchopra666.github.io/excelscrubx/"
End Function

Function TextVersion()
    TextVersion = "Current Version: V1. For Latest version, visit https://pythonvishal.github.io/excelscrubx/"
End Function

Function TextDeveloperInfo()
    TextDeveloperInfo = "Developed by Vishal Chopra, visit https://vishalchopra.in for more info."
End Function


'                                   Debugging Notes
' ###################################################################################
