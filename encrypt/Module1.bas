Attribute VB_Name = "Module1"
' note: this module is optional. if you are a new vb programmer,
' you can paste this code in your main form.

Public ShowEncryptedText As String      'this will return the value of encrypted text
Public ShowDecryptedText As String      'this will return the value of decrypted text
   
   Public Function Encrypt(ByVal source As String) As String
        '----- Encrypt Method -----'

        Dim x As Integer        'used for counting purposes only
        Dim Hold As String      'this will hold the converted string to ascii temporarily
        Dim countfirststring As Long    'you need to count the number of strings to be converted for easy conversion
        Dim enc As String               'this will hold the encrypted text inside the function

        countfirststring = Len(source)  'count the number of strings

        For x = 1 To countfirststring   'start the loop

            Hold = Asc(Mid(source, x, 1))   'convert the first string into ascii then
                                            'pass it to Hold variable
            If Len(Hold) <> 3 Then          'note that we must standardize our encryption
                Hold = "0" & Hold           'for easy decryption of text
                enc = enc & Hold            'then concatinate it to the whole encrypted variable
            Else
                enc = enc & Hold            'if it is already standard, just go on with the concatination
            End If
        Next
        ShowEncryptedText = enc     'if you are using vb.net, simply remove this comment to
        'Return enc                 'return the value from the function
    End Function                    'and comment the ShowEncryptedText = enc line

    Public Function Decrypt(ByVal Encrypted As String) As String
        '----- Decrypt Method -----'
        
        Dim CountEncryptedString As Long    'before decrypting, first count the encrypted strings
        Dim x, y As Integer                 'for counting purposes only
        Dim a As String                     'this will split all concatenated strings
        Dim dec As String                   'this will hold the decrypted text inside the function

        CountEncryptedString = Len(Encrypted)   'count the number of strings
        y = 1                                   'tell Y to begin splitting at first string
        For x = 1 To CountEncryptedString       'start the loop
            a = Val(Mid(Encrypted, y, 3))       'we must split every 3 strings together as stated in our
                                                'standardized format
            y = y + 3                           'add 3 as stated in the format

            dec = dec & Chr(a)                  'then concatenate it to the decrypted variable
        Next
        ShowDecryptedText = dec     'if you are using vb.net, simply remove this comment to
        'Return dec                 'return the value from the function
                                    'and comment the ShowDecryptedText = dec line
    End Function

'there you are! simple and effective way of encrypting and decrypting
'in this way, you can go on to a higher level with more complex conversion

