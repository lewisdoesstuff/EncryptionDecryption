Imports System
Imports System.IO 'imports
Module Module1
    Dim temp
    Dim key ' global declarations
    Dim ascii
    Sub encrypt()
        For k = 1 To intlength
            temp = CInt(Math.Floor((126 - 33 + 1) * Rnd())) + 33 ' generates random integer between 33 and 126
            key += temp
            ascii = ascii + Chr(temp).ToString
        Next
        key = key / 8
        key = Math.Floor(key * 100) / 100 ' math.
        key = key - 32
        Dim unencrypt As String
        unencrypt = My.Computer.FileSystem.ReadAllText(path) ' read file to string
        Dim encrypt As String = ""
        Dim offset As Integer
        For i = 0 To 100
            i = i + 1
            Console.WriteLine(i & "%")
            Threading.Thread.Sleep(50)
            Console.Clear()
        Next
        For i = 1 To unencrypt.Length
            offset = Asc(GetChar(unencrypt, i)) ' generate the offset
            If offset <> 32 Then
                If offset > 126 Then
                    encrypt = encrypt + Chr(offset + key)
                Else
                    encrypt = encrypt + Chr(offset + key) ' more offset generation
                End If
            Else
                encrypt = encrypt + GetChar(unencrypt, i)
            End If
        Next
        My.Computer.FileSystem.WriteAllText(save & ".txt", encrypt, False) ' save the encrypted string to file ending with txt
    End Sub
    Sub decrypt()
        For i = 1 To ascii.length
            key += Asc(GetChar(ascii, i))
        Next
        key = key / 8 ' does the same thing in reverse
        key = Math.Floor(key * 100) / 100
        key = key - 32
        Dim unencrypt As String = ""
        Dim encrypt As String = My.Computer.FileSystem.ReadAllText(path) ' reads encrypted to string
        Dim offset As Integer
        For i = 1 To encrypt.Length
            offset = Asc(GetChar(encrypt, i))
            If offset <> 32 Then ' finds offset value and gets the original file
                If offset < 33 Then
                    unencrypt = unencrypt + Chr((offset - key) + 94)
                Else
                    unencrypt = unencrypt + Chr(offset - key)
                End If
            Else
                unencrypt = unencrypt + GetChar(encrypt, i)
            End If
        Next
        Console.WriteLine(unencrypt) 'prints to console  
        My.Computer.FileSystem.WriteAllText(save, unencrypt, False) ' saves file with given name
    End Sub
    Sub removewhitespace()
        Dim whitespace As String = My.Computer.FileSystem.ReadAllText(path & ".txt")
        Dim final As String = 0
        final = whitespace.Replace(" ", "")
        Dim i As Integer = 5
        While i < final.Length
            final = final.Insert(i, " ")
            i += 6
        End While
        Console.WriteLine(final)
        My.Computer.FileSystem.WriteAllText(save & ".txt", final, False)
        Console.ReadKey()
    End Sub
    Dim choice As String
    Dim path As String
    Dim save As String
    Dim length As String ' more declarations
    Dim intlength As Integer
    Public Sub main()
        Console.Clear()
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("~~~~~~File Encryption~~~~~")
        Console.WriteLine("Menu")
        Console.WriteLine("")
        Console.WriteLine("1: Encrypt sample.txt")
        Console.WriteLine("2: Decrypt sample.txt") ' menu
        Console.WriteLine("3: Add spaces every 5 letters, WARNING, THIS WILL BREAK ORIGINAL SPACING")
        Console.WriteLine("4: Quit Program")
        choice = Console.ReadLine()
        If IsNumeric(choice) Then
            If choice = 1 Or choice = 2 Or choice = 3 Or choice = 4 Then ' Menu
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Incorrect value") ' a bunch of if statements to mostly do the same thing up ahead
                Console.ForegroundColor = ConsoleColor.White
                Threading.Thread.Sleep(1000)
                main()
            End If
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Incorrect value")
            Console.ForegroundColor = ConsoleColor.White
            Threading.Thread.Sleep(1000)
            main() ' checks if its numeric else fail.
        End If
        If choice = 1 Then
            Console.WriteLine("Please enter how many characters the Secret Key will be.")
            length = Console.ReadLine
            If Not IsNumeric(length) Or length <= 0 Or length >= 65565 Then ' get intlength
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Error: Incorrect Value") 'check 
                Threading.Thread.Sleep(2000)
                Console.ForegroundColor = ConsoleColor.White
                main()
            Else
                intlength = length
                Console.WriteLine("Please enter the file name stored in %project%\bin\debug\")
                path = Console.ReadLine() ' finds file name
                path = path & ".txt"
                If File.Exists(path) Then
                    Console.WriteLine("Please enter the name of the file to save the encrypted text to (File will be created if it doesn't exist")
                    save = Console.ReadLine
                    encrypt() ' runs encryption subroutine
                    Console.WriteLine("Your encrypton key is " & ascii) ' prints key to console
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Write this down else you will not be able to decrypt the file later on")
                    Console.WriteLine("Would you like to copy the key to clipboard?")
                    Console.WriteLine("")
                    Console.WriteLine("1: Yes.")
                    Console.WriteLine("2: No.") ' copy to clipboard
                    choice = Console.ReadLine()
                    If Not IsNumeric(choice) Or choice < 1 Or choice > 2 Then
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Error: Incorrect Value")
                        Threading.Thread.Sleep(2000)
                        Console.ForegroundColor = ConsoleColor.White
                        main()
                    ElseIf choice = 1 Then
                        My.Computer.Clipboard.SetText(ascii) ' best function ever.
                        Console.WriteLine("Copied to Clipboard, to paste it in to the console, right click the top bar > edit > paste.")
                        Threading.Thread.Sleep(2000)
                        main()
                    ElseIf choice = 2 Then
                        Console.WriteLine("Please copy down the sample key somewere")
                        Console.Read()
                        main()
                    End If
                    Console.WriteLine("file does not exist")
                    Threading.Thread.Sleep(2000)
                End If
            End If
        ElseIf choice = 2 Then
            ' Decrypt the file.
            Console.WriteLine("Please enter your encryption key.")
            ascii = Console.ReadLine() 'enter the secret key
            Console.WriteLine("Please enter the name of the file to decrypt stored in %project%\bin\debug\")
            path = Console.ReadLine()
            path = path & ".txt" ' get path and append .txt
            If File.Exists(path) Then ' if file exists then do this
                Console.WriteLine("Please enter the name of the file to save the decrypted text to (File will be created if it doesn't exist")
                save = Console.ReadLine() 'get save location
                save = save & ".txt"
                decrypt()
            End If
        ElseIf choice = 3 Then
            Console.WriteLine("This will remove all whitespace and add a space every 5 letters, However. This will make the decrypted file harder to read as it will break original spacing.")
            Console.WriteLine("Enter the name of the Encrypted file saved in %project%\bin\debug\.")
            path = Console.ReadLine()
            Console.WriteLine("Please enter the name of the file to be encrypted")
            save = Console.ReadLine()
            removewhitespace()

        ElseIf choice = 4 Then
            Console.WriteLine("Are you sure you want to exit?")
            choice = Console.ReadLine()
            If choice = 1 Then
                main()
            ElseIf choice = 2 Then 'exiting the program
                Environment.Exit(0.0)
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Incorrect Value")
                Threading.Thread.Sleep(1000)
                Console.ForegroundColor = ConsoleColor.White
            End If
            End If
    End Sub
End Module