Imports System.IO
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Module Module1
    Public Function GenerateRandomString(ByRef len As Integer, ByRef upper As Boolean) As String ' random string generator
        Dim rand As New Random() 'Create a new random value
        Dim allowableChars() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray() 'Characters that can be in the code
        Dim final As String = String.Empty ' Create new empty string
        For i As Integer = 0 To len - 1 ' Loop for the amount of length (I set it to 8 when i call it)
            final += allowableChars(rand.Next(allowableChars.Length - 1)) ' add random values
        Next

        Return IIf(upper, final.ToUpper(), final) ' return string
    End Function
    Sub EncryptFile(ByVal sInputFilename As String,
                   ByVal sOutputFilename As String,
                   ByVal sKey As String)

        Dim fsInput As New FileStream(sInputFilename,
                                    FileMode.Open, FileAccess.Read)
        Dim fsEncrypted As New FileStream(sOutputFilename,
                                    FileMode.Create, FileAccess.Write)

        Dim DES As New DESCryptoServiceProvider()

        'Secret key is set later on
        'Create the DES encryptor 
        Dim desencrypt As ICryptoTransform = DES.CreateEncryptor()
        'Create the crypto stream to encrypt the file
        Dim cryptostream As New CryptoStream(fsEncrypted,
                                            desencrypt,
                                            CryptoStreamMode.Write)

        'Read the file text to the byte array.      cryptostreams are way better than reading every value in a text file
        Dim bytearrayinput(fsInput.Length - 1) As Byte
        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
        'Write out the DES encrypted file.
        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Close()
    End Sub

    Sub DecryptFile(ByVal sInputFilename As String,
        ByVal sOutputFilename As String,
        ByVal sKey As String)

        Dim DES As New DESCryptoServiceProvider()
        'secret key is set from user input

        'Create the file stream to read the encrypted file back.
        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
        'Create the DES decryptor from the DES instance.
        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
        'Create the crypto stream set to read and to do a DES decryption transform on incoming bytes.
        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
        'Print out the contents of the decrypted file.
        Dim fsDecrypted As New StreamWriter(sOutputFilename)
        Try
            fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)

        Catch CrypographyException As Exception
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Error. Incorrect encryption key. Please try again")
            Console.ForegroundColor = ConsoleColor.White
            Threading.Thread.Sleep("2000")
            Environment.Exit(1.0)
        End Try
        fsDecrypted.Flush()
        fsDecrypted.Close()
    End Sub
    Dim choice As String
    Dim path As String
    Dim save As String
    Dim length As String
    Public Sub main()
        Console.Clear()
        Console.ForegroundColor = ConsoleColor.White
        Dim sSecretKey As String 'Create the Secret Key
        Console.WriteLine("~~~~~~File Encryption~~~~~")
        Console.WriteLine("Menu")
        Console.WriteLine("")
        Console.WriteLine("1: Encrypt sample.txt")
        Console.WriteLine("2: Decrypt sample.txt")
        Console.WriteLine("3: Quit Program")

        choice = Console.ReadLine()
        If IsNumeric(choice) Then
            If choice = 1 Or choice = 2 Or choice = 3 Then ' Menu
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Incorrect value")
                Console.ForegroundColor = ConsoleColor.White
                Threading.Thread.Sleep(1000)
                main()
            End If
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Incorrect value")
            Console.ForegroundColor = ConsoleColor.White
            Threading.Thread.Sleep(1000)
            main()
        End If

        If choice = 1 Then
            Console.WriteLine("Please enter how many characters the Secret Key will be.")
            length = Console.ReadLine
            If Not IsNumeric(length) Or length <= 0 Or length >= 65565 Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Error: Incorrect Value")
                Threading.Thread.Sleep(2000)
                Console.ForegroundColor = ConsoleColor.White
                main()
            Else
                Console.WriteLine("Please enter the file name stored in %project%\bin\debug\ (If the file is in another location, please enter the FULL path)")
                path = Console.ReadLine()
                If File.Exists(path) Then
                    sSecretKey = GenerateRandomString(length, True) ' generate 8 digit key with numbers, using this instead of a 64 bit key because its human readable
                    Console.WriteLine("Please enter the name of the file to save the encrypted text to (File will be created if it doesn't exist) File will be stored in %project%\bin\debug\")
                    save = Console.ReadLine
                    EncryptFile(path,
                                     save,
                     sSecretKey) ' get file and encrypt it with the given paths and secret key

                    Console.WriteLine("Your encrypton key is " & sSecretKey)
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Write this down else you will not be able to decrypt the file later on")
                    Console.WriteLine("Would you like to copy the key to clipboard?")
                    Console.WriteLine("")
                    Console.WriteLine("1: Yes.")
                    Console.WriteLine("2: No.")
                    choice = Console.ReadLine()
                    If Not IsNumeric(choice) Or choice < 1 Or choice > 2 Then
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Error: Incorrect Value")
                        Threading.Thread.Sleep(2000)
                        Console.ForegroundColor = ConsoleColor.White
                        main()
                    ElseIf choice = 1 Then
                        My.Computer.Clipboard.SetText(sSecretKey)
                        Console.WriteLine("Copied to Clipboard, to paste it in to the console, right click the top bar > edit > paste.")
                        Threading.Thread.Sleep(2000)
                        main()
                    ElseIf choice = 2 Then
                        Console.WriteLine("Please copy down the encryption key somewere")
                        Console.Read()
                        main()

                    End If
                Else
                    Console.WriteLine("file does not exist")
                    Threading.Thread.Sleep(2000)
                    main()
                End If
            End If
        ElseIf choice = 2 Then
            ' Decrypt the file.
            Console.WriteLine("Please enter your encryption key.")
            sSecretKey = Console.ReadLine() 'enter the secret key
            Dim gch As GCHandle = GCHandle.Alloc(sSecretKey, GCHandleType.Pinned) ' add secret key to DES
            Console.WriteLine("Please enter the name of the file to decrypt")
            path = Console.ReadLine()
            If File.Exists(path) Then
                Console.WriteLine("Please enter the name of the file to save the decrypted text to (File will be created if it doesn't exist) File will be stored in %project%\bin\debug\")
                save = Console.ReadLine()
                DecryptFile(path,
                            save,
                sSecretKey)
            ElseIf choice = 3 Then
                Console.WriteLine("Are you sure you want to exit?")
                choice = Console.ReadLine()
                If choice = 1 Then
                    main()
                ElseIf choice = 2 Then
                    Environment.Exit(0.0)
                Else
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Incorrect Value")
                    Threading.Thread.Sleep(1000)
                    Console.ForegroundColor = ConsoleColor.White
                End If
            End If
        End If




    End Sub
End Module