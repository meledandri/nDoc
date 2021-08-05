Imports System.Drawing
Imports System.IO
Imports Microsoft.Win32


Module Module1
    'Prevenire che la console venga ridimensionata o minimizzata
    Private Const MF_BYCOMMAND As Integer = &H0
    Public Const SC_CLOSE As Integer = &HF060
    Public Const SC_MINIMIZE As Integer = &HF020
    Public Const SC_MAXIMIZE As Integer = &HF030
    Public Const SC_SIZE As Integer = &HF000

    Friend Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As IntPtr, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Friend Declare Function GetSystemMenu Lib "user32.dll" (hWnd As IntPtr, bRevert As Boolean) As IntPtr
    '---------------------------

    Dim WindowWidth As Integer = Console.WindowWidth
    Dim WindowHeight As Integer = Console.WindowHeight




    Dim variabili As New local_db("D:\upgrade\values.db")

    Sub Main()
        fixConsole()
        ''HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word
        ''HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems
        'Dim RegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE").OpenSubKey("Microsoft").OpenSubKey("Office")
        'For Each KeyName As String In RegKey.GetSubKeyNames
        '    If IsNumeric(KeyName) Then
        '        Try
        '            Dim rkey As RegistryKey = RegKey.OpenSubKey(KeyName).OpenSubKey("Word").OpenSubKey("Resiliency").OpenSubKey("DisabledItems", RegistryKeyPermissionCheck.ReadWriteSubTree, Security.AccessControl.RegistryRights.FullControl)
        '            'Dim rkey As RegistryKey = RegKey.OpenSubKey(KeyName).OpenSubKey("Word").OpenSubKey("Resiliency").OpenSubKey("DisabledItems", True)
        '            Console.WriteLine("Word: DisabledItems (trovato)")
        '            Console.WriteLine(rkey.ToString)
        '            For Each ValueName As String In rkey.GetValueNames
        '                Console.WriteLine("Valore: " & ValueName & " (..)")
        '                rkey.DeleteValue(ValueName)
        '                Console.WriteLine("Valore: " & ValueName & " (Cancellato)")
        '            Next
        '        Catch ex As Exception
        '            Console.WriteLine("#KeyName: " & KeyName)
        '        End Try
        '    End If
        'Next





        If Not variabili.Exist Then
            variabili.add("[COMPUTER_NAME]", My.Computer.Name)
            variabili.add("[DATABASE_NAME]", "DBNAM")
            variabili.add("[DATABASE_ACCOUNT]", "DBACC")
            variabili.add("[DATABASE_PASSWORD]", "DBPSW")

            variabili.add("[MAIL_ACCOUNT]", "mail_account")
            variabili.add("[MAIL_PASSWORD]", "mail_password")
            variabili.add("[MAIL_FROM_ADDR]", "mail_from_address")
            variabili.add("[MAIL_TO_ADDR]", "mail_to_address")
            variabili.add("[MAIL_SUBJECT]", "mail_subject")
            variabili.add("[MAIL_FROM_DISPLAYNAME]", "mail_from_displayName")


            variabili.add("[TOKEN_APP]", "AB123-CD456-EF789-GH012-IF345")

            variabili.save()

        Else
            variabili.load()
        End If



        'Update config file
        Dim t As String = File.ReadAllText("d:\upgrade\sendGmail.exe.config.inst")
        t = text_replace(t)
        File.WriteAllText("d:\upgrade\sendGmail.exe.config", t)



        display_menu_main()
    End Sub

    Private Function text_replace(ByVal text As String) As String
        Dim list As List(Of local_db._record) = variabili.data
        For Each r As local_db._record In list
            text = text.Replace(r.Key, r.Value)
        Next
        Return text
    End Function


    Sub display_menu_main()
display_menu_main_start:

        Dim mm As New MenuConsole
        mm.setTitle("MENU PRINCIPALE (" & My.Application.Info.AssemblyName & ")", MenuConsole.horizontal_align.center, ConsoleColor.Yellow)
        mm.setDescription("Questa è una descrizione del menu che supera la riga di testo per fare le prove di impostazione dell'altezza massima dedicata", ConsoleColor.White, ConsoleColor.DarkGray, 1)
        Dim ms As New List(Of String)

        'VOCI DEL MENU
        ' mm.Options.Add(New MenuConsoleOption("CFG", "Configurazione generale Applicazione."))
        mm.Options.Add(New MenuConsoleOption("CFG", "Configurazione generale Applicazione.", "Sistema di configurazione generale per la gestione dei poogetti e degli aggiornamenti dei dati e struttura dati connessa."))

        mm.Options.Add(New MenuConsoleOption("VAR", "Configurazione variabili locali.", Nothing))
        mm.Options.Add(New MenuConsoleOption("UTIL", "Strumenti di utilità."))
        mm.Options.Add(New MenuConsoleOption("EXIT", "Esci dall'applicazione."))
        'For o As Integer = 0 To 20
        '    Dim m As New MenuConsoleOption("MENU" & o, "MENU 0" & o, "Descrizione del menu 0" & o)
        '    m.menuColor = ConsoleColor.DarkGreen
        '    'm.menuBackColor = ConsoleColor.DarkBlue
        '    m.menuSelBackColor = ConsoleColor.DarkGreen
        '    m.menuSelColor = ConsoleColor.Green
        '    mm.Options.Add(m)
        'Next

        Dim di As New DirectoryInfo(My.Application.Info.DirectoryPath)
        For Each fi As FileInfo In di.GetFiles
            Dim m As New MenuConsoleOption(fi.Name, fi.Name, fi.FullName & vbTab & " [" & fi.Length & "]")
            m.menuColor = ConsoleColor.DarkYellow
            'm.menuBackColor = ConsoleColor.DarkBlue
            m.menuSelBackColor = ConsoleColor.DarkYellow
            m.menuSelColor = ConsoleColor.Yellow
            mm.Options.Add(m)
        Next

        'For o As Integer = 0 To 25
        '    Dim m As New MenuConsoleOption("MENU" & o, "MENU 1" & o, "Descrizione del menu 1" & o)
        '    m.menuColor = ConsoleColor.DarkYellow
        '    'm.menuBackColor = ConsoleColor.DarkBlue
        '    m.menuSelBackColor = ConsoleColor.DarkYellow
        '    m.menuSelColor = ConsoleColor.Yellow
        '    mm.Options.Add(m)
        'Next


        mm.addFunction(New FunctionKey(ConsoleKey.F10, "SELEZIONA", True))
        mm.addFunction(New FunctionKey(ConsoleKey.Escape, "ESCI"))
        mm.addFunction(New FunctionKey(ConsoleKey.F12, "ACCETTA", False, special_functions.confirm_selection))
        mm.addFunction(New FunctionKey(ConsoleKey.Q, "QUIT", False, special_functions.select_item))
        mm.addFunction(New FunctionKey(ConsoleKey.Q, "QUIT"))
        mm.addFunction(New FunctionKey(ConsoleKey.Q, "QUIT"))
        ' mm.addFunction(New FunctionKey(ConsoleKey.Enter, "INVIO", False, special_functions.confirm_selection))


        AddHandler mm.functionKey, AddressOf runFunction

        mm.multiselection = True

        Dim r As String = mm.Show

        Select Case r
            Case "CFG"
                display_menu_cfg()
            Case "VAR"
                display_menu_var()
            Case "UTIL"
            Case "EXIT"
            Case Else
                GoTo display_menu_main_start
        End Select

    End Sub

    Private Sub runFunction(ByVal fn As FunctionKey, ByVal Selected As ArrayList, ByRef sender As MenuConsole)
        setLog(log_type.info, fn.Value.ToString & " Premuto --> " & fn.Description)
        setLog(log_type.info, "SELECTED (" & Selected.Count & "): " & String.Join(",", Selected.ToArray))
        If fn.special_function = special_functions.confirm_selection Then
            For Each s As String In Selected
                Process.Start("notepad++", s)
            Next
        End If
        If fn.Value = ConsoleKey.Q Then
            sender.multiselection = Not sender.multiselection
        End If


    End Sub


    Sub display_menu_cfg()
display_menu_cfg_start:

        Dim mm As New MenuConsole
        mm.setTitle("MENU CONFIGURAZIONE (" & My.Application.Info.AssemblyName & ")")
        Dim ms As New List(Of String)

        'VOCI DEL MENU
        mm.Options.Add(New MenuConsoleOption("WEB", "Configurazione impostazioni web."))
        mm.Options.Add(New MenuConsoleOption("MAIL", "Configurazione Mail."))
        mm.Options.Add(New MenuConsoleOption("API", "Configurazioni API."))
        mm.Options.Add(New MenuConsoleOption("MAIN", "Torna alla configurazione principale."))


        mm.multiselection = False

        Dim r As String = mm.Show

        Select Case r
            Case "WEB"
            Case "MAIL"
            Case "API"
            Case "MAIN"
                display_menu_main()
            Case Else
                GoTo display_menu_cfg_start
        End Select
    End Sub

    Sub display_menu_var()
display_menu_var_start:

        Dim mm As New MenuConsole
        mm.setTitle("MENU VARIABILI (" & My.Application.Info.AssemblyName & ")")
        Dim ms As New List(Of String)

        'VOCI DEL MENU
        mm.Options.Add(New MenuConsoleOption("VIEW", "Visualizza variabili."))
        mm.Options.Add(New MenuConsoleOption("MOD", "Modifica variabili."))
        mm.Options.Add(New MenuConsoleOption("DEL", "Cancella variabili."))
        mm.Options.Add(New MenuConsoleOption("MAIN", "Torna alla configurazione principale."))


        mm.multiselection = False

        Dim r As String = mm.Show

        Select Case r
            Case "VIEW"
                display_variabiles()
            Case "MOD"
                display_menu_var_mod()
            Case "DEL"
            Case "MAIN"
                display_menu_main()
            Case Else
                GoTo display_menu_var_start
        End Select
    End Sub




    Sub display_variabiles()
        Console.Clear()
        Dim list As List(Of local_db._record) = variabili.data

        Dim max_space = 0
        For Each r As local_db._record In list
            If r.Key.Length > max_space Then max_space = r.Key.Length
        Next


        For Each r As local_db._record In list
            Console.WriteLine(r.Key & Space(max_space - r.Key.Length) & " = " & r.Value)
        Next
        Dim x = Console.ReadKey
        display_menu_var()
    End Sub

    Sub display_menu_var_mod()
display_menu_var_modr_start:

        Dim mm As New MenuConsole
        mm.setTitle("MODIFICA VARIABILI (" & My.Application.Info.AssemblyName & ")")
        Dim ms As New List(Of String)

        Dim list As List(Of local_db._record) = variabili.data
        For Each rv As local_db._record In list
            mm.Options.Add(New MenuConsoleOption(rv.Key, rv.Key))
        Next

        'VOCI DEL MENU
        mm.Options.Add(New MenuConsoleOption("CANCEL", "Cancella variabili."))
        mm.Options.Add(New MenuConsoleOption("MAIN", "Torna alla configurazione principale."))


        mm.multiselection = False

        Dim r As String = mm.Show

        Select Case r
            Case "CANCEL"
                display_menu_var()
            Case "MAIN"
                display_menu_main()
            Case Else
                Console.ReadLine()
                display_menu_var()

        End Select
    End Sub





    Function input_menu_main() As Char
        Console.WriteLine("Seleziona l'azione")
        Dim r As ConsoleKeyInfo = Console.ReadKey
        Return r.KeyChar
    End Function

    Sub display_menu_new_program()

        Dim intInput As Integer = 0
        Do
            Console.Clear()
            Console.WriteLine("############# MENU NUOVO ##############")
            Console.WriteLine()
            Console.WriteLine("1) PROGRAMMA")
            Console.WriteLine()
            Console.WriteLine("2) VERSIONE")
            Console.WriteLine()
            Console.WriteLine("3) DATABASE")
            Console.WriteLine()
            Integer.TryParse(input_menu_main, intInput)
            Select Case intInput
                Case 1
                    'read_File()
                    Exit Select
                Case 2





                    Exit Select
                Case 3
                    Exit Sub
                Case Else
                    Exit Sub
            End Select



        Loop





    End Sub


    Public Class MenuConsole

        Public Property BackColor As ConsoleColor = ConsoleColor.Black
        Public Property ForeColor As ConsoleColor = ConsoleColor.DarkRed

        'Header
        '   Titolo
        Dim TitleAlign As horizontal_align = horizontal_align.center
        Dim TitleText As String = "Menu"
        Dim TitleColor As ConsoleColor = ConsoleColor.Red
        Dim TitleBackColor As ConsoleColor = Nothing
        '   Descrizione
        Dim DescriptionText As String = ""
        Dim DescriptionColor As ConsoleColor = ConsoleColor.White
        Dim DescriptionBackColor As ConsoleColor = Nothing
        Dim DescriptionFixHeight As Integer = -1

        'MENU
        Dim MenuAlign As horizontal_align = horizontal_align.center
        Public DefaultMenuColor As ConsoleColor = ConsoleColor.DarkRed



        Public Property LabelColor As ConsoleColor = ConsoleColor.White
        Public Property Options As List(Of MenuConsoleOption) = New List(Of MenuConsoleOption)
        Public Property Options_selected As New List(Of String)
        Public Property Spacing As Integer = 3
        Public Property Selection As String = " >> "
        Public Property Escape_enable As Boolean = True
        Public Property Escape_Label As String = "[ESC] Esci"
        Public Property f10_label As String = "[F10] Conferma"
        Public Property multiselection As Boolean = True

        ''' <summary>
        ''' Elenco dei tasti funzione da implementare all'interno del codice
        ''' </summary>
        ''' <value>Classe FunctionKey(ConsoleKey, Descrizione)</value>
        ''' <returns>Lista dei tasti funzione programmati</returns>
        ''' <remarks></remarks>
        Private Property functions As List(Of FunctionKey) = New List(Of FunctionKey)

        Private selected_collection As New Collection
        Private selected_last As Integer = 0
        Private area_border As Integer = 1
        Private area_header_height As Integer = 2
        Private area_footer_height As Integer = 1
        Private area_max_rows As Integer = 0
        Private area_max_column_num As Integer = 0
        Private area_fn_rows As Integer = 0
        Private area_column_width As Integer = 0
        Private area_desc_rows As Integer = 0


        Dim number_max_length As Integer = 0    '   Numero di caratteri utilizzata dai numeri delle vocie del menu
        Dim option_length As Integer = 0
        Dim num_tot_rows As Integer = Console.WindowHeight
        Dim num_tot_chars As Integer = Console.WindowWidth
        Dim num_col_width As Integer = 0
        Dim fn_menu As String = ""
        Dim selected As Integer = -1
        Dim selected_text As String = ""
        Dim curItem As Integer = -1

        Dim Colors As New Collection

        Public Event functionKey(ByVal fn As FunctionKey, ByVal Selected As ArrayList, ByRef sender As MenuConsole)



        Public Enum horizontal_align
            Left = 0
            center = 1
            Right = 2
        End Enum

        Public Enum vertical_align
            top = 0
            middle = 1
            bottom = 2
        End Enum

        Sub New()
            Console.SetWindowSize(Console.WindowWidth, Console.WindowHeight)
            functions.Add(New FunctionKey(ConsoleKey.UpArrow, "", False, special_functions.selection_move_up))
            functions.Add(New FunctionKey(ConsoleKey.DownArrow, "", False, special_functions.selection_move_down))
            functions.Add(New FunctionKey(ConsoleKey.LeftArrow, "", False, special_functions.selection_move_left))
            functions.Add(New FunctionKey(ConsoleKey.RightArrow, "", False, special_functions.selection_move_right))
            functions.Add(New FunctionKey(ConsoleKey.Spacebar, "", False, special_functions.select_item))
            functions.Add(New FunctionKey(ConsoleKey.Enter, "", False, special_functions.select_item_and_move_next))
        End Sub

        Public Function Show() As String
refresh_all:

            Console.ResetColor()
            Console.Clear()
            Dim response As String = Nothing
            'Set the color
            Console.BackgroundColor = Me.BackColor
            Console.ForegroundColor = Me.ForeColor

            number_max_length = Options.Count.ToString.Length
            Dim desc_length As Integer = 0
            'Determino la larghezza massima della voce di menu..
            For Each opt As MenuConsoleOption In Options
                Dim lbl As String = opt.label
                If lbl.Length > option_length Then option_length = lbl.Length
                If String.Format("{0}", opt.Description).Length > desc_length Then desc_length = opt.Description.Length
            Next
            'Determino la grandezza del menu funzioni
            area_fn_rows = IIf(functions.Count > 0, 1, 0) ' Numero di riga del menu
            For Each fn As FunctionKey In functions
                Dim elem As String = "[" & fn.Value.ToString & "] " & fn.Description
                If fn_menu = "" Then
                    fn_menu = Space(area_border)
                Else
                    fn_menu &= "   "
                End If
                If elem.Length > Console.WindowWidth - (area_border * 2) Then elem = Left(elem, Console.WindowWidth - (area_border * 2))
                If (elem + fn_menu).Length >= (Console.WindowWidth * area_fn_rows) - (area_border * 2) Then
                    fn_menu &= Space((Console.WindowWidth * area_fn_rows) - fn_menu.Length)
                    area_fn_rows += 1
                    fn_menu &= Space(area_border)
                End If
                fn_menu &= "[" & fn.Value.ToString & "] " & fn.Description
            Next
            If fn_menu <> "" Then
                fn_menu &= Space((Console.WindowWidth * area_fn_rows) - fn_menu.Length)
            End If
            area_footer_height += area_fn_rows

            'Calcolo l'altezza di un campo descrizione (max 2 righe)
            If desc_length > 0 Then
                If desc_length < WindowWidth - 2 Then
                    area_desc_rows = 1
                Else
                    area_desc_rows = 2
                End If
            End If
            area_footer_height += area_desc_rows


            'Le descrizioni non possono essere maggiori della larghezza della console
            If option_length > Console.WindowWidth - (area_border * 2) Then option_length = Console.WindowWidth - (area_border * 2)
            area_max_rows = Console.WindowHeight - area_footer_height - area_header_height   '   Numero massimo di righe del menu
            area_max_column_num = Options.Count / (area_max_rows)    '   Numero di colonne calcolato in base al numero di voci   del menu
            If area_max_column_num < Options.Count / (area_max_rows) Then area_max_column_num += 1
            area_column_width = (Console.WindowWidth - (area_border * 2)) / area_max_column_num






            'Dim num_length As Integer = Options.Count.ToString.Length
            Me.Spacing = number_max_length + 1
            num_tot_rows = Console.WindowHeight
            num_tot_chars = Console.WindowWidth

            Dim num_cols As Integer = area_max_column_num
            'If num_cols < Options.Count / (num_tot_rows - 2) Then num_cols += 1
            num_col_width = area_column_width

            If Not Options_selected.Count = 0 Then
                For Each o As String In Options_selected
                    If selected_collection.Contains(o) Then selected_collection.Remove(o)
                    selected_collection.Add(o, o)
                Next
            End If


inizio:
            Dim key As New ConsoleKeyInfo
            'Dim curItem As Short = 0, c
            Dim rows_step = num_tot_rows - 3

            write_header()
            write_footer()
menu:

            write_menu(curItem)


            '########## SELEZIONE
            Console.CursorTop = 0
            Console.BackgroundColor = Me.BackColor
            Console.ForegroundColor = Me.LabelColor
            Console.CursorLeft = area_border
            Console.CursorTop = Console.WindowHeight - area_fn_rows - 1
            Dim ln As String = "Selezione: " & selected_text
            Console.Write(ln & Space(WindowWidth - (area_border * 2) - ln.Length))
            Console.CursorLeft = area_border + ln.Length




            key = Console.ReadKey(True)

            'If key.Key.ToString() = "DownArrow" Then
            '    'curItem += 1
            '    'If curItem > Me.Options.Count - 1 Then curItem = 0
            '    'selected = curItem + 1
            '    'GoTo menu

            '    'ElseIf key.Key.ToString() = "UpArrow" Then
            '    '    curItem -= 1
            '    '    If curItem < 0 Then curItem = Convert.ToInt16(Me.Options.Count - 1)
            '    '    selected = curItem + 1
            '    '    GoTo menu
            'ElseIf key.Key.ToString() = "RightArrow" Then
            '    curItem += num_tot_rows - area_header_height - area_footer_height

            '    If curItem > Me.Options.Count - 1 Then curItem = curItem - (area_max_rows) * (CInt(curItem / (area_max_rows)))
            '    If curItem < 0 Then curItem = Me.Options.Count - 1
            '    selected = curItem + 1
            '    GoTo menu
            'ElseIf key.Key.ToString() = "LeftArrow" Then
            '    curItem -= area_max_rows
            '    If curItem < 0 Then curItem = curItem + (area_max_rows * area_max_column_num)
            '    If curItem > Me.Options.Count - 1 Then curItem = Convert.ToInt16(Me.Options.Count - 1)

            '    selected = curItem + 1
            '    GoTo menu
            'Else
            If key.Key.ToString() = "Backspace" Then
                If selected.ToString.Length > 0 Then
                    Dim res As String = Left(selected, selected.ToString.Length - 1)
                    selected = CInt(0 & res)
                    selected_text = res
                    curItem = selected - 1
                End If
                GoTo menu
            ElseIf isFunction(key.Key) Then '   --                                              is a function
                Console.Write(key.Key.ToString)
                Dim fn As FunctionKey = getFunction(key.Key)
                If fn.special_function <> special_functions.none Then
                    Select Case fn.special_function '   --                                      is a special function
                        Case special_functions.selection_move_up '   move up
                            curItem -= 1
                            If curItem < 0 Then curItem = Convert.ToInt16(Me.Options.Count - 1)
                            selected = curItem + 1
                            GoTo menu
                        Case special_functions.selection_move_down  '   nome down
                            curItem += 1
                            If curItem > Me.Options.Count - 1 Then curItem = 0
                            selected = curItem + 1
                            GoTo menu
                        Case special_functions.selection_move_right
                            curItem += num_tot_rows - area_header_height - area_footer_height

                            If curItem > Me.Options.Count - 1 Then curItem = curItem - (area_max_rows) * (CInt(curItem / (area_max_rows)))
                            If curItem < 0 Then curItem = Me.Options.Count - 1
                            selected = curItem + 1
                            GoTo menu
                        Case special_functions.selection_move_left
                            curItem -= area_max_rows
                            If curItem < 0 Then curItem = curItem + (area_max_rows * area_max_column_num)
                            If curItem > Me.Options.Count - 1 Then curItem = Convert.ToInt16(Me.Options.Count - 1)

                            selected = curItem + 1
                            GoTo menu
                        Case special_functions.select_item
                            If multiselection Then
                                If curItem >= 0 Then
                                    If Not selected_collection.Contains(Me.Options.Item(curItem).Value) Then
                                        selected_collection.Add(Me.Options.Item(curItem).Value, Me.Options.Item(curItem).Value)
                                    Else
                                        selected_collection.Remove(Me.Options.Item(curItem).Value)
                                    End If
                                    selected_text = ""
                                End If
                                GoTo menu

                            Else
                                If Not curItem = -1 Then
                                    response = Me.Options.Item(curItem).Value
                                    GoTo fine
                                End If

                            End If
                        Case special_functions.select_item_and_move_next
                            If multiselection Then
                                If curItem >= 0 Then
                                    If Not selected_collection.Contains(Me.Options.Item(curItem).Value) Then
                                        selected_collection.Add(Me.Options.Item(curItem).Value, Me.Options.Item(curItem).Value)
                                    Else
                                        selected_collection.Remove(Me.Options.Item(curItem).Value)
                                    End If
                                    selected_text = ""
                                End If
                                If (key.Key.ToString() = "Enter") Then
                                    curItem += 1
                                    If curItem > Me.Options.Count - 1 Then curItem = 0
                                    selected = curItem + 1
                                    GoTo menu
                                End If
                                'selected = ""
                                GoTo menu
                            End If

                    End Select
                End If
                Dim sel As ArrayList = New ArrayList
                If multiselection Then
                    sel = getSelected2ArrayList()
                Else
                    sel.Add(Options(curItem).Value)
                End If
                RaiseEvent functionKey(fn, sel, Me)
                If Not fn.exitMenu Then GoTo menu
            ElseIf IsNumeric(key.KeyChar) Then
                selected_text &= key.KeyChar
                selected = selected_text
                If selected > (Options.Count) Then
                    Dim res As String = Left(selected, selected.ToString.Length - 1)
                    selected = CInt(0 & res)
                    selected_text = selected
                End If


                curItem = selected - 1
                GoTo menu
            ElseIf key.Key.ToString() = "F10" And multiselection = True Then
                response = list()
            ElseIf (key.Key.ToString() = "Enter" Or key.Key.ToString() = "F10") And multiselection = False Then
                If Not curItem = -1 Then
                    response = Me.Options.Item(curItem).Value
                    GoTo fine

                End If
            ElseIf key.Key.ToString() = "Escape" Then
                If Escape_enable Then
                    response = Nothing
                End If
            ElseIf (key.Key.ToString() = "Enter" Or key.Key.ToString() = "Spacebar") And multiselection = True Then
                If curItem >= 0 Then
                    If Not selected_collection.Contains(Me.Options.Item(curItem).Value) Then
                        selected_collection.Add(Me.Options.Item(curItem).Value, Me.Options.Item(curItem).Value)
                    Else
                        selected_collection.Remove(Me.Options.Item(curItem).Value)
                    End If
                    selected_text = ""
                End If
                If (key.Key.ToString() = "Enter") Then
                    curItem += 1
                    If curItem > Me.Options.Count - 1 Then curItem = 0
                    selected = curItem + 1
                    GoTo menu
                End If
                'selected = ""
                GoTo menu
            Else


            End If



fine:
            'Finish
            ' Return Console.ReadLine()
            Return response
        End Function

        Private Sub specialFunction(sel As special_functions)
            Select Case sel
                Case special_functions.select_item

                Case special_functions.select_all

                Case special_functions.clear_selections


                Case special_functions.confirm_selection



            End Select
        End Sub

        Public Sub addFunction(ByVal key As FunctionKey)
            Select Case key.Value
                Case ConsoleKey.DownArrow, ConsoleKey.LeftArrow, ConsoleKey.RightArrow, ConsoleKey.UpArrow
                    'Non viene aggiunta la funzionalità
                Case Else
                    Dim presente As Boolean = False
                    For Each k As FunctionKey In functions
                        If k.Value = key.Value Then
                            k.Description = key.Description
                            k.exitMenu = key.exitMenu
                            k.special_function = key.special_function
                            presente = True
                            Exit For
                        End If
                    Next
                    If Not presente Then
                        functions.Add(key)
                    End If
            End Select
        End Sub

        Private Function isFunction(key As ConsoleKey) As Boolean
            Dim r As Boolean = False
            For Each f As FunctionKey In functions
                If f.Value = key Then
                    r = True
                    Exit For
                End If
            Next
            Return r
        End Function

        Private Function getFunction(key As ConsoleKey) As FunctionKey
            Dim fn As New FunctionKey(ConsoleKey.NoName, "")
            For Each f As FunctionKey In functions
                If f.Value = key Then
                    fn = f
                    Exit For
                End If
            Next
            Return fn
        End Function

        Private Function getSelected2ArrayList() As ArrayList
            Dim a As New ArrayList
            For i = 1 To selected_collection.Count
                Dim item = selected_collection(i)
                a.Add(item)
            Next
            Return a
        End Function

        Private Function write_header()
            If IsNothing(TitleBackColor) Then
                Console.BackgroundColor = BackColor
            Else
                Console.BackgroundColor = TitleBackColor
            End If
            Console.ForegroundColor = Me.TitleColor
            'Print the title
            Console.CursorTop = 0
            Console.CursorLeft = 0
            Console.Write(Space(Console.WindowWidth))
            Console.CursorTop = 0
            If TitleText.Length > (Console.WindowWidth) Then TitleText = Left(TitleText, Console.WindowWidth - 2)
            Select Case TitleAlign
                Case horizontal_align.Left
                    Console.CursorLeft = 1
                Case horizontal_align.center
                    Console.CursorLeft = (Console.WindowWidth \ 2) - (Me.TitleText.Length \ 2)
                Case horizontal_align.Right
                    Console.CursorLeft = Console.WindowWidth - TitleText.Length - 1
            End Select
            Console.Write(Me.TitleText)

            If DescriptionText <> "" Then
                Console.BackgroundColor = IIf(IsNothing(DescriptionBackColor), BackColor, DescriptionBackColor)
                Console.ForegroundColor = IIf(IsNothing(DescriptionColor), ForeColor, DescriptionColor)
                Console.CursorTop = 1
                Console.CursorLeft = 0
                If DescriptionText.Length > WindowWidth * DescriptionFixHeight Then DescriptionText = Left(DescriptionText, (WindowWidth * DescriptionFixHeight) - 3) & " .."
                Console.Write(DescriptionText & Space((WindowWidth * DescriptionFixHeight) - DescriptionText.Length))
                area_header_height = DescriptionFixHeight + 1
            End If


        End Function

        Private Sub write_footer()
            Console.BackgroundColor = Me.ForeColor
            Console.ForegroundColor = Me.BackColor
            Console.CursorLeft = 0
            Console.CursorTop = Console.WindowHeight - area_fn_rows
            Console.Write(fn_menu)



            Console.BackgroundColor = Me.BackColor
            Console.ForegroundColor = Me.ForeColor

        End Sub

        Private Sub write_description(desc As String)
            Console.BackgroundColor = DescriptionBackColor
            Console.ForegroundColor = DescriptionColor

            Dim d As String = ""
            If area_desc_rows > 0 Then
                Dim start_pos As Integer = 1
                For r As Integer = 1 To area_desc_rows
                    d &= Space(area_border) & Mid(desc, start_pos, WindowWidth - (area_border * 2)) & Space(area_border)
                    start_pos += WindowWidth - (area_border * 2)
                    If d.Length < (start_pos - 1) + ((area_border * 2) * r) Then d = d & Space(((start_pos - 1) + ((area_border * 2) * r)) - d.Length)
                Next

            End If
            Console.CursorLeft = 0
            Console.CursorTop = WindowHeight - area_footer_height
            Console.Write(d)
        End Sub

        Private Sub write_menu(sel As Integer)
            If sel > -1 Then
                Dim s As String = selected_last & "," & sel
                For Each op As String In s.Split(",")
                    Dim i As Integer = CInt(op)
                    Dim item As MenuConsoleOption = Me.Options.Item(i)
                    Console.CursorLeft = Options(i).X
                    Console.CursorTop = Options(i).Y
                    If i = sel Then
                        Console.BackgroundColor = IIf(IsNothing(Options(i).menuSelBackColor), ForeColor, Options(i).menuSelBackColor)
                        Console.ForegroundColor = IIf(IsNothing(Options(i).menuSelColor), BackColor, Options(i).menuSelColor)
                    Else
                        Console.BackgroundColor = IIf(IsNothing(Options(i).menuBackColor), BackColor, Options(i).menuBackColor)
                        Console.ForegroundColor = IIf(IsNothing(Options(i).menuColor), DefaultMenuColor, Options(i).menuColor)
                    End If


                    If multiselection Then
                        If isSelected(Options(op).Value) Then
                            Console.Write("[X] ")
                        Else
                            Console.Write(Space(Selection.Length))
                        End If
                    Else
                        If i = sel Then
                            Console.Write(Selection)
                        Else
                            Console.Write(Space(Selection.Length))
                        End If
                    End If
                    Dim label As String
                    label = (i + 1).ToString & ") " + Space(number_max_length - (i + 1).ToString.Length) & item.label

                    If label.Length + Selection.Length > num_col_width Then label = Left(label, num_col_width - 4 - Selection.Length) & " .."

                    If label.Length + Selection.Length < num_col_width Then label = label + Space(num_col_width - Selection.Length - label.Length)


                    Console.WriteLine(label)

                Next


            Else

                Dim act_col As Integer = 0
                'Dim curItem As Short = 0, c
                Dim rows_step = num_tot_rows - area_footer_height - area_header_height

                Console.CursorLeft = 0 : Console.CursorTop = area_header_height
                'Dim cleft As Integer = Me.Options.Max(Function(m) m.Value.Length) + Me.Spacing
                For i As Integer = 0 To Me.Options.Count - 1
                    If i = 0 Then
                        Console.BackgroundColor = IIf(IsNothing(Options(i).menuSelBackColor), ForeColor, Options(i).menuSelBackColor)
                        Console.ForegroundColor = IIf(IsNothing(Options(i).menuSelColor), BackColor, Options(i).menuSelColor)
                    Else
                        Console.BackgroundColor = IIf(IsNothing(Options(i).menuBackColor), BackColor, Options(i).menuBackColor)
                        Console.ForegroundColor = IIf(IsNothing(Options(i).menuColor), DefaultMenuColor, Options(i).menuColor)
                    End If

                    If i >= (1 + act_col) * rows_step Then
                        act_col += 1 : Console.CursorTop = area_header_height
                    Else
                        If i > 0 Then
                            Console.CursorTop = Options(i - 1).Y + 1
                        End If
                    End If

                    Dim _left As Integer = (act_col * num_col_width)


                    Console.CursorLeft = _left

                    Dim label As String = ""

                    Dim item As MenuConsoleOption = Me.Options.Item(i)

                    'If selected = -1 Then
                    '    curItem = 0
                    '    selected = 0
                    'Else
                    '    If CInt(selected) <= Options.Count And CInt(selected) > -1 Then
                    '        curItem = CInt(selected) - 1
                    '    Else
                    '        curItem = 0
                    '        selected = 0
                    '    End If
                    'End If

                    If curItem = -1 Then curItem = 0

                    If curItem = i Then
                        Console.BackgroundColor = Me.ForeColor
                        Console.ForegroundColor = Me.BackColor
                    End If


                    Options(i).X = Console.CursorLeft
                    Options(i).Y = Console.CursorTop


                    If multiselection Then
                        If isSelected(Options(i).Value) Then
                            Console.Write("[X] ")
                            'Console.BackgroundColor = Me.ForeColor
                            'Console.ForegroundColor = Me.BackColor

                        Else
                            Console.Write(Space(Selection.Length))
                        End If
                    Else
                        If curItem = i Then
                            Console.Write(Selection)
                            'Console.BackgroundColor = Me.ForeColor
                            'Console.ForegroundColor = Me.BackColor

                        Else
                            Console.Write(Space(Selection.Length))
                        End If
                    End If


                    'Console.CursorLeft = _left + Selection.Length


                    label = (i + 1).ToString & ") " + Space(number_max_length - (i + 1).ToString.Length) & item.label

                    If label.Length + Selection.Length > num_col_width Then label = Left(label, num_col_width - 4 - Selection.Length) & " .."

                    If label.Length + Selection.Length < num_col_width Then label = label + Space(num_col_width - Selection.Length - label.Length)


                    'Console.CursorLeft = _left
                    'Console.Write(num)
                    'Console.CursorLeft = left + (i + 1).ToString.Length + 1
                    'Dim col As Integer = _left + num.Length
                    'If col > num_tot_chars Then col = num_tot_chars - 1
                    'Console.CursorLeft = col
                    Console.WriteLine(label)
                Next
            End If


            selected_last = curItem
            write_description(Options(curItem).Description)


        End Sub


        Private Function isSelected(ByVal currentItem As Object) As Boolean
            Return selected_collection.Contains(currentItem)
        End Function

        Private Function list() As String
            Dim r As String = ""
            For Each itm In Me.selected_collection
                If r = "" Then
                    r = itm
                Else
                    r = r & "," & itm
                End If
            Next

            Return r
        End Function

        Public Sub setTitle(_titolo As String, Optional _title_align As horizontal_align = horizontal_align.center, Optional _color As ConsoleColor = ConsoleColor.Red, Optional _backColor As ConsoleColor = Nothing)
            TitleText = _titolo
            TitleAlign = _title_align
            TitleColor = _color
            TitleBackColor = _backColor
        End Sub

        Public Sub setDescription(description As String, Optional descColor As ConsoleColor = Nothing, Optional descBackColor As ConsoleColor = Nothing, Optional fixHeight As Integer = -1)
            DescriptionText = description
            DescriptionColor = descColor
            DescriptionBackColor = descBackColor
            If fixHeight = -1 Or fixHeight = 0 Or fixHeight >= WindowHeight - 3 Then
                DescriptionFixHeight = DescriptionText.Length / WindowWidth
                If DescriptionFixHeight < (DescriptionText.Length / WindowHeight) Then DescriptionFixHeight += 1
            Else
                DescriptionFixHeight = fixHeight
            End If

        End Sub



    End Class

    Public Class MenuConsoleOption
        Dim _x As Integer = Nothing
        Dim _y As Integer = Nothing
        Public Property X As Integer
            Set(value As Integer)
                _x = value
            End Set
            Get
                Return _x
            End Get
        End Property

        Public Property Y As Integer
            Set(value As Integer)
                _y = value
            End Set
            Get
                Return _y
            End Get
        End Property

        Dim _color As Object = Nothing
        Dim _backColor As Object = Nothing
        Dim _selColor As Object = Nothing
        Dim _selBackColor As Object = Nothing

        Public Property menuColor As Object
            Set(value As Object)
                If Not IsNothing(value) AndAlso isColor(value) Then
                    _color = value
                Else
                    _color = Nothing
                End If
            End Set
            Get
                Return _color
            End Get
        End Property

        Private Function isColor(ByVal value As Object) As Boolean
            Dim r As Boolean = False
            Try
                If value.GetType.FullName = GetType(ConsoleColor).FullName Then
                    If CInt(value) >= 0 And CInt(value) <= 15 Then
                        r = True
                    End If
                End If
            Catch ex As Exception

            End Try
            Return r
        End Function


        Public Property menuBackColor As Object
            Set(value As Object)
                If Not IsNothing(value) AndAlso isColor(value) Then
                    _backColor = value
                Else
                    _backColor = Nothing
                End If
            End Set
            Get
                Return _backColor
            End Get
        End Property

        Public Property menuSelColor As Object
            Set(value As Object)
                If Not IsNothing(value) AndAlso isColor(value) Then
                    _selColor = value
                Else
                    _selColor = Nothing
                End If
            End Set
            Get
                Return _selColor
            End Get
        End Property

        Public Property menuSelBackColor As Object
            Set(value As Object)
                If Not IsNothing(value) AndAlso isColor(value) Then
                    _selBackColor = value
                Else
                    _selBackColor = Nothing
                End If
            End Set
            Get
                Return _selBackColor
            End Get
        End Property

        Sub New(ByVal _value As String, ByVal _label As String, Optional _description As String = "")
            Value = _value
            label = _label
            Description = _description
        End Sub
        Public Property Value As String
        Public Property label As String
        Public Property Description As String

    End Class

    ''' <summary>
    ''' Special functions enbedded
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum special_functions
        none = -1
        clear_selections = 0
        selection_move_up = 1
        selection_move_left = 2
        selection_move_right = 3
        selection_move_down = 4
        select_all = 10
        confirm_selection = 1000
        select_item = 100
        select_item_and_move_next = 110
    End Enum

    Public Class FunctionKey
        Sub New(ByVal _value As ConsoleKey, ByVal _description As String, Optional ByVal _exitMenu As Boolean = False, Optional ByVal _special_function As special_functions = special_functions.none)
            Value = _value
            Description = _description
            exitMenu = _exitMenu
            special_function = _special_function
        End Sub
        Public Property Value As ConsoleKey
        Public Property Description As String
        Public Property exitMenu As Boolean
        Public Property special_function As special_functions
    End Class

    Private Class mColor
        Public Property backColor As ConsoleColor
        Public Property foreColor As ConsoleColor

    End Class

    Private Sub fixConsole()
        Dim handle As IntPtr
        handle = Process.GetCurrentProcess.MainWindowHandle ' Get the handle to the console window

        Dim sysMenu As IntPtr
        sysMenu = GetSystemMenu(handle, False) ' Get the handle to the system menu of the console window

        If handle <> IntPtr.Zero Then
            'DeleteMenu(sysMenu, SC_CLOSE, MF_BYCOMMAND) ' To prevent user from closing console window
            DeleteMenu(sysMenu, SC_MINIMIZE, MF_BYCOMMAND) 'To prevent user from minimizing console window
            DeleteMenu(sysMenu, SC_MAXIMIZE, MF_BYCOMMAND) 'To prevent user from maximizing console window
            DeleteMenu(sysMenu, SC_SIZE, MF_BYCOMMAND) 'To prevent the use from re-sizing console window
        End If



    End Sub




End Module
