Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports System.Collections.Generic

'---------------- Bolts and Nuts ------------------

' ter verificatie van getallen zie: http://www.werktuigbouw.nl/calculators/e3_6a.htm
Public Class Form1
    Public Shared bolttype() As String =
     {"Description; Kerndia;  dia_thread; width head; spoed; buitendiameter",                                      'width head: http://stsindustrial.com/a4-hex-cap-screw-technical-data/
    "M 6; 4.773	;   5.350	;9.78	;1.00;	6	",
    "M 8; 6.466	;   7.188	;12.73	;1.25;	8	",
    "M 10;8.160	;   9.026	;15.73	;1.50;	10	",
    "M 12;9.853	;   10.863	;17.73	;1.75;	12	",
    "M 14;11.546	;12.701	;20.67	;2.00;	14	",
    "M 16;13.546	;14.701	;23.67	;2.00;	16	",
    "M 18;14.933	;16.376	;26.67	;2.50;	18	",
    "M 20;16.933	;18.376	;29.67	;2.50;	20	",
    "M 22;18.933	;20.376	;32.61	;2.50;	22	",
    "M 24;20.319	;22.051	;35.61	;3.00;	24	",
    "M 27;23.319	;25.051	;39.61	;3.00;	27	",
    "M 30;25.706	;27.727	;44.61	;3.50;	30	",
    "M 33;28.706	;30.727	;49.61	;3.50;	33	",
    "M 36;31.093	;33.402	;53.54	;4.00;	36	",
    "M 42;36.479	;39.077	;62.54	;4.50;	42	"}

    'SKF book page 1278
    '"Description; Kerndia;  spoed; OD",
    Public Shared Lock_nut() As String =
     {
    "KM19, M95x2;   95; 2; 125",
    "KM20, M100x2; 100; 2; 130",
    "KM22, M110x2; 110; 2; 145",
    "kM24, M120x2; 120; 2; 155",
    "kM26, M130x2; 130; 2; 165",
    "kM28, M140x2; 140; 2; 180",
    "kM30, M150x2; 150; 2; 195",
    "kM32, M160x3; 160; 3; 210",
    "kM34, M170x3; 170; 3; 220",
    "kM36, M180x3; 180; 3; 230",
    "kM38, M190x3; 190; 3; 240",
    "kM40, M120x3; 200; 3; 250",
    "HM44T, Tr220x4; 220;4;280",
    "HM48T, Tr240x4; 240;4;300",
    "HM52T, Tr260x4; 260;4;330",
    "HM56T, Tr280x4; 280;4;350",
    "HM3160,Tr300x4; 300;4;380",
    "HM3164,Tr320x5; 320;5;400"
    }

    Public Shared boltgrade() As String =
     {"Description; treksterkte;rekgrens",
     " 3.6;300;180",
    " 4.6;400;240",
    " 4.8;400;320",
    " 5.6;500;300",
    " 5.8;500;400",
    " 6.8;600;480",
    " 8.8;800;640",
    " 9.8;900;720",
    "10.9;1000;900",
    "12.9;1200;1080",
    "A4-50;500;210",
    "A4-70;700;450",
    "A4-80;800;600"}

    Public Shared geinstaantbout() As String =
     {"Aantal",
    "1",
    "2",
    "4",
    "8",
    "12",
    "16",
     "20"}
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")      'Decimal separator "."
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")    'Decimal separator "."

        Bouten_combo()
        Geinstbout_combo()
        Grade_combo()
        Lock_nut_combo()
    End Sub
    Private Sub Bouten_combo()
        Dim words() As String

        ComboBox1.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 0 To (UBound(bolttype))                'Fill combobox 1 with bolt data
            words = bolttype(hh).Split(";")
            ComboBox1.Items.Add(Trim(words(0)))
        Next hh
        ComboBox1.SelectedIndex = 2

    End Sub
    Private Sub Lock_nut_combo()
        Dim words() As String

        ComboBox4.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 0 To Lock_nut.Length - 1            'Fill combobox 
            words = Lock_nut(hh).Split(";")
            ComboBox4.Items.Add(Trim(words(0)))
        Next hh
        ComboBox4.SelectedIndex = 1
    End Sub
    Private Sub Geinstbout_combo()
        Dim words() As String

        ComboBox2.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 0 To (UBound(geinstaantbout))                'Fill combobox 1 with bolt data
            words = geinstaantbout(hh).Split(";")
            ComboBox2.Items.Add(Trim(words(0)))
        Next hh
        ComboBox2.SelectedIndex = 2
    End Sub
    Private Sub Grade_combo()
        Dim words() As String

        ComboBox3.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 0 To (UBound(boltgrade))                'Fill combobox 1 with bolt data
            words = boltgrade(hh).Split(";")
            ComboBox3.Items.Add(Trim(words(0)))
        Next hh
        ComboBox3.SelectedIndex = 1
    End Sub
    Private Sub Aantalbouten()
        Dim motorverm, toerntal, Torque, dia, Fmotor, frictiecoefficient, F_fric, veiligfacmot, aantwaai As Double
        Dim safetyfact, rekgrens, kerndia, toegsp, dia_thread, d0 As Double
        Dim oppbout, Totoppbout, aantbout, F_bout As Double
        motorverm = NumericUpDown1.Value
        toerntal = NumericUpDown2.Value
        veiligfacmot = NumericUpDown6.Value
        aantwaai = NumericUpDown7.Value
        Torque = motorverm * veiligfacmot * 9550 / (aantwaai * toerntal)
        dia = NumericUpDown3.Value

        Fmotor = Torque / (1000 * dia / 2000)              'in [kN] 

        frictiecoefficient = NumericUpDown5.Value
        F_fric = Fmotor / frictiecoefficient      'in [kN] 

        Try
            Dim words2() As String = boltgrade(ComboBox3.SelectedIndex).Split(";")
            rekgrens = words2(2)
            Dim words1() As String = bolttype(ComboBox1.SelectedIndex).Split(";")
            kerndia = words1(1)
            dia_thread = words1(2)
        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Line 1290")  ' Show the exception's message.
        End Try

        d0 = (kerndia + dia_thread) / 2
        safetyfact = NumericUpDown17.Value
        NumericUpDown17.Enabled = False
        toegsp = safetyfact * rekgrens
        oppbout = 0.25 * PI * d0 ^ 2

        Totoppbout = F_fric * 1000 / toegsp           'van [kN] naar [N]

        aantbout = Totoppbout / oppbout
        F_bout = F_fric / aantbout                         'in [kN] 

        TextBox1.Text = Round(Torque, 0).ToString
        TextBox2.Text = Round(Fmotor, 0).ToString
        TextBox5.Text = Round(F_fric, 0).ToString
        TextBox6.Text = Round(rekgrens, 0).ToString

        TextBox3.Text = Round(toegsp, 0).ToString
        TextBox7.Text = Round(kerndia, 1).ToString
        TextBox8.Text = Round(oppbout, 2).ToString
        TextBox9.Text = Round(Totoppbout, 2).ToString
        TextBox12.Text = Round(aantbout, 2).ToString
        TextBox28.Text = Round(F_bout, 2).ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged, ComboBox3.SelectedIndexChanged
        Aantalbouten()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim qq, sigma02 As Double

        If (ComboBox1.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = bolttype(ComboBox1.SelectedIndex).Split(";")

            '--------------- select the strength @ temperature
            qq = NumericUpDown4.Value

            Select Case True
                Case (qq > 0 AndAlso qq < 100)
                    Double.TryParse(words(2), sigma02)     'Sigma 0.2 [N/mm]
                Case (qq >= 100 AndAlso qq < 200)
                    Double.TryParse(Math.Round(0.85 * words(2)), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq >= 200 AndAlso qq < 300)
                    Double.TryParse(Math.Round(0.8 * words(2)), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq >= 300 AndAlso qq < 400)
                    Double.TryParse(Math.Round(0.75 * words(2)), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq >= 400)
                    Double.TryParse(Math.Round(0.7 * words(2)), sigma02)    'Sigma 0.2 [N/mm]
            End Select
            TextBox6.Text = sigma02.ToString

        End If
    End Sub
    Private Sub Zetting()
        Dim Ra_head, Ra_nut, Ra_plate1, Ra_plate2, no_rings, Ra_ring As Double
        Dim zet_head_plate, zet_nut_plate, zet_plate_plate, zet_head_ring, zet_plate_ring, zet_ring_ring, zet_tot As Double
        Dim length_bolt, E_modulus, oppbout, F_bout, elong_force As Double

        Ra_head = NumericUpDown13.Value
        Ra_nut = NumericUpDown14.Value
        Ra_plate1 = NumericUpDown8.Value
        Ra_plate2 = NumericUpDown10.Value
        no_rings = NumericUpDown11.Value
        Ra_ring = NumericUpDown9.Value

        zet_head_plate = (Ra_head + Ra_plate1) / 4
        zet_nut_plate = (Ra_plate2 + Ra_nut) / 4
        zet_plate_plate = (Ra_plate1 + Ra_plate2) / 4
        zet_head_ring = (Ra_head + Ra_ring) / 4
        zet_plate_ring = (Ra_plate1 + Ra_ring) / 4
        zet_ring_ring = (Ra_ring + Ra_ring) / 4
        If no_rings > 0 Then
            zet_tot = zet_head_plate + zet_nut_plate + zet_plate_plate + zet_head_ring + zet_plate_ring + (no_rings - 1) * zet_ring_ring
        Else
            zet_tot = zet_head_plate + zet_nut_plate + zet_plate_plate
        End If
        length_bolt = NumericUpDown18.Value
        E_modulus = NumericUpDown12.Value
        Double.TryParse(TextBox8.Text, oppbout)
        Double.TryParse(TextBox28.Text, F_bout)
        'elong_force = (1 / 1000) * (F_fric * 1000) * (length_bolt / 1000) / ((oppbout * 10 ^ -6) * E_modulus * 10 ^ 9)   'elongation due to force in [mm]
        elong_force = F_bout * length_bolt / (oppbout * E_modulus)

        TextBox10.Text = Round(zet_head_plate, 1).ToString
        TextBox18.Text = Round(zet_nut_plate, 1).ToString
        TextBox17.Text = Round(zet_plate_plate, 1).ToString
        TextBox16.Text = Round(zet_head_ring, 1).ToString
        TextBox4.Text = Round(zet_plate_ring, 1).ToString
        TextBox13.Text = Round(zet_ring_ring, 1).ToString
        TextBox11.Text = Round(zet_tot, 1).ToString
        TextBox26.Text = Round(oppbout, 2).ToString
        TextBox14.Text = Round(F_bout, 2).ToString
        TextBox20.Text = Round(elong_force, 4).ToString   'Round(elong_force, 4).ToString

        If (zet_tot / 1000) < 0.8 * elong_force Then
            TextBox24.BackColor = Color.White
        Else
            TextBox24.BackColor = Color.Red
        End If
        TextBox24.Text = Round(zet_tot / 1000, 3).ToString
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, TabPage2.Enter
        Zetting()
    End Sub
    Private Sub Vastmoment()
        Dim arm_sleutel, frict_bout, dia_thread, dia_head, arm_fric_bout, dia_buiten As Double
        Dim F_a, F_f, M_f, spoed, M_netto, M_totaal, F_netto As Double
        Dim M_WD, beta, phi, rho_acc, MG_vast, MG_los, M_A As Double

        Double.TryParse(TextBox28.Text, F_a)      'Van F_fric naar F_a(=F_bout)=====voorspankracht
        arm_sleutel = NumericUpDown15.Value       'arm van de sleutel in [mm]
        frict_bout = NumericUpDown16.Value        'frictie factor die op bout werkt
        F_f = frict_bout * F_a                     'frictiekracht werkend halverwegen uitsteeksel head
        Try
            Dim words4() As String = bolttype(ComboBox1.SelectedIndex).Split(";")
            dia_thread = words4(2)
            dia_head = words4(3)
            spoed = words4(4)
            dia_buiten = words4(5)
        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Line 1290")  ' Show the exception's message.
        End Try

        arm_fric_bout = (dia_head + dia_thread) / 4               'arm van frictiekracht op bout
        M_f = F_f * arm_fric_bout           'moment door frictiekracht bout van [kN.mm] naar [Nm]

        'evenwicht: spoed*F_a=2*PI*arm_sleutel*F_netto          'M_netto=M_totaal=M_f
        F_netto = spoed * F_a / (2 * PI * arm_sleutel)          'in [kN]
        M_netto = F_netto * arm_sleutel                         'van [kN.mm] naar [Nm]
        M_totaal = M_netto + M_f                                'in [Nm]

        'Zie http://www.werktuigbouw.nl/
        M_WD = frict_bout * F_a * dia_buiten / 2         'Draagvlakwrijvingsmoment in [Nm]
        beta = 60 * PI / 180                             '[deg]
        phi = Atan(spoed / (PI * dia_thread))      '[deg]
        rho_acc = Atan(frict_bout / Cos(beta / 2))
        MG_vast = F_a * 0.5 * dia_thread * Tan(phi + rho_acc)    'draadwrijvingsmoment in [Nm]
        MG_los = F_a * 0.5 * dia_thread * Tan(phi - rho_acc)        'draadwrijvingsmoment bij losdraaien in [Nm]
        M_A = M_WD + MG_vast                            'totaal vastmoment in [Nm]

        'MessageBox.Show(dia_thread, dia_buiten)
        TextBox19.Text = Round(F_a, 2).ToString
        TextBox25.Text = Round(F_f, 2).ToString
        TextBox22.Text = Round(M_f, 2).ToString
        TextBox21.Text = Round(F_netto, 4).ToString
        TextBox15.Text = Round(M_netto, 2).ToString
        TextBox27.Text = Round(M_totaal, 2).ToString

        TextBox31.Text = Round(M_WD, 2).ToString
        TextBox34.Text = Round(MG_vast, 2).ToString
        TextBox33.Text = Round(MG_los, 2).ToString
        TextBox32.Text = Round(M_A, 4).ToString
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, TabPage3.Enter
        Vastmoment()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim oWord As Word.Application = Nothing
        oWord = New Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer

        'Start Word and open the document template. 
        font_sizze = 9
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 3
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze + 1
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = False
        oPara2.Range.Text = "Determination number of bolts" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox23.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox35.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Description "
        oTable.Cell(row, 2).Range.Text = TextBox53.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input Data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Installed motorvermogen"
        oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value
        oTable.Cell(row, 3).Range.Text = "[kW]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Veiligheidsfactor motor"
        oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Aantal waaiers"
        oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Toerental"
        oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value
        oTable.Cell(row, 3).Range.Text = "[rpm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter naaf, bout"
        oTable.Cell(row, 2).Range.Text = NumericUpDown3.Value
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Friction coefficient"
        oTable.Cell(row, 2).Range.Text = NumericUpDown5.Value
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Ontwerptemperatuur"
        oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value
        oTable.Cell(row, 3).Range.Text = "[C]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Type bout"
        oTable.Cell(row, 2).Range.Text = ComboBox1.SelectedItem


        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.9)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.6)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


        'Insert a 5 x 7 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 11, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Output"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Torque"
        oTable.Cell(row, 2).Range.Text = TextBox1.Text
        oTable.Cell(row, 3).Range.Text = "[Nm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Force due to motor"
        oTable.Cell(row, 2).Range.Text = TextBox2.Text
        oTable.Cell(row, 3).Range.Text = "[kN]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Force with friction"
        oTable.Cell(row, 2).Range.Text = TextBox5.Text
        oTable.Cell(row, 3).Range.Text = "[kN]"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Treksterkte"
        oTable.Cell(row, 2).Range.Text = TextBox6.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Toegestane spanning"
        oTable.Cell(row, 2).Range.Text = TextBox3.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Kerndiameter"
        oTable.Cell(row, 2).Range.Text = TextBox7.Text
        oTable.Cell(row, 3).Range.Text = "[mm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Opp bout"
        oTable.Cell(row, 2).Range.Text = TextBox8.Text
        oTable.Cell(row, 3).Range.Text = "[mm2]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Tot opp bout"
        oTable.Cell(row, 2).Range.Text = TextBox9.Text
        oTable.Cell(row, 3).Range.Text = "[mm2]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Aantal bouten"
        oTable.Cell(row, 2).Range.Text = TextBox12.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Geinstalleerd aantal bouten"
        oTable.Cell(row, 2).Range.Text = ComboBox2.SelectedItem

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.9)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.6)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown20.ValueChanged, ComboBox4.SelectedIndexChanged
        Vastmoment2()
    End Sub
    Private Sub Vastmoment2()
        'Zie http://www.werktuigbouw.nl/
        Dim area As Double
        Dim frict_bout, dia_thread As Double
        Dim F_a, F_f, spoed, lock_od As Double
        Dim M_WD, beta, phi, rho_acc, MG_vast, MG_los, M_A As Double

        frict_bout = NumericUpDown20.Value      'frictie factor die op bout werkt
        'frictiekracht werkend halverwegen uitsteeksel head
        Try
            Dim words4() As String = Lock_nut(ComboBox4.SelectedIndex).Split(";")
            Double.TryParse(words4(1), dia_thread)
            Double.TryParse(words4(2), spoed)
            Double.TryParse(words4(3), lock_od)
        Catch ex As Exception

        End Try

        area = PI / 4 * dia_thread ^ 2    '[mm2] area shaft 
        F_a = area * 800 * 0.8 * 0.6    '[N] voorspankracht in shaft
        F_a *= 0.1                      '10% van normale voorspan kracht
        F_f = frict_bout * F_a

        'Machine onderdelen pag 84, 24e druk
        M_WD = frict_bout * F_a * (dia_thread + lock_od) / 4    'Draagvlakwrijvingsmoment in [Nm]
        beta = 60 * PI / 180                                    'tophoek draad in [radians]
        phi = Atan(spoed / (PI * dia_thread))                   'hellingshoek door spoed
        rho_acc = Atan(frict_bout / Cos(beta / 2))
        MG_vast = F_a * (dia_thread / 2) * Tan(phi + rho_acc)   'draadwrijvingsmoment vast [Nm]
        MG_los = F_a * (dia_thread / 2) * Tan(phi - rho_acc)    'draadwrijvingsmoment los [Nm]
        M_A = M_WD + MG_vast                            'totaal aanhaal moment in [Nm]

        'MessageBox.Show(dia_thread, dia_buiten)
        TextBox30.Text = (dia_thread).ToString("0")   '[mm]
        TextBox40.Text = area.ToString("0")                 '[mm2]
        TextBox36.Text = (F_a / 10 ^ 3).ToString("0")       '[kN]

        TextBox38.Text = (M_WD / 10 ^ 3).ToString("0")      '[N.m]
        TextBox37.Text = (MG_vast / 10 ^ 3).ToString("0")   '[N.m]
        TextBox29.Text = (M_A / 10 ^ 3).ToString("0")       '[N.m]
    End Sub
End Class
