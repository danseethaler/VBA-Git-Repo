Attribute VB_Name = "CustomFunctions"
Option Explicit

Public Function RandEmpID() As Long

RandEmpID = WorksheetFunction.RandBetween(1000, 770000)

End Function

Public Function RandSSN() As String

RandSSN = WorksheetFunction.RandBetween(400000000, 600000000)

End Function

Public Function RandEmail(Name As String) As String

RandEmail = Left(Name, WorksheetFunction.Find(" ", Name) - 1) & "@test.com"

End Function

Public Function RandPhone() As String

RandPhone = WorksheetFunction.RandBetween(100, 999) & "-" & WorksheetFunction.RandBetween(100, 999) & "-" & WorksheetFunction.RandBetween(1000, 9999)

End Function

Public Function RandName() As String
Dim Location As Integer
Dim FirstNames As String
Dim LastNames As String

FirstNames = "Sophia;Emma;Olivia;Isabella;Mia;Ava;Lily;Zoe;Emily;Chloe;Layla;Madison;Madelyn;Abigail;Aubrey;Charlotte;Amelia;Ella;Kaylee;Avery;Aaliyah;Hailey;Hannah;Addison;Riley;Harper;Aria;Arianna;Mackenzie;Lila;Evelyn;Adalyn;Grace;Brooklyn;Ellie;Anna;Kaitlyn;Isabelle;Sophie;Scarlett;Natalie;Leah;Sarah;Nora;Mila;Elizabeth;Lillian;Kylie;Audrey;Lucy;Maya;Annabelle;" & _
    "Makayla;Gabriella;Elena;Victoria;Claire;Savannah;Peyton;Maria;Alaina;Kennedy;Stella;Liliana;Allison;Samantha;Keira;Alyssa;Reagan;Molly;Alexandra;Violet;Charlie;Julia;Sadie;Ruby;Eva;Alice;Eliana;Taylor;Callie;Penelope;Camilla;Bailey;Kaelyn;Alexis;Kayla;Katherine;Sydney;Lauren;Jasmine;London;Bella;Adeline;Caroline;Vivian;Juliana;Gianna;Skyler;Jordyn;Jackson;Aiden;Liam;Lucas;Noah;Mason;Jayden;Ethan;Jacob;Jack;Caden;Logan;Benjamin;Michael;Caleb;Ryan;Alexander;Elijah;James;William;Oliver;Connor;Matthew;Daniel;Luke;Brayden;Jayce;Henry;Carter;Dylan;Gabriel;Joshua;Nicholas;Isaac;Owen;Nathan;Grayson;Eli;Landon;Andrew;Max;Samuel;Gavin;Wyatt;Christian;Hunter;Cameron;Evan;" & _
    "Charlie;David;Sebastian;Joseph;Dominic;Anthony;Colton;John;Tyler;Zachary;Thomas;Julian;Levi;Adam;Isaiah;Alex;Aaron;Parker;Cooper;Miles;Chase;Muhammad;Christopher;Blake;Austin;Jordan;Leo;Jonathan;Adrian;Colin;Hudson;Ian;Xavier;Camden;Tristan;Carson;Jason;Nolan;Riley;Lincoln;Brody;Bentley;Nathaniel;Josiah;Declan;Jake;Asher;Jeremiah;Cole;Mateo;Micah;Elliot"

FirstNames = Right(FirstNames, Len(FirstNames) - WorksheetFunction.Find(";", FirstNames, WorksheetFunction.RandBetween(1, 1360)))

LastNames = "Johnson;Williams;Jones;Brown;Davis;Miller;Wilson;Moore;Taylor;Anderson;Thomas;Jackson;White;Harris;Martin;Thompson;Garcia;Martinez;Robinson;Clark;Rodriguez;Lewis;Lee;Walker;Hall;Allen;Young;Hernandez;King;Wright;Lopez;Hill;Scott;Green;Adams;Baker;Gonzalez;Nelson;Carter;Mitchell;Perez;Roberts;Turner;Phillips;Campbell;Parker;Evans;Edwards;Collins;Stewart;Sanchez;Morris;Rogers;Reed;Cook;Morgan;Bell;Murphy;Bailey;Rivera;Cooper;Richardson;Cox;Howard;Ward;Torres;Peterson;Gray;Ramirez;James;Watson;Brooks;Kelly;Sanders;Price;Bennett;Wood;Barnes;Ross;Henderson;Coleman;Jenkins;Perry;Powell;Long;Patterson;Hughes;Flores;Washington;Butler;Simmons;Foster;Gonzales;Bryant;Alexander;Russell;Griffin;Diaz;Hayes"
LastNames = Right(LastNames, Len(LastNames) - WorksheetFunction.Find(";", LastNames, WorksheetFunction.RandBetween(1, 685)))

RandName = Left(FirstNames, WorksheetFunction.Find(";", FirstNames) - 1) & " " & Left(LastNames, WorksheetFunction.Find(";", LastNames) - 1)

End Function


Public Function RandAddress() As String
Dim Location As Integer
Dim StreetName As String
Dim StreetType As String

StreetName = "Second;Third;First;Fourth;Park;Fifth;Main;Sixth;Oak;Seventh;Pine;Maple;Cedar;Eighth;Elm;View;Washington;Ninth;Lake;Hill;"
StreetName = Right(StreetName, Len(StreetName) - WorksheetFunction.Find(";", StreetName, WorksheetFunction.RandBetween(1, 115)))

StreetType = "Plaza;Place;Lane;Drive;Avenue;Street;Road;Boulevard;"
StreetType = Right(StreetType, Len(StreetType) - WorksheetFunction.Find(";", StreetType, WorksheetFunction.RandBetween(1, 41)))

RandAddress = WorksheetFunction.RandBetween(100, 9999) & " " & Left(StreetName, WorksheetFunction.Find(";", StreetName) - 1) & " " & Left(StreetType, WorksheetFunction.Find(";", StreetType) - 1)

End Function

Public Function RandTitle() As String
Dim Titles As String

Titles = "Account Manager;Accountant;Actor;Actuary;Adjustment Clerk;Admin;Administrative Law Judge;Administrative Services Manager;" & _
            "Administrative Support Supervisors;Advertising Manager OR Promotions Manager;Advertising Sales Agent;Aerospace Engineer;" & _
            "Agricultural Crop Farm Manager;Agricultural Crop Worker;Agricultural Engineer;Agricultural Equipment Operator;Agricultural Inspector;" & _
            "Agricultural Manager;Agricultural Product Grader Sorter;Agricultural Sales Representative;Agricultural Science Technician;Agricultural Sciences Teacher;" & _
            "Agricultural Technician;Agricultural Worker;Air Crew Member;Air Crew Officer;Air Traffic Controller;Aircraft Assembler;Aircraft Body Repairer;" & _
            "Aircraft Cargo Handling Supervisor;Aircraft Engine Specialist;Aircraft Launch and Recovery Officer;Aircraft Launch Specialist;" & _
            "Aircraft Mechanics OR Aircraft Service Technician;Aircraft Rigging Assembler;Aircraft Structure Assemblers;Airfield Operations Specialist;" & _
            "Admin Assistant 2;Administrative Asst 1, FM;Bishop Storehouse Manager;Clerk, Sales;Clerk,Clothing-Temple;Concord NH ARP;Coord, Travel 1;" & _
            "Coord,Housing 1;Coordinator;Coordinator, Mission Housing;Counseling Specialist;Counselor 1;Counselor 2;Customer Service Rep II;Cutter;" & _
            "DI Business Partner Associate;DS Business Partner Associate;Engineer, QA 3;Group Manager;IT Project Mgr. 1;Intern;Intern - LDSFS;" & _
            "Job Coach Trainer 1;Job Coach Trainer 2;Mgr 2, Self-Reliance;Mgr, 3;Mgr, Administration;Mgr, Counseling;Mgr,Counseling;Mgr,Family History 2;" & _
            "PEF Loan Review Consolidator;Part Time Clerk - LDSFS;Part-Time Caseworker 2;Part-Time Counselor 3;Piece Rate - Sewer;Project Manager 3;" & _
            "Quality Control Facilitator 3;Regional Manager-Family Svcs;SC Institute Instructor;SC Institute Supervisor;SC Seminary Instructor;" & _
            "Software Developer;Spec,PEF;Tech Product Mgr. 3;Tech,A&E 1;Web Design Eng. 2;Worker, Special Projects"

Titles = Right(Titles, Len(Titles) - WorksheetFunction.Find(";", Titles, WorksheetFunction.RandBetween(1, 1500)))

RandTitle = Left(Titles, WorksheetFunction.Find(";", Titles) - 1)

End Function

Public Function MyVLookup(Criteria As Variant, LookupRange As Range, Optional ResultRange As Variant) As Variant

If IsMissing(ResultRange) Then

MyVLookup = WorksheetFunction.Index(LookupRange, WorksheetFunction.Match(Criteria, LookupRange, 0))

Else
MyVLookup = WorksheetFunction.Index(ResultRange, WorksheetFunction.Match(Criteria, LookupRange, 0))

End If

End Function

Function CellType(cell As Range) As String
'   Returns the cell type of the upper left cell in a range
    Dim UpperLeft As Range
    Application.Volatile True
    Set UpperLeft = cell.Range("A1")
    Select Case True
        Case UpperLeft.NumberFormat = "@"
            CellType = "Text"
        Case IsEmpty(UpperLeft.Value)
            CellType = "Blank"
        Case WorksheetFunction.IsText(UpperLeft)
            CellType = "Text"
        Case WorksheetFunction.IsLogical(UpperLeft.Value)
            CellType = "Logical"
        Case WorksheetFunction.IsErr(UpperLeft.Value)
            CellType = "Error"
        Case IsDate(UpperLeft.Value)
            CellType = "Date"
        Case InStr(1, UpperLeft.Text, ":") <> 0
            CellType = "Time"
        Case IsNumeric(UpperLeft.Value)
            CellType = "Value"
    End Select
End Function

Public Function FederalWithholdings2014(TaxableWages As Double, FilingStatus As String, Allowances As Integer)
Dim BaseWages As Double
Dim AllowanceAmount As Double

AllowanceAmount = 151.9

If FilingStatus <> "Exempt" And FilingStatus <> "Married" And FilingStatus <> "Single" Then

    FederalWithholdings2014 = "Please enter a valid filing status."
    Exit Function
    
End If

BaseWages = TaxableWages - (Allowances * AllowanceAmount)

If FilingStatus = "Exempt" Then FederalWithholdings2014 = 0

If FilingStatus = "Single" Then
    Select Case BaseWages
        Case Is < 87: FederalWithholdings2014 = 0
        Case 87 To 436: FederalWithholdings2014 = (BaseWages - 87) * 0.1
        Case 436 To 1506: FederalWithholdings2014 = ((BaseWages - 436) * 0.15) + 34.9
        Case 1506 To 3523: FederalWithholdings2014 = ((BaseWages - 1506) * 0.25) + 195.4
        Case 3523 To 7254: FederalWithholdings2014 = ((BaseWages - 3523) * 0.28) + 699.65
        Case 7254 To 15667: FederalWithholdings2014 = ((BaseWages - 7254) * 0.33) + 1744.33
        Case 15667 To 15731: FederalWithholdings2014 = ((BaseWages - 15667) * 0.35) + 4520.62
        Case Is > 15731: FederalWithholdings2014 = ((BaseWages - 15731) * 0.396) + 4543.02
    End Select
End If

If FilingStatus = "Married" Then
    Select Case BaseWages
        Case Is < 325: FederalWithholdings2014 = 0
        Case 325 To 1023: FederalWithholdings2014 = (BaseWages - 325) * 0.1
        Case 1023 To 3163: FederalWithholdings2014 = ((BaseWages - 1023) * 0.15) + 69.8
        Case 3163 To 6050: FederalWithholdings2014 = ((BaseWages - 3163) * 0.25) + 390.8
        Case 6050 To 9050: FederalWithholdings2014 = ((BaseWages - 6050) * 0.28) + 1112.55
        Case 9050 To 15906: FederalWithholdings2014 = ((BaseWages - 9050) * 0.33) + 1952.55
        Case 15906 To 17925: FederalWithholdings2014 = ((BaseWages - 15906) * 0.35) + 4215.03
        Case Is > 17925: FederalWithholdings2014 = ((BaseWages - 17925) * 0.396) + 4921.68
    End Select
End If

End Function


Public Function UtahWithholdings2014(TaxableWages As Double, FilingStatus As String, Allowances As Integer)
Dim Line1 As Double
Dim Line2 As Double
Dim Line6 As Integer
Dim Line8 As Double
Dim Line9 As Double


If FilingStatus <> "Exempt" And FilingStatus <> "Married" And FilingStatus <> "Single" Then

    UtahWithholdings2014 = "Please enter a valid filing status."
    Exit Function
    
End If

If FilingStatus = "Exempt" Then

    UtahWithholdings2014 = 0

End If

If FilingStatus = "Single" Then
    
    Line2 = TaxableWages * 0.05
    Line6 = (Allowances * 5) + 10
    
    If TaxableWages < 462 Then
            Line8 = 0
        Else
            Line8 = (TaxableWages - 462) * 0.013
    End If
    If Line6 - Line8 < 0 Then
            Line9 = 0
        Else
            Line9 = Line6 - Line8
    End If
    If Line2 - Line9 < 0 Then
            UtahWithholdings2014 = 0
        Else
            UtahWithholdings2014 = (Line2 - Line9)
    End If
End If

If FilingStatus = "Married" Then
    
    Line2 = TaxableWages * 0.05
    Line6 = (Allowances * 5) + 14
    
    If TaxableWages < 692 Then
            Line8 = 0
        Else
            Line8 = (TaxableWages - 692) * 0.013
    End If
    If Line6 - Line8 < 0 Then
            Line9 = 0
        Else
            Line9 = Line6 - Line8
    End If
    If Line2 - Line9 < 0 Then
            UtahWithholdings2014 = 0
        Else
            UtahWithholdings2014 = (Line2 - Line9)
    End If
End If

End Function

Sub GeneratePPNumber(control As IRibbonControl, PPDate As Date)
    Dim PPNumber As String, PPBeginDate As Date, PPEndDate As Date
    Dim Msg As String, Deduction As String, DedACD As String, DedBCD As String, DedD As String
    Dim PPinMonth As Integer
    Dim PPMonth As String
    
    
DedACD = _
            "- US ER Paid DMBA Admin Fee" & vbNewLine & _
            "- DMBA And Altius Health Insurance" & vbNewLine & _
            "- Supplemental Life Insurance" & vbNewLine & _
            "- 24 - Hour AD&D Insurance" & vbNewLine & _
            "- Flexible Spending" & vbNewLine & vbNewLine & _
            "- Canadian Provincial Health Insurance" & vbNewLine & vbNewLine & _
            "- Thrift Plan Loan Payments Thrift Plan Contributions (U.S.)" & vbNewLine & _
            "- Retirement Savings Plan Contributions (Canada)" & vbNewLine & vbNewLine & _
            "- Metpay And Liberty  Mutual Insurance" & vbNewLine & vbNewLine & _
            "- United Way" & vbNewLine & _
            "- Laundry" & vbNewLine & _
            "- Bus Pass" & vbNewLine & _
            "- Van Pool" & vbNewLine

DedBCD = _
            "- DMBA And Altius Health Insurance" & vbNewLine & _
            "- Supplemental Life Insurance" & vbNewLine & _
            "- 24 - Hour AD&D Insurance" & vbNewLine & _
            "- Flexible Spending" & vbNewLine & vbNewLine & _
            "- Thrift Plan Loan Payments Thrift Plan Contributions (U.S.)" & vbNewLine & _
            "- Retirement Savings Plan Contributions (Canada)" & vbNewLine & vbNewLine & _
            "- Metpay And Liberty  Mutual Insurance" & vbNewLine & vbNewLine & _
            "- Employee Funds" & vbNewLine & _
            "- LDS Philanthropies" & vbNewLine & _
            "- Bus Pass" & vbNewLine & _
            "- Van Pool" & vbNewLine

DedD = _
            "- Thrift Plan Loan Payments Thrift Plan Contributions (U.S.) " & vbNewLine & _
            "- Retirement Savings Plan Contributions (Canada) " & vbNewLine & vbNewLine & _
            "- BUS Pass " & vbNewLine & _
            "- Van Pool " & vbNewLine

    On Error Resume Next
    
    Select Case PPDate
    Case Is < 40901: MsgBox "Please enter a valide date.": Exit Sub
    
        Case Is < 40915: PPNumber = "PP01 2012": PPMonth = "January PP01": PPBeginDate = "12/24/2011": PPEndDate = "01/06/2012"
        Case Is < 40929: PPNumber = "PP02 2012": PPBeginDate = "01/07/2012": PPEndDate = "01/20/2012"
        Case Is < 40943: PPNumber = "PP03 2012": PPBeginDate = "01/21/2012": PPEndDate = "02/03/2012"
        Case Is < 40957: PPNumber = "PP04 2012": PPBeginDate = "02/04/2012": PPEndDate = "02/17/2012"
        Case Is < 40971: PPNumber = "PP05 2012": PPBeginDate = "02/18/2012": PPEndDate = "03/02/2012"
        Case Is < 40985: PPNumber = "PP06 2012": PPBeginDate = "03/03/2012": PPEndDate = "03/16/2012"
        Case Is < 40999: PPNumber = "PP07 2012": PPBeginDate = "03/17/2012": PPEndDate = "03/30/2012"
        Case Is < 41013: PPNumber = "PP08 2012": PPBeginDate = "03/31/2012": PPEndDate = "04/13/2012"
        Case Is < 41027: PPNumber = "PP09 2012": PPBeginDate = "04/14/2012": PPEndDate = "04/27/2012"
        Case Is < 41041: PPNumber = "PP10 2012": PPBeginDate = "04/28/2012": PPEndDate = "05/11/2012"
        Case Is < 41055: PPNumber = "PP11 2012": PPBeginDate = "05/12/2012": PPEndDate = "05/25/2012"
        Case Is < 41069: PPNumber = "PP12 2012": PPBeginDate = "05/26/2012": PPEndDate = "06/08/2012"
        Case Is < 41083: PPNumber = "PP13 2012": PPBeginDate = "06/09/2012": PPEndDate = "06/22/2012"
        Case Is < 41097: PPNumber = "PP14 2012": PPBeginDate = "06/23/2012": PPEndDate = "07/06/2012"
        Case Is < 41111: PPNumber = "PP15 2012": PPBeginDate = "07/07/2012": PPEndDate = "07/20/2012"
        Case Is < 41125: PPNumber = "PP16 2012": PPBeginDate = "07/21/2012": PPEndDate = "08/03/2012"
        Case Is < 41139: PPNumber = "PP17 2012": PPBeginDate = "08/04/2012": PPEndDate = "08/17/2012"
        Case Is < 41153: PPNumber = "PP18 2012": PPBeginDate = "08/18/2012": PPEndDate = "08/31/2012"
        Case Is < 41167: PPNumber = "PP19 2012": PPBeginDate = "09/01/2012": PPEndDate = "09/14/2012"
        Case Is < 41181: PPNumber = "PP20 2012": PPBeginDate = "09/15/2012": PPEndDate = "09/28/2012"
        Case Is < 41195: PPNumber = "PP21 2012": PPBeginDate = "09/29/2012": PPEndDate = "10/12/2012"
        Case Is < 41209: PPNumber = "PP22 2012": PPBeginDate = "10/13/2012": PPEndDate = "10/26/2012"
        Case Is < 41223: PPNumber = "PP23 2012": PPBeginDate = "10/27/2012": PPEndDate = "11/09/2012"
        Case Is < 41237: PPNumber = "PP24 2012": PPBeginDate = "11/10/2012": PPEndDate = "11/23/2012"
        Case Is < 41251: PPNumber = "PP25 2012": PPBeginDate = "11/24/2012": PPEndDate = "12/07/2012"
        Case Is < 41265: PPNumber = "PP26 2012": PPBeginDate = "12/08/2012": PPEndDate = "12/21/2012"
        
        Case Is < 41279: PPNumber = "PP01 2013": PPBeginDate = "12/22/2012": PPEndDate = "01/04/2013"
        Case Is < 41293: PPNumber = "PP02 2013": PPBeginDate = "01/05/2013": PPEndDate = "01/18/2013"
        Case Is < 41307: PPNumber = "PP03 2013": PPBeginDate = "01/19/2013": PPEndDate = "02/01/2013"
        Case Is < 41321: PPNumber = "PP04 2013": PPBeginDate = "02/02/2013": PPEndDate = "02/15/2013"
        Case Is < 41335: PPNumber = "PP05 2013": PPBeginDate = "02/16/2013": PPEndDate = "03/01/2013"
        Case Is < 41349: PPNumber = "PP06 2013": PPBeginDate = "03/02/2013": PPEndDate = "03/15/2013"
        Case Is < 41363: PPNumber = "PP07 2013": PPBeginDate = "03/16/2013": PPEndDate = "03/29/2013"
        Case Is < 41377: PPNumber = "PP08 2013": PPBeginDate = "03/30/2013": PPEndDate = "04/12/2013"
        Case Is < 41391: PPNumber = "PP09 2013": PPBeginDate = "04/13/2013": PPEndDate = "04/26/2013"
        Case Is < 41405: PPNumber = "PP10 2013": PPBeginDate = "04/27/2013": PPEndDate = "05/10/2013"
        Case Is < 41419: PPNumber = "PP11 2013": PPBeginDate = "05/11/2013": PPEndDate = "05/24/2013"
        Case Is < 41433: PPNumber = "PP12 2013": PPBeginDate = "05/25/2013": PPEndDate = "06/07/2013"
        Case Is < 41447: PPNumber = "PP13 2013": PPBeginDate = "06/08/2013": PPEndDate = "06/21/2013"
        Case Is < 41461: PPNumber = "PP14 2013": PPBeginDate = "06/22/2013": PPEndDate = "07/05/2013"
        Case Is < 41475: PPNumber = "PP15 2013": PPBeginDate = "07/06/2013": PPEndDate = "07/19/2013"
        Case Is < 41489: PPNumber = "PP16 2013": PPBeginDate = "07/20/2013": PPEndDate = "08/02/2013"
        Case Is < 41503: PPNumber = "PP17 2013": PPBeginDate = "08/03/2013": PPEndDate = "08/16/2013"
        Case Is < 41517: PPNumber = "PP18 2013": PPBeginDate = "08/17/2013": PPEndDate = "08/30/2013"
        Case Is < 41531: PPNumber = "PP19 2013": PPBeginDate = "08/31/2013": PPEndDate = "09/13/2013"
        Case Is < 41545: PPNumber = "PP20 2013": PPBeginDate = "09/14/2013": PPEndDate = "09/27/2013"
        Case Is < 41559: PPNumber = "PP21 2013": PPBeginDate = "09/28/2013": PPEndDate = "10/11/2013"
        Case Is < 41573: PPNumber = "PP22 2013": PPBeginDate = "10/12/2013": PPEndDate = "10/25/2013"
        Case Is < 41587: PPNumber = "PP23 2013": PPBeginDate = "10/26/2013": PPEndDate = "11/08/2013"
        Case Is < 41601: PPNumber = "PP24 2013": PPBeginDate = "11/09/2013": PPEndDate = "11/22/2013"
        Case Is < 41615: PPNumber = "PP25 2013": PPBeginDate = "11/23/2013": PPEndDate = "12/06/2013"
        Case Is < 41629: PPNumber = "PP26 2013": PPBeginDate = "12/07/2013": PPEndDate = "12/20/2013"

        Case Is < 41643: PPNumber = "PP01 2014": PPMonth = "January PP01": PPBeginDate = "12/21/2013": PPEndDate = "01/03/2014": Deduction = DedACD
        Case Is < 41657: PPNumber = "PP02 2014": PPMonth = "January PP02": PPBeginDate = "01/04/2014": PPEndDate = "01/17/2014": Deduction = DedBCD
        Case Is < 41671: PPNumber = "PP03 2014": PPMonth = "February PP01": PPBeginDate = "01/18/2014": PPEndDate = "01/31/2014": Deduction = DedACD
        Case Is < 41685: PPNumber = "PP04 2014": PPMonth = "February PP02": PPBeginDate = "02/01/2014": PPEndDate = "02/14/2014": Deduction = DedBCD
        Case Is < 41699: PPNumber = "PP05 2014": PPMonth = "March PP01": PPBeginDate = "02/15/2014": PPEndDate = "02/28/2014": Deduction = DedACD
        Case Is < 41713: PPNumber = "PP06 2014": PPMonth = "March PP02": PPBeginDate = "03/01/2014": PPEndDate = "03/14/2014": Deduction = DedBCD
        Case Is < 41727: PPNumber = "PP07 2014": PPMonth = "April PP01": PPBeginDate = "03/15/2014": PPEndDate = "03/28/2014": Deduction = DedACD
        Case Is < 41741: PPNumber = "PP08 2014": PPMonth = "April PP02": PPBeginDate = "03/29/2014": PPEndDate = "04/11/2014": Deduction = DedBCD
        Case Is < 41755: PPNumber = "PP09 2014": PPMonth = "May PP01": PPBeginDate = "04/12/2014": PPEndDate = "04/25/2014": Deduction = DedACD
        Case Is < 41769: PPNumber = "PP10 2014": PPMonth = "May PP02": PPBeginDate = "04/26/2014": PPEndDate = "05/09/2014": Deduction = DedBCD
        Case Is < 41783: PPNumber = "PP11 2014": PPMonth = "May PP03": PPBeginDate = "05/10/2014": PPEndDate = "05/23/2014": Deduction = DedD
        Case Is < 41797: PPNumber = "PP12 2014": PPMonth = "June PP01": PPBeginDate = "05/24/2014": PPEndDate = "06/06/2014": Deduction = DedACD
        Case Is < 41811: PPNumber = "PP13 2014": PPMonth = "June PP02": PPBeginDate = "06/07/2014": PPEndDate = "06/20/2014": Deduction = DedBCD
        Case Is < 41825: PPNumber = "PP14 2014": PPMonth = "July PP01": PPBeginDate = "06/21/2014": PPEndDate = "07/04/2014": Deduction = DedACD
        Case Is < 41839: PPNumber = "PP15 2014": PPMonth = "July PP02": PPBeginDate = "07/05/2014": PPEndDate = "07/18/2014": Deduction = DedBCD
        Case Is < 41853: PPNumber = "PP16 2014": PPMonth = "August PP01": PPBeginDate = "07/19/2014": PPEndDate = "08/01/2014": Deduction = DedACD
        Case Is < 41867: PPNumber = "PP17 2014": PPMonth = "August PP02": PPBeginDate = "08/02/2014": PPEndDate = "08/15/2014": Deduction = DedBCD
        Case Is < 41881: PPNumber = "PP18 2014": PPMonth = "September PP01": PPBeginDate = "08/16/2014": PPEndDate = "08/29/2014": Deduction = DedACD
        Case Is < 41895: PPNumber = "PP19 2014": PPMonth = "September PP02": PPBeginDate = "08/30/2014": PPEndDate = "09/12/2014": Deduction = DedBCD
        Case Is < 41909: PPNumber = "PP20 2014": PPMonth = "October PP01": PPBeginDate = "09/13/2014": PPEndDate = "09/26/2014": Deduction = DedACD
        Case Is < 41923: PPNumber = "PP21 2014": PPMonth = "October PP02": PPBeginDate = "09/27/2014": PPEndDate = "10/10/2014": Deduction = DedBCD
        Case Is < 41937: PPNumber = "PP22 2014": PPMonth = "October PP03": PPBeginDate = "10/11/2014": PPEndDate = "10/24/2014": Deduction = DedD
        Case Is < 41951: PPNumber = "PP23 2014": PPMonth = "November PP01": PPBeginDate = "10/25/2014": PPEndDate = "11/07/2014": Deduction = DedACD
        Case Is < 41965: PPNumber = "PP24 2014": PPMonth = "November PP02": PPBeginDate = "11/08/2014": PPEndDate = "11/21/2014": Deduction = DedBCD
        Case Is < 41979: PPNumber = "PP25 2014": PPMonth = "December PP01": PPBeginDate = "11/22/2014": PPEndDate = "12/05/2014": Deduction = DedACD
        Case Is < 41993: PPNumber = "PP26 2014": PPMonth = "December PP02": PPBeginDate = "12/06/2014": PPEndDate = "12/19/2014": Deduction = DedBCD

        Case Is < 42007: PPNumber = "PP01 2015": PPBeginDate = "12/20/2014": PPEndDate = "01/02/2015"
        Case Is < 42021: PPNumber = "PP02 2015": PPBeginDate = "01/03/2015": PPEndDate = "01/16/2015"
        Case Is < 42035: PPNumber = "PP03 2015": PPBeginDate = "01/17/2015": PPEndDate = "01/30/2015"
        Case Is < 42049: PPNumber = "PP04 2015": PPBeginDate = "01/31/2015": PPEndDate = "02/13/2015"
        Case Is < 42063: PPNumber = "PP05 2015": PPBeginDate = "02/14/2015": PPEndDate = "02/27/2015"
        Case Is < 42077: PPNumber = "PP06 2015": PPBeginDate = "02/28/2015": PPEndDate = "03/13/2015"
        Case Is < 42091: PPNumber = "PP07 2015": PPBeginDate = "03/14/2015": PPEndDate = "03/27/2015"
        Case Is < 42105: PPNumber = "PP08 2015": PPBeginDate = "03/28/2015": PPEndDate = "04/10/2015"
        Case Is < 42119: PPNumber = "PP09 2015": PPBeginDate = "04/11/2015": PPEndDate = "04/24/2015"
        Case Is < 42133: PPNumber = "PP10 2015": PPBeginDate = "04/25/2015": PPEndDate = "05/08/2015"
        Case Is < 42147: PPNumber = "PP11 2015": PPBeginDate = "05/09/2015": PPEndDate = "05/22/2015"
        Case Is < 42161: PPNumber = "PP12 2015": PPBeginDate = "05/23/2015": PPEndDate = "06/05/2015"
        Case Is < 42175: PPNumber = "PP13 2015": PPBeginDate = "06/06/2015": PPEndDate = "06/19/2015"
        Case Is < 42189: PPNumber = "PP14 2015": PPBeginDate = "06/20/2015": PPEndDate = "07/03/2015"
        Case Is < 42203: PPNumber = "PP15 2015": PPBeginDate = "07/04/2015": PPEndDate = "07/17/2015"
        Case Is < 42217: PPNumber = "PP16 2015": PPBeginDate = "07/18/2015": PPEndDate = "07/31/2015"
        Case Is < 42231: PPNumber = "PP17 2015": PPBeginDate = "08/01/2015": PPEndDate = "08/14/2015"
        Case Is < 42245: PPNumber = "PP18 2015": PPBeginDate = "08/15/2015": PPEndDate = "08/28/2015"
        Case Is < 42259: PPNumber = "PP19 2015": PPBeginDate = "08/29/2015": PPEndDate = "09/11/2015"
        Case Is < 42273: PPNumber = "PP20 2015": PPBeginDate = "09/12/2015": PPEndDate = "09/25/2015"
        Case Is < 42287: PPNumber = "PP21 2015": PPBeginDate = "09/26/2015": PPEndDate = "10/09/2015"
        Case Is < 42301: PPNumber = "PP22 2015": PPBeginDate = "10/10/2015": PPEndDate = "10/23/2015"
        Case Is < 42315: PPNumber = "PP23 2015": PPBeginDate = "10/24/2015": PPEndDate = "11/06/2015"
        Case Is < 42329: PPNumber = "PP24 2015": PPBeginDate = "11/07/2015": PPEndDate = "11/20/2015"
        Case Is < 42343: PPNumber = "PP25 2015": PPBeginDate = "11/21/2015": PPEndDate = "12/04/2015"
        Case Is < 42357: PPNumber = "PP26 2015": PPBeginDate = "12/05/2015": PPEndDate = "12/18/2015"
        
    Case Else: MsgBox "Please enter a valid date.": Exit Sub
    
    End Select

    If Err.Number = 0 And PPNumber <> "" Then
    
    If Deduction = "" Then
    Msg = PPNumber & vbNewLine & vbNewLine & _
        "PP of the Month: " & PPMonth & vbNewLine & _
        "Begin Date: " & PPBeginDate & vbNewLine & _
        "End Date: " & PPEndDate
    Else
        Msg = PPNumber & vbNewLine & _
        "PP of the Month: " & PPMonth & vbNewLine & vbNewLine & _
        "Begin Date: " & PPBeginDate & vbNewLine & _
        "End Date: " & PPEndDate & vbNewLine & vbNewLine & _
        Deduction
    End If
        
    MsgBox Msg, vbInformation
        
    Else
        If PPNumber <> "Empty" Then MsgBox "Please enter a valid date.", vbCritical
    End If
    
Call UsageLog("PP Info Generator")
    
End Sub


Function LastNameMatch(Name1 As String, Name2 As String) As Boolean

'Removing any trailing spaces from the two names
Do Until Right(Name1, 1) <> " "
    Name1 = Left(Name1, Len(Name1) - 1)
Loop

Do Until Right(Name2, 1) <> " "
    Name1 = Left(Name2, Len(Name2) - 1)
Loop

'Determine if the last name in each string are the same, regardless of case
If UCase(Right(Name1, Len(Name1) - InStrRev(Name1, " "))) = UCase(Right(Name2, Len(Name2) - InStrRev(Name2, " "))) Then
    LastNameMatch = True
    Else
    LastNameMatch = False
End If

End Function

Function CloseMatch(Value1 As Double, Value2 As Double, Difference As Double) As Boolean

If Abs(Round(Value1 - Value2)) <= Round(Difference, 2) Then
CloseMatch = True
Else
CloseMatch = False
End If

End Function

Function CountCellsByColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In Intersect(rData, ActiveSheet.UsedRange)
        If indRefColor = cellCurrent.Interior.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByColor = cntRes
End Function

Public Function RecentPP() As Integer
'This function returns the most recently completed pay period number as an integer.

Dim PP01 As Date
Dim lastPPEndDate As Date

'Determine the most recent PP end date based on a static date in the past
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Determine the PPEnd date for PP01 of the current year
PP01 = CDate("1/3/2014")
Do Until Year(PP01 + 7) = Year(Date)
    PP01 = PP01 + 14
Loop

'If the first PP end date this year is after the current date then we assume the most recent PP
'is pay period 26.
If PP01 >= Date Then
    RecentPP = 26
    Exit Function
End If

'Determine which pay period just ended.
RecentPP = 1
Do Until lastPPEndDate = PP01
    PP01 = PP01 + 14
    RecentPP = RecentPP + 1
Loop

End Function
 
Public Sub ExportSourceFiles()

    Dim destPath As String
    Dim component As VBComponent
    
    destPath = "C:\Users\danseethaler\Documents\GitHub\VBA-Git-Repo\Excel VBA\"
    
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
     
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
     
End Function
