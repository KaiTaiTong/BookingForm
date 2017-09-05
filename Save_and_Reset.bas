Attribute VB_Name = "Save_and_Reset"
Sub SaveButton()

    Dim ActSheet As Worksheet
    Dim ActBook As Workbook
    Dim CurrentFile As String
    Dim NewFileType As String
    Dim NewFile As String

'--------------------------------------------------------------------------
'Checking the validation of the user input form

'Class Modules Queue for later use
'Please Do Not Delete or change any declaration below
Dim Q As New Queue
Dim trigger As Integer
trigger = 0

'Initializing the spreadsheet to transparent
'Please Only Change this Set of code when The structure of the excel form has been changed
'In order to change this Set of code, you need to use the Record Macro Function from Excel
    Range("A5:Q49").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range("A13").Interior.ColorIndex = 15
Range("A14").Interior.ColorIndex = 15
Range("F13").Interior.ColorIndex = 15
Range("F14").Interior.ColorIndex = 15
Range("K14:Q16").Interior.ColorIndex = 15
Range("A9:Q10").Interior.ColorIndex = 15

PtName = Range("A6").Value & "_" & Range("F6").Value
Service = Range("I24").Value

'To imporve, we are making an array and store all the address of cells      In the 2D Array, First Column is used to store name of address of the Cell
'                                                                                            Second Column is used to store the value in the corresponding Cell
'                                                                           In the 1D Array, Name of the field of that cell is stored
'that represents the mandatory fields

'Before you Make any Maintainance, Please Make Sure you Change the Size of the Array First
Dim SizeOfStringArray As Integer
SizeOfStringArray = 8       'SizeOfStringArray = Total Number of Boxes of Cells (for which the input is considered to be a String Type) - 1
                            'The - 1 is essential because array index starts at 0

Dim SizeOfNumberArray As Integer
SizeOfNumberArray = 5       'SizeOfNumberArray represents the Array for the boxes of cells that has an input of Number Type
                            'SizeOfNumberArray =  Total Number of Boxes of Cells (for which as above described) - 1
                            'Again due to the mother nature of Array, that -1 is essential

Dim SizeOfTFButtonsWithNoRequiredText As Integer
SizeOfTFButtonsWithNoRequiredText = 1       'SizeOfTFButtonsWithNoRequiredText = Total Number of Boxes of Cells (for which input type is a Selection, thus Boolean) - 1
'Initializing the Arrays according to the types of input values
'Dim CellForString As Variant

ReDim CellForString(0 To SizeOfStringArray, 0 To 1)
ReDim StringforString(0 To SizeOfStringArray) As String
ReDim CellForNumber(0 To SizeOfNumberArray, 0 To 1)
ReDim StringforNumber(0 To SizeOfNumberArray) As String
ReDim CellforTFButton(0 To SizeOfTFButtonsWithNoRequiredText, 0 To 2) 'Ths is for the buttons that do NOT require text to be filled
'                                                                      The number of size 2 of the columns are considered due to the reason that longest option box has a total of 3 options
'                                                                      For Maintainance Purposes, you will soon see that the key idea here that is being used is to Take the SUM as hashcode
'                                                                      We check if the user has input the correct option(here, I really mean if the user has selected one) by take the SUM of the
'                                                                      Boolean Values of each option
'                                                                      Notice again that option button returns -4146 if turned out to be NOT SELECTED
'                                                                                                       Other Number if SELECTED (We Don't Care)
ReDim StringforTFbutton(0 To 3) As String

'For maintainance purpose, we will not loop, but manually input the address
'--------------------------------------------------------------------------

'Patient Surname
CellForString(0, 0) = "A6"
CellForString(0, 1) = Range("A6").Value
StringforString(0) = "Patient Surname"
'First Name
CellForString(1, 0) = "G6"
CellForString(1, 1) = Range("G6").Value
StringforString(1) = "First Name"
'Date of Birth
CellForString(2, 0) = "J8"
CellForString(2, 1) = Range("J8").Value
StringforString(2) = "Date of Birth"
'Surgery Decision Date
CellForString(3, 0) = "F15"
CellForString(3, 1) = Range("F15").Value
StringforString(3) = "Surgery Decision Date"
'Surgeons
CellForString(4, 0) = "E24"
CellForString(4, 1) = Range("E24").Value
StringforString(4) = "Surgeons"
'Services
CellForString(5, 0) = "I24"
CellForString(5, 1) = Range("I24").Value
StringforString(5) = "Services"
'Procedures
CellForString(6, 0) = "K24"
CellForString(6, 1) = Range("K24").Value
StringforString(6) = "Procedures"
'Diagnosis
CellForString(7, 0) = "K30"
CellForString(7, 1) = Range("K30").Value
StringforString(7) = "Diagnosis"

'Legal Guardian
CellForString(8, 0) = "A12"
CellForString(8, 1) = Range("A12").Value
StringforString(8) = "Legal Guardian"

'If you are the maintainer and are trying to add in new mandatory fields, follow the following format please
'Before doing so, please make sure that you have already changed the size of the Array, because this will affect the looping in the later code

'Name of the Mandatory Field you are currently working on
'CellForString(8, 0) = "XXX"                 The Address of the Cell
'CellForString(8, 1) = Range("XXX").Value    The Value that the user has input in the Cell
'StringforString(8) = "XXXXXXXX"            Name of the Mandatory Field you are currently working on

'Same Format could be used for the later Arrays

'--------------------------------------------------------------------------
'PHN, Personal Health Number
CellForNumber(0, 0) = "A8"
CellForNumber(0, 1) = Range("A8").Value
StringforNumber(0) = "Personal Health Number"
'MRN, Medical Record Number
CellForNumber(1, 0) = "F8"
CellForNumber(1, 1) = Range("F8").Value
StringforNumber(1) = "Medical Record Number"
'Contact Number of Legal Guardian
CellForNumber(2, 0) = "K12"
CellForNumber(2, 1) = Range("K12").Value
StringforNumber(2) = "Contact Number of Legal Guardian"
'Procedure Code(s)
CellForNumber(3, 0) = "A24"
CellForNumber(3, 1) = Range("A24").Value
StringforNumber(3) = "Procedure Codes"
'Diagnosis Code or PCATS Code
CellForNumber(4, 0) = "F30"
CellForNumber(4, 1) = Range("F30").Value
StringforNumber(4) = "PCATS Code"
'SKIN TO SKIN(Minutes)
CellForNumber(5, 0) = "D30"
CellForNumber(5, 1) = Range("D30").Value
StringforNumber(5) = "SKIN TO SKIN (Minutes)"


'--------------------------------------------------------------------------
'Gender
CellforTFButton(0, 0) = ActiveSheet.OptionButtons("Option Button 3124").Value
CellforTFButton(0, 1) = ActiveSheet.OptionButtons("Option Button 3127").Value
CellforTFButton(0, 2) = -4146                   'We manually set this value to be -4146 to match the length of Array of Field Cancer
StringforTFbutton(0) = "Gender"                 'This way we are allowed to check the validation of the option boxes By the SUM of these 3 values
'Cancer
CellforTFButton(1, 0) = ActiveSheet.OptionButtons("Option Button 3443").Value
CellforTFButton(1, 1) = ActiveSheet.OptionButtons("Option Button 3444").Value
CellforTFButton(1, 2) = ActiveSheet.OptionButtons("Option Button 3445").Value
StringforTFbutton(1) = "Cancer Suspection"


'Please Do Not Change the following code, unless you are re-structuring the program
'---------------------------------------------------------------------------------
For Index = 0 To SizeOfStringArray
    
    'String Array Validation
    '----------------
    If CellForString(Index, 1) = "" Then
    'Yellow Color Box
    Range(CellForString(Index, 0)).Interior.ColorIndex = 6
    
    'Pushing the invalid input cell into a Queue for later MsgBox use
    Q.Enqueue StringforString(Index)
    
    End If
    
    'Now Testing the validation of Number Array
    '------------------
     If Index < SizeOfNumberArray + 1 Then
     '**************
     'Here we Could implement format of the number to indicate invalid input - This function is not completed or shown, Just an idea...
     '**************
        If CellForNumber(Index, 1) = "" Then
            'Yellow Color Box
            Range(CellForNumber(Index, 0)).Interior.ColorIndex = 6
            
            'Pushing the invalid input cell into a Queue for later MsgBox use
            Q.Enqueue StringforNumber(Index)
        End If
    End If
    
    'Button Selection Validation
    '----------------
    If Index < SizeOfTFButtonsWithNoRequiredText + 1 Then
        Dim TestNumber As Integer
        TestNumber = 0
    
        For counter = 0 To 2
            TestNumber = TestNumber + CellforTFButton(Index, counter)
        Next counter
        
        
        'To see if the sum of the three boxes match -4146 * 3
        If TestNumber = -12438 Then
        Q.Enqueue StringforTFbutton(Index)
        End If
            
    End If
Next Index

'--------------------------------------------------------------------------
'Testing for Admission Status
'This part of the Program possibly needs to be adjusted when adding in new mandatory fields
'Since there is very few obvious common points between these option buttons, we are just gonna do it using If Statements instead of Looping through

'Check to see if at least one box is checked
If ActiveSheet.CheckBoxes("Check Box 313") = -4146 And ActiveSheet.CheckBoxes("Check Box 314") = -4146 And ActiveSheet.CheckBoxes("Check Box 315") = -4146 Then
    'Again as previously mentioned, if the box is NOT SELECTED, then the Boolean Value will return -4146
    'Thus this part we are checking to see if there is ANY box being selected
    
    'if True, We push it into the Queue
    Q.Enqueue "Admission Status"


Else

    'The following part of the code you see deals with the situation that The box is checked, BUT left the String Cell next to it blank
    'This does not make sense and have to be checked because if you have tick the box, that means you must fill in the String box next to it. Otherwise why would the user click the box?
    'Check that if the box is checked, the cell has input (Admit Day of Procedure)
    If ActiveSheet.CheckBoxes("Check Box 314").Value = 1 And Range("G19") = "" Then
        Q.Enqueue "Admission Status: Please specify ELOS days"
        Range("G19").Interior.ColorIndex = 6
    End If

    'Check that if the is checked, the cell has input (Inpatient)
    If ActiveSheet.CheckBoxes("Check Box 315").Value = 1 Then
        If Range("G20") = "" Then
            Q.Enqueue "Admission Status: Please specify location of inpatient"
            Range("G20").Interior.ColorIndex = 6
        End If
        
        If Range("C21") = "" Then
            Q.Enqueue "Admission Status: Please specify ELOS days"
            Range("C21").Interior.ColorIndex = 6
        End If
        
        If Range("F21") = "" Then
            Q.Enqueue "Admission Status: Please specify days Prior OR date"
            Range("F21").Interior.ColorIndex = 6
        End If
    End If

End If

'--------------------------------------------------------------------------
'Testing for Other Text Required Button

'Check if the checkbox for Other Under Special Post OP is checked and then remind user to specify what other post op bed requirement
If ActiveSheet.CheckBoxes("Check Box 318").Value = 1 And Range("M21") = "" Then
    Q.Enqueue "Specify Other Special Post Op Bed Requirements"
    Range("M21").Interior.ColorIndex = 6
End If


'INTERPRETER Required Language
If ActiveSheet.CheckBoxes("Check Box 2870").Value = 1 And Range("O35") = "" Then
    Q.Enqueue "Specify the Language to be Interpreted"
    Range("O35").Interior.ColorIndex = 6
End If


'--------------------------------------------------------------------------
'Building MsgBox and Ready to Send Message to the user

Dim MissingDataException As String
MissingDataException = "You are missing the following information: " & vbNewLine & vbNewLine

Dim ListOfMissingData As String
ListOfMissingData = "   - "

Dim FinalString As String

If Q.Count <> 0 Then
    While Q.Count <> 0
        FinalString = FinalString & ListOfMissingData & Q.Dequeue & vbNewLine
        trigger = 1
    Wend
    MsgBox (MissingDataException + FinalString & vbNewLine & vbNewLine)
    MsgBox ("Save Request Declined")

End If

'--------------------------------------------------------------------------
If trigger = 0 Then

        Dim Surname_FileName As String
        Dim FirstName_FileName As String
        Dim Services_FileName As String
            
            'Checking Surname For FileName
            If Trim(Range("A6").Value) = "" Then
                Surname_FileName = "SurnameIncomplete"
            Else
                Surname_FileName = Trim(Range("A6").Value)
            End If
            'Checking FirstName For FileName
            If Trim(Range("G6").Value) = "" Then
                FirstName_FileName = "FirstNameIncomplete"
            Else
                FirstName_FileName = Trim(Range("G6").Value)
            End If
            'Checking Services_FileName For FileName
            If Trim(Range("I24").Value) = "" Then
                Services_FileName = "ServiceIncomplete"
            Else
                Services_FileName = Trim(Range("I24").Value)
            End If
            
            NewFileName = Surname_FileName + "_" + FirstName_FileName + "_" + Services_FileName
         
            Application.ScreenUpdating = False    ' Prevents screen refreshing.
        
            CurrentFile = ThisWorkbook.FullName
         
            NewFileType = "Excel Files 1997-2003 (*.xls), *.xls," & _
                       "Excel Files 2007 (*.xlsx), *.xlsx," & _
                       "All files (*.*), *.*"
         
            NewFile = Application.GetSaveAsFilename( _
                InitialFileName:=NewFileName, _
                fileFilter:=NewFileType)
         
            If NewFile <> "" And NewFile <> "False" Then
                ActiveWorkbook.SaveAs Filename:=NewFile, _
                    FileFormat:=xlNormal, _
                    Password:="", _
                    WriteResPassword:="", _
                    ReadOnlyRecommended:=False, _
                    CreateBackup:=False
         
                Set ActBook = ActiveWorkbook
                'Workbooks.Open CurrentFile
                'Sheets("BookingForm").Unprotect Password:="secret"
                'ActBook.Unprotect Password:="secret"
                'Sheets("BookingForm").Protect Password:="secret"
                'ActBook.Protect Password:="secret"
        
            End If
         
            Application.ScreenUpdating = True
        
    End If

End Sub

Sub Reset_Click()
Dim msg As String
Dim title As String
Dim response As String

msg = "Information will be deleted in this worksheet. Do you wish to continue? "
title = "Clear Form"
response = MsgBox(msg, vbYesNo, title)

If response = vbYes Then
   Range("A6:E6") = ""
   Range("F6:J6") = ""
   Range("K6:N6") = ""
   Range("O6:Q6") = ""
   Range("A8:D8") = ""
   Range("E8:I8") = ""
   Range("J8:M8") = ""
   Range("A10:J10") = ""
   Range("K10:L10") = ""
   Range("M10") = "BC"
   Range("N10:O10") = ""
   Range("P10:Q10") = "Canada"
   Range("A12:J12") = ""
   Range("K12:M12") = ""
   Range("N12:Q12") = ""
   Range("F13, F14,F15,F16") = ""
   Range("N14,N15,N16") = ""
   Range("G19,G20,C21,M21,F21") = ""
   Range("A24,E24,I24,A25,E25,I25,A26,E26,I26,A27,E27,I27,A28,E28,I28,A29,E29,I29,K24,D30,F30,K30,N32,O35,A44,F47,A48,A38") = ""
    

    Range("H19").Font.Bold = False


Dim sh As Worksheet
For Each sh In Sheets
On Error Resume Next
sh.CheckBoxes.Value = False
ActiveSheet.OptionButtons("Option Button 336").Value = -4146
ActiveSheet.OptionButtons("Option Button 337").Value = -4146
ActiveSheet.OptionButtons("Option Button 3443").Value = -4146
ActiveSheet.OptionButtons("Option Button 3444").Value = -4146
ActiveSheet.OptionButtons("Option Button 3445").Value = -4146
ActiveSheet.OptionButtons("Option Button 3124").Value = -4146
ActiveSheet.OptionButtons("Option Button 3127").Value = -4146
ActiveSheet.CheckBox(" 3127").Value = -4146
On Error GoTo 0
Next sh

Range("A5:Q49").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range("A13").Interior.ColorIndex = 15
Range("A14").Interior.ColorIndex = 15
Range("F13").Interior.ColorIndex = 15
Range("F14").Interior.ColorIndex = 15
Range("K14:Q16").Interior.ColorIndex = 15
Range("A9:Q10").Interior.ColorIndex = 15



'Now cleaning the Bolded Text
Range("B18:E18").Font.Bold = False
Range("B19:E19").Font.Bold = False
Range("B20").Font.Bold = False
Range("B21").Font.Bold = False
Range("D21").Font.Bold = False
Range("H21").Font.Bold = False
Range("B21:E21").Font.Bold = False
Range("K19,K20,K21").Font.Bold = False
Range("C34,C35,C36").Font.Bold = False
Range("F34,F35,F36").Font.Bold = False
Range("K34,K35,K36").Font.Bold = False
Range("B42,G42,J42,O42").Font.Bold = False

End If
End Sub
