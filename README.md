# activexdllvb6office64bit
Pomoću ove ActiveX Exe biblioteke se može izvršiti "ubacivanje" i korišćenje bilo kog ActiveX Dll napisanog u VB6 u pre svega MS Office 64-bit verzzije VBA a da se ne menja već urađena ActiveX dll biblioteka.

Potrebno je ovaj ActiveX Exe dodati kao referencu u VBA projekat kao i svaki drugi ActiveX Dll kako bi se isti koristio pored ActiveX dll koji se koristi u VBA projektu.

U nastavku je primer za ActiveX Dll MUPLKLib.

Potrebno je registrovati MUPLKLib.dll sa regsvr32 odnosno kako je opisano u procitajme.txt.

CelikApi.dll je potrebno da bude vidljiv od strane operativnog sistema kako bi biblioteka MUPLKLib.dll mogla da pristupi CelikApi.

NAPOMENE:
Nije testirano sa poslednjom verzijom CelikApi.dll 1.3.3 koja verovatno neće biti podržana zbog izvršenih poslednjih promena od strane MUP RS.
Iz tog razloga ali i kada je napravljen ActiveX Dll MUPLKLib su date obe verzije CelikApi.dll sa kojima je razvijan tada ActiveX dll.

Private Const ProgID = "MUPLKLib.cMupLK"
Private host As New VB6Office64.loader

Private Function GetInstance() As MUPLKLib.cMupLK

    On Error GoTo ErrHandler

    Set GetInstance = host.CreateInstance(ProgID)
    

Exit Function
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
End Function

' Primer 1
Public Sub ReadLKVersion1()

    On Error GoTo ErrHandler

    Dim x As MUPLKLib.cMupLK
    Dim r As Integer
    Dim v As String
    Dim DoUTF8Decode As Boolean
    
    v = ""
    
    ' Create new instance of object
    Set x = GetInstance()
    
    ' Try to read data from smart card device
    r = x.ReadData()
    
    ' If reading success then
    If r = 0 Then
    
        ' Document data
        With x
        
            ' Document data
            Cells(2, 2) = .UTF8Decode(.DocumentData.docRegNo)
            Cells(3, 2) = .UTF8Decode(.DocumentData.documentType)
            Cells(4, 2) = .UTF8Decode(.DocumentData.expiryDate)
            Cells(5, 2) = .UTF8Decode(.DocumentData.issuingDate)
            Cells(6, 2) = .UTF8Decode(.DocumentData.issuingAuthority)
            
            ' Fixed personal data
            Cells(8, 2) = .UTF8Decode(.FixedPersonalData.personalNumber)
            Cells(9, 2) = .UTF8Decode(.FixedPersonalData.surname)
            Cells(10, 2) = .UTF8Decode(.FixedPersonalData.GivenName)
            Cells(11, 2) = .UTF8Decode(.FixedPersonalData.parentGivenName)
            Cells(12, 2) = .UTF8Decode(.FixedPersonalData.sex)
            Cells(13, 2) = .UTF8Decode(.FixedPersonalData.placeOfBirth)
            Cells(14, 2) = .UTF8Decode(.FixedPersonalData.stateOfBirth)
            Cells(15, 2) = .UTF8Decode(.FixedPersonalData.dateOfBirth)
            Cells(16, 2) = .UTF8Decode(.FixedPersonalData.communityOfBirth)
            
            ' Variable personal data
            Cells(18, 2) = .UTF8Decode(.VariablePersonalData.State)
            Cells(19, 2) = .UTF8Decode(.VariablePersonalData.community)
            Cells(20, 2) = .UTF8Decode(.VariablePersonalData.place)
            Cells(21, 2) = .UTF8Decode(.VariablePersonalData.Street)
            Cells(22, 2) = .UTF8Decode(.VariablePersonalData.houseNumber)
            Cells(23, 2) = .UTF8Decode(.VariablePersonalData.houseLetter)
            Cells(24, 2) = .UTF8Decode(.VariablePersonalData.entrance)
            Cells(25, 2) = .UTF8Decode(.VariablePersonalData.Floor)
            Cells(26, 2) = .UTF8Decode(.VariablePersonalData.apartmentNumber)
            Cells(27, 2) = .UTF8Decode(.VariablePersonalData.addressDate)
            Cells(28, 2) = .UTF8Decode(.VariablePersonalData.addressLabel)
    
        ' Picture
        ' -> x.Picture.PersonPictureB ' Ovo je binarni niz koji iz koga je konvertovan sadrzaj i napravljena slika IPicture tipa
        ' -> x.Picture.PersonPicture  ' Ovo je IPicture tip objekta | treba videti kako se to moze iskoristiti u MS Access-u
       
       End With
       
        MsgBox "Success."
        
    Else
        
        MsgBox "Error reading data from card."
        
    End If
    
    
    
    Set x = Nothing
    
Exit Sub
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
End Sub


' Primer 2
Public Sub ReadLKVersion2()

    ' ----
    ' Vezija 2 - sa/bez UTF8 podrške
    ' ----

    ' VBA -> Menu 'Tools' -> References -> MUP LK Library (Browse... to file to select it if not listed)

    On Error GoTo ErrHandler

    Dim x As MUPLKLib.cMupLK
    Dim r As Integer
    Dim v As String
    Dim DoUTF8Decode As Boolean
    
    v = ""
    
    ' Set value for doing auto. UTF8 decoding
    DoUTF8Decode = True ' False
    
    ' Create new instance of object
    Set x = GetInstance()
    
    ' Try to read data from smart card device
    r = x.ReadData(DoUTF8Decode)
    
    ' If reading success then
    If r = 0 Then
    
        With x
    
            ' Document data
            Cells(2, 2) = .DocumentData.docRegNo
            Cells(3, 2) = .DocumentData.documentType
            Cells(4, 2) = .DocumentData.expiryDate
            Cells(5, 2) = .DocumentData.issuingDate
            Cells(6, 2) = .DocumentData.issuingAuthority
            
            ' Fixed personal data
            Cells(8, 2) = .FixedPersonalData.personalNumber
            Cells(9, 2) = .FixedPersonalData.surname
            Cells(10, 2) = .FixedPersonalData.GivenName
            Cells(11, 2) = .FixedPersonalData.parentGivenName
            Cells(12, 2) = .FixedPersonalData.sex
            Cells(13, 2) = .FixedPersonalData.placeOfBirth
            Cells(14, 2) = .FixedPersonalData.stateOfBirth
            Cells(15, 2) = .FixedPersonalData.dateOfBirth
            Cells(16, 2) = .FixedPersonalData.communityOfBirth
            
            ' Variable personal data
            Cells(18, 2) = .VariablePersonalData.State
            Cells(19, 2) = .VariablePersonalData.community
            Cells(20, 2) = .VariablePersonalData.place
            Cells(21, 2) = .VariablePersonalData.Street
            Cells(22, 2) = .VariablePersonalData.houseNumber
            Cells(23, 2) = .VariablePersonalData.houseLetter
            Cells(24, 2) = .VariablePersonalData.entrance
            Cells(25, 2) = .VariablePersonalData.Floor
            Cells(26, 2) = .VariablePersonalData.apartmentNumber
            Cells(27, 2) = .VariablePersonalData.addressDate
            Cells(28, 2) = .VariablePersonalData.addressLabel
        
            ' Picture
            ' -> x.Picture.PersonPictureB ' Ovo je binarni niz koji iz koga je konvertovan sadrzaj i napravljena slika IPicture tipa
            ' -> x.Picture.PersonPicture  ' Ovo je IPicture tip objekta | treba videti kako se to moze iskoristiti u MS Access-u
            
       End With
       
        MsgBox "Success."
        
    Else
        
        MsgBox "Error reading data from card."
        
    End If
    
    ' Free memory resource
    Set x = Nothing
    
    
Exit Sub
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
End Sub
