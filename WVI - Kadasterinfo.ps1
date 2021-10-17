<#  ************************************
    Author: Jan Speecke
    Last Updated : 15/10/2021 - 20:19:32
    Version	: 0.71
    Comments :
    Changes:
        01/10/2021 - Changed naming CSV and JSON files
        05/10/2021 - Added CAPAKEY calculation
        07/10/2021 - Changed value "Oppervlakte belast"
        15/10/2021 - Change CAPAKEY: "/"" is replaced by "-" (and not by "_")
    ************************************#>

<#

.SYNOPSIS
    This Powershell will take a single folder
    Read content of PDF files in this folder
    export the content to a CSV and JSON file

.DESCRIPTION
    How to use the script:
        1. Put itextsharp.dll in the same folder as the script.  This script uses itextsharp.dll to extract information from the PDF file.
        2. Change the directories and locations (parameter section) below.  At least "$PathBase" and "$PathToProcess" must be changed to your needs.
        3. Put all the PDF files to process in the $PathToProcess folder.
        4. Run the script.  If you see a lot of red alert messages passing by, don't panic.  Not all PDF files will contain all the values,
           so this results in empty values and alerts.  Noting to worry about.
        5. Put a little smile on your face and enjoy the result

.EXAMPLE

.NOTES
Put some notes here.

.LINK
#>


# ===============================================================================================
# PARAMETERS
# ===============================================================================================

# DIRECTORIES AND LOCATIONS
$PathBase     = "D:\05 - Temp\Kadasterinfo"                 # Base path for files and processing - No training slash !
$PathToProcess  = "$PathBase\ToProcess\"                    # Location where original PDF files must be places
$PathPDF  = "$PathBase\PDF\"                                # Location where processed PDF files will be moved to
$PathCSV  = "$PathBase\CSV\"                                # Location where CSV result will be placed
$PathJSON  = "$PathBase\JSON\"                              # Location where  JSON result will be placed


# ===============================================================================================
# FUNCTIONS
# ===============================================================================================

# FUNCTION READ PDF FILE
function ReadPdfFile {
    param(
        [string]$fileName
    )
    If(!(Test-Path .\itextsharp.dll))
    {
        throw "Change into directory contain itextsharp.dll before running this function."
    }
    Add-Type -Path .\itextsharp.dll
    #$text= New-Object -TypeName System.Text.StringBuilder
    if(Test-Path $fileName){
        $pdfReader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $fileName
            ForEach($page in (1..$pdfReader.NumberOfPages))
            {
                #$strategy = [iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy]
                #$PdfTextExtractor = [iTextSharp.text.pdf.parser.PdfTextExtractor]
                $currentText=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdfReader,$page,[iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy]::new())
                $UTF8 = New-Object System.Text.UTF8Encoding
                $ASCII = New-Object System.Text.ASCIIEncoding
                $EndText = $UTF8.GetString($ASCII::Convert([System.Text.Encoding]::Default, [System.Text.Encoding]::UTF8, [System.Text.Encoding]::DEFAULT.GetBytes($currentText)))
                $EndText

                # SPLIT THE LINES
                $lines = [char[]]$pdfReader.GetPageContent($page) -join "" -split "`n"
                foreach ($line in $lines) {
                    if ($line -match "^\[") {
                    #$line = $line -replace \\([\S]), $matches[1]
                    #$line -replace "^\[\(|\)\]TJ$", "" -split "\)\-?\d+\.?\d*\(" -join ""
                    }
                }

                #$currentText
                } #end for loop
    $pdfReader.Close()
        }
}

function CreateFolders {
    New-Item -Path $PathToProcess -ItemType Directory -Force
    New-Item -Path $PathPDF -ItemType Directory -Force
    New-Item -Path $PathCSV -ItemType Directory -Force
    New-Item -Path $PathJSON -ItemType Directory -Force
}
function Get-TotalJobTime {
    # $start_time = Get-Date
    $JobTimeDays = (Get-Date).Subtract($start_time).Days
    $JobTimeDays = If ($JobTimeDays -eq 0) { "" } else { " $JobTimeDays days" }
    $JobTimeHours = (Get-Date).Subtract($start_time).Hours
    $JobTimeHours = If ($JobTimeHours -eq 0) { "" } else { " $JobTimeHours hours" }
    $JobTimeMinutes = (Get-Date).Subtract($start_time).Minutes
    $JobTimeMinutes = If ($JobTimeMinutes -eq 0) { "" } else { " $JobTimeMinutes minutes" }
    $JobTimeseconds = (Get-Date).Subtract($start_time).Seconds
    $JobTimeseconds = If ($JobTimeseconds -eq 0) { "" } else { " $JobTimeseconds seconds" }
    $TotalJobTime = "This action took$JobTimeDays$JobTimeHours$JobTimeMinutes$JobTimeseconds to complete."
    $TotalJobTime
}

# ===============================================================================================
# PROGRAM
# ===============================================================================================
$start_time = Get-Date

# Create folders
CreateFolders

# Retrieve source documents to process
$sourcefiles = Get-ChildItem $PathToProcess -File -Filter "*.PDF"

# Process all files found...
foreach ($sourcefile in $sourcefiles) {
    Write-host "`nThe file '$sourcefile' will be processed...`n------------------------------------------------"

    # Get content of the PDF as single lines of text
    # $sourcefile = "D:\05 - Temp\Kadasterinfo\sourcePDF\export_81e82813-2508-441e-9807-e37d9911c566.pdf"
    $PDFContent = ReadPdfFile -fileName $sourcefile.FullName
    $PDFContent = $PDFContent -split "`n"           #Split content in multi line string

    # Extract single values from content
        # EigendomID
        $EigendomID = $PDFContent | Select-String -Pattern "EigendomID" -encoding ASCII
        $EigendomID = [Regex]::Match($EigendomID, '\d+').value
        Write-Host "   - EigendomID: $EigendomID"

        # Datum bevraging
        $Datumopvraging = ($PDFContent | Select-String -Pattern '\d+\/\d+\/\d+\s\d+:\d+' -encoding ASCII)[0]
        Write-Host "   - Datum Opvraging: $Datumopvraging"

        # InschrijvingVoorgaand
        $InschrijvingVoorgaand = $PDFContent | Select-String -Pattern "ArtikelInschrijving Voorgaand " -encoding ASCII
        $InschrijvingVoorgaand = (($InschrijvingVoorgaand -split 'ArtikelInschrijving Voorgaand ')[1]).Trim()
        Write-Host "   - InschrijvingVoorgaand: $InschrijvingVoorgaand"

        # InshrijvingHuidig
        $InshrijvingHuidig = $PDFContent | Select-String -Pattern "ArtikelInschrijving Voorgaand " -context 1 -encoding ASCII
        $InshrijvingHuidig = (($InshrijvingHuidig -split 'Huidig ')[1]).Trim()
        Write-Host "   - InshrijvingHuidig: $InshrijvingHuidig"

        # OppervlakteBelast
        $OppervlakteBelast = $PDFContent | Select-String -Pattern "AlgemeneInformaties Oppervlakte Belast" -encoding ASCII
        $OppervlakteBelast = (($OppervlakteBelast -split 'AlgemeneInformaties Oppervlakte Belast ')[1]).Trim()
        Write-Host "   - OppervlakteBelast: $OppervlakteBelast"

        # KIBedrag
        $KIBedrag = $PDFContent | Select-String -Pattern "Bedrag" -encoding ASCII
        $KIBedrag = (($KIBedrag -split 'Bedrag ')[1]).Trim()
        Write-Host "   - KIBedrag: $KIBedrag"

        # Partitie
        $Partitie = $PDFContent | Select-String -Pattern "Identificatie Partitie" -encoding ASCII
        $Partitie = (($Partitie -split 'Identificatie Partitie ')[1]).Trim()
        Write-Host "   - Partitie: $Partitie"

        # Letterexponent
        $Letterexponent = $PDFContent | Select-String -Pattern "Letterexponent" -encoding ASCII
        $Letterexponent = (($Letterexponent -split 'Letterexponent ')[1]).Trim()
        Write-Host "   - Letterexponent: $Letterexponent"

        # Cijferexponent
        $Cijferexponent = $PDFContent | Select-String -Pattern "Cijferexponent" -encoding ASCII
        $Cijferexponent = (($Cijferexponent -split 'Cijferexponent ')[1]).Trim()
        Write-Host "   - Cijferexponent: $Cijferexponent"

        # Bisnummer
        $Bisnummer = $PDFContent | Select-String -Pattern "Bisnummer" -encoding ASCII
        $Bisnummer = (($Bisnummer -split 'Bisnummer ')[1]).Trim()
        Write-Host "   - Bisnummer: $Bisnummer"

        # Grondnummer
        $Grondnummer = $PDFContent | Select-String -Pattern "Grondnummer" -encoding ASCII
        $Grondnummer = (($Grondnummer -split 'Grondnummer ')[1]).Trim()
        Write-Host "   - Grondnummer: $Grondnummer"

        # Sectie
        $Sectie = $PDFContent | Select-String -Pattern "Sectie" -encoding ASCII
        $Sectie = (($Sectie -split 'Sectie ')[1]).Trim()
        Write-Host "   - Sectie: $Sectie"

        # KadastraleAfdelingCode
        $KadastraleAfdelingCode = $PDFContent | Select-String -Pattern "KadastraleAfdeling Code" -encoding ASCII
        $KadastraleAfdelingCode = (($KadastraleAfdelingCode -split 'KadastraleAfdeling Code ')[1]).Trim()
        Write-Host "   - KadastraleAfdelingCode: $KadastraleAfdelingCode"

        # KadastraleAfdelingOmschrijving
        $KadastraleAfdelingOmschrijving = $PDFContent | Select-String -Pattern "KadastraleAfdeling Code" -context 1 -encoding ASCII
        $KadastraleAfdelingOmschrijving = (($KadastraleAfdelingOmschrijving -split 'Omschrijving ')[1]).Trim()
        Write-Host "   - KadastraleAfdelingOmschrijving: $KadastraleAfdelingOmschrijving"

        # LiggingIdPerceel
        $LiggingIdPerceel = $PDFContent | Select-String -Pattern "LiggingIdPerceel" -encoding ASCII
        $LiggingIdPerceel = (($LiggingIdPerceel -split 'LiggingIdPerceel ')[1]).Trim()
        Write-Host "   - LiggingIdPerceel: $LiggingIdPerceel"

        # StatusPatrimoniaalPerceel
        $StatusPatrimoniaalPerceel = $PDFContent | Select-String -Pattern "StatusPatrimoniaalPerceel" -encoding ASCII
        $StatusPatrimoniaalPerceel = (($StatusPatrimoniaalPerceel -split 'StatusPatrimoniaalPerceel ')[1]).Trim()
        Write-Host "   - StatusPatrimoniaalPerceel: $StatusPatrimoniaalPerceel "

        # AardKadastraalPerceel
        $AardKadastraalPerceel = $PDFContent | Select-String -Pattern "AardKadastraalPerceel" -encoding ASCII
        $AardKadastraalPerceel = (($AardKadastraalPerceel -split 'AardKadastraalPerceel ')[1]).Trim()
        Write-Host "   - AardKadastraalPerceel: $AardKadastraalPerceel"

        # AardKadastraalPerceelOmschrijving
        $AardKadastraalPerceelOmschrijving = $PDFContent | Select-String -Pattern "AardKadastraalPerceel" -context 0,1 -encoding ASCII
        $AardKadastraalPerceelOmschrijving = (($AardKadastraalPerceelOmschrijving -split 'Omschrijving ')[1]).Trim()
        Write-Host "   - AardKadastraalPerceelOmschrijving: $AardKadastraalPerceelOmschrijving"

        # Calculate CAPAKEY
            # prepare and convert parts of CAPAKEY
            $Grondnummer = '{0:d4}' -f [int]$Grondnummer                # Convert Grondnummer to format of 4 digits
            $Bisnummer = '{0:d2}' -f [int]$Bisnummer                    # Check if Bisnummer exists, if so, convert to two digits. If not, "00"
            $Cijferexponent = '{0:d3}' -f [int]$Cijferexponent          # Check if Bisnummer exists, if so, convert to three digits. if not, "000"
            $Letterexponent = If ($IsNull -eq $Letterexponent) {"_"} else {$Letterexponent}   # Check if Letterexponent exists, if not, use "_"

            # Build CAPAKEY
            $Capakey = $KadastraleAfdelingCode + $Sectie + $Grondnummer + "/" + $Bisnummer + $Letterexponent + $Cijferexponent
            $CapakeyFileName = $KadastraleAfdelingCode + $Sectie + $Grondnummer + "-" + $Bisnummer + $Letterexponent + $Cijferexponent
            Write-Host "   - Capakey: $Capakey"

    # extract 'partijen'

        # Create CSV file containing extracted detail information
        $CSVfileDetail = "$PathCSV" + "KadasterInfo_$EigendomID" + "_$CapakeyFileName" + "_Owner.CSV"
        New-Item $CSVfileDetail -Force | Out-Null
        Add-Content -Path $CSVfileDetail  -Value 'EigendomID;Capakey;PartijID;INSZ;Achternaam;Voornaam;Straat;Huisnummer;Postcode;Gemeente;NISCodeGemeente;Land;ISOCodeLand;Taal;ZakelijkRecht;Aandeel;Type;TypeNatuurlijkPersoon'


        $Partijen = $PDFContent | Select-String '^PartijID (\d+)' -Context 0, 11
        foreach ($Partij in $Partijen) {
            Write-Host "`nPartij`n---------------------------------------"
            $Partij = $Partij  -split "`n"           #Split content in multi line string

            # PrtijID
            $PartijID = $Partij | Select-String -Pattern "PartijID" -encoding ASCII
            $PartijID = [Regex]::Match($PartijID, '\d+').value
            Write-Host "   - PartijID: $PartijID"

            # INSZ
            $INSZ = $Partij | Select-String -Pattern "INSZ" -encoding ASCII
            $INSZ = [Regex]::Match($INSZ, '\d+').value
            Write-Host "   - INSZ: $INSZ"

            # Achternaam
            $Achternaam = $Partij | Select-String -Pattern 'Achternamen(.*?)Voornamen' -encoding ASCII  # Find line matching the pattern
            $Achternaam = [Regex]::Match($Achternaam, 'Achternamen(.*?)Voornamen' ).Groups[1].value     # find value between 'achternamen' and 'Voornamen'
            $Achternaam = ($Achternaam).Trim()                                                          # Trim spaces
            Write-Host "   - Achternaam: $Achternaam"

            # Voornaam
            $Voornaam = $Partij | Select-String -Pattern 'Voornamen ' -encoding ASCII                   # Find line matching the pattern
            $Voornaam = ($Voornaam -split 'Voornamen ')[1]
            $Voornaam = ($Voornaam).Trim()                                              # find value after 'Voornamen'
            Write-Host "   - Voornaam: $Voornaam"

            # Postcode
            $Postcode = $Partij | Select-String -Pattern 'Postcode' -encoding ASCII                     # Find line matching the pattern
            $Postcode = ($Postcode -split 'Postcode')[1]
            $Postcode = ($Postcode).Trim()                                                # find value after 'Voornamen'
            Write-Host "   - Postcode: $Postcode"

            # Straat
            $Straat = $Partij | Select-String -Pattern 'Adres Straat'  -encoding ASCII                 # Find line matching the pattern
            $Straat = ($straat -split 'Adres Straat ')[1]
            $Straat = $Straat -replace '[^\p{L}\p{Nd}]', ''
            Write-Host "   - Straat: $Straat"

            # Huisnummer
            $Huisnummer = $Partij | Select-String -Pattern 'Huisnummer ' -encoding ASCII                # Find line matching the pattern
            $Huisnummer = ($Huisnummer -split 'Huisnummer ')[1]
            $Huisnummer = ($Huisnummer).Trim()
            Write-Host "   - Huisnummer: $Huisnummer"

            # Gemeente
            $Gemeente = $Partij | Select-String -Pattern 'Gemeente [a-zA-Z]+' -encoding ASCII         # Find line matching the pattern
            $Gemeente = $Gemeente[1]
            $Gemeente = ($Gemeente -split 'Gemeente ')[1]
            $Gemeente = ($Gemeente).Trim()
            Write-Host "   - Gemeente: $Gemeente"

            # NISCodeGemeente
            $NISCodeGemeente = $Partij | Select-String -Pattern 'NISCodeGemeente' -encoding ASCII         # Find line matching the pattern
            $NISCodeGemeente = ($NISCodeGemeente -split 'NISCodeGemeente ')[1]
            $NISCodeGemeente = ($NISCodeGemeente).Trim()
            Write-Host "   - NISCodeGemeente: $NISCodeGemeente"

            # Land
            $Land = $Partij | Select-String -Pattern 'Land' -encoding ASCII                             # Find line matching the pattern
            $Land = [Regex]::Match($Land, 'Land Naam (.*?)NISCode' ).Groups[1].value                    # find value between
            $Land = ($Land).Trim()
            Write-Host "   - Land: $Land"

            # Taal
            $Taal = $Partij | Select-String -Pattern 'Taal' -encoding ASCII         # Find line matching the pattern
            $Taal = ($Taal -split 'Taal ')[1]
            $Taal = ($Taal).Trim()
            Write-Host "   - Land: $Taal"

            # ISOCodeLand
            $ISOCodeLand = $Partij | Select-String -Pattern 'ISOCodeLand' -encoding ASCII         # Find line matching the pattern
            $ISOCodeLand = ($ISOCodeLand -split 'ISOCodeLand ')[1]
            $ISOCodeLand = ($ISOCodeLand).Trim()
            Write-Host "   - ISOCodeLand: $ISOCodeLand"

        # extract 'Eigenaarsrechten'
            $pattern = "LijstId PartijID $PartijID"
            $Eigenaarsrecht = $PDFContent  | Select-String $pattern -Context 6, 0
            $Eigenaarsrecht = $Eigenaarsrecht  -split "`n"           #Split content in multi line string

            # ZakelijkRecht
            $ZakelijkRecht = $Eigenaarsrecht | Select-String -Pattern 'ZakelijkRecht Ongestructureerd Rechten' -encoding ASCII         # Find line matching the pattern
            $ZakelijkRecht = ($ZakelijkRecht -split 'ZakelijkRecht Ongestructureerd Rechten ')[1]
            $ZakelijkRecht = ($ZakelijkRecht).Trim()
            Write-Host "   - ZakelijkRecht: $ZakelijkRecht"

            # Aandeel
            $Aandeel = $Eigenaarsrecht | Select-String -Pattern 'Aandeel' -encoding ASCII         # Find line matching the pattern
            $Aandeel = ($Aandeel -split 'Aandeel ')[1]
            $Aandeel = ($Aandeel).Trim()
            Write-Host "   - Aandeel: $Aandeel"

            # Type
            $Type = $Eigenaarsrecht | Select-String -Pattern 'Code' -encoding ASCII         # Find line matching the pattern
            $Type = ($Type -split 'Code ')[1]
            $Type = ($Type).Trim()
            Write-Host "   - Type: $Type"

            # TypeNatuurlijkPersoon
            $TypeNatuurlijkPersoon = $Eigenaarsrecht | Select-String -Pattern 'Natuurlijk Persoon' -encoding ASCII         # Find line matching the pattern
            $TypeNatuurlijkPersoon = ($TypeNatuurlijkPersoon -split 'Natuurlijk Persoon ')[1]
            $TypeNatuurlijkPersoon = ($TypeNatuurlijkPersoon).Trim()
            Write-Host "   - TypeNatuurlijkPersoon: $TypeNatuurlijkPersoon"


        # Add content to the CSV file
            # Add content
            $contentdetail = "$EigendomID;$Capakey;$PartijID;$INSZ;$Achternaam;$Voornaam;$Straat;$Huisnummer;$Postcode;$Gemeente;$NISCodeGemeente;$Land;$ISOCodeLand;$Taal;$ZakelijkRecht;$Aandeel;$Type;$TypeNatuurlijkPersoon"
            Add-Content -Path $CSVfileDetail  -Value $contentdetail
        }

    # Add content to the CSV file
        # Create CSV file containing extracted information
        $CSVfile = "$PathCSV" + "KadasterInfo_$EigendomID" + "_$CapakeyFileName" + "_Parcel.CSV"
        New-Item $CSVfile -Force | Out-Null
        Add-Content -Path $CSVfile  -Value 'EigendomID;Datumopvraging;InschrijvingVoorgaand;InshrijvingHuidig;OppervlakteBelast;KIBedrag;Partitie;Letterexponent;Grondnummer;Sectie;KadastraleAfdelingCode;KadastraleAfdelingOmschrijving;LiggingIdPerceel;StatusPatrimoniaalPerceel;AardKadastraalPerceel;AardKadastraalPerceelOmschrijving;Capakey'
        # Add content
        $content = "$EigendomID;$Datumopvraging;$InschrijvingVoorgaand;$InshrijvingHuidig;$OppervlakteBelast;$KIBedrag;$Partitie;$Letterexponent;$Grondnummer;$Sectie;$KadastraleAfdelingCode;$KadastraleAfdelingOmschrijving;$LiggingIdPerceel;$StatusPatrimoniaalPerceel;$AardKadastraalPerceel;$AardKadastraalPerceelOmschrijving;$Capakey"
        Add-Content -Path $CSVfile  -Value $content

    # Save content as JSON
        $JSONfile = "$PathJSON" + "KadasterInfo_$EigendomID" + "_$CapakeyFileName" + "_Parcel.JSON"
        Import-CSV $CSVFile -Delimiter ";" | ConvertTo-Json | Out-File $JSONfile

        $JSONfileDetail = "$PathJSON" + "KadasterInfo_$EigendomID"+ "_$CapakeyFileName" + "_Owner.JSON"
        Import-CSV $CSVFileDetail -Delimiter ";" | ConvertTo-Json | Out-File $JSONfileDetail

    # Copy the current file to new file name
        # Compose new file name
        $WVI_NewName = "KadasterInfo_$EigendomID"+ "_$CapakeyFileName" + ".pdf"
        $WVI_NewFullname = "$PathToProcess$WVI_NewName"

        # Rename current file based on content of the file
        rename-Item -path $sourcefile.FullName -NewName $WVI_NewFullname -Force

        # Move item to PDF folder
        Move-Item $WVI_NewFullname $PathPDF -Force

}

Get-TotalJobTime #  Put following command at beginning of the job execution:   $start_time = Get-Date





