*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

#https://robocorp.com/docs/developer-tools/visual-studio-code/extension-features#linking-to-control-room
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.FileSystem
Library             Dialogs
Library             RPA.Tables
Library             RPA.Archive
Library             RPA.PDF
Library             RPA.Dialogs
Library             RPA.Excel.Files
Library             OperatingSystem
Library             String
Library             BuiltIn
Library             Collections
Library             RPA.Robocorp.Vault


*** Variables ***
${SourceData_URL}           https://robotsparebinindustries.com/orders.csv
${SourceData_Download}      ${CURDIR}${/}data
${Robot_Orders}             ${CURDIR}${/}orders
${secret}


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    [Documentation]    aka: 'Follow Orders Like A Boss'
    Log To Console    Here We Go AGAIN..
    Input form dialog CSV File
    Modifying secrets
    Reading secrets
    Get CSV Source Data
    Log To Console    ..Source Data recieved..
    Open URL to Robot order
    Get Reference Data
    Log To Console    ..Robot Model Info recieved..
    Log To Console    ..LETS ORDER:
    RPA.FileSystem.Empty Directory    ${Robot_Orders}
    FOR    ${rivi}    IN    @{GLOBAL_SourceData}
        Log To Console    Ordernumber: ${rivi}[Order number]
        Fill Order In    ${rivi}
        Hae Preview Pic
        Send the Order and Get the Receipt    ${rivi}
        Build Receipt with Preview
        Go to Oder Another Robot
    END
    Log To Console    ..All Orders Done and receipts Printed..
    Close All Browsers
    Create a ZIP file of the receipts
    RPA.FileSystem.Empty Directory    ${Robot_Orders}
    Log To Console    ...ALL DONE. Find the archived zip-file of all the reciepts from output -directory.
    Success dialog


*** Keywords ***
Get CSV Source Data
    [Documentation]    Haetaan cvs/excel tiedosto linkistä ja listään se GLOBAL variableen
    Set Download Directory    directory=${SourceData_Download}
    RPA.FileSystem.Empty Directory    ${SourceData_Download}
    #Open Available Browser    ${SourceData_URL}    #VANHA tapa
    Open Available Browser    ${GLOBAL_saatu_CSV_linkki}
    RPA.FileSystem.Wait Until Created    ${SourceData_Download}${/}orders.csv    timeout=5
    Close All Browsers
    ${csv_table} =    Read table from CSV    ${SourceData_Download}${/}orders.csv
    Log To Console    ${csv_table}
    Set Suite Variable    ${GLOBAL_SourceData}    ${csv_table}

Open URL to Robot order
    [Documentation]    Avataan robot oder -sivu ja piilotetaan POPUP
    Open Chrome Browser    https://robotsparebinindustries.com/#/robot-order
    Wait Until Element Is Visible    //div[@class="modal" and @role="dialog"]    5s
    Execute Javascript    document.querySelector("#root > div > div.modal").setAttribute("style", "display: none;");
    Wait Until Element Is Not Visible    //div[@class="modal" and @role="dialog"]    5s

Get Reference Data
    [Documentation]    Hateaan Referenssi -taulukko "Show model info" kohdasta.
    Wait Until Element Is Visible    //button[contains(.,"Show model info")]    2s
    Click Button    //button[contains(.,"Show model info")]
    Wait Until Element Is Visible    //button[contains(.,"Hide model info")]    2s
    Wait Until Element Is Visible    //table[@id="model-info"]    2s
    ${Riveja_kpl} =    Get Element Count    //table[@id="model-info"]//tr
    @{Table_Data_1} =    Create List
    @{Table_Data_2} =    Create List
    FOR    ${rivi}    IN RANGE    ${Riveja_kpl}
        ${rivi} =    Evaluate    int(${rivi}) + int(1)
        ${Taulukko_rivi_1} =    RPA.Browser.Selenium.Get Table Cell    //table[@id="model-info"]    ${rivi}    1
        ${Taulukko_rivi_2} =    RPA.Browser.Selenium.Get Table Cell    //table[@id="model-info"]    ${rivi}    2
        Append To List    ${Table_Data_1}    ${Taulukko_rivi_1}
        Append To List    ${Table_Data_2}    ${Taulukko_rivi_2}
    END
    &{Table_Data} =    Create Dictionary
    ...    column1=${Table_Data_1}
    ...    column2=${Table_Data_2}
    ${Robot_Model_Info} =    Create Table    ${Table_Data}
    Set Row As Column Names    ${Robot_Model_Info}    0
    Log To Console    ${Robot_Model_Info}
    Set Suite Variable    ${GLOBAL_Robot_Model_Info}    ${Robot_Model_Info}

Fill Order In
    [Documentation]    Täytä rivin tiedot sivulle ja paina lähetä/send!
    [Arguments]    ${row}
    Set Suite Variable    ${GLOBAL_ORDERNRO}    ${row}[Order number]
    Set Suite Variable    ${GLOBAL_ORDER_ADDRESS}    ${row}[Address]
    Wait Until Element Is Visible    //select[@id="head"]    5s
    Wait Until Element Is Enabled    //select[@id="head"]    2s
    Select From List By Value    //select[@id="head"]    ${row}[Head]
    Wait Until Element Is Visible    //input[@name="body" and @id="id-body-${row}[Body]"]    5s
    Wait Until Element Is Enabled    //input[@name="body" and @id="id-body-${row}[Body]"]    2s
    Select Radio Button    body    ${row}[Body]
    #Huom! Jalkoja voi olla 1-6 kpl!
    Wait Until Element Is Visible    //input[@placeholder="Enter the part number for the legs"]    5s
    Wait Until Element Is Enabled    //input[@placeholder="Enter the part number for the legs"]    2s
    Input Text    //input[@placeholder="Enter the part number for the legs"]    ${row}[Legs]
    Wait Until Element Is Visible    //input[@placeholder="Shipping address"]    5s
    Wait Until Element Is Enabled    //input[@placeholder="Shipping address"]    2s
    Input Text    //input[@placeholder="Shipping address"]    ${row}[Address]

Hae Preview Pic
    [Documentation]    Haetaan ennakko-kuva talteen
    Sleep    2s
    Wait Until Element Is Visible    //button[@id="preview"]    5s
    Wait Until Element Is Enabled    //button[@id="preview"]    2s
    Click Button    //button[@id="preview"]
    Sleep    2s
    Wait Until Element Is Enabled    //div[@id="robot-preview-image"]/img[@alt="Legs"]    5s
    Execute Javascript    window.scrollTo(0, 1000)
    Sleep    1s
    RPA.Browser.Selenium.Screenshot
    ...    locator=//div[@id="robot-preview-image"]
    ...    filename=${Robot_Orders}${/}Preview_${GLOBAL_ORDERNRO}.png
    RPA.FileSystem.Wait Until Created    ${Robot_Orders}${/}Preview_${GLOBAL_ORDERNRO}.png    timeout=10
    Set Suite Variable    ${GLOBAL_Robot_Preview_PIC}    ${Robot_Orders}${/}Preview_${GLOBAL_ORDERNRO}.png

Send the Order and Get the Receipt
    [Documentation]    Lähetä valmis tilaus.
    [Arguments]    ${rivi}
    ${status} =    Run Keyword And Return Status    Send the Order
    IF    ${status} == ${FALSE}
        Wait Until Keyword Succeeds    5 times    2s    Send the Order by Force
    END
    ${Receipt_html} =    Get Element Attribute    //div[@id="order-completion"]/div[@id="receipt"]    outerHTML
    Html To Pdf    ${Receipt_html}    ${Robot_Orders}${/}Receipt_${GLOBAL_ORDERNRO}.pdf
    Set Suite Variable    ${GLOBAL_Robot_Reciept}    ${Robot_Orders}${/}Receipt_${GLOBAL_ORDERNRO}.pdf

Send the Order
    [Documentation]    Yritä lähettää tilausta KERRAN
    Wait Until Element Is Visible    //button[@id="order"]    5s
    Wait Until Element Is Enabled    //button[@id="order"]    2s
    Click Button    //button[@id="order"]
    Wait Until Element Is Visible    //div[@id="order-completion"]    5s

Send the Order by Force
    [Documentation]    Yritä lähettää tilausta MONTA kertaa
    Execute Javascript    window.scrollTo(0, 0)
    Sleep    2s
    Wait Until Element Is Visible    //button[@id="preview"]    5s
    Wait Until Element Is Enabled    //button[@id="preview"]    2s
    Click Button    //button[@id="preview"]
    Sleep    2s
    Wait Until Element Is Visible    //button[@id="order"]    5s
    Wait Until Element Is Enabled    //button[@id="order"]    2s
    Click Button    //button[@id="order"]
    Wait Until Element Is Visible    //div[@id="order-completion"]    5s

Build Receipt with Preview
    [Documentation]    Yhdistä kuittiin kuva tilauksesta
    #${files} =    Create List    ${GLOBAL_Robot_Preview_PIC}:align=center
    Open Pdf    ${GLOBAL_Robot_Reciept}
    #Add Files To PDF    ${files}    ${GLOBAL_Robot_Reciept}    True
    Add Watermark Image To PDF
    ...    image_path=${GLOBAL_Robot_Preview_PIC}
    ...    source_path=${GLOBAL_Robot_Reciept}
    ...    output_path=${GLOBAL_Robot_Reciept}
    Close All Pdfs
    RPA.FileSystem.Remove File    ${GLOBAL_Robot_Preview_PIC}

Go to Oder Another Robot
    [Documentation]    Palataan takaisin alkuun jotta voidaan tilata uusi robotti/nollata sivu
    Wait Until Element Is Visible    //button[@id="order-another"]    5s
    Wait Until Element Is Enabled    //button[@id="order-another"]    2s
    Click Button    //button[@id="order-another"]
    Wait Until Element Is Visible    //div[@class="modal" and @role="dialog"]    10s
    Execute Javascript    document.querySelector("#root > div > div.modal").setAttribute("style", "display: none;");
    Wait Until Element Is Not Visible    //div[@class="modal" and @role="dialog"]    10s

Create a ZIP file of the receipts
    [Documentation]    Kokoa kuitit/pdf tiedostot ZIP-tiedostoksi.
    Archive Folder With Zip    ${Robot_Orders}    ${CURDIR}${/}output${/}order_receipts.zip

Input form dialog CSV File
    Add heading    Give a robot a link to a csv-file that contains the orders
    Add text input    link    label=The Order Link    placeholder=Enter csv-link here
    ${saatu_csv_linkki} =    Run dialog
    Set Suite Variable    ${GLOBAL_saatu_CSV_linkki}    ${saatu_csv_linkki.link}

Success dialog
    Add icon    Success
    Add heading    Your orders have been processed
    Add file    ${CURDIR}${/}output${/}order_receipts.zip    label=Order Recipts
    Run dialog    title=Successful_Orders

Modifying secrets
    ${secret} =    Get Secret    order_link
    ${level} =    Set Log Level    NONE
    Set To Dictionary    ${secret}    link    ${GLOBAL_saatu_CSV_linkki}
    Set Log Level    ${level}
    Set Secret    ${secret}

Reading secrets
    ${secret} =    Get Secret    order_link
    Log To Console    SALAISUUDET: ${secret}
