*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.PDF
Library             RPA.Tables
Library             RPA.Archive
Library             RPA.RobotLogListener
Library             Screenshot
Library             Dialogs
Library             OperatingSystem
Library             Collections
Library             RPA.Robocloud.Secrets


*** Variables ***
${receipt_dir}      ${OUTPUT_DIR}${/}receipts/
${sshot_dir}        ${OUTPUT_DIR}${/}screenshot/
${zip_dir}          ${OUTPUT_DIR}${/}


*** Tasks ***
Order Robot
    Log    Process to order Robots from Robot Sparebin Industries.
    Open Order Robot Website
    ${Orders}=    Download Excel file
    FOR    ${row}    IN    @{orders}
        Click Button    OK
        Wait Until Page Contains Element    class:form-group
        Fill the form using the data from the Excel file    ${row}
        Wait Until Keyword Succeeds    10x    2s    Preview the robot
        Wait Until Keyword Succeeds    10x    2s    Submit the order
        ${pdf}=    Store the receipt as a PDF file    ${row}[Order number]
        ${screenshot}=    Take a screenshot of the robot    ${row}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
        Go to order another robot
    END
    Create a ZIP file of the receipts
    Delete screenshots and receipts
    Log out and close the browser


*** Keywords ***
Open Order Robot Website
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order

Download Excel File
    Download    https://robotsparebinindustries.com/orders.csv    overwrite=True
    ${Orders}=    Read table from CSV    orders.csv
    RETURN    ${Orders}

Fill the form using the data from the Excel file
    [Arguments]    ${ordernum}
    Select From List By Value    head    ${ordernum}[Head]
    Select Radio Button    body    ${ordernum}[Body]
    Input Text    xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${ordernum}[Legs]
    Input Text    address    ${ordernum}[Address]

Preview the robot
    Click Button    preview
    Wait Until Element Is Visible    robot-preview-image

Submit the order
    Mute Run On Failure    Page Should Contain Element
    Click Button    order
    Page Should Contain Element    receipt

Store the receipt as a PDF file
    [Arguments]    ${order_id}
    Set Local Variable    ${receipt_filename}    ${receipt_dir}receipt_${order_id}.pdf
    ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    content=${receipt_html}    output_path=${receipt_filename}
    RETURN    ${receipt_filename}

    # Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf
    # ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    # Wait Until Element Is Visible    id:sales-results

Take a screenshot of the robot
    [Arguments]    ${order_id}
    Set Local Variable    ${sshot_filename}    ${sshot_dir}robot_${order_id}.png
    Screenshot    id:robot-preview-image    ${sshot_filename}
    RETURN    ${sshot_filename}

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${sshot}    ${receipt}
    # dont know why dont need to open and close pdf
    # Open Pdf    ${receipt}

    # Create the list of files that is to be added to the PDF (1 file)
    @{myfiles}=    Create List    ${sshot}:x=0,y=0
    Add Files To PDF    ${myfiles}    ${receipt}    ${True}
    # Close PDF    ${receipt}

Go to order another robot
    Wait Until Keyword Succeeds    10x    1s    Click Button When Visible    order-another

Create a ZIP file of the receipts
    ${name}=    Get Value From User    Give name for ZIP Folder
    Create the ZIP    ${name}

Create the ZIP
    [Arguments]    ${name}
    Archive Folder With Zip    ${receipt_dir}    ${zip_dir}${name}

Delete screenshots and receipts
    Empty Directory    ${sshot_dir}
    Empty Directory    ${receipt_dir}

Log out and close the browser
    Close Browser
