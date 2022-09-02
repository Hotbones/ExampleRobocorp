*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.
Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.PDF
Library    RPA.Desktop
   

*** Keywords ***
Open Intranet Robocorp
    Open Available Browser    https://robotsparebinindustries.com/#/

Log-in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${True}

Fill and Submit form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[Last Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Input Text    salesresult    ${sales_rep}[Sales]
    Click Button    Submit

Fill form using data from excel file
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook

    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and Submit form for one person    ${sales_rep}
        
    END
Collect the results
    RPA.Browser.Selenium.Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Exports the table as PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_result_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_result_html}    ${OUTPUT_DIR}${/}sales_results.pdf

 Open PDF file
     Open File  ${OUTPUT_DIR}${/}sales_results.pdf
     #Close Pdf    sales_results.pdf 

Log out and close
    Click Button    logout
    Close Browser

*** Tasks ***
Open Browser and Log in
    Open Intranet Robocorp
    Log-in    
    Download excel file
    #Fill and Submit form for one person
    Fill form using data from excel file
    Collect the results
    Exports the table as PDF
    Open PDF file
    [teardown]    Log out and close