*** Settings ***
Documentation    QAD CRM Accounts
...     Need to Open QAD website and login with credentials through vault
...     Bot will search for CRM Accounts
...     After need to fill the details 
...     Bot will scrap all the details and save it into excel sheet
...     Send mail to those who are not filling all the details
...     submit 

Library    RPA.Robocloud.Secrets
Library    RPA.Browser.Selenium    auto_close=${FALSE}    
Library    RPA.Excel.Files
Library    RPA.FileSystem
Library    RPA.Robocorp.WorkItems
Library    XML
Library    RPA.Outlook.Application
*** Variables ***
${Directory}=    C:\\Users\\priyanka.madugula\\Documents\\visual studio code\\QAD task
${Account}=    priyanka.madugula@yash.com
${Global}=    ${True}



*** Keywords ***

Send Mail
    [Arguments]    ${Subject}    ${Body}
    Log To Console    Exception is :----${Body}
    Send Email    recipients=${Account}   subject=${Subject}    body=${Body}
Get Data From Config File
    ${Check1}=    Is Directory Empty    ${Directory}
    IF    ${Check1}
        Set Global Variable    ${Global}    ${False}
        Set Global Variable    ${Exception}    ${Directory} directory is Empty.
        Send Email    recipients=${Account}    subject=Business Exception    body=${Exception}
    ELSE
        ${Check2}=    Does File Exist    config.xml
        IF    ${Check2}
            ${Check3}=    Is File Empty    config.xml
            IF    ${Check3}
                Set Global Variable    ${Global}    ${False}
                Send Email    recipients=${Account}    subject=Business Exception   body=Config file is empty
            ELSE
                 ${config}=    Parse Xml    config.xml
                 ${input_Excel_file}=    Get Element Text   ${config}[1]
                 Open Workbook    ${input_Excel_file}     
                ${table}=    Read Worksheet As Table    header=${True}
                Close Workbook
                Log To Console    ${table}
                FOR    ${element}    IN    @{table}
                    Log    ${element}
                    Create Output Work Item    ${element}
                    Save Work Item
            
                END 
            END
        ELSE
            Set Global Variable    ${Global}    ${False}
            Send Email    recipients=${Account}    subject=Business Exception    body=File is not exists
        END
    END
 
   

*** Tasks ***
Opening QAD website
    Get Data From Config File