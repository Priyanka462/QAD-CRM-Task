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
Library    RPA.Tables
Library    RPA.Outlook.Application

*** Variables ***
${Account_Name}    priyanka.madugula@yash.com

*** Keywords ***
Opening QAD website
    Open Available Browser     https://gl034jaerpc496.qad.com/clouderp/resources/login.jsp#/view/webshell/home    maximized=True    browser_selection=chrome

Login with credentials
    ${secret}=    Get Secret    QAD
    Input Text    id:username    ${secret}[username]
    Sleep    2s
    Input Text    id:password    ${secret}[password]
    Sleep    2s
    Click Button    //*[@id="logInBtn"]
    Sleep    5s
Searching for details
    Click Element  xpath=//span[@class="fa fa-search"]
    Sleep    2s
    Input Text    //*[@id="webshellMenu_kAutoCompleteMenuSearch"]    CRM Accounts
    Sleep    3s
    Click Element    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li[1]/a/div/div/span[2]
    Sleep    10s
    Click Element     xpath=//a[@id="ToolBtnNew"]
    Sleep    2s
Create form in CRM Account details
    Open Workbook    QAD input.xlsx
    Set Active Worksheet    Sheet1
    ${table}=    Read Worksheet As Table    header=True
    [Return]    ${table}
    Log To Console    ${table}
    FOR    ${element}    IN    @{table}
            Input Text    //*[@id="NameAutoField"]    ${element}[Account Name]
            Input Text    //*[@id="Address1AutoField"]    ${element}[Address]
            Input Text    //*[@id="PostCodeAutoField1"]    ${element}[Pincode]
            Input Text    //*[@id="CityAutoField"]    ${element}[City]
            Input Text    //*[@id="StateAutoField"]    ${element}[State]
            Input Text    //*[@id="PhoneAutoField"]    ${element}[Phone Number]
            Input Text    //*[@id="EmailAutoField"]    ${element}[Email]
            Input Text    //*[@id="AccountRegionAutoField"]    ${element}[Region]
            Click Button    //*[@id="ParentNameAutoField_lookup"]
            Sleep    3s
            Click Button    //*[@id="lookUpOkBtn"]
            Sleep    2s
            Click Button    //*[@id="ToolBtnSave"]
            Sleep    2s
            Click Element     xpath=//a[@id="ToolBtnNew"]
            Sleep    2s
            Log To Console    ${table}
        
    END  
    Click Button    //*[@id="btnViewFormPane"] 
    #Click Element    //*[@id="qHybridViewNavMaint"]/div/span  
Fetching CRM Account details

    ${count}=    Set Variable     2
    ${Account Name}=    Get WebElements    xpath=//td[contains(@class,"qGridDataString qFieldName-crmAccount.primaryContactName")]   
        FOR    ${element}    IN    @{Account Name}
            ${text}=    Get Text    ${element}
            ${name}=    Log To Console    ${text}
            Open Workbook    QAD output.xlsx    
            Read Worksheet    Sheet1
            Set Cell Value    ${count}    A   ${text}
            Save Workbook
            ${count}=    Evaluate    ${count} + 1
            
        END

    ${count}=    Set Variable     2
    ${Email}=    Get WebElements    xpath=//td[contains(@class,"qGridDataString qFieldName-crmAccount.primaryEmail")]
        FOR    ${PN}    IN    @{Email}
            ${num}=    Get Text    ${PN}
            ${pnum}=    Log To Console    ${num}
            Open Workbook    QAD output.xlsx     
            Read Worksheet    Sheet1
            Set Cell Value    ${count}    B   ${num}
            Save Workbook
            ${count}=    Evaluate    ${count} + 1
        
        END

    ${count}=    Set Variable     2    
    ${Ultimate acc name}=    Get WebElements    xpath=//td[contains(@class,"qGridDataString qFieldName-joinTable_8e49bfd24977eea.name")]   
        FOR     ${PA}    IN    @{Ultimate acc name}
                ${acct}=    Get Text    ${PA}
                ${pacct}=    Log To Console    ${acct}
                Open Workbook    QAD output.xlsx     
                Read Worksheet    Sheet1
                Set Cell Value    ${count}    C   ${acct}
                Save Workbook
                ${count}=    Evaluate    ${count} + 1
        
        END

    ${count}=    Set Variable     2
    ${Adress}=    Get WebElements    xpath=//td[contains(@class,"qGridDataString qFieldName-crmAccount.address1")]   
        FOR     ${Addres}    IN    @{Adress}
                ${adressoutput}=    Get Text    ${Addres}
                ${add}=    Log To Console    ${adressoutput}
                Open Workbook    QAD output.xlsx     
                Read Worksheet    Sheet1
                Set Cell Value    ${count}    D   ${adressoutput}
                Save Workbook
                ${count}=    Evaluate    ${count} + 1
        
        END

Get status update for data in excel
    Open Workbook    QAD output.xlsx
    ${table}=    Read Worksheet As Table    header=True
    ${count}=    Set Variable    2
    FOR    ${element}    IN    @{table}
            ${acc name}=    Set Variable    ${element}[Account Name]
            ${Email}=    Set Variable    ${element}[Email ID]
            ${ultimatename}=    Set Variable    ${element}[Ultimate Account Name]
            ${address}=    Set Variable    ${element}[ADDRESS]
            IF    "${acc_name}" != "None"
                    IF    "${Email}" != "None"
                        IF    "${ultimatename}" != "None"
                            IF    "${address}" != "None"
                                 Open Workbook    QAD output.xlsx
                                 Set Cell Value    ${count}    E   Completed
                                 ${count}=    Evaluate    ${count} + 1
                                 Save Workbook
                            ELSE
                                 Open Workbook    QAD output.xlsx
                                 Set Cell Value    ${count}    E   Incomplete
                                 ${count}=    Evaluate    ${count} + 1
                                 Save Workbook
                            END
                        ELSE
                                 Open Workbook    QAD output.xlsx
                                 Set Cell Value    ${count}    E    Incomplete 
                                 ${count}=    Evaluate    ${count} + 1 
                                 Save Workbook
                        END
                    ELSE
                                 Open Workbook    QAD output.xlsx
                                 Set Cell Value    ${count}    E   Incomplete
                                 ${count}=    Evaluate    ${count} + 1
                                 Save Workbook
                    END
            ELSE
                                 Open Workbook    QAD output.xlsx
                                 Set Cell Value    ${count}    E   Incomplete
                                 ${count}=    Evaluate    ${count} + 1
                                 Save Workbook
            END
        
    END  
      
Send mail to the incomplete users  
    Open Workbook    QAD output.xlsx    
    ${table}=    Read Worksheet As Table    Sheet1    header=True 
    ${count}=    Set Variable    2
      FOR        ${row}    IN        @{table}     
                Log   ${row}
                 ${Email}=    Set Variable    ${row}[Email ID]
                 IF   "${Email}" != "None"
                     Open Application
                    Send Message  ${row}[Email ID]    subject=Hi   body=Please update the account details  
                     Open Workbook    QAD output.xlsx
                    Set Cell Value    ${count}    F   mail sent successfully
                    ${count}=    Evaluate    ${count} + 1
                     Save Workbook
                 ELSE
                    Open Workbook    QAD output.xlsx
                    Set Cell Value    ${count}    F   mail not found
                    ${count}=    Evaluate    ${count} + 1
                    Save Workbook
                 END    
     END
                
         
    
*** Tasks ***
QAD demo
    
             Opening QAD website
             Login with credentials
             Searching for details
              Open Workbook    QAD output.xlsx            
              Read Worksheet    Sheet1    header=True
              Set Cell Value    1    A    Account Name
              Set Cell Value    1    B    Email ID
              Set Cell Value    1    C    Ultimate Account Name
              Set Cell Value    1    D    ADDRESS
              Set Cell Value    1    E    Status
              Set Cell Value    1    F    Mail status
              Save Workbook    QAD output.xlsx    
         Fetching CRM Account details
         Get status update for data in excel
         Send mail to the incomplete users

    
    
        
    
    