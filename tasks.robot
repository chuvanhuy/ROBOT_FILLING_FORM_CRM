*** Settings ***
Documentation       Robot Login CRM

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP


*** Variables ***
${WEB_ADDRESS}      https://robotsparebinindustries.com/


*** Tasks ***
Robot Login LMS
    Mở trình duyệt & Website
    Đăng nhập
    # [Teardown]    Thoát và đóng trình duyệt


*** Keywords ***
Mở trình duyệt & Website
    Open Available Browser    ${WEB_ADDRESS}

Đăng nhập
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form

# Thoát và đóng trình duyệt
    # Click Element    ${WEB_ADDRESS}/login/logout.php
    # Close Browser
