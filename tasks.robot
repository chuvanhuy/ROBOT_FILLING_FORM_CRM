*** Settings ***
Documentation       Robot Filling Form CRM

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Dialogs


*** Variables ***
${WEB_ADDRESS}      https://robotsparebinindustries.com/


*** Tasks ***
Robot Filling Form CRM
    Mở trình duyệt & Website
    Đăng nhập
    ${excel_file_path}=    Lấy file dữ liệu Excel cần nhập thông tin từ phía người dùng
    Nhập dữ liệu    ${excel_file_path}
    Báo nhập dữ liệu thành công
    # [Teardown]    Thoát và đóng trình duyệt


*** Keywords ***
Mở trình duyệt & Website
    Open Available Browser    ${WEB_ADDRESS}

Đăng nhập
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form

Lấy file dữ liệu Excel cần nhập thông tin từ phía người dùng
    Add heading    Tải File dữ liệu Excel muốn hệ thống nhập tự động
    Add file input
    ...    label=Tải File dữ liệu Excel muốn hệ thống nhập tự động
    ...    name=fileupload
    ...    file_type=Excel files (*.xls;*.xlsx)
    ...    destination=${OUTPUT_DIR}
    ${response}=    Run dialog
    RETURN    ${response.fileupload}[0]

Nhập dữ liệu
    [Arguments]    ${excel_file_path}
    Open Workbook    ${excel_file_path}
    ${nhap_records}=    Read Worksheet As Table    header=True
    Close Workbook

    FOR    ${nhap_record}    IN    @{nhap_records}
        Input Text    firstname    ${nhap_record}[Họ và tên đệm]
        Input Text    lastname    ${nhap_record}[Tên]
        Input Text    salesresult    ${nhap_record}[Kết quả bán hàng thực tế]
        Select From List By Value    salestarget    ${nhap_record}[Mục tiêu bán hàng]
        Click Button    Submit
    END

Báo nhập dữ liệu thành công
    Add icon    Success
    Add heading    Bạn đã nhập tự động tất cả các dữ liệu trong file Excel lên hệ thống !!!
    Run dialog    title=Thông báo hoàn tất nhập dữ liệu

# Thoát và đóng trình duyệt
    # Click Element    ${WEB_ADDRESS}/login/logout.php
    # Close Browser
