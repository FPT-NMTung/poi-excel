# Hướng dẫn config file 
## Excel

![image](https://github.com/FPT-NMTung/poi-excel/assets/63092029/60521b7b-0e13-4b7a-ad1d-54cfac4470e5)

### Thông tin sheet config

**Thông tin chung**

* `totalGroup`: Tổng số bảng lồng nhau
* `isHasGeneralData`: File có thông tin chung như thông tin KH, cổ phiếu, đợt thực hiện quyền, ...
* `isMergeCell`: Bảng cần tạo ra các khoảng được merge với nhau

**Bảng bên dưới bao gồm thông tin chi tiết về cấu hình các khoảng bảng mẫu:**

* `range_*`: Thứ tự bảng
* `begin`: Điểm bắt đầu của bảng
* `end`: Điểm kết thức của bảng
* `name_col`: Thông tin cần được fill của bảng, nếu có nhiều hơn một thông tin thì ngăn cách bới dấu `,`. VD: `name` or `name,age,address`

### Thông tin sheet Report

* Các thông tin thông thường sẽ được biểu diễn dưới dạng `<#table.(name_column)>` với `name_column` là tên của cột dữ liệu cần điều. VD: `<#table.address>`.
* Các thông tin chung sẽ được biểu diễn dưới dạng `<#general.(name_column)>` với `name_column` là tên của cột dữ liệu cần điều. VD: `<#general.name_report>`.
* Với những cột cần merge sẽ thêm hậu tố `<#merge.<#table.(name_column)>>` với `name_column` là tên của dữ liệu để phân biệt các khoảng merge với nhau.

### Performance
![image](https://github.com/FPT-NMTung/poi-excel/assets/63092029/458a20da-c7aa-42a0-8585-73682e91b9a5)
