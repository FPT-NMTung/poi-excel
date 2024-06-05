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

## Word

### Thông tin comment config

Tạo một comment vào nội dung bất kì trên file, có chứa nội dung là một JSON có dạng như sau:

```json
{
  "general": [
    {
      "name": "GENERAL_BOND_PRICE_VI",
      "data": "GENERAL_BOND_PRICE",
      "format": "number_char_vi"
    }
  ],
  "table": [
    {
      "name": "TABLE_GENERATE_0",
      "row": {
        "index": "ROW_NUM",
        "range": "2|2",
        "column": [
          {
            "name": "ROW_NUM",
            "data": "ROW_NUM",
            "format": "number"
          }
        ],
        "row": {}
      }
    }
  ]
}
```

* Thuộc tính `general` sẽ là các thông tin chung trong file, không phải thông tin trong bảng thông tin.
  * `name`: Tên của dữ liệu, nơi cần được điền trong file, trên file sẽ điền nội dung dạng `<#general.GENERAL_BOND_PRICE_VI>` để mapping với thông tin.
  * `data`: Data của dữ liệu tương ứng.
  * `format`: Có 4 dạng format sẽ đề cập bên dưới.

* Thuộc tính `table` là danh sách các bảng lưu.
  * `name`: Tên của bảng, tên này cần phải trùng với thông tin trong data với trường thông tin là `NAME_TABLE`.
  * `row`: Config của dòng trong bảng
    * `index`: Thông tin duy nhất theo từng dòng.
    * `range`: Khoảng của dòng (bắt đầu | kết thúc)
    * `column`: Config của từng cột trong dòng.
    * `row`: Đây là config của deep table
 
Xem mẫu file [tại đây](./sampleConfigDocx.json).

Để dánh dấu những bảng nào cần fill dữ liệu, tạo thêm một dòng dưới cùng của bảng và điền `<#TBG>` vào, cột nào không quan trọng. VD như sau:

![image](https://github.com/FPT-NMTung/poi-excel/assets/63092029/f8a89556-da47-44e0-8af5-3e3fbf3692b9)

### Format code

|   |Data|Format|
|---|---|---|
|`number`|123456.098|123,456.089|
|`string`|123456.098|123456.098|
|`number_char_vi`|123456|một trăm hai mươi ba nghìn bốn trăm năm mươi sáu|
|`number_char_Vi`|123456|Một trăm hai mươi ba nghìn bốn trăm năm mươi sáu|
|`number_char_VI`|123456|MỘT TRĂM HAI MƯƠI BA NGHÌN BỐN TRĂM NĂM MƯƠI SÁU|
|`number_char_en`|123456|one hundred twenty-three thousand four hundred fifty-six|
|`number_char_En`|123456|One hundred twenty-three thousand four hundred fifty-six|
|`number_char_EN`|123456|ONE HUNDRED TWENTY-THREE THOUSAND FOUR HUNDRED FIFTY-SIX|

### Performance
![image](https://github.com/FPT-NMTung/poi-excel/assets/63092029/b081e0ce-3c94-41cd-a799-5ca3b7e7fca4)
