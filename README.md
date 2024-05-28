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
  "table": {
    "TABLE_GENERATE_0": {
      "name": "TABLE_GENERATE_0",
      "index": "ROW_NUM",
      "start": 1,
      "column": [
        {
          "name": "ROW_NUM",
          "format": "number"
        }
      ]
    }
  }
}
```

* Thuộc tính `general` sẽ là các thông tin chung trong file, không phải thông tin trong bảng thông tin.
  * `name`: Tên của dữ liệu, nơi cần được điền trong file, trên file sẽ điền nội dung dạng `<#general.GENERAL_BOND_PRICE_VI>` để mapping với thông tin.
  * `data`: Data của dữ liệu tương ứng.
  * `format`: Có 4 dạng format sẽ đề cập bên dưới.

* Thuộc tính `table` là danh sách các bảng lưu theo object để dàng truy xuất dữ liệu. Có key của object là tên của bảng, tên này cần phải trùng với thông tin trong data với trường thông tin là `NAME_TABLE`
  * `name`: Tên của bảng.
  * `index`: Thông tin duy nhất của từng dòng dữ liệu trong bảng.
  * `start`: Nơi bắt đầu fill thông tin và cũng là dòng chứa mẫu của bảng cần fill.
  * `column`: Danh sách config data các cột của bảng, **Tương tự như general nhưng không cần thuộc tính** `data`
 
Xem mẫu file [tại đây](./sampleConfigDocx.json).

Để dánh dấu những bảng nào cần fill dữ liệu, điền `<#TBG>` vào dòng template của bảng, cột nào không quan trọng. VD như sau:

![image](https://github.com/FPT-NMTung/poi-excel/assets/63092029/86714e90-f991-4e5c-a04e-fcd25fac7cf5)

### Format code

|   |Data|Format|
|---|---|---|
|`number`|123456.098|123,456.089|
|`string`|123456.098|123456.098|
|`number_char_vi`|123456|một trăm hai mươi ba nghìn bốn trăm năm mươi sáu|
|`number_char_en`|123456|one hundred twenty-three thousand four hundred fifty-six|
