# xlsx-exporter
Export excel files to other format and check errors.

## Dependence
[openpyxl]: https://openpyxl.readthedocs.io/


## Usage
copy main.py and util directory to your excel file directory an run
python main.py or ./mail.py

## Excel
**file name**: same rule as python variables.

**sheet name**:
1. same rule as python variables.
2. unique in all files, same as json file name.
3. those start with "_" are comment sheet, will be ignored.

**head lines**:
4 lines at head:
1. field name
2. describle name
3. comment
4. type describe json

**data description**:
dataType: int, float, string, ref
ref: ref sheet(bookName:sheetName)
minValue: min value when dataType is int
maxValue: max value when dataType is int
minLen: min length when dataType is string
maxLen: max length when dataType is string 
regExp: regular expression to valid string
notNull: true or false
idType: id or combineid
isArray: true or false, if true, values can't be string and seperated by ","
allowdValues: values allowed array, if not empty, value must be one in the array