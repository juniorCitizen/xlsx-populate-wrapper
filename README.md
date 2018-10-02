# xlsx-populate Wrapper

[xlsx-populate](https://www.npmjs.com/package/xlsx-populate) wrapper library for working with excel files

## Install

```bash
npm install --save xlsx-populate-wrapper
```

## Usage

```javascript
// 1. require the wrapper class
const ExcelWrapper = require("xlsx-populate-wrapper")
const path = require("path")

// 2. create an instance by passing in path to an excel file
const filePath = path.resolve(someFilePath)
new ExcelWrapper(filePath)
  .then(wb => {
    const ws = wb.worksheet("example")
    console.log(ws.data())
  })
  .catch(error => {
    throw error
  })
```

## Classes

<dl>
<dt><a href="#Workbook">Workbook</a></dt>
<dd></dd>
<dt><a href="#Worksheet">Worksheet</a></dt>
<dd></dd>
</dl>

## Functions

<dl>
<dt><a href="#sanitize">sanitize(object, properties)</a></dt>
<dd><p>compare the object&#39;s existing properties against a list of strings</p>
</dd>
</dl>

## Typedefs

<dl>
<dt><a href="#worksheetData">worksheetData</a> : <code>Object</code></dt>
<dd><p>custom data object, representing an excel worksheet data in array of arrays or array of objects</p>
</dd>
<dt><a href="#Workbook">Workbook</a> : <code>Class</code></dt>
<dd><p>represents an excel workbook</p>
</dd>
<dt><a href="#Worksheet">Worksheet</a> : <code>Class</code></dt>
<dd><p>represents an Excel worksheet</p>
</dd>
</dl>

<a name="Workbook"></a>

## Workbook

**Kind**: global class

- [Workbook](#Workbook)
  - [new Workbook(filePath)](#new_Workbook_new)
  - _instance_
    - [.init()](#Workbook+init) ⇒ [<code>Promise.&lt;Workbook&gt;</code>](#Workbook)
    - [.data()](#Workbook+data) ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)
    - [.update(worksheetName, dataset)](#Workbook+update)
    - [.worksheetNames()](#Workbook+worksheetNames) ⇒ <code>Array.&lt;string&gt;</code>
    - [.worksheet(worksheetName)](#Workbook+worksheet) ⇒ [<code>Worksheet</code>](#Worksheet)
  - _static_
    - [.convertJson(jsonData, headings)](#Workbook.convertJson) ⇒ [<code>worksheetData</code>](#worksheetData)

<a name="new_Workbook_new"></a>

### new Workbook(filePath)

instantiate a Workbook class object

| Param    | Type                | Description                    |
| -------- | ------------------- | ------------------------------ |
| filePath | <code>string</code> | absolute path to an excel file |

<a name="Workbook+init"></a>

### workbook.init() ⇒ [<code>Promise.&lt;Workbook&gt;</code>](#Workbook)

initialize a workbook

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Promise.&lt;Workbook&gt;</code>](#Workbook) - Workbook instance
<a name="Workbook+data"></a>

### workbook.data() ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)

get dataset of all existing worksheets

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) - array of worksheetData
<a name="Workbook+update"></a>

### workbook.update(worksheetName, dataset)

update and persist data of a particular worksheet

**Kind**: instance method of [<code>Workbook</code>](#Workbook)

| Param         | Type                                                                                            | Description                                                                                                  |
| ------------- | ----------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| worksheetName | <code>string</code>                                                                             | name of target worksheet to update                                                                           |
| dataset       | <code>Array.&lt;Object&gt;</code> \| [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) | can be an array of objects or an worksheetData object. Only 'jsonData' property is required to use the later |

<a name="Workbook+worksheetNames"></a>

### workbook.worksheetNames() ⇒ <code>Array.&lt;string&gt;</code>

get the worksheet names of existing worksheets

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: <code>Array.&lt;string&gt;</code> - - names of existing worksheets
<a name="Workbook+worksheet"></a>

### workbook.worksheet(worksheetName) ⇒ [<code>Worksheet</code>](#Worksheet)

get a specific worksheet by worksheet name

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Worksheet</code>](#Worksheet) - Worksheet instance

| Param         | Type                | Description              |
| ------------- | ------------------- | ------------------------ |
| worksheetName | <code>string</code> | name of target worksheet |

<a name="Workbook.convertJson"></a>

### Workbook.convertJson(jsonData, headings) ⇒ [<code>worksheetData</code>](#worksheetData)

convert an array of objects to a worksheetData object

**Kind**: static method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>worksheetData</code>](#worksheetData) - converted data

| Param    | Type                              | Default       | Description                                                                              |
| -------- | --------------------------------- | ------------- | ---------------------------------------------------------------------------------------- |
| jsonData | <code>Array.&lt;Object&gt;</code> |               | array of json objects                                                                    |
| headings | <code>Array.&lt;string&gt;</code> | <code></code> | (optional) column headings, props of first json object is used when headings are omitted |

<a name="Worksheet"></a>

## Worksheet

**Kind**: global class

- [Worksheet](#Worksheet)
  - [new Worksheet(worksheet)](#new_Worksheet_new)
  - [.name()](#Worksheet+name) ⇒ <code>string</code>
  - [.data()](#Worksheet+data) ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)
  - [.headings()](#Worksheet+headings) ⇒ <code>Array.&lt;string&gt;</code>
  - [.update(dataset)](#Worksheet+update)

<a name="new_Worksheet_new"></a>

### new Worksheet(worksheet)

instantiate a Worksheet class object

| Param     | Type                | Description                                 |
| --------- | ------------------- | ------------------------------------------- |
| worksheet | <code>Object</code> | Xlsx-populate.sheet([name of sheet]) object |

<a name="Worksheet+name"></a>

### worksheet.name() ⇒ <code>string</code>

get the worksheet name

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: <code>string</code> - name of worksheet
<a name="Worksheet+data"></a>

### worksheet.data() ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)

get data from the worksheet

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) - data from worksheet
<a name="Worksheet+headings"></a>

### worksheet.headings() ⇒ <code>Array.&lt;string&gt;</code>

get headings of all columns

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: <code>Array.&lt;string&gt;</code> - column headings
<a name="Worksheet+update"></a>

### worksheet.update(dataset)

update worksheet data (changes are not persisted with this method)

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)

| Param            | Type                                                      | Description                       |
| ---------------- | --------------------------------------------------------- | --------------------------------- |
| dataset          | [<code>worksheetData</code>](#worksheetData)              | data to update the worksheet with |
| dataset.headings | <code>Array.&lt;string&gt;</code>                         | heading of columns                |
| dataset.aoaData  | <code>Array.&lt;Array.&lt;(string\|number)&gt;&gt;</code> | data in array of arrays format    |

<a name="sanitize"></a>

## sanitize(object, properties)

compare the object's existing properties against a list of strings

**Kind**: global function

| Param      | Type                              | Description                 |
| ---------- | --------------------------------- | --------------------------- |
| object     | <code>Object</code>               | json object to be sanitized |
| properties | <code>Array.&lt;string&gt;</code> | property keys to keep       |

<a name="worksheetData"></a>

## worksheetData : <code>Object</code>

custom data object, representing an excel worksheet data in array of arrays or array of objects

**Kind**: global typedef
**Properties**

| Name           | Type                                                      | Description                 |
| -------------- | --------------------------------------------------------- | --------------------------- |
| headings       | <code>Array.&lt;string&gt;</code>                         | first value of every column |
| aoaData        | <code>Array.&lt;Array.&lt;(string\|number)&gt;&gt;</code> | array of array data         |
| jsonData       | <code>Array.&lt;Object&gt;</code>                         | array of data object        |
| jsonData.propA | <code>string</code> \| <code>number</code>                | value of propA              |
| jsonData.propB | <code>string</code> \| <code>number</code>                | value of propB              |
| jsonData.propN | <code>string</code> \| <code>number</code>                | value of propN, etc...      |

<a name="Workbook"></a>

## Workbook : <code>Class</code>

represents an excel workbook

**Kind**: global typedef
**Example**

```js
// 1. require the wrapper class
const ExcelWrapper = require("xlsx-populate-wrapper")
const path = require("path")

// 2. create an instance by passing in path to an excel file
const filePath = path.resolve(someFilePath)
new ExcelWrapper(filePath)
  .then(workbook => {
    const worksheet = workbook.worksheet("example")
    console.log(worksheet.data())
  })
  .catch(error => {
    throw error
  })
```

- [Workbook](#Workbook) : <code>Class</code>
  - [new Workbook(filePath)](#new_Workbook_new)
  - _instance_
    - [.init()](#Workbook+init) ⇒ [<code>Promise.&lt;Workbook&gt;</code>](#Workbook)
    - [.data()](#Workbook+data) ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)
    - [.update(worksheetName, dataset)](#Workbook+update)
    - [.worksheetNames()](#Workbook+worksheetNames) ⇒ <code>Array.&lt;string&gt;</code>
    - [.worksheet(worksheetName)](#Workbook+worksheet) ⇒ [<code>Worksheet</code>](#Worksheet)
  - _static_
    - [.convertJson(jsonData, headings)](#Workbook.convertJson) ⇒ [<code>worksheetData</code>](#worksheetData)

<a name="new_Workbook_new"></a>

### new Workbook(filePath)

instantiate a Workbook class object

| Param    | Type                | Description                    |
| -------- | ------------------- | ------------------------------ |
| filePath | <code>string</code> | absolute path to an excel file |

<a name="Workbook+init"></a>

### workbook.init() ⇒ [<code>Promise.&lt;Workbook&gt;</code>](#Workbook)

initialize a workbook

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Promise.&lt;Workbook&gt;</code>](#Workbook) - Workbook instance
<a name="Workbook+data"></a>

### workbook.data() ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)

get dataset of all existing worksheets

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) - array of worksheetData
<a name="Workbook+update"></a>

### workbook.update(worksheetName, dataset)

update and persist data of a particular worksheet

**Kind**: instance method of [<code>Workbook</code>](#Workbook)

| Param         | Type                                                                                            | Description                                                                                                  |
| ------------- | ----------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| worksheetName | <code>string</code>                                                                             | name of target worksheet to update                                                                           |
| dataset       | <code>Array.&lt;Object&gt;</code> \| [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) | can be an array of objects or an worksheetData object. Only 'jsonData' property is required to use the later |

<a name="Workbook+worksheetNames"></a>

### workbook.worksheetNames() ⇒ <code>Array.&lt;string&gt;</code>

get the worksheet names of existing worksheets

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: <code>Array.&lt;string&gt;</code> - - names of existing worksheets
<a name="Workbook+worksheet"></a>

### workbook.worksheet(worksheetName) ⇒ [<code>Worksheet</code>](#Worksheet)

get a specific worksheet by worksheet name

**Kind**: instance method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>Worksheet</code>](#Worksheet) - Worksheet instance

| Param         | Type                | Description              |
| ------------- | ------------------- | ------------------------ |
| worksheetName | <code>string</code> | name of target worksheet |

<a name="Workbook.convertJson"></a>

### Workbook.convertJson(jsonData, headings) ⇒ [<code>worksheetData</code>](#worksheetData)

convert an array of objects to a worksheetData object

**Kind**: static method of [<code>Workbook</code>](#Workbook)
**Returns**: [<code>worksheetData</code>](#worksheetData) - converted data

| Param    | Type                              | Default       | Description                                                                              |
| -------- | --------------------------------- | ------------- | ---------------------------------------------------------------------------------------- |
| jsonData | <code>Array.&lt;Object&gt;</code> |               | array of json objects                                                                    |
| headings | <code>Array.&lt;string&gt;</code> | <code></code> | (optional) column headings, props of first json object is used when headings are omitted |

<a name="Worksheet"></a>

## Worksheet : <code>Class</code>

represents an Excel worksheet

**Kind**: global typedef

- [Worksheet](#Worksheet) : <code>Class</code>
  - [new Worksheet(worksheet)](#new_Worksheet_new)
  - [.name()](#Worksheet+name) ⇒ <code>string</code>
  - [.data()](#Worksheet+data) ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)
  - [.headings()](#Worksheet+headings) ⇒ <code>Array.&lt;string&gt;</code>
  - [.update(dataset)](#Worksheet+update)

<a name="new_Worksheet_new"></a>

### new Worksheet(worksheet)

instantiate a Worksheet class object

| Param     | Type                | Description                                 |
| --------- | ------------------- | ------------------------------------------- |
| worksheet | <code>Object</code> | Xlsx-populate.sheet([name of sheet]) object |

<a name="Worksheet+name"></a>

### worksheet.name() ⇒ <code>string</code>

get the worksheet name

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: <code>string</code> - name of worksheet
<a name="Worksheet+data"></a>

### worksheet.data() ⇒ [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData)

get data from the worksheet

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: [<code>Array.&lt;worksheetData&gt;</code>](#worksheetData) - data from worksheet
<a name="Worksheet+headings"></a>

### worksheet.headings() ⇒ <code>Array.&lt;string&gt;</code>

get headings of all columns

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)
**Returns**: <code>Array.&lt;string&gt;</code> - column headings
<a name="Worksheet+update"></a>

### worksheet.update(dataset)

update worksheet data (changes are not persisted with this method)

**Kind**: instance method of [<code>Worksheet</code>](#Worksheet)

| Param            | Type                                                      | Description                       |
| ---------------- | --------------------------------------------------------- | --------------------------------- |
| dataset          | [<code>worksheetData</code>](#worksheetData)              | data to update the worksheet with |
| dataset.headings | <code>Array.&lt;string&gt;</code>                         | heading of columns                |
| dataset.aoaData  | <code>Array.&lt;Array.&lt;(string\|number)&gt;&gt;</code> | data in array of arrays format    |
