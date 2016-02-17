
# docxtemplater-link-module
Hyperlink module for docxtemplater

[![Download count](https://img.shields.io/npm/dt/docxtemplater-link-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-link-module)

## Installation:
You will need docxtemplater v1: `npm install docxtemplater`

install this module: `npm install docxtemplater-link-module`

## Usage
Your docx should contain the text: {^link}

```js

var fs = require('fs');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
  .attachModule(linkModule)
  .load(content)
  .setData({
    link : {
      text : "link to Google",
      url : "http://google.com"
    }
    }).
    render();
var buffer = docx
  .getZip()
  .generate({type:"nodebuffer"});
 
fs.writeFile("test.docx", buffer);

```

## Testing 

You can test that everything works fine using the command `mocha`. This will also create 2 docx files under the root directory that you can open to check if the docx are correct
