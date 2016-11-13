
# docxtemplater-link-module
Hyperlink module for [docxtemplater](https://github.com/open-xml-templating/docxtemplater)

[![Download count](https://img.shields.io/npm/dt/docxtemplater-link-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-link-module)
[![ghit.me](https://ghit.me/badge.svg?repo=sujith3g/docxtemplater-link-module)](https://ghit.me/repo/sujith3g/docxtemplater-link-module)
[![Build Status](https://travis-ci.org/sujith3g/docxtemplater-link-module.svg?branch=master)](https://travis-ci.org/sujith3g/docxtemplater-link-module)

## Installation:
You will need docxtemplater v2.1.1: `npm install docxtemplater`

Install this module: `npm install docxtemplater-link-module`

## Usage: Text and URL

The example below will displays the following hyperlink:
> Lorem ipsum [dolor sit](http://google.com) amet.

Your docx should contain the text: `Lorem ipsum {^link} amet.`.

```js
var fs = require('fs');
var content = fs.readFileSync(__dirname + "/example-href.docx", "binary");
var DocxGen = require('docxtemplater');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
	.attachModule(linkModule)
	.load(content)
	.setData({
		link : {
			text : "dolor sit",
			url : "http://google.com"
		}
	}).
	render();
var buffer = docx
	.getZip()
	.generate({type:"nodebuffer"});
fs.writeFile("test.docx", buffer);
```

## Usage: URL only

The example below will displays the following hyperlink:
> Lorem ipsum [http://google.com](http://google.com) amet.

Your docx should contain the text: `Lorem ipsum {^link} amet.`.

```js
var fs = require('fs');
var content = fs.readFileSync(__dirname + "/example-href.docx", "binary");
var DocxGen = require('docxtemplater');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
	.attachModule(linkModule)
	.load(content)
	.setData({
		link : "http://google.com"
	}).
	render();
var buffer = docx
	.getZip()
	.generate({type:"nodebuffer"});
fs.writeFile("test.docx", buffer);
```

## Usage: Email address support

The example below will displays the following hyperlink:
> Lorem ipsum [john.smith@example.com](mailto:john.smith@example.com) amet.

Your docx should contain the text: `Lorem ipsum {^link} amet.`.

```js
var fs = require('fs');
var content = fs.readFileSync(__dirname + "/example-mailto.docx", "binary");
var DocxGen = require('docxtemplater');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
	.attachModule(linkModule)
	.load(content)
	.setData({
		link : "john.smith@example.com"
	}).
	render();
var buffer = docx
	.getZip()
	.generate({type:"nodebuffer"});
fs.writeFile("test.docx", buffer);
```

## Usage: Text and URL in powerpoint pptx

The example below will displays the following hyperlink powerpoint:
> Lorem ipsum [dolor sit](http://google.com) amet.

Your pptx should contain the text: `Lorem ipsum {^link} amet.`.

```js
var fs = require('fs');
var content = fs.readFileSync(__dirname + "/example-href.pptx", "binary");
var DocxGen = require('docxtemplater');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
	.attachModule(linkModule)
	.setOptions({ fileType : "pptx" })
	.load(content)
	.setData({
		link : {
			text : "dolor sit",
			url : "http://google.com"
		}
	}).
	render();
var buffer = docx
	.getZip()
	.generate({type:"nodebuffer"});
fs.writeFile("output-href.pptx", buffer);
```

## Usage: Email address support in powerpoint

The example below will displays the following hyperlink:
> Lorem ipsum [john.smith@example.com](mailto:john.smith@example.com) amet.

Your pptx should contain the text: `Lorem ipsum {^link} amet.`.

```js
var fs = require('fs');
var content = fs.readFileSync(__dirname + "/example-mailto.pptx", "binary");
var DocxGen = require('docxtemplater');
var LinkModule = require('docxtemplater-link-module');
var linkModule = new LinkModule();
 
var docx = new DocxGen()
	.attachModule(linkModule)
	.setOptions({ fileType : "pptx" })
	.load(content)
	.setData({
		link : "john.smith@example.com"
	}).
	render();
var buffer = docx
	.getZip()
	.generate({type:"nodebuffer"});
fs.writeFile("output-mailto.pptx", buffer);
```


## Testing 

You can test that everything works fine using the command `mocha`. This will also create 2 docx files under the root directory that you can open to check if the docx are correct

