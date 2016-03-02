  var LinkModule, DocxGen, content, docX, expect, fileNames, fs, loadFile, name, _i, _len;

  fs = require('fs');

  DocxGen = require('docxtemplater');

  expect = require('chai').expect;

  fileNames = ['linkExample.docx', 'loopLinkExample.docx'];

  LinkModule = require('../index.js');

  docX = {};

  loadFile = function(name) {
    var xhrDoc;
    if (fs.readFileSync != null) {
      return fs.readFileSync(__dirname + "/../examples/" + name, "binary");
    }
    xhrDoc = new XMLHttpRequest();
    xhrDoc.open('GET', "../examples/" + name, false);
    if (xhrDoc.overrideMimeType) {
      xhrDoc.overrideMimeType('text/plain; charset=x-user-defined');
    }
    xhrDoc.send();
    return xhrDoc.response;
  };

  for (_i = 0, _len = fileNames.length; _i < _len; _i++) {
    name = fileNames[_i];
    content = loadFile(name);
    docX[name] = new DocxGen();
    docX[name].loadedContent = content;
  }

  describe('adding with {^ link} syntax', function() {
    var linkModule, out, zip;
    name = 'linkExample.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      link: {
        text : "Link to Google",
        url : "http://google.com"
      }
    }).render();
    zip = out.getZip();
    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects\" Target=\"stylesWithEffects.xml\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/></Relationships>");

    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('test.docx', zip.generate({
      type: "nodebuffer"
    }));
  });
  describe('adding with {^ link} syntax inside a loop', function() {
    var linkModule, out, zip;
    name = 'loopLinkExample.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      subsidiaries: [
        {
          title: "Euro Giant Ltd",
          link : {
            text : "link to Euro Giant",
            url : "http://google.com/?q=Euro%20Giant"
          }
        }, {
          title: "USA Giant Inc",
          link : {
            text : "link to USA Giant",
            url : "http://googl.com/?q=USA%20Giant"
          }
        }
      ]
    }).render();
    zip = out.getZip();
    it('should create two relationships in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects\" Target=\"stylesWithEffects.xml\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com/?q=Euro%20Giant\" TargetMode=\"External\"/><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://googl.com/?q=USA%20Giant\" TargetMode=\"External\"/></Relationships>");
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('looptest.docx', zip.generate({
      type: "nodebuffer"
    }));
  });

