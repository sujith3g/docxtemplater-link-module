  var LinkModule, DocxGen, content, docX, expect, fileNames, fs, loadFile, name, _i, _len;

  fs = require('fs');

  DocxGen = require('docxtemplater');

  expect = require('chai').expect;

  fileNames = ['example-text+href.docx', 'example-text+href+loop.docx', 'example-href.docx', 'example-href+loop.docx', 'example-mailto.docx', 'example-mailto+loop.docx', 'example-href.pptx', 'example-mailto.pptx','example-mailto+loop.pptx'];

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

  describe('adding with {^ link} syntax (with text & href)', function() {
    var linkModule, out, zip;
    name = 'example-text+href.docx';
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

      return expect(relsFileContent).to.equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/></Relationships>");
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-text+href.docx', zip.generate({
      type: "nodebuffer"
    }));
  });
  describe('adding with {^ link} syntax inside a loop (with text & href)', function() {
    var linkModule, out, zip;
    name = 'example-text+href+loop.docx';
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
      expect(relsFileContent).to.contain('Target=\"http://googl.com/?q=USA%20Giant\"');
      expect(relsFileContent).to.contain('Target=\"http://google.com/?q=Euro%20Giant\"');
      expect(relsFileContent).to.contain('Relationship Id=\"rId6\"');
      return expect(relsFileContent).to.contain('Relationship Id=\"rId7\"');
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-text+href+loop.docx', zip.generate({
      type: "nodebuffer"
    }));
  });


  describe('adding with {^ link} syntax (href only)', function() {
    var linkModule, out, zip;
    name = 'example-href.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      link: "http://google.com"//The particularity of this test is that we are passing in a string rather than an object.
    }).render();
    zip = out.getZip();
    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      expect(relsFileContent.indexOf('<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/>')).to.be.above(-1);
      expect(relsFileContent.indexOf('<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/>')).to.be.above(-1);
      // expect(relsFileContent.indexOf('')).to.be.above(-1);
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-href.docx', zip.generate({
      type: "nodebuffer"
    }));
  });

  describe('adding with {^ link} syntax inside a loop (href only)', function() {
    var linkModule, out, zip;
    name = 'example-href+loop.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      subsidiaries: [
        {
          title: "Google",
          link: "http://google.com"
        }, {
          title: "Bing",
          link: "https://www.bing.com"
        }
      ]
    }).render();
    zip = out.getZip();
    it('should create two relationships in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      expect(relsFileContent.indexOf('<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/>')).to.be.above(-1);
      return expect(relsFileContent.indexOf('<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://www.bing.com\" TargetMode=\"External\"/>')).to.be.above(-1);
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-href+loop.docx', zip.generate({
      type: "nodebuffer"
    }));
  });




  describe('adding with {^ link} syntax (email address)', function() {
    var linkModule, out, zip;
    name = 'example-mailto.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      link: "john-smith@example.com"
    }).render();
    zip = out.getZip();
    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent.indexOf('<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:john-smith@example.com\" TargetMode=\"External\"/>')).to.be.above(-1);

    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-mailto.docx', zip.generate({
      type: "nodebuffer"
    }));
  });

  describe('adding with {^ link} syntax inside a loop (email address)', function() {
    var linkModule, out, zip;
    name = 'example-mailto+loop.docx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    out = docX[name].load(docX[name].loadedContent).setData({
      subsidiaries: [
        {
          title: "John Smith",
          link: "john-smith@example.com"
        }, {
          title: "Bill Knott",
          link: "bill.knott@example.com"
        }
      ]
    }).render();
    zip = out.getZip();
    it('should create two relationships in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['word/_rels/document.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      expect(relsFileContent.indexOf('<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:john-smith@example.com\" TargetMode=\"External\"/>')).to.be.above(-1);
      return expect(relsFileContent.indexOf('<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:bill.knott@example.com\" TargetMode=\"External\"/>')).to.be.above(-1);
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['word/styles.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("w:styleId=\"Hyperlink\"");
      return expect(styleFileContent).to.contain("w:val=\"Hyperlink\"");
    });
    return fs.writeFile('output-mailto+loop.docx', zip.generate({
      type: "nodebuffer"
    }));
  });


  describe('adding with {^ link} syntax in a powerpoint', function() {
    var linkModule, out, zip;
    name = 'example-href.pptx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    docX[name].setOptions({ fileType: 'pptx' });
    out = docX[name].load(docX[name].loadedContent).setData({
      description: "Testing the link feature",
      link: {
        TEXT : "Hakuna matata",
        URL : "http://google.com"
      }
    }).render();
    zip = out.getZip();
    it('should insert the label in the slide file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/slide2.xml'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.contain("<a:hlinkClick r:id=\"rId2\"/></a:rPr><a:t>Hakuna matata</a:t>");
    });

    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/_rels/slide2.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.contain("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://google.com\" TargetMode=\"External\"/>");
    });

    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['ppt/presentation.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("<p:extLst>");
      return expect(styleFileContent).to.contain("<p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\">");
    });
    return fs.writeFile('output-text+href.pptx', zip.generate({
      type: "nodebuffer"
    }));
  });

  describe('adding with {^ link} syntax (email address) in power-point', function() {
    var linkModule, out, zip;
    name = 'example-mailto.pptx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    docX[name].setOptions({ fileType: 'pptx' });
    out = docX[name].load(docX[name].loadedContent).setData({
      link: "email@example.com"
    }).render();
    zip = out.getZip();
    it('should insert the label in the slide file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/slide1.xml'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.contain("<a:hlinkClick r:id=\"rId2\"/></a:rPr><a:t>email@example.com</a:t>");
    });
    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/_rels/slide1.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      return expect(relsFileContent).to.contain("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:email@example.com\" TargetMode=\"External\"/>");
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['ppt/presentation.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("<p:extLst>");
      return expect(styleFileContent).to.contain("<p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\">");
    });
    return fs.writeFile('output-mailto.pptx', zip.generate({
      type: "nodebuffer"
    }));
  });

  describe('adding with {^ link} syntax inside a loop (email address)', function() {
    var linkModule, out, zip;
    name = 'example-mailto+loop.pptx';
    linkModule = new LinkModule();
    docX[name].attachModule(linkModule);
    docX[name].setOptions({ fileType: 'pptx' });
    out = docX[name].load(docX[name].loadedContent).setData({
      subsidiaries: [
        {
          title: "John Smith",
          link: "john-smith@example.com"
        }, {
          title: "Bill Knott",
          link: "bill.knott@example.com"
        }
      ]
    }).render();
    zip = out.getZip();
    it('should insert the label in the slide file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/slide1.xml'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();

      expect(relsFileContent).to.contain("<a:t>John Smith Sales for Q1</a:t>");
      expect(relsFileContent).to.contain("<a:t>Bill Knott Sales for Q1</a:t>");
      expect(relsFileContent).to.contain("<a:hlinkClick r:id=\"rId2\"/></a:rPr><a:t>john-smith@example.com</a:t>");
      return expect(relsFileContent).to.contain("<a:hlinkClick r:id=\"rId3\"/></a:rPr><a:t>bill.knott@example.com</a:t>");
    });
    it('should create relationship in rels file', function() {
      var relsFile, relsFileContent;
      relsFile = zip.files['ppt/slides/_rels/slide1.xml.rels'];
      expect(relsFile != null).to.equal(true);
      relsFileContent = relsFile.asText();
      expect(relsFileContent).to.contain("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:john-smith@example.com\" TargetMode=\"External\"/>");
      return expect(relsFileContent).to.contain("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"mailto:bill.knott@example.com\" TargetMode=\"External\"/>");
    });
    it('should add HyperlinkStyle if it is not present', function() {
      var styleFile, styleFileContent;
      styleFile = zip.files['ppt/presentation.xml'];
      expect(styleFile != null).to.equal(true);
      styleFileContent = styleFile.asText();
      expect(styleFileContent).to.contain("<p:extLst>");
      return expect(styleFileContent).to.contain("<p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\">");
    });
    return fs.writeFile('output-mailto+loop.pptx', zip.generate({
      type: "nodebuffer"
    }));
  });

