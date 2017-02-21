var LinkManager, DocUtils;

DocUtils = require('./docUtils');

module.exports = LinkManager = (function() {
  function LinkManager(zip, fileName) {
    this.zip = zip;
    this.fileName = fileName;
    this.endFileName = this.fileName.replace(/^.*?([a-z0-9]+)\.xml$/, "$1");
    this.relsLoaded = false;
    this.pptx = this.fileName.indexOf('ppt') !== -1;
  }
  LinkManager.prototype.loadLinkRels = function() {

    /**
    		 * load file, save path
    		 * @param  {String} filePath path to current file
    		 * @return {Object}          file
     */
    var RidArray, content, file, loadFile, tag;
    loadFile = (function(_this) {
      return function(filePath) {
        _this.filePath = filePath;
        return _this.zip.files[_this.filePath];
      };
    })(this);
    file = loadFile("word/_rels/" + this.endFileName + ".xml.rels") || loadFile("word/_rels/document.xml.rels")
         || loadFile("ppt/slides/_rels/" + this.endFileName + ".xml.rels") || loadFile("ppt/_rels/presentation.xml.rels");
    if (file === void 0) {
      return;
    }
    content = DocUtils.decodeUtf8(file.asText());
    this.xmlDoc = DocUtils.Str2xml(content);
    RidArray = (function() {
      var _i, _len, _ref, _results;
      _ref = this.xmlDoc.getElementsByTagName('Relationship');
      _results = [];
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        tag = _ref[_i];
        _results.push(parseInt(tag.getAttribute("Id").substr(3)));
      }
      return _results;
    }).call(this);
    this.maxRid = DocUtils.maxArray(RidArray);
    this.linkRels = [];
    this.relsLoaded = true;
    return this;
  };
  LinkManager.prototype.addLinkRels = function(linkName, linkUrl) {
    if (!this.relsLoaded) {
      return;
    }
    this.maxRid++;
    this._addLinkRelationship(this.maxRid, linkName, linkUrl);
    this.zip.file(this.filePath, DocUtils.encodeUtf8(DocUtils.xml2Str(this.xmlDoc)), {});
    return this.maxRid;
  };
  LinkManager.prototype.addLinkStyle = function(){
    var file, loadFile, tag, stylePath;
    if (this.pptx) {
      stylePath = "ppt/presentation.xml";
    } else {
      stylePath = "word/styles.xml";
    }
    file = this.zip.files[stylePath];
    if (file === void 0) {
      return;
    }
    content = DocUtils.decodeUtf8(file.asText());
    styleXml = DocUtils.Str2xml(content);
    var isLinkStyle = false;
    if (this.pptx) {
      _ref = styleXml.getElementsByTagName('p:ext');
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        tag = _ref[_i];
        isLinkStyle = (tag.getAttribute("uri") == "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}");
        if (isLinkStyle) {
          break;
        }
      }
      if(!isLinkStyle) {
        var styles, newStyle, nameTag, basedOn, uiPrio, unHide, rsId;
        styles = styleXml.getElementsByTagName("p:extLst")[0];
        if(styles) {
          newStyleStr = "<p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\"><p15:sldGuideLst xmlns:p15=\"http://schemas.microsoft.com/office/powerpoint/2012/main\"/></p:ext>";
          newStyleXml = DocUtils.Str2xml(newStyleStr);
          styles.appendChild(newStyleXml);
        } else {
          var presentation = styleXml.getElementsByTagName("p:presentation")[0];
          newStyleStr = "<p:extLst><p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\"><p15:sldGuideLst xmlns:p15=\"http://schemas.microsoft.com/office/powerpoint/2012/main\"/></p:ext></p:extLst>";
          newStyleXml = DocUtils.Str2xml(newStyleStr);
          presentation.appendChild(newStyleXml);
        }
      }
    } else {
      _ref = styleXml.getElementsByTagName('w:style');
      // console.log("ref.length", _ref.length);
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        tag = _ref[_i];
        isLinkStyle = (tag.getAttribute("w:styleId") == "Hyperlink") ? true : isLinkStyle;
      }
      //add Hyperlink style if doesn't exist
      //<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="00052F25"/><w:rPr><w:color w:val="0000FF" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>
      // console.log("isLinkStyle", isLinkStyle);
      if(!isLinkStyle) {
        var styles, newStyle, nameTag, basedOn, uiPrio, unHide, rsId;
        styles = styleXml.getElementsByTagName("w:styles")[0];
        newStyleStr = "<w:style w:type=\"character\" w:styleId=\"Hyperlink\"><w:name w:val=\"Hyperlink\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:rsid w:val=\"00052F25\"/><w:rPr><w:color w:val=\"0000FF\" w:themeColor=\"hyperlink\"/><w:u w:val=\"single\"/></w:rPr></w:style>";
        newStyleXml = DocUtils.Str2xml(newStyleStr);
        styles.appendChild(newStyleXml);
      }
    }
    this.zip.file(stylePath, DocUtils.encodeUtf8(DocUtils.xml2Str(styleXml)), {});
    return this;
  };
  LinkManager.prototype._addLinkRelationship = function(id, name, url) {
    var newTag, relationships;
    relationships = this.xmlDoc.getElementsByTagName("Relationships")[0];
    newTag = this.xmlDoc.createElement('Relationship');
    newTag.namespaceURI = null;
    newTag.setAttribute('Id', "rId" + id);
    newTag.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
    newTag.setAttribute('Target', url);
    newTag.setAttribute('TargetMode', "External");
    return relationships.appendChild(newTag);
  };

  return LinkManager;
})();
