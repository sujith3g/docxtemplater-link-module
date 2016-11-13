var LinkManager, LinkModule, SubContent;

SubContent = require('docxtemplater').SubContent;

LinkManager = require('./src/linkManager');

LinkModule = (function () {
  LinkModule.prototype.name = 'link';
  function LinkModule (options) {
    this.options = options | {};
  }
  LinkModule.prototype.handleEvent = function (event, eventData) {
    var gen;
    if (event === 'rendering-file') {
      this.renderingFileName = eventData;
      gen = this.manager.getInstance('gen');
      this.linkManager = new LinkManager(gen.zip, this.renderingFileName);
      return this.linkManager.loadLinkRels();
    } else if (event === 'rendered') {
      return this.finished();
    }
  };

  LinkModule.prototype.handle = function (type, data) {
    if (type === 'replaceTag' && data === this.name) {
      this.replaceTag();
    }
    return null;
  };

  LinkModule.prototype.get = function (data) {
    var templaterState;
    if (data === 'loopType') {
      templaterState = this.manager.getInstance('templaterState');
      if (templaterState.textInsideTag[0] === '^') {
        return this.name;
      }
    }
    return null;
  };

  LinkModule.prototype.replaceTag = function () {
    var scopeManager, templaterState, gen, tag, linkData, linkRels, linkId, filename;
    scopeManager = this.manager.getInstance('scopeManager');
    templaterState = this.manager.getInstance('templaterState');
    gen = this.manager.getInstance('gen');
    tag = templaterState.textInsideTag.substr(1);
    linkData = scopeManager.getValueFromScope(tag);
    if (linkData == null) {
      return;
    }
    filename = tag + (this.linkManager.maxRid + 1);
    linkRels = this.linkManager.loadLinkRels();
    if (!linkRels) {
      return;
    }
    var url, text;
    if(typeof linkData === "string") {
        var emailRegex = /^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$/;
        url = emailRegex.test(linkData) ? "mailto:" + linkData : linkData;
        text = linkData;
    }
    else {
        url = linkData.url || linkData.URL;
        text = linkData.text || linkData.TEXT;
    }
    linkId = this.linkManager.addLinkRels(filename, url);
    this.linkManager.addLinkStyle();
    var xmlTemplater = this.manager.getInstance('xmlTemplater');
    tagXml = xmlTemplater.fileTypeConfig.tagsXmlArray[0];
    newText = this.getLinkXml({
      linkID : linkId,
      linkText : text,
      size: this.getLinkFontSize(xmlTemplater, templaterState.fullTextTag)
    });
    return this.replaceBy(newText, tagXml);
  };

  LinkModule.prototype.getLinkFontSize = function(xmlTemplater, fullTextTag) {
    var beforeTheTag = xmlTemplater.content.slice(0, xmlTemplater.content.indexOf(fullTextTag));
    beforeTheTag = beforeTheTag.slice(beforeTheTag.lastIndexOf("<a:endParaRPr"));
    var indexOfSz = beforeTheTag.indexOf("sz=\"");
    if (indexOfSz !== -1 && beforeTheTag.indexOf("extLst") === -1) {
      var szRegex = /sz="(\d+)"/;
      var size = szRegex.exec(beforeTheTag);
      return size[1];
    }
    return -1;
  }

  LinkModule.prototype.getLinkXml = function(_arg) {
    var linkId = _arg.linkID, linkText = _arg.linkText, size = _arg.size;
    if (this.linkManager.pptx) {
      return "</a:t></a:r><a:r><a:rPr " + (size !== -1 ? "sz=\"" + size + "\" " : "") + "lang=\"en-US\" dirty=\"0\" smtClean=\"0\"><a:hlinkClick r:id=\"rId" + linkId + "\"/></a:rPr><a:t>" + linkText + "</a:t></a:r><a:r><a:t>";
    }
    return  "</w:t></w:r><w:hyperlink r:id=\"rId" + linkId + "\" w:history=\"1\"><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/><w:r w:rsidR=\"00052F25\" w:rsidRPr=\"00052F25\"><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>" + linkText + "</w:t></w:r></w:hyperlink><w:r><w:t xml:space=\"preserve\">";
  }

  LinkModule.prototype.replaceBy = function (text, outsideElement) {
    var innerTag = this.getInnerTag().text;
    var subContent = this.linkManager.pptx ? this.getOuterXml(outsideElement) : this.addAttribute(outsideElement, 'xml:space', '"preserve"');
    return this.manager.getInstance('xmlTemplater').replaceXml(subContent, subContent.text.replace(innerTag, text));
  };

  LinkModule.prototype.addAttribute = function(outsideElement, attribute, value){
    var subContent = this.getOuterXml(outsideElement);
    if (subContent.text.indexOf(attribute) < 0) {
      var replaceRegex = '<' + outsideElement;
      var replaceWith = '<' + outsideElement + ' ' + attribute + '=' + value;
      subContent.text = subContent.text.replace(new RegExp(replaceRegex, 'g'), replaceWith);
    }
    return subContent;
  };

  LinkModule.prototype.getInnerTag = function () {
    var xmlTemplater = this.manager.getInstance('xmlTemplater');
    var templaterState = this.manager.getInstance('templaterState');
    return new SubContent(xmlTemplater.content).getInnerTag(templaterState);
  }

  LinkModule.prototype.getOuterXml = function (outsideElement) {
    return this.getInnerTag().getOuterXml(outsideElement);
  }

  LinkModule.prototype.finished = function() {};

  return LinkModule;
})();

module.exports = LinkModule;
