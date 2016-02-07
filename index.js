var LinkManager, LinkModule, SubContent, fs;

SubContent = require('docxtemplater').SubContent;

LinkManager = require('./src/linkManager');

fs = require('fs');

LinkModule = (function() {

  LinkModule.prototype.name = 'link';
  function LinkModule(options){
    this.options = options | {};
  }
  LinkModule.prototype.handleEvent = function(event, eventData) {
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

  LinkModule.prototype.handle = function(type, data) {
    if (type === 'replaceTag' && data === this.name) {
      this.replaceTag();
    }
    return null;
  };

  LinkModule.prototype.get = function(data) {
    var templaterState;
    if (data === 'loopType') {
      templaterState = this.manager.getInstance('templaterState');
      if (templaterState.textInsideTag[0] === '^') {
        return this.name;
      }
    }
    return null;
  };

  LinkModule.prototype.replaceTag = function() {
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
    // console.log(tag, linkData, filename);
    linkId = this.linkManager.addLinkRels(filename, linkData.url);
    this.linkManager.addLinkStyle();
    tagXml = this.manager.getInstance('xmlTemplater').tagXml;
    newText = this.getLinkXml({
      linkID : linkId,
      linkText : linkData.text
    });
    // console.log("tag",tagXml);
    return this.replaceBy(newText, tagXml);
  };

  LinkModule.prototype.getLinkXml = function(_arg) {
    var linkId = _arg.linkID, linkText = _arg.linkText;
    return  "<w:hyperlink r:id=\"rId" + linkId + "\" w:history=\"1\"><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/><w:r w:rsidR=\"00052F25\" w:rsidRPr=\"00052F25\"><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>" + linkText + "</w:t></w:r></w:hyperlink>";
  }
  LinkModule.prototype.replaceBy = function(text, outsideElement) {
    var subContent, templaterState, xmlTemplater;
    xmlTemplater = this.manager.getInstance('xmlTemplater');
    templaterState = this.manager.getInstance('templaterState');
    subContent = new SubContent(xmlTemplater.content).getInnerTag(templaterState).getOuterXml(outsideElement);
    // console.log("subContent", subContent);
    return xmlTemplater.replaceXml(subContent, text);
  };

  LinkModule.prototype.finished = function() {};

  return LinkModule;
})();

module.exports = LinkModule;
