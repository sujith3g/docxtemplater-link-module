var DOMParser, DocUtils, XMLSerializer,
  __slice = [].slice;

DOMParser = require('xmldom').DOMParser;

XMLSerializer = require('xmldom').XMLSerializer;

DocUtils = {};

DocUtils.xml2Str = function(xmlNode) {
  var a;
  a = new XMLSerializer();
  return a.serializeToString(xmlNode);
};

DocUtils.Str2xml = function(str, errorHandler) {
  var parser, xmlDoc;
  parser = new DOMParser({
    errorHandler: errorHandler
  });
  return xmlDoc = parser.parseFromString(str, "text/xml");
};

DocUtils.maxArray = function(a) {
  return Math.max.apply(null, a);
};

DocUtils.decodeUtf8 = function(s) {
  var e;
  try {
    if (s === void 0) {
      return void 0;
    }
    return decodeURIComponent(escape(DocUtils.convertSpaces(s)));
  } catch (_error) {
    e = _error;
    console.error(s);
    console.error('could not decode');
    throw new Error('end');
  }
};

DocUtils.encodeUtf8 = function(s) {
  return unescape(encodeURIComponent(s));
};

DocUtils.convertSpaces = function(s) {
  return s.replace(new RegExp(String.fromCharCode(160), "g"), " ");
};

DocUtils.pregMatchAll = function(regex, content) {

  /*regex is a string, content is the content. It returns an array of all matches with their offset, for example:
  	regex=la
  	content=lolalolilala
  	returns: [{0:'la',offset:2},{0:'la',offset:8},{0:'la',offset:10}]
   */
  var matchArray, replacer;
  if (!(typeof regex === 'object')) {
    regex = new RegExp(regex, 'g');
  }
  matchArray = [];
  replacer = function() {
    var match, offset, pn, string, _i;
    match = arguments[0], pn = 4 <= arguments.length ? __slice.call(arguments, 1, _i = arguments.length - 2) : (_i = 1, []), offset = arguments[_i++], string = arguments[_i++];
    pn.unshift(match);
    pn.offset = offset;
    return matchArray.push(pn);
  };
  content.replace(regex, replacer);
  return matchArray;
};

module.exports = DocUtils;
