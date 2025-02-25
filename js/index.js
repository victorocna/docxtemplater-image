"use strict";

function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function _defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, _toPropertyKey(o.key), o); } }
function _createClass(e, r, t) { return r && _defineProperties(e.prototype, r), t && _defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
var templates = require("./templates");
var DocUtils = require("docxtemplater").DocUtils;
var DOMParser = require("@xmldom/xmldom").DOMParser;
function isNaN(number) {
  return !(number === number);
}
var ImgManager = require("./imgManager");
var moduleName = "open-xml-templating/docxtemplater-image-module";
function getInnerDocx(_ref) {
  var part = _ref.part;
  return part;
}
function getInnerPptx(_ref2) {
  var part = _ref2.part,
    left = _ref2.left,
    right = _ref2.right,
    postparsed = _ref2.postparsed;
  var xmlString = postparsed.slice(left + 1, right).reduce(function (concat, item) {
    return concat + item.value;
  }, "");
  var xmlDoc = new DOMParser().parseFromString("<xml>" + xmlString + "</xml>");
  part.offset = {
    x: 0,
    y: 0
  };
  part.ext = {
    cx: 0,
    cy: 0
  };
  var offset = xmlDoc.getElementsByTagName("a:off");
  var ext = xmlDoc.getElementsByTagName("a:ext");
  if (ext.length > 0) {
    part.ext.cx = parseInt(ext[ext.length - 1].getAttribute("cx"), 10);
    part.ext.cy = parseInt(ext[ext.length - 1].getAttribute("cy"), 10);
  }
  if (offset.length > 0) {
    part.offset.x = parseInt(offset[offset.length - 1].getAttribute("x"), 10);
    part.offset.y = parseInt(offset[offset.length - 1].getAttribute("y"), 10);
  }
  return part;
}
var ImageModule = /*#__PURE__*/function () {
  function ImageModule(options) {
    _classCallCheck(this, ImageModule);
    this.name = "ImageModule";
    this.options = options || {};
    this.imgManagers = {};
    if (this.options.centered == null) {
      this.options.centered = false;
    }
    if (this.options.getImage == null) {
      throw new Error("You should pass getImage");
    }
    if (this.options.getSize == null) {
      throw new Error("You should pass getSize");
    }
    this.imageNumber = 1;
  }
  return _createClass(ImageModule, [{
    key: "optionsTransformer",
    value: function optionsTransformer(options, docxtemplater) {
      var relsFiles = docxtemplater.zip.file(/\.xml\.rels/).concat(docxtemplater.zip.file(/\[Content_Types\].xml/)).map(function (file) {
        return file.name;
      });
      this.fileTypeConfig = docxtemplater.fileTypeConfig;
      this.fileType = docxtemplater.fileType;
      this.zip = docxtemplater.zip;
      options.xmlFileNames = options.xmlFileNames.concat(relsFiles);
      return options;
    }
  }, {
    key: "set",
    value: function set(options) {
      if (options.zip) {
        this.zip = options.zip;
      }
      if (options.xmlDocuments) {
        this.xmlDocuments = options.xmlDocuments;
      }
    }
  }, {
    key: "parse",
    value: function parse(placeHolderContent) {
      var module = moduleName;
      var type = "placeholder";
      if (this.options.setParser) {
        return this.options.setParser(placeHolderContent);
      }
      if (placeHolderContent.substring(0, 2) === "%%") {
        return {
          type: type,
          value: placeHolderContent.substr(2),
          module: module,
          centered: true
        };
      }
      if (placeHolderContent.substring(0, 1) === "%") {
        return {
          type: type,
          value: placeHolderContent.substr(1),
          module: module,
          centered: false
        };
      }
      return null;
    }
  }, {
    key: "postparse",
    value: function postparse(parsed) {
      var expandTo;
      var getInner;
      if (this.fileType === "pptx") {
        expandTo = "p:sp";
        getInner = getInnerPptx;
      } else {
        expandTo = this.options.centered ? "w:p" : "w:t";
        getInner = getInnerDocx;
      }
      return DocUtils.traits.expandToOne(parsed, {
        moduleName: moduleName,
        getInner: getInner,
        expandTo: expandTo
      });
    }
  }, {
    key: "render",
    value: function render(part, options) {
      if (!part.type === "placeholder" || part.module !== moduleName) {
        return null;
      }
      var tagValue = options.scopeManager.getValue(part.value, {
        part: part
      });
      if (!tagValue) {
        return {
          value: this.fileTypeConfig.tagTextXml
        };
      } else if (_typeof(tagValue) === "object") {
        return this.getRenderedPart(part, tagValue.rId, tagValue.sizePixel);
      }
      var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments, this.fileType);
      var imgBuffer = this.options.getImage(tagValue, part.value);
      var rId = imgManager.addImageRels(this.getNextImageName(), imgBuffer);
      var sizePixel = this.options.getSize(imgBuffer, tagValue, part.value);
      return this.getRenderedPart(part, rId, sizePixel);
    }
  }, {
    key: "resolve",
    value: function resolve(part, options) {
      var _this = this;
      var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments, this.fileType);
      if (!part.type === "placeholder" || part.module !== moduleName) {
        return null;
      }
      var value = options.scopeManager.getValue(part.value, {
        part: part
      });
      if (!value) {
        return {
          value: this.fileTypeConfig.tagTextXml
        };
      }
      return new Promise(function (resolve) {
        var imgBuffer = _this.options.getImage(value, part.value);
        resolve(imgBuffer);
      }).then(function (imgBuffer) {
        var rId = imgManager.addImageRels(_this.getNextImageName(), imgBuffer);
        return new Promise(function (resolve) {
          var sizePixel = _this.options.getSize(imgBuffer, value, part.value);
          resolve(sizePixel);
        }).then(function (sizePixel) {
          return {
            rId: rId,
            sizePixel: sizePixel
          };
        });
      });
    }
  }, {
    key: "getRenderedPart",
    value: function getRenderedPart(part, rId, sizePixel) {
      if (isNaN(rId)) {
        throw new Error("rId is NaN, aborting");
      }
      var size = [DocUtils.convertPixelsToEmus(sizePixel[0]), DocUtils.convertPixelsToEmus(sizePixel[1])];
      var centered = this.options.centered || part.centered;
      var newText;
      if (this.fileType === "pptx") {
        newText = this.getRenderedPartPptx(part, rId, size, centered);
      } else {
        newText = this.getRenderedPartDocx(rId, size, centered);
      }
      return {
        value: newText
      };
    }
  }, {
    key: "getRenderedPartPptx",
    value: function getRenderedPartPptx(part, rId, size, centered) {
      var offset = {
        x: parseInt(part.offset.x, 10),
        y: parseInt(part.offset.y, 10)
      };
      var cellCX = parseInt(part.ext.cx, 10) || 1;
      var cellCY = parseInt(part.ext.cy, 10) || 1;
      var imgW = parseInt(size[0], 10) || 1;
      var imgH = parseInt(size[1], 10) || 1;
      if (centered) {
        offset.x = Math.round(offset.x + cellCX / 2 - imgW / 2);
        offset.y = Math.round(offset.y + cellCY / 2 - imgH / 2);
      }
      return templates.getPptxImageXml(rId, [imgW, imgH], offset);
    }
  }, {
    key: "getRenderedPartDocx",
    value: function getRenderedPartDocx(rId, size, centered) {
      return centered ? templates.getImageXmlCentered(rId, size) : templates.getImageXml(rId, size);
    }
  }, {
    key: "getNextImageName",
    value: function getNextImageName() {
      var name = "image_generated_".concat(this.imageNumber, ".png");
      this.imageNumber++;
      return name;
    }
  }]);
}();
module.exports = ImageModule;