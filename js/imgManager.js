"use strict";

function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function _defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, _toPropertyKey(o.key), o); } }
function _createClass(e, r, t) { return r && _defineProperties(e.prototype, r), t && _defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
var DocUtils = require("./docUtils");
var extensionRegex = /[^.]+\.([^.]+)/;
var rels = {
  getPrefix: function getPrefix(fileType) {
    return fileType === "docx" ? "word" : "ppt";
  },
  getFileTypeName: function getFileTypeName(fileType) {
    return fileType === "docx" ? "document" : "presentation";
  },
  getRelsFileName: function getRelsFileName(fileName) {
    return fileName.replace(/^.*?([a-zA-Z0-9]+)\.xml$/, "$1") + ".xml.rels";
  },
  getRelsFilePath: function getRelsFilePath(fileName, fileType) {
    var relsFileName = rels.getRelsFileName(fileName);
    var prefix = fileType === "pptx" ? "ppt/slides" : "word";
    return "".concat(prefix, "/_rels/").concat(relsFileName);
  }
};
module.exports = /*#__PURE__*/function () {
  function ImgManager(zip, fileName, xmlDocuments, fileType) {
    _classCallCheck(this, ImgManager);
    this.fileName = fileName;
    this.prefix = rels.getPrefix(fileType);
    this.zip = zip;
    this.xmlDocuments = xmlDocuments;
    this.fileTypeName = rels.getFileTypeName(fileType);
    this.mediaPrefix = fileType === "pptx" ? "../media" : "media";
    var relsFilePath = rels.getRelsFilePath(fileName, fileType);
    this.relsDoc = xmlDocuments[relsFilePath] || this.createEmptyRelsDoc(xmlDocuments, relsFilePath);
  }
  return _createClass(ImgManager, [{
    key: "createEmptyRelsDoc",
    value: function createEmptyRelsDoc(xmlDocuments, relsFileName) {
      var mainRels = this.prefix + "/_rels/" + this.fileTypeName + ".xml.rels";
      var doc = xmlDocuments[mainRels];
      if (!doc) {
        var err = new Error("Could not copy from empty relsdoc");
        err.properties = {
          mainRels: mainRels,
          relsFileName: relsFileName,
          files: Object.keys(this.zip.files)
        };
        throw err;
      }
      var relsDoc = DocUtils.str2xml(DocUtils.xml2str(doc));
      var relationships = relsDoc.getElementsByTagName("Relationships")[0];
      var relationshipChilds = relationships.getElementsByTagName("Relationship");
      for (var i = 0, l = relationshipChilds.length; i < l; i++) {
        relationships.removeChild(relationshipChilds[i]);
      }
      xmlDocuments[relsFileName] = relsDoc;
      return relsDoc;
    }
  }, {
    key: "loadImageRels",
    value: function loadImageRels() {
      var iterable = this.relsDoc.getElementsByTagName("Relationship");
      return Array.prototype.reduce.call(iterable, function (max, relationship) {
        var id = relationship.getAttribute("Id");
        if (/^rId[0-9]+$/.test(id)) {
          return Math.max(max, parseInt(id.substr(3), 10));
        }
        return max;
      }, 0);
    }
    // Add an extension type in the [Content_Types.xml], is used if for example you want word to be able to read png files (for every extension you add you need a contentType)
  }, {
    key: "addExtensionRels",
    value: function addExtensionRels(contentType, extension) {
      var contentTypeDoc = this.xmlDocuments["[Content_Types].xml"];
      var defaultTags = contentTypeDoc.getElementsByTagName("Default");
      var extensionRegistered = Array.prototype.some.call(defaultTags, function (tag) {
        return tag.getAttribute("Extension") === extension;
      });
      if (extensionRegistered) {
        return;
      }
      var types = contentTypeDoc.getElementsByTagName("Types")[0];
      var newTag = contentTypeDoc.createElement("Default");
      newTag.namespaceURI = null;
      newTag.setAttribute("ContentType", contentType);
      newTag.setAttribute("Extension", extension);
      types.appendChild(newTag);
    }
    // Add an image and returns it's Rid
  }, {
    key: "addImageRels",
    value: function addImageRels(imageName, imageData, i) {
      if (i == null) {
        i = 0;
      }
      var realImageName = i === 0 ? imageName : imageName + "(".concat(i, ")");
      var imagePath = "".concat(this.prefix, "/media/").concat(realImageName);
      if (this.zip.files[imagePath] != null) {
        return this.addImageRels(imageName, imageData, i + 1);
      }
      var image = {
        name: imagePath,
        data: imageData,
        options: {
          binary: true
        }
      };
      this.zip.file(image.name, image.data, image.options);
      var extension = realImageName.replace(extensionRegex, "$1");
      this.addExtensionRels("image/".concat(extension), extension);
      var relationships = this.relsDoc.getElementsByTagName("Relationships")[0];
      var newTag = this.relsDoc.createElement("Relationship");
      newTag.namespaceURI = null;
      var maxRid = this.loadImageRels() + 1;
      newTag.setAttribute("Id", "rId".concat(maxRid));
      newTag.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
      newTag.setAttribute("Target", "".concat(this.mediaPrefix, "/").concat(realImageName));
      relationships.appendChild(newTag);
      return maxRid;
    }
  }]);
}();