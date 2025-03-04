"use strict";

/* eslint-disable no-console */
var fs = require("fs");
var Docxtemplater = require("docxtemplater");
var path = require("path");
var JSZip = require("jszip");
var ImageModule = require("./index.js");
var testutils = require("docxtemplater/js/tests/utils");
var shouldBeSame = testutils.shouldBeSame;
var sizeOf = require("image-size");
var fileNames = ["imageExample.docx", "imageHeaderFooterExample.docx", "imageLoopExample.docx", "imageInlineExample.docx", "expectedInline.docx", "expectedNoImage.docx", "expectedHeaderFooter.docx", "expectedOneImage.docx", "expectedCentered.docx", "expectedLoopCentered.docx", "withoutRels.docx", "expectedWithoutRels.docx", "expectedBase64.docx", "tagImage.pptx", "expectedTagImage.pptx", "tagImageCentered.pptx", "expectedTagImageCentered.pptx", "expectedInlineResize.docx"];
beforeEach(function () {
  this.opts = {
    getImage: function getImage(tagValue) {
      return fs.readFileSync(tagValue);
    },
    getSize: function getSize() {
      return [150, 150];
    },
    centered: false
  };
  this.loadAndRender = function () {
    var file = testutils.createDoc(this.name);
    this.doc = new Docxtemplater();
    var inputZip = new JSZip(file.loadedContent);
    this.doc.loadZip(inputZip).setData(this.data);
    var imageModule = new ImageModule(this.opts);
    this.doc.attachModule(imageModule);
    this.renderedDoc = this.doc.render();
    var doc = this.renderedDoc;
    shouldBeSame({
      doc: doc,
      expectedName: this.expectedName
    });
  };
});
function testStart() {
  describe("{%image}", function () {
    it("should work with one image", function () {
      this.name = "imageExample.docx";
      this.expectedName = "expectedOneImage.docx";
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work without initial rels", function () {
      this.name = "withoutRels.docx";
      this.expectedName = "expectedWithoutRels.docx";
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with image tag == null", function () {
      this.name = "imageExample.docx";
      this.expectedName = "expectedNoImage.docx";
      this.data = {};
      this.loadAndRender();
    });
    it("should work with inline", function () {
      this.name = "imageInlineExample.docx";
      this.expectedName = "expectedInline.docx";
      this.data = {
        firefox: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with centering", function () {
      this.name = "imageExample.docx";
      this.expectedName = "expectedCentered.docx";
      this.opts.centered = true;
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with loops", function () {
      this.name = "imageLoopExample.docx";
      this.expectedName = "expectedLoopCentered.docx";
      this.opts.centered = true;
      this.data = {
        images: ["examples/image.png", "examples/image2.png"]
      };
      this.loadAndRender();
    });
    it("should work with image in header/footer", function () {
      this.name = "imageHeaderFooterExample.docx";
      this.expectedName = "expectedHeaderFooter.docx";
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with PPTX documents", function () {
      this.name = "tagImage.pptx";
      this.expectedName = "expectedTagImage.pptx";
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with PPTX documents centered", function () {
      this.name = "tagImageCentered.pptx";
      this.expectedName = "expectedTagImageCentered.pptx";
      this.data = {
        image: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with auto resize", function () {
      this.name = "imageInlineExample.docx";
      this.expectedName = "expectedInlineResize.docx";
      this.opts.getSize = function (img) {
        var sizeObj = sizeOf(img);
        return [sizeObj.width, sizeObj.height];
      };
      this.data = {
        firefox: "examples/image.png"
      };
      this.loadAndRender();
    });
    it("should work with base64 data", function () {
      var base64Image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH4QIJBywfp3IOswAAAB1pVFh0Q29tbWVudAAAAAAAQ3JlYXRlZCB3aXRoIEdJTVBkLmUHAAAAkUlEQVQY052PMQqDQBREZ1f/d1kUm3SxkeAF/FdIjpOcw2vpKcRWCwsRPMFPsaIQSIoMr5pXDGNUFd9j8TOn7kRW71fvO5HTq6qqtnWtzh20IqE3YXtL0zyKwAROQLQ5l/c9gHjfKK6wMZjADE6s49Dver4/smEAc2CuqgwAYI5jU9NcxhHEy60sni986H9+vwG1yDHfK1jitgAAAABJRU5ErkJggg==";
      this.name = "imageExample.docx";
      function base64DataURLToArrayBuffer(dataURL) {
        var stringBase64 = dataURL.replace(/^data:image\/(png|jpg);base64,/, "");
        var binaryString;
        if (typeof window !== "undefined") {
          binaryString = window.atob(stringBase64);
        } else {
          binaryString = new Buffer(stringBase64, "base64").toString("binary");
        }
        var len = binaryString.length;
        var bytes = new Uint8Array(len);
        for (var i = 0; i < len; i++) {
          var ascii = binaryString.charCodeAt(i);
          bytes[i] = ascii;
        }
        return bytes.buffer;
      }
      this.opts.getImage = function (image) {
        return image;
      };
      this.expectedName = "expectedBase64.docx";
      this.data = {
        image: base64DataURLToArrayBuffer(base64Image)
      };
      this.loadAndRender();
    });
  });
}
testutils.setExamplesDirectory(path.resolve(__dirname, "..", "examples"));
testutils.setStartFunction(testStart);
fileNames.forEach(function (filename) {
  testutils.loadFile(filename, testutils.loadDocument);
});
testutils.start();