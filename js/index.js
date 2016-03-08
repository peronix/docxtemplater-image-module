"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var SubContent = require("docxtemplater").SubContent;
var ImgManager = require("./imgManager");
var ImgReplacer = require("./imgReplacer");

var ImageModule = function () {
	function ImageModule() {
		var options = arguments.length <= 0 || arguments[0] === undefined ? {} : arguments[0];

		_classCallCheck(this, ImageModule);

		this.options = options;
		if (!(this.options.centered != null)) {
			this.options.centered = false;
		}
		if (!(this.options.getImage != null)) {
			throw new Error("You should pass getImage");
		}
		if (!(this.options.getSize != null)) {
			throw new Error("You should pass getSize");
		}
		this.qrQueue = [];
		this.imageNumber = 1;
	}

	_createClass(ImageModule, [{
		key: "handleEvent",
		value: function handleEvent(event, eventData) {
			if (event === "rendering-file") {
				this.renderingFileName = eventData;
				var gen = this.manager.getInstance("gen");
				this.imgManager = new ImgManager(gen.zip, this.renderingFileName);
				this.imgManager.loadImageRels();
			}
			if (event === "rendered") {
				if (this.qrQueue.length === 0) {
					return this.finished();
				}
			}
		}
	}, {
		key: "get",
		value: function get(data) {
			if (data === "loopType") {
				var templaterState = this.manager.getInstance("templaterState");
				if (templaterState.textInsideTag[0] === "%") {
					return "image";
				}
			}
			return null;
		}
	}, {
		key: "getNextImageName",
		value: function getNextImageName() {
			var name = "image_generated_" + this.imageNumber + ".png";
			this.imageNumber++;
			return name;
		}
	}, {
		key: "replaceBy",
		value: function replaceBy(text, outsideElement) {
			var xmlTemplater = this.manager.getInstance("xmlTemplater");
			var templaterState = this.manager.getInstance("templaterState");
			var subContent = new SubContent(xmlTemplater.content);
			subContent = subContent.getInnerTag(templaterState);
			subContent = subContent.getOuterXml(outsideElement);
			return xmlTemplater.replaceXml(subContent, text);
		}
	}, {
		key: "convertPixelsToEmus",
		value: function convertPixelsToEmus(pixel) {
			return Math.round(pixel * 9525);
		}
	}, {
		key: "replaceTag",
		value: function replaceTag() {
			var scopeManager = this.manager.getInstance("scopeManager");
			var templaterState = this.manager.getInstance("templaterState");
			var xmlTemplater = this.manager.getInstance("xmlTemplater");
			var tagXml = xmlTemplater.fileTypeConfig.tagsXmlArray[0];

			var tag = templaterState.textInsideTag.substr(1);
			var tagValue = scopeManager.getValue(tag);

			if (tagValue == null) {
				return this.replaceBy(startEnd, tagXml);
			}

			var tagXmlParagraph = tagXml.substr(0, 1) + ":p";

			var startEnd = "<" + tagXml + "></" + tagXml + ">";
			var imgBuffer;
			try {
				imgBuffer = this.options.getImage(tagValue, tag);
			} catch (e) {
				return this.replaceBy(startEnd, tagXml);
			}
			var imageRels = this.imgManager.loadImageRels();
			if (!imageRels) {
				return;
			}
			var rId = imageRels.addImageRels(this.getNextImageName(), imgBuffer);
			var sizePixel = this.options.getSize(imgBuffer, tagValue, tag);
			var size = [this.convertPixelsToEmus(sizePixel[0]), this.convertPixelsToEmus(sizePixel[1])];
			var newText = this.options.centered ? this.getImageXmlCentered(rId, size) : this.getImageXml(rId, size);
			var outsideElement = this.options.centered ? tagXmlParagraph : tagXml;
			return this.replaceBy(newText, outsideElement);
		}
	}, {
		key: "replaceQr",
		value: function replaceQr() {
			var _this = this;

			var xmlTemplater = this.manager.getInstance("xmlTemplater");
			var imR = new ImgReplacer(xmlTemplater, this.imgManager);
			imR.getDataFromString = function (result, cb) {
				if (_this.options.getImageAsync != null) {
					return _this.options.getImageAsync(result, cb);
				}
				return cb(null, _this.options.getImage(result));
			};
			imR.pushQrQueue = function (num) {
				return _this.qrQueue.push(num);
			};
			imR.popQrQueue = function (num) {
				var found = _this.qrQueue.indexOf(num);
				if (found !== -1) {
					_this.qrQueue.splice(found, 1);
				} else {
					_this.on("error", new Error("qrqueue " + num + " is not in qrqueue"));
				}
				if (_this.qrQueue.length === 0) {
					return _this.finished();
				}
			};
			var num = parseInt(Math.random() * 10000, 10);
			imR.pushQrQueue("rendered-" + num);
			try {
				imR.findImages().replaceImages();
			} catch (e) {
				this.on("error", e);
			}
			var f = function f() {
				return imR.popQrQueue("rendered-" + num);
			};
			return setTimeout(f, 1);
		}
	}, {
		key: "finished",
		value: function finished() {}
	}, {
		key: "on",
		value: function on(event, data) {
			if (event === "error") {
				throw data;
			}
		}
	}, {
		key: "handle",
		value: function handle(type, data) {
			if (type === "replaceTag" && data === "image") {
				this.replaceTag();
			}
			if (type === "xmlRendered" && this.options.qrCode) {
				this.replaceQr();
			}
			return null;
		}
	}, {
		key: "getImageXml",
		value: function getImageXml(rId, size) {
			return "<w:drawing>\n  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n    <wp:extent cx=\"" + size[0] + "\" cy=\"" + size[1] + "\"/>\n    <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n    <wp:docPr id=\"2\" name=\"Image 2\" descr=\"image\"/>\n    <wp:cNvGraphicFramePr>\n      <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n    </wp:cNvGraphicFramePr>\n    <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n      <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n        <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n          <pic:nvPicPr>\n            <pic:cNvPr id=\"0\" name=\"Picture 1\" descr=\"image\"/>\n            <pic:cNvPicPr>\n              <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n            </pic:cNvPicPr>\n          </pic:nvPicPr>\n          <pic:blipFill>\n            <a:blip r:embed=\"rId" + rId + "\">\n              <a:extLst>\n                <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n                  <a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>\n                </a:ext>\n              </a:extLst>\n            </a:blip>\n            <a:srcRect/>\n            <a:stretch>\n              <a:fillRect/>\n            </a:stretch>\n          </pic:blipFill>\n          <pic:spPr bwMode=\"auto\">\n            <a:xfrm>\n              <a:off x=\"0\" y=\"0\"/>\n              <a:ext cx=\"" + size[0] + "\" cy=\"" + size[1] + "\"/>\n            </a:xfrm>\n            <a:prstGeom prst=\"rect\">\n              <a:avLst/>\n            </a:prstGeom>\n            <a:noFill/>\n            <a:ln>\n              <a:noFill/>\n            </a:ln>\n          </pic:spPr>\n        </pic:pic>\n      </a:graphicData>\n    </a:graphic>\n  </wp:inline>\n</w:drawing>\n\t\t";
		}
	}, {
		key: "getImageXmlCentered",
		value: function getImageXmlCentered(rId, size) {
			return "\t\t<w:p>\n\t\t  <w:pPr>\n\t\t\t<w:jc w:val=\"center\"/>\n\t\t  </w:pPr>\n\t\t  <w:r>\n\t\t\t<w:rPr/>\n\t\t\t<w:drawing>\n\t\t\t  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n\t\t\t\t<wp:extent cx=\"" + size[0] + "\" cy=\"" + size[1] + "\"/>\n\t\t\t\t<wp:docPr id=\"0\" name=\"Picture\" descr=\"\"/>\n\t\t\t\t<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n\t\t\t\t  <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n\t\t\t\t\t<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n\t\t\t\t\t  <pic:nvPicPr>\n\t\t\t\t\t\t<pic:cNvPr id=\"0\" name=\"Picture\" descr=\"\"/>\n\t\t\t\t\t\t<pic:cNvPicPr>\n\t\t\t\t\t\t  <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n\t\t\t\t\t\t</pic:cNvPicPr>\n\t\t\t\t\t  </pic:nvPicPr>\n\t\t\t\t\t  <pic:blipFill>\n\t\t\t\t\t\t<a:blip r:embed=\"rId" + rId + "\"/>\n\t\t\t\t\t\t<a:stretch>\n\t\t\t\t\t\t  <a:fillRect/>\n\t\t\t\t\t\t</a:stretch>\n\t\t\t\t\t  </pic:blipFill>\n\t\t\t\t\t  <pic:spPr bwMode=\"auto\">\n\t\t\t\t\t\t<a:xfrm>\n\t\t\t\t\t\t  <a:off x=\"0\" y=\"0\"/>\n\t\t\t\t\t\t  <a:ext cx=\"" + size[0] + "\" cy=\"" + size[1] + "\"/>\n\t\t\t\t\t\t</a:xfrm>\n\t\t\t\t\t\t<a:prstGeom prst=\"rect\">\n\t\t\t\t\t\t  <a:avLst/>\n\t\t\t\t\t\t</a:prstGeom>\n\t\t\t\t\t\t<a:noFill/>\n\t\t\t\t\t\t<a:ln w=\"9525\">\n\t\t\t\t\t\t  <a:noFill/>\n\t\t\t\t\t\t  <a:miter lim=\"800000\"/>\n\t\t\t\t\t\t  <a:headEnd/>\n\t\t\t\t\t\t  <a:tailEnd/>\n\t\t\t\t\t\t</a:ln>\n\t\t\t\t\t  </pic:spPr>\n\t\t\t\t\t</pic:pic>\n\t\t\t\t  </a:graphicData>\n\t\t\t\t</a:graphic>\n\t\t\t  </wp:inline>\n\t\t\t</w:drawing>\n\t\t  </w:r>\n\t\t</w:p>\n\t\t";
		}
	}]);

	return ImageModule;
}();

module.exports = ImageModule;