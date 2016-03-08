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
		if (!(this.options.getImage != null)) {
			throw new Error("You should pass getImage");
		}
		if (!(this.options.getSize != null)) {
			throw new Error("You should pass getSize");
		}
		this.qrQueue = [];
		this.imageNumber = 1;
		this.relmap = {};
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
		value: function getNextImageName(imgId) {
			if(imgId){
				return "image_generated_" + imgId + ".png";
			}
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

			var startEnd = "<" + tagXml + "></" + tagXml + ">";
			var imgBuffer;
			var imgId = this.options.getID ? this.options.getID(tagValue, tag) : false;
			var rId = '';
			if(imgId && this.relmap[imgId]){
				rId = this.relmap[imgId];
			}else{
				try {
					imgBuffer = this.options.getImage(tagValue, tag);
				} catch (e) {
					return this.replaceBy(startEnd, tagXml);
				}
				var imageRels = this.imageRels;
				if (!imageRels) {
					imageRels = this.imageRels = this.imgManager.loadImageRels();
					if (!imageRels) {
						return;
					}
				}
				if(imgId){
					rId = imageRels.addImageRels(this.getNextImageName(imgId), imgBuffer);
					this.relmap[imgId] = rId;
				}else{
					rId = imageRels.addImageRels(this.getNextImageName(), imgBuffer);
				}
			}
			
			var sizePixel = this.options.getSize(imgBuffer, tagValue, tag);
			var size = [this.convertPixelsToEmus(sizePixel[0]), this.convertPixelsToEmus(sizePixel[1])];
			var newText = this.getImageXml(rId, size);
			var outsideElement = tagXml;
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
			var xml = [
				'<w:drawing>',
				'<wp:inline distT="0" distB="0" distL="0" distR="0">',
				'<wp:extent cx="' + size[0] + '" cy="' + size[1] + '"/>',
				'<wp:effectExtent l="0" t="0" r="0" b="0"/>',
				'<wp:docPr id="2" name="Image 2" descr="image"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>',
				'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1" descr="image"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId' + rId + '">',
				'<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + size[0] + '" cy="' + size[1] + '"/>',
				'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic>',
				'</wp:inline>',
				'</w:drawing>'
			];
			return xml.join('');
		}
	}]);

	return ImageModule;
}();

module.exports = ImageModule;