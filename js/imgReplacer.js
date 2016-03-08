"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var DocUtils = require("./docUtils");
var DocxQrCode = require("./docxQrCode");
var PNG = require("png-js");
var base64encode = require("./base64").encode;

module.exports = function () {
	function ImgReplacer(xmlTemplater, imgManager) {
		_classCallCheck(this, ImgReplacer);

		this.xmlTemplater = xmlTemplater;
		this.imgManager = imgManager;
		this.imageSetter = this.imageSetter.bind(this);
		this.imgMatches = [];
		this.xmlTemplater.numQrCode = 0;
	}

	_createClass(ImgReplacer, [{
		key: "findImages",
		value: function findImages() {
			this.imgMatches = DocUtils.pregMatchAll(/<w:drawing[^>]*>.*?<a:blip.r:embed.*?<\/w:drawing>/g, this.xmlTemplater.content);
			return this;
		}
	}, {
		key: "replaceImages",
		value: function replaceImages() {
			this.qr = [];
			this.xmlTemplater.numQrCode += this.imgMatches.length;
			var iterable = this.imgMatches;
			for (var imgNum = 0, match; imgNum < iterable.length; imgNum++) {
				match = iterable[imgNum];
				this.replaceImage(match, imgNum);
			}
			return this;
		}
	}, {
		key: "imageSetter",
		value: function imageSetter(docxqrCode) {
			if (docxqrCode.callbacked === true) {
				return;
			}
			docxqrCode.callbacked = true;
			docxqrCode.xmlTemplater.numQrCode--;
			this.imgManager.setImage("word/media/" + docxqrCode.imgName, docxqrCode.data, { binary: true });
			return this.popQrQueue(this.imgManager.fileName + "-" + docxqrCode.num, false);
		}
	}, {
		key: "getXmlImg",
		value: function getXmlImg(match) {
			var baseDocument = "<?xml version=\"1.0\" ?>\n\t\t<w:document\n\t\tmc:Ignorable=\"w14 wp14\"\n\t\txmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"\n\t\t\txmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"\n\t\t\txmlns:o=\"urn:schemas-microsoft-com:office:office\"\n\t\txmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n\t\t\txmlns:v=\"urn:schemas-microsoft-com:vml\"\n\t\txmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"\n\t\t\txmlns:w10=\"urn:schemas-microsoft-com:office:word\"\n\t\txmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"\n\t\t\txmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\"\n\t\t\txmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"\n\t\t\txmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"\n\t\t\txmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\"\n\t\t\txmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"\n\t\t\txmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\"\n\t\t\txmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" + match[0] + "</w:document>\n\t\t\t";
			var f = function f(e) {
				if (e === "fatalError") {
					throw new Error("fatalError");
				}
			};
			return DocUtils.str2xml(baseDocument, f);
		}
	}, {
		key: "replaceImage",
		value: function replaceImage(match, imgNum) {
			var num = parseInt(Math.random() * 10000, 10);
			var xmlImg;
			try {
				xmlImg = this.getXmlImg(match);
			} catch (e) {
				return;
			}
			var tagrId = xmlImg.getElementsByTagName("a:blip")[0];
			if (tagrId === undefined) {
				throw new Error("tagRiD undefined !");
			}
			var rId = tagrId.getAttribute("r:embed");
			var tag = xmlImg.getElementsByTagName("wp:docPr")[0];
			if (tag === undefined) {
				throw new Error("tag undefined");
			}
			// if image is already a replacement then do nothing
			if (tag.getAttribute("name").substr(0, 6) === "Copie_") {
				return;
			}
			var imgName = this.imgManager.getImageName();
			this.pushQrQueue(this.imgManager.fileName + "-" + num, true);
			var newId = this.imgManager.addImageRels(imgName, "");
			this.xmlTemplater.imageId++;
			var oldFile = this.imgManager.getImageByRid(rId);
			this.imgManager.setImage(this.imgManager.getFullPath(imgName), oldFile.data, { binary: true });
			tag.setAttribute("name", "" + imgName);
			tagrId.setAttribute("r:embed", "rId" + newId);
			var imageTag = xmlImg.getElementsByTagName("w:drawing")[0];
			if (imageTag === undefined) {
				throw new Error("imageTag undefined");
			}
			var replacement = DocUtils.xml2Str(imageTag);
			this.xmlTemplater.content = this.xmlTemplater.content.replace(match[0], replacement);

			return this.decodeImage(oldFile, imgName, num, imgNum);
		}
	}, {
		key: "decodeImage",
		value: function decodeImage(oldFile, imgName, num, imgNum) {
			var _this = this;

			var mockedQrCode = { xmlTemplater: this.xmlTemplater, imgName: imgName, data: oldFile.asBinary(), num: num };
			if (!/\.png$/.test(oldFile.name)) {
				return this.imageSetter(mockedQrCode);
			}
			return function (imgName) {
				var base64 = base64encode(oldFile.asBinary());
				var binaryData = new Buffer(base64, "base64");
				var png = new PNG(binaryData);
				var finished = function finished(a) {
					png.decoded = a;
					try {
						_this.qr[imgNum] = new DocxQrCode(png, _this.xmlTemplater, imgName, num, _this.getDataFromString);
						return _this.qr[imgNum].decode(_this.imageSetter);
					} catch (e) {
						return _this.imageSetter(mockedQrCode);
					}
				};
				return png.decode(finished);
			}(imgName);
		}
	}]);

	return ImgReplacer;
}();