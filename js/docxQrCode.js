"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var XmlTemplater = require("docxtemplater").XmlTemplater;

var QrCode = require("qrcode-reader");

module.exports = function () {
	function DocxQrCode(imageData, xmlTemplater) {
		var imgName = arguments.length <= 2 || arguments[2] === undefined ? "" : arguments[2];
		var num = arguments[3];
		var getDataFromString = arguments[4];

		_classCallCheck(this, DocxQrCode);

		this.xmlTemplater = xmlTemplater;
		this.imgName = imgName;
		this.num = num;
		this.getDataFromString = getDataFromString;
		this.callbacked = false;
		this.data = imageData;
		if (this.data === undefined) {
			throw new Error("data of qrcode can't be undefined");
		}
		this.ready = false;
		this.result = null;
	}

	_createClass(DocxQrCode, [{
		key: "decode",
		value: function decode(callback) {
			this.callback = callback;
			var self = this;
			this.qr = new QrCode();
			this.qr.callback = function () {
				self.ready = true;
				self.result = this.result;
				var testdoc = new XmlTemplater(this.result, { fileTypeConfig: self.xmlTemplater.fileTypeConfig,
					tags: self.xmlTemplater.tags,
					Tags: self.xmlTemplater.Tags,
					parser: self.xmlTemplater.parser
				});
				testdoc.render();
				self.result = testdoc.content;
				return self.searchImage();
			};
			return this.qr.decode({ width: this.data.width, height: this.data.height }, this.data.decoded);
		}
	}, {
		key: "searchImage",
		value: function searchImage() {
			var _this = this;

			var cb = function cb(_err) {
				var data = arguments.length <= 1 || arguments[1] === undefined ? _this.data.data : arguments[1];

				_this.data = data;
				return _this.callback(_this, _this.imgName, _this.num);
			};
			if (!(this.result != null)) {
				return cb();
			}
			return this.getDataFromString(this.result, cb);
		}
	}]);

	return DocxQrCode;
}();