/*
* @license
* @apollo-design/docx-preview <https://github.com/apollocmh/docx-preview>
* Released under Apache License 2.0 — Copyright apollocmh
* Based on docxjs by Volodymyr Baydalka <https://github.com/VolodymyrBaydalka/docxjs>
*/
(function(global, factory) {
	typeof exports === "object" && typeof module !== "undefined" ? factory(exports, require("jszip")) : typeof define === "function" && define.amd ? define(["exports", "jszip"], factory) : (global = typeof globalThis !== "undefined" ? globalThis : global || self, factory(global.docx = {}, global.JSZip));
})(this, function(exports, jszip) {
	Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
	//#region \0rolldown/runtime.js
	var __create = Object.create;
	var __defProp = Object.defineProperty;
	var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
	var __getOwnPropNames = Object.getOwnPropertyNames;
	var __getProtoOf = Object.getPrototypeOf;
	var __hasOwnProp = Object.prototype.hasOwnProperty;
	var __copyProps = (to, from, except, desc) => {
		if (from && typeof from === "object" || typeof from === "function") for (var keys = __getOwnPropNames(from), i = 0, n = keys.length, key; i < n; i++) {
			key = keys[i];
			if (!__hasOwnProp.call(to, key) && key !== except) __defProp(to, key, {
				get: ((k) => from[k]).bind(null, key),
				enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable
			});
		}
		return to;
	};
	var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(isNodeMode || !mod || !mod.__esModule || !__hasOwnProp.call(mod, "default") ? __defProp(target, "default", {
		value: mod,
		enumerable: true
	}) : target, mod));
	//#endregion
	jszip = __toESM(jszip, 1);
	//#region src/common/relationship.ts
	var RelationshipTypes = /* @__PURE__ */ function(RelationshipTypes) {
		RelationshipTypes["OfficeDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		RelationshipTypes["FontTable"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
		RelationshipTypes["Image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
		RelationshipTypes["Numbering"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
		RelationshipTypes["Styles"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
		RelationshipTypes["StylesWithEffects"] = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
		RelationshipTypes["Theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
		RelationshipTypes["Settings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
		RelationshipTypes["WebSettings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
		RelationshipTypes["Hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
		RelationshipTypes["Footnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
		RelationshipTypes["Endnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
		RelationshipTypes["Footer"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
		RelationshipTypes["Header"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
		RelationshipTypes["ExtendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
		RelationshipTypes["CoreProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
		RelationshipTypes["CustomProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties";
		RelationshipTypes["Comments"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
		RelationshipTypes["CommentsExtended"] = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
		RelationshipTypes["AltChunk"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
		return RelationshipTypes;
	}({});
	function parseRelationships(root, xml) {
		return xml.elements(root).map((e) => ({
			id: xml.attr(e, "Id"),
			type: xml.attr(e, "Type"),
			target: xml.attr(e, "Target"),
			targetMode: xml.attr(e, "TargetMode")
		}));
	}
	//#endregion
	//#region src/utils.ts
	function escapeClassName(className) {
		return className === null || className === void 0 ? void 0 : className.replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase();
	}
	function encloseFontFamily(fontFamily) {
		return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
	}
	function splitPath(path) {
		let si = path.lastIndexOf("/") + 1;
		return [si == 0 ? "" : path.substring(0, si), si == 0 ? path : path.substring(si)];
	}
	function resolvePath(path, base) {
		try {
			return new URL(path, "http://docx/" + base).toString().substring(12);
		} catch {
			return `${base}${path}`;
		}
	}
	function keyBy(array, by) {
		return array.reduce((a, x) => {
			a[by(x)] = x;
			return a;
		}, {});
	}
	function blobToBase64(blob) {
		return new Promise((resolve, reject) => {
			const reader = new FileReader();
			reader.onloadend = () => resolve(reader.result);
			reader.onerror = () => reject();
			reader.readAsDataURL(blob);
		});
	}
	function isObject(item) {
		return item && typeof item === "object" && !Array.isArray(item);
	}
	function isString(item) {
		return typeof item === "string" || item instanceof String;
	}
	function mergeDeep(target, ...sources) {
		if (!sources.length) return target;
		const source = sources.shift();
		if (isObject(target) && isObject(source)) for (const key in source) if (isObject(source[key])) {
			var _target$key;
			mergeDeep((_target$key = target[key]) !== null && _target$key !== void 0 ? _target$key : target[key] = {}, source[key]);
		} else target[key] = source[key];
		return mergeDeep(target, ...sources);
	}
	function asArray(val) {
		return Array.isArray(val) ? val : [val];
	}
	function clamp(val, min, max) {
		return min > val ? min : max < val ? max : val;
	}
	//#endregion
	//#region src/document/common.ts
	var ns$1 = {
		wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
		picture: "http://schemas.openxmlformats.org/drawingml/2006/picture",
		compatibility: "http://schemas.openxmlformats.org/markup-compatibility/2006",
		math: "http://schemas.openxmlformats.org/officeDocument/2006/math"
	};
	var LengthUsage = {
		Dxa: {
			mul: .05,
			unit: "pt"
		},
		Emu: {
			mul: 1 / 12700,
			unit: "pt"
		},
		FontSize: {
			mul: .5,
			unit: "pt"
		},
		Border: {
			mul: .125,
			unit: "pt",
			min: .25,
			max: 12
		},
		Point: {
			mul: 1,
			unit: "pt"
		},
		Percent: {
			mul: .02,
			unit: "%"
		},
		LineHeight: {
			mul: 1 / 240,
			unit: ""
		},
		VmlEmu: {
			mul: 1 / 12700,
			unit: ""
		}
	};
	function convertLength(val, usage = LengthUsage.Dxa) {
		if (val == null || /.+(p[xt]|[%])$/.test(val)) return val;
		var num = parseInt(val) * usage.mul;
		if (usage.min && usage.max) num = clamp(num, usage.min, usage.max);
		return `${num.toFixed(2)}${usage.unit}`;
	}
	function convertBoolean(v, defaultValue = false) {
		switch (v) {
			case "1": return true;
			case "0": return false;
			case "on": return true;
			case "off": return false;
			case "true": return true;
			case "false": return false;
			default: return defaultValue;
		}
	}
	function parseCommonProperty(elem, props, xml) {
		if (elem.namespaceURI != ns$1.wordml) return false;
		switch (elem.localName) {
			case "color":
				props.color = xml.attr(elem, "val");
				break;
			case "sz":
				props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
				break;
			default: return false;
		}
		return true;
	}
	//#endregion
	//#region src/parser/xml-parser.ts
	function parseXmlString(xmlString, trimXmlDeclaration = false) {
		if (trimXmlDeclaration) xmlString = xmlString.replace(/<[?].*[?]>/, "");
		xmlString = removeUTF8BOM(xmlString);
		const result = new DOMParser().parseFromString(xmlString, "application/xml");
		const errorText = hasXmlParserError(result);
		if (errorText) throw new Error(errorText);
		return result;
	}
	function hasXmlParserError(doc) {
		var _doc$getElementsByTag;
		return (_doc$getElementsByTag = doc.getElementsByTagName("parsererror")[0]) === null || _doc$getElementsByTag === void 0 ? void 0 : _doc$getElementsByTag.textContent;
	}
	function removeUTF8BOM(data) {
		return data.charCodeAt(0) === 65279 ? data.substring(1) : data;
	}
	function serializeXmlString(elem) {
		return new XMLSerializer().serializeToString(elem);
	}
	var XmlParser = class {
		elements(elem, localName = null) {
			const result = [];
			for (let i = 0, l = elem.childNodes.length; i < l; i++) {
				let c = elem.childNodes.item(i);
				if (c.nodeType == Node.ELEMENT_NODE && (localName == null || c.localName == localName)) result.push(c);
			}
			return result;
		}
		element(elem, localName) {
			for (let i = 0, l = elem.childNodes.length; i < l; i++) {
				let c = elem.childNodes.item(i);
				if (c.nodeType == 1 && c.localName == localName) return c;
			}
			return null;
		}
		elementAttr(elem, localName, attrLocalName) {
			var el = this.element(elem, localName);
			return el ? this.attr(el, attrLocalName) : void 0;
		}
		attrs(elem) {
			return Array.from(elem.attributes);
		}
		attr(elem, localName) {
			for (let i = 0, l = elem.attributes.length; i < l; i++) {
				let a = elem.attributes.item(i);
				if (a.localName == localName) return a.value;
			}
			return null;
		}
		intAttr(node, attrName, defaultValue = null) {
			var val = this.attr(node, attrName);
			return val ? parseInt(val) : defaultValue;
		}
		hexAttr(node, attrName, defaultValue = null) {
			var val = this.attr(node, attrName);
			return val ? parseInt(val, 16) : defaultValue;
		}
		floatAttr(node, attrName, defaultValue = null) {
			var val = this.attr(node, attrName);
			return val ? parseFloat(val) : defaultValue;
		}
		boolAttr(node, attrName, defaultValue = null) {
			return convertBoolean(this.attr(node, attrName), defaultValue);
		}
		lengthAttr(node, attrName, usage = LengthUsage.Dxa) {
			return convertLength(this.attr(node, attrName), usage);
		}
	};
	var globalXmlParser = new XmlParser();
	//#endregion
	//#region src/common/part.ts
	var Part = class {
		constructor(_package, path) {
			this._package = _package;
			this.path = path;
		}
		async load() {
			this.rels = await this._package.loadRelationships(this.path);
			const xmlText = await this._package.load(this.path);
			const xmlDoc = this._package.parseXmlDocument(xmlText);
			if (this._package.options.keepOrigin) this._xmlDocument = xmlDoc;
			this.parseXml(xmlDoc.firstElementChild);
		}
		save() {
			this._package.update(this.path, serializeXmlString(this._xmlDocument));
		}
		parseXml(root) {}
	};
	//#endregion
	//#region src/font-table/fonts.ts
	var embedFontTypeMap = {
		embedRegular: "regular",
		embedBold: "bold",
		embedItalic: "italic",
		embedBoldItalic: "boldItalic"
	};
	function parseFonts(root, xml) {
		return xml.elements(root).map((el) => parseFont(el, xml));
	}
	function parseFont(elem, xml) {
		let result = {
			name: xml.attr(elem, "name"),
			embedFontRefs: []
		};
		for (let el of xml.elements(elem)) switch (el.localName) {
			case "family":
				result.family = xml.attr(el, "val");
				break;
			case "altName":
				result.altName = xml.attr(el, "val");
				break;
			case "embedRegular":
			case "embedBold":
			case "embedItalic":
			case "embedBoldItalic": result.embedFontRefs.push(parseEmbedFontRef(el, xml));
		}
		return result;
	}
	function parseEmbedFontRef(elem, xml) {
		return {
			id: xml.attr(elem, "id"),
			key: xml.attr(elem, "fontKey"),
			type: embedFontTypeMap[elem.localName]
		};
	}
	//#endregion
	//#region src/font-table/font-table.ts
	var FontTablePart = class extends Part {
		parseXml(root) {
			this.fonts = parseFonts(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/common/content-types.ts
	function parseContentTypes(root, xml) {
		return xml.elements(root).map((e) => ({
			extension: xml.attr(e, "Extension"),
			partName: xml.attr(e, "PartName"),
			contentType: xml.attr(e, "ContentType")
		}));
	}
	//#endregion
	//#region src/common/open-xml-package.ts
	var OpenXmlPackage = class OpenXmlPackage {
		constructor(_zip, options) {
			this._zip = _zip;
			this.options = options;
			this.xmlParser = new XmlParser();
		}
		get(path) {
			var _this$_zip$files$p;
			const p = normalizePath(path);
			return (_this$_zip$files$p = this._zip.files[p]) !== null && _this$_zip$files$p !== void 0 ? _this$_zip$files$p : this._zip.files[p.replace(/\//g, "\\")];
		}
		update(path, content) {
			this._zip.file(path, content);
		}
		static async load(input, options) {
			const zip = await jszip.default.loadAsync(input);
			return new OpenXmlPackage(zip, options);
		}
		save(type = "blob") {
			return this._zip.generateAsync({ type });
		}
		load(path, type = "string") {
			var _this$get$async, _this$get;
			return (_this$get$async = (_this$get = this.get(path)) === null || _this$get === void 0 ? void 0 : _this$get.async(type)) !== null && _this$get$async !== void 0 ? _this$get$async : Promise.resolve(null);
		}
		async loadRelationships(path = null) {
			let relsPath = `_rels/.rels`;
			if (path != null) {
				const [f, fn] = splitPath(path);
				relsPath = `${f}_rels/${fn}.rels`;
			}
			const txt = await this.load(relsPath);
			return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
		}
		async loadContentTypes() {
			const txt = await this.load("[Content_Types].xml");
			return txt ? parseContentTypes(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : [];
		}
		/** @internal */
		parseXmlDocument(txt) {
			return parseXmlString(txt, this.options.trimXmlDeclaration);
		}
	};
	function normalizePath(path) {
		return path.startsWith("/") ? path.substr(1) : path;
	}
	//#endregion
	//#region src/document/document-part.ts
	var DocumentPart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
		parseXml(root) {
			this.body = this._documentParser.parseDocumentFile(root);
		}
	};
	//#endregion
	//#region src/document/border.ts
	function parseBorder(elem, xml) {
		return {
			type: xml.attr(elem, "val"),
			color: xml.attr(elem, "color"),
			size: xml.lengthAttr(elem, "sz", LengthUsage.Border),
			offset: xml.lengthAttr(elem, "space", LengthUsage.Point),
			frame: xml.boolAttr(elem, "frame"),
			shadow: xml.boolAttr(elem, "shadow")
		};
	}
	function parseBorders(elem, xml) {
		var result = {};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "left":
				result.left = parseBorder(e, xml);
				break;
			case "top":
				result.top = parseBorder(e, xml);
				break;
			case "right":
				result.right = parseBorder(e, xml);
				break;
			case "bottom": result.bottom = parseBorder(e, xml);
		}
		return result;
	}
	//#endregion
	//#region src/document/section.ts
	function parseSectionProperties(elem, xml = globalXmlParser) {
		var section = {};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "pgSz":
				section.pageSize = {
					width: xml.lengthAttr(e, "w"),
					height: xml.lengthAttr(e, "h"),
					orientation: xml.attr(e, "orient")
				};
				break;
			case "type":
				section.type = xml.attr(e, "val");
				break;
			case "pgMar":
				section.pageMargins = {
					left: xml.lengthAttr(e, "left"),
					right: xml.lengthAttr(e, "right"),
					top: xml.lengthAttr(e, "top"),
					bottom: xml.lengthAttr(e, "bottom"),
					header: xml.lengthAttr(e, "header"),
					footer: xml.lengthAttr(e, "footer"),
					gutter: xml.lengthAttr(e, "gutter")
				};
				break;
			case "cols":
				section.columns = parseColumns(e, xml);
				break;
			case "headerReference":
				var _section$headerRefs;
				((_section$headerRefs = section.headerRefs) !== null && _section$headerRefs !== void 0 ? _section$headerRefs : section.headerRefs = []).push(parseFooterHeaderReference(e, xml));
				break;
			case "footerReference":
				var _section$footerRefs;
				((_section$footerRefs = section.footerRefs) !== null && _section$footerRefs !== void 0 ? _section$footerRefs : section.footerRefs = []).push(parseFooterHeaderReference(e, xml));
				break;
			case "titlePg":
				section.titlePage = xml.boolAttr(e, "val", true);
				break;
			case "pgBorders":
				section.pageBorders = parseBorders(e, xml);
				break;
			case "pgNumType": section.pageNumber = parsePageNumber(e, xml);
		}
		return section;
	}
	function parseColumns(elem, xml) {
		return {
			numberOfColumns: xml.intAttr(elem, "num"),
			space: xml.lengthAttr(elem, "space"),
			separator: xml.boolAttr(elem, "sep"),
			equalWidth: xml.boolAttr(elem, "equalWidth", true),
			columns: xml.elements(elem, "col").map((e) => ({
				width: xml.lengthAttr(e, "w"),
				space: xml.lengthAttr(e, "space")
			}))
		};
	}
	function parsePageNumber(elem, xml) {
		return {
			chapSep: xml.attr(elem, "chapSep"),
			chapStyle: xml.attr(elem, "chapStyle"),
			format: xml.attr(elem, "fmt"),
			start: xml.intAttr(elem, "start")
		};
	}
	function parseFooterHeaderReference(elem, xml) {
		return {
			id: xml.attr(elem, "id"),
			type: xml.attr(elem, "type")
		};
	}
	//#endregion
	//#region src/document/line-spacing.ts
	function parseLineSpacing(elem, xml) {
		return {
			before: xml.lengthAttr(elem, "before"),
			after: xml.lengthAttr(elem, "after"),
			line: xml.intAttr(elem, "line"),
			lineRule: xml.attr(elem, "lineRule")
		};
	}
	//#endregion
	//#region src/document/run.ts
	function parseRunProperties(elem, xml) {
		let result = {};
		for (let el of xml.elements(elem)) parseRunProperty(el, result, xml);
		return result;
	}
	function parseRunProperty(elem, props, xml) {
		if (parseCommonProperty(elem, props, xml)) return true;
		return false;
	}
	//#endregion
	//#region src/document/paragraph.ts
	function parseParagraphProperties(elem, xml) {
		let result = {};
		for (let el of xml.elements(elem)) parseParagraphProperty(el, result, xml);
		return result;
	}
	function parseParagraphProperty(elem, props, xml) {
		if (elem.namespaceURI != ns$1.wordml) return false;
		if (parseCommonProperty(elem, props, xml)) return true;
		switch (elem.localName) {
			case "tabs":
				props.tabs = parseTabs(elem, xml);
				break;
			case "sectPr":
				props.sectionProps = parseSectionProperties(elem, xml);
				break;
			case "numPr":
				props.numbering = parseNumbering$1(elem, xml);
				break;
			case "spacing":
				props.lineSpacing = parseLineSpacing(elem, xml);
				return false;
			case "textAlignment":
				props.textAlignment = xml.attr(elem, "val");
				return false;
			case "keepLines":
				props.keepLines = xml.boolAttr(elem, "val", true);
				break;
			case "keepNext":
				props.keepNext = xml.boolAttr(elem, "val", true);
				break;
			case "pageBreakBefore":
				props.pageBreakBefore = xml.boolAttr(elem, "val", true);
				break;
			case "outlineLvl":
				props.outlineLevel = xml.intAttr(elem, "val");
				break;
			case "pStyle":
				props.styleName = xml.attr(elem, "val");
				break;
			case "rPr":
				props.runProps = parseRunProperties(elem, xml);
				break;
			default: return false;
		}
		return true;
	}
	function parseTabs(elem, xml) {
		return xml.elements(elem, "tab").map((e) => ({
			position: xml.lengthAttr(e, "pos"),
			leader: xml.attr(e, "leader"),
			style: xml.attr(e, "val")
		}));
	}
	function parseNumbering$1(elem, xml) {
		var result = {};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "numId":
				result.id = xml.attr(e, "val");
				break;
			case "ilvl": result.level = xml.intAttr(e, "val");
		}
		return result;
	}
	//#endregion
	//#region src/numbering/numbering.ts
	function parseNumberingPart(elem, xml) {
		let result = {
			numberings: [],
			abstractNumberings: [],
			bulletPictures: []
		};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "num":
				result.numberings.push(parseNumbering(e, xml));
				break;
			case "abstractNum":
				result.abstractNumberings.push(parseAbstractNumbering(e, xml));
				break;
			case "numPicBullet": result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
		}
		return result;
	}
	function parseNumbering(elem, xml) {
		let result = {
			id: xml.attr(elem, "numId"),
			overrides: []
		};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "abstractNumId":
				result.abstractId = xml.attr(e, "val");
				break;
			case "lvlOverride": result.overrides.push(parseNumberingLevelOverrride(e, xml));
		}
		return result;
	}
	function parseAbstractNumbering(elem, xml) {
		let result = {
			id: xml.attr(elem, "abstractNumId"),
			levels: []
		};
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "name":
				result.name = xml.attr(e, "val");
				break;
			case "multiLevelType":
				result.multiLevelType = xml.attr(e, "val");
				break;
			case "numStyleLink":
				result.numberingStyleLink = xml.attr(e, "val");
				break;
			case "styleLink":
				result.styleLink = xml.attr(e, "val");
				break;
			case "lvl": result.levels.push(parseNumberingLevel(e, xml));
		}
		return result;
	}
	function parseNumberingLevel(elem, xml) {
		let result = { level: xml.intAttr(elem, "ilvl") };
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "start":
				result.start = xml.attr(e, "val");
				break;
			case "lvlRestart":
				result.restart = xml.intAttr(e, "val");
				break;
			case "numFmt":
				result.format = xml.attr(e, "val");
				break;
			case "lvlText":
				result.text = xml.attr(e, "val");
				break;
			case "lvlJc":
				result.justification = xml.attr(e, "val");
				break;
			case "lvlPicBulletId":
				result.bulletPictureId = xml.attr(e, "val");
				break;
			case "pStyle":
				result.paragraphStyle = xml.attr(e, "val");
				break;
			case "pPr":
				result.paragraphProps = parseParagraphProperties(e, xml);
				break;
			case "rPr": result.runProps = parseRunProperties(e, xml);
		}
		return result;
	}
	function parseNumberingLevelOverrride(elem, xml) {
		let result = { level: xml.intAttr(elem, "ilvl") };
		for (let e of xml.elements(elem)) switch (e.localName) {
			case "startOverride":
				result.start = xml.intAttr(e, "val");
				break;
			case "lvl": result.numberingLevel = parseNumberingLevel(e, xml);
		}
		return result;
	}
	function parseNumberingBulletPicture(elem, xml) {
		var pict = xml.element(elem, "pict");
		var shape = pict && xml.element(pict, "shape");
		var imagedata = shape && xml.element(shape, "imagedata");
		return imagedata ? {
			id: xml.attr(elem, "numPicBulletId"),
			referenceId: xml.attr(imagedata, "id"),
			style: xml.attr(shape, "style")
		} : null;
	}
	//#endregion
	//#region src/numbering/numbering-part.ts
	var NumberingPart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
		parseXml(root) {
			Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
			this.domNumberings = this._documentParser.parseNumberingFile(root);
		}
	};
	//#endregion
	//#region src/styles/styles-part.ts
	var StylesPart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
		parseXml(root) {
			this.styles = this._documentParser.parseStylesFile(root);
		}
	};
	//#endregion
	//#region src/document/dom.ts
	var DomType = /* @__PURE__ */ function(DomType) {
		DomType["Document"] = "document";
		DomType["Paragraph"] = "paragraph";
		DomType["Run"] = "run";
		DomType["Break"] = "break";
		DomType["NoBreakHyphen"] = "noBreakHyphen";
		DomType["Table"] = "table";
		DomType["Row"] = "row";
		DomType["Cell"] = "cell";
		DomType["Hyperlink"] = "hyperlink";
		DomType["SmartTag"] = "smartTag";
		DomType["Drawing"] = "drawing";
		DomType["Image"] = "image";
		DomType["Text"] = "text";
		DomType["Tab"] = "tab";
		DomType["Symbol"] = "symbol";
		DomType["BookmarkStart"] = "bookmarkStart";
		DomType["BookmarkEnd"] = "bookmarkEnd";
		DomType["Footer"] = "footer";
		DomType["Header"] = "header";
		DomType["FootnoteReference"] = "footnoteReference";
		DomType["EndnoteReference"] = "endnoteReference";
		DomType["Footnote"] = "footnote";
		DomType["Endnote"] = "endnote";
		DomType["SimpleField"] = "simpleField";
		DomType["ComplexField"] = "complexField";
		DomType["Instruction"] = "instruction";
		DomType["VmlPicture"] = "vmlPicture";
		DomType["MmlMath"] = "mmlMath";
		DomType["MmlMathParagraph"] = "mmlMathParagraph";
		DomType["MmlFraction"] = "mmlFraction";
		DomType["MmlFunction"] = "mmlFunction";
		DomType["MmlFunctionName"] = "mmlFunctionName";
		DomType["MmlNumerator"] = "mmlNumerator";
		DomType["MmlDenominator"] = "mmlDenominator";
		DomType["MmlRadical"] = "mmlRadical";
		DomType["MmlBase"] = "mmlBase";
		DomType["MmlDegree"] = "mmlDegree";
		DomType["MmlSuperscript"] = "mmlSuperscript";
		DomType["MmlSubscript"] = "mmlSubscript";
		DomType["MmlPreSubSuper"] = "mmlPreSubSuper";
		DomType["MmlSubArgument"] = "mmlSubArgument";
		DomType["MmlSuperArgument"] = "mmlSuperArgument";
		DomType["MmlNary"] = "mmlNary";
		DomType["MmlDelimiter"] = "mmlDelimiter";
		DomType["MmlRun"] = "mmlRun";
		DomType["MmlEquationArray"] = "mmlEquationArray";
		DomType["MmlLimit"] = "mmlLimit";
		DomType["MmlLimitLower"] = "mmlLimitLower";
		DomType["MmlMatrix"] = "mmlMatrix";
		DomType["MmlMatrixRow"] = "mmlMatrixRow";
		DomType["MmlBox"] = "mmlBox";
		DomType["MmlBar"] = "mmlBar";
		DomType["MmlGroupChar"] = "mmlGroupChar";
		DomType["VmlElement"] = "vmlElement";
		DomType["Inserted"] = "inserted";
		DomType["Deleted"] = "deleted";
		DomType["DeletedText"] = "deletedText";
		DomType["Comment"] = "comment";
		DomType["CommentReference"] = "commentReference";
		DomType["CommentRangeStart"] = "commentRangeStart";
		DomType["CommentRangeEnd"] = "commentRangeEnd";
		DomType["AltChunk"] = "altChunk";
		return DomType;
	}({});
	var OpenXmlElementBase = class {
		constructor() {
			this.children = [];
			this.cssStyle = {};
		}
	};
	//#endregion
	//#region src/header-footer/elements.ts
	var WmlHeader = class extends OpenXmlElementBase {
		constructor(..._args) {
			super(..._args);
			this.type = DomType.Header;
		}
	};
	var WmlFooter = class extends OpenXmlElementBase {
		constructor(..._args2) {
			super(..._args2);
			this.type = DomType.Footer;
		}
	};
	//#endregion
	//#region src/header-footer/parts.ts
	var BaseHeaderFooterPart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
		parseXml(root) {
			this.rootElement = this.createRootElement();
			this.rootElement.children = this._documentParser.parseBodyElements(root);
		}
	};
	var HeaderPart = class extends BaseHeaderFooterPart {
		createRootElement() {
			return new WmlHeader();
		}
	};
	var FooterPart = class extends BaseHeaderFooterPart {
		createRootElement() {
			return new WmlFooter();
		}
	};
	//#endregion
	//#region src/document-props/extended-props.ts
	function parseExtendedProps(root, xmlParser) {
		const result = {};
		for (let el of xmlParser.elements(root)) switch (el.localName) {
			case "Template":
				result.template = el.textContent;
				break;
			case "Pages":
				result.pages = safeParseToInt(el.textContent);
				break;
			case "Words":
				result.words = safeParseToInt(el.textContent);
				break;
			case "Characters":
				result.characters = safeParseToInt(el.textContent);
				break;
			case "Application":
				result.application = el.textContent;
				break;
			case "Lines":
				result.lines = safeParseToInt(el.textContent);
				break;
			case "Paragraphs":
				result.paragraphs = safeParseToInt(el.textContent);
				break;
			case "Company":
				result.company = el.textContent;
				break;
			case "AppVersion": result.appVersion = el.textContent;
		}
		return result;
	}
	function safeParseToInt(value) {
		if (typeof value === "undefined") return;
		return parseInt(value);
	}
	//#endregion
	//#region src/document-props/extended-props-part.ts
	var ExtendedPropsPart = class extends Part {
		parseXml(root) {
			this.props = parseExtendedProps(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/document-props/core-props.ts
	function parseCoreProps(root, xmlParser) {
		const result = {};
		for (let el of xmlParser.elements(root)) switch (el.localName) {
			case "title":
				result.title = el.textContent;
				break;
			case "description":
				result.description = el.textContent;
				break;
			case "subject":
				result.subject = el.textContent;
				break;
			case "creator":
				result.creator = el.textContent;
				break;
			case "keywords":
				result.keywords = el.textContent;
				break;
			case "language":
				result.language = el.textContent;
				break;
			case "lastModifiedBy":
				result.lastModifiedBy = el.textContent;
				break;
			case "revision": el.textContent && (result.revision = parseInt(el.textContent));
		}
		return result;
	}
	//#endregion
	//#region src/document-props/core-props-part.ts
	var CorePropsPart = class extends Part {
		parseXml(root) {
			this.props = parseCoreProps(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/theme/theme.ts
	var DmlTheme = class {};
	function parseTheme(elem, xml) {
		var result = new DmlTheme();
		var themeElements = xml.element(elem, "themeElements");
		for (let el of xml.elements(themeElements)) switch (el.localName) {
			case "clrScheme":
				result.colorScheme = parseColorScheme(el, xml);
				break;
			case "fontScheme": result.fontScheme = parseFontScheme(el, xml);
		}
		return result;
	}
	function parseColorScheme(elem, xml) {
		var result = {
			name: xml.attr(elem, "name"),
			colors: {}
		};
		for (let el of xml.elements(elem)) {
			var srgbClr = xml.element(el, "srgbClr");
			var sysClr = xml.element(el, "sysClr");
			if (srgbClr) result.colors[el.localName] = xml.attr(srgbClr, "val");
			else if (sysClr) result.colors[el.localName] = xml.attr(sysClr, "lastClr");
		}
		return result;
	}
	function parseFontScheme(elem, xml) {
		var result = { name: xml.attr(elem, "name") };
		for (let el of xml.elements(elem)) switch (el.localName) {
			case "majorFont":
				result.majorFont = parseFontInfo(el, xml);
				break;
			case "minorFont": result.minorFont = parseFontInfo(el, xml);
		}
		return result;
	}
	function parseFontInfo(elem, xml) {
		return {
			latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
			eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
			csTypeface: xml.elementAttr(elem, "cs", "typeface")
		};
	}
	//#endregion
	//#region src/theme/theme-part.ts
	var ThemePart = class extends Part {
		constructor(pkg, path) {
			super(pkg, path);
		}
		parseXml(root) {
			this.theme = parseTheme(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/notes/elements.ts
	var WmlBaseNote = class {};
	var WmlFootnote = class extends WmlBaseNote {
		constructor(..._args) {
			super(..._args);
			this.type = DomType.Footnote;
		}
	};
	var WmlEndnote = class extends WmlBaseNote {
		constructor(..._args2) {
			super(..._args2);
			this.type = DomType.Endnote;
		}
	};
	//#endregion
	//#region src/notes/parts.ts
	var BaseNotePart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
	};
	var FootnotesPart = class extends BaseNotePart {
		constructor(pkg, path, parser) {
			super(pkg, path, parser);
		}
		parseXml(root) {
			this.notes = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
		}
	};
	var EndnotesPart = class extends BaseNotePart {
		constructor(pkg, path, parser) {
			super(pkg, path, parser);
		}
		parseXml(root) {
			this.notes = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
		}
	};
	//#endregion
	//#region src/settings/settings.ts
	function parseSettings(elem, xml) {
		var result = {};
		for (let el of xml.elements(elem)) switch (el.localName) {
			case "defaultTabStop":
				result.defaultTabStop = xml.lengthAttr(el, "val");
				break;
			case "footnotePr":
				result.footnoteProps = parseNoteProperties(el, xml);
				break;
			case "endnotePr":
				result.endnoteProps = parseNoteProperties(el, xml);
				break;
			case "autoHyphenation":
				result.autoHyphenation = xml.boolAttr(el, "val");
				break;
			case "compat": for (let c of xml.elements(el)) {
				var _xml$attr;
				if (c.localName == "compatSetting" && xml.attr(c, "name") == "compatibilityMode" && ((_xml$attr = xml.attr(c, "uri")) !== null && _xml$attr !== void 0 ? _xml$attr : "").includes("schemas.microsoft.com/office/word")) result.compatMode = xml.intAttr(c, "val", null);
			}
		}
		return result;
	}
	function parseNoteProperties(elem, xml) {
		var result = { defaultNoteIds: [] };
		for (let el of xml.elements(elem)) switch (el.localName) {
			case "numFmt":
				result.nummeringFormat = xml.attr(el, "val");
				break;
			case "footnote":
			case "endnote": result.defaultNoteIds.push(xml.attr(el, "id"));
		}
		return result;
	}
	//#endregion
	//#region src/settings/settings-part.ts
	var SettingsPart = class extends Part {
		constructor(pkg, path) {
			super(pkg, path);
		}
		parseXml(root) {
			this.settings = parseSettings(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/document-props/custom-props.ts
	function parseCustomProps(root, xml) {
		return xml.elements(root, "property").map((e) => {
			const firstChild = e.firstChild;
			return {
				formatId: xml.attr(e, "fmtid"),
				name: xml.attr(e, "name"),
				type: firstChild.nodeName,
				value: firstChild.textContent
			};
		});
	}
	//#endregion
	//#region src/document-props/custom-props-part.ts
	var CustomPropsPart = class extends Part {
		parseXml(root) {
			this.props = parseCustomProps(root, this._package.xmlParser);
		}
	};
	//#endregion
	//#region src/comments/comments-part.ts
	var CommentsPart = class extends Part {
		constructor(pkg, path, parser) {
			super(pkg, path);
			this._documentParser = parser;
		}
		parseXml(root) {
			this.comments = this._documentParser.parseComments(root);
			this.commentMap = keyBy(this.comments, (x) => x.id);
		}
	};
	//#endregion
	//#region src/comments/comments-extended-part.ts
	var CommentsExtendedPart = class extends Part {
		constructor(pkg, path) {
			super(pkg, path);
			this.comments = [];
		}
		parseXml(root) {
			const xml = this._package.xmlParser;
			for (let el of xml.elements(root, "commentEx")) this.comments.push({
				paraId: xml.attr(el, "paraId"),
				paraIdParent: xml.attr(el, "paraIdParent"),
				done: xml.boolAttr(el, "done")
			});
			this.commentMap = keyBy(this.comments, (x) => x.paraId);
		}
	};
	//#endregion
	//#region src/word-document.ts
	var topLevelRels = [
		{
			type: RelationshipTypes.OfficeDocument,
			target: "word/document.xml"
		},
		{
			type: RelationshipTypes.ExtendedProperties,
			target: "docProps/app.xml"
		},
		{
			type: RelationshipTypes.CoreProperties,
			target: "docProps/core.xml"
		},
		{
			type: RelationshipTypes.CustomProperties,
			target: "docProps/custom.xml"
		}
	];
	var WordDocument = class WordDocument {
		constructor() {
			this.parts = [];
			this.partsMap = {};
			this.contentTypes = [];
		}
		static async load(blob, parser, options) {
			var d = new WordDocument();
			d._options = options;
			d._parser = parser;
			d._package = await OpenXmlPackage.load(blob, options);
			d.rels = await d._package.loadRelationships();
			d.contentTypes = await d._package.loadContentTypes();
			await Promise.all(topLevelRels.map((rel) => {
				var _d$rels$find;
				const r = (_d$rels$find = d.rels.find((x) => x.type === rel.type)) !== null && _d$rels$find !== void 0 ? _d$rels$find : rel;
				return d.loadRelationshipPart(r.target, r.type);
			}));
			return d;
		}
		save(type = "blob") {
			return this._package.save(type);
		}
		async loadRelationshipPart(path, type) {
			var _part$rels;
			if (this.partsMap[path]) return this.partsMap[path];
			if (!this._package.get(path)) return null;
			let part = null;
			switch (type) {
				case RelationshipTypes.OfficeDocument:
					this.documentPart = part = new DocumentPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.FontTable:
					this.fontTablePart = part = new FontTablePart(this._package, path);
					break;
				case RelationshipTypes.Numbering:
					this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.Styles:
					this.stylesPart = part = new StylesPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.Theme:
					this.themePart = part = new ThemePart(this._package, path);
					break;
				case RelationshipTypes.Footnotes:
					this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.Endnotes:
					this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.Footer:
					part = new FooterPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.Header:
					part = new HeaderPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.CoreProperties:
					this.corePropsPart = part = new CorePropsPart(this._package, path);
					break;
				case RelationshipTypes.ExtendedProperties:
					this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
					break;
				case RelationshipTypes.CustomProperties:
					part = new CustomPropsPart(this._package, path);
					break;
				case RelationshipTypes.Settings:
					this.settingsPart = part = new SettingsPart(this._package, path);
					break;
				case RelationshipTypes.Comments:
					this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
					break;
				case RelationshipTypes.CommentsExtended: this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
			}
			if (part == null) return Promise.resolve(null);
			this.partsMap[path] = part;
			this.parts.push(part);
			await part.load();
			if (((_part$rels = part.rels) === null || _part$rels === void 0 ? void 0 : _part$rels.length) > 0) {
				const [folder] = splitPath(part.path);
				await Promise.all(part.rels.map((rel) => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
			}
			return part;
		}
		async loadDocumentImage(id, part) {
			const path = this.getPathById(part !== null && part !== void 0 ? part : this.documentPart, id);
			return path ? this.blobToURL(await this._package.load(path, "blob"), path) : null;
		}
		async loadNumberingImage(id) {
			const path = this.getPathById(this.numberingPart, id);
			return path ? this.blobToURL(await this._package.load(path, "blob"), path) : null;
		}
		async loadFont(id, key) {
			const path = this.getPathById(this.fontTablePart, id);
			if (!path) return null;
			const x = await this._package.load(path, "uint8array");
			return x ? this.blobToURL(new Blob([deobfuscate(x, key)]), path) : x;
		}
		async loadAltChunk(id, part) {
			const path = this.getPathById(part !== null && part !== void 0 ? part : this.documentPart, id);
			return path ? this._package.load(path, "string") : Promise.resolve(null);
		}
		blobToURL(blob, path) {
			if (!blob) return null;
			if (path) {
				const ct = this.contentTypes.find((x) => x.partName === path || x.extension && path.endsWith(`.${x.extension}`));
				blob = ct ? new Blob([blob], { type: ct.contentType }) : blob;
			}
			if (this._options.useBase64URL) return blobToBase64(blob);
			return URL.createObjectURL(blob);
		}
		findPartByRelId(id, basePart = null) {
			var _basePart$rels;
			var rel = ((_basePart$rels = basePart.rels) !== null && _basePart$rels !== void 0 ? _basePart$rels : this.rels).find((r) => r.id == id);
			const folder = basePart ? splitPath(basePart.path)[0] : "";
			return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
		}
		getPathById(part, id) {
			const rel = part.rels.find((x) => x.id == id);
			const [folder] = splitPath(part.path);
			return rel ? resolvePath(rel.target, folder) : null;
		}
	};
	function deobfuscate(data, guidKey) {
		const len = 16;
		const trimmed = guidKey.replace(/{|}|-/g, "");
		const numbers = new Array(len);
		for (let i = 0; i < len; i++) numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);
		for (let i = 0; i < 32; i++) data[i] = data[i] ^ numbers[i % len];
		return data;
	}
	//#endregion
	//#region src/document/bookmarks.ts
	function parseBookmarkStart(elem, xml) {
		return {
			type: DomType.BookmarkStart,
			id: xml.attr(elem, "id"),
			name: xml.attr(elem, "name"),
			colFirst: xml.intAttr(elem, "colFirst"),
			colLast: xml.intAttr(elem, "colLast")
		};
	}
	function parseBookmarkEnd(elem, xml) {
		return {
			type: DomType.BookmarkEnd,
			id: xml.attr(elem, "id")
		};
	}
	//#endregion
	//#region src/vml/vml.ts
	var VmlElement = class extends OpenXmlElementBase {
		constructor(..._args) {
			super(..._args);
			this.type = DomType.VmlElement;
			this.attrs = {};
		}
	};
	function parseVmlElement(elem, parser) {
		var result = new VmlElement();
		switch (elem.localName) {
			case "rect":
				result.tagName = "rect";
				Object.assign(result.attrs, {
					width: "100%",
					height: "100%"
				});
				break;
			case "oval":
				result.tagName = "ellipse";
				Object.assign(result.attrs, {
					cx: "50%",
					cy: "50%",
					rx: "50%",
					ry: "50%"
				});
				break;
			case "line":
				result.tagName = "line";
				break;
			case "shape":
				result.tagName = "g";
				break;
			case "textbox":
				result.tagName = "foreignObject";
				Object.assign(result.attrs, {
					width: "100%",
					height: "100%"
				});
				break;
			default: return null;
		}
		for (const at of globalXmlParser.attrs(elem)) switch (at.localName) {
			case "style":
				result.cssStyleText = at.value;
				break;
			case "fillcolor":
				result.attrs.fill = at.value;
				break;
			case "from":
				const [x1, y1] = parsePoint(at.value);
				Object.assign(result.attrs, {
					x1,
					y1
				});
				break;
			case "to":
				const [x2, y2] = parsePoint(at.value);
				Object.assign(result.attrs, {
					x2,
					y2
				});
		}
		for (const el of globalXmlParser.elements(elem)) switch (el.localName) {
			case "stroke":
				Object.assign(result.attrs, parseStroke(el));
				break;
			case "fill":
				Object.assign(result.attrs, parseFill(el));
				break;
			case "imagedata":
				result.tagName = "image";
				Object.assign(result.attrs, {
					width: "100%",
					height: "100%"
				});
				result.imageHref = {
					id: globalXmlParser.attr(el, "id"),
					title: globalXmlParser.attr(el, "title")
				};
				break;
			case "txbxContent":
				result.children.push(...parser.parseBodyElements(el));
				break;
			default:
				const child = parseVmlElement(el, parser);
				child && result.children.push(child);
		}
		return result;
	}
	function parseStroke(el) {
		var _xml$lengthAttr;
		return {
			"stroke": globalXmlParser.attr(el, "color"),
			"stroke-width": (_xml$lengthAttr = globalXmlParser.lengthAttr(el, "weight", LengthUsage.Emu)) !== null && _xml$lengthAttr !== void 0 ? _xml$lengthAttr : "1px"
		};
	}
	function parseFill(el) {
		return {};
	}
	function parsePoint(val) {
		return val.split(",");
	}
	//#endregion
	//#region src/comments/elements.ts
	var WmlComment = class extends OpenXmlElementBase {
		constructor(..._args) {
			super(..._args);
			this.type = DomType.Comment;
		}
	};
	var WmlCommentReference = class extends OpenXmlElementBase {
		constructor(id) {
			super();
			this.id = id;
			this.type = DomType.CommentReference;
		}
	};
	var WmlCommentRangeStart = class extends OpenXmlElementBase {
		constructor(id) {
			super();
			this.id = id;
			this.type = DomType.CommentRangeStart;
		}
	};
	var WmlCommentRangeEnd = class extends OpenXmlElementBase {
		constructor(id) {
			super();
			this.id = id;
			this.type = DomType.CommentRangeEnd;
		}
	};
	//#endregion
	//#region src/document-parser.ts
	var autos = {
		shd: "inherit",
		color: "black",
		borderColor: "black",
		highlight: "transparent"
	};
	var supportedNamespaceURIs = [];
	var mmlTagMap = {
		"oMath": DomType.MmlMath,
		"oMathPara": DomType.MmlMathParagraph,
		"f": DomType.MmlFraction,
		"func": DomType.MmlFunction,
		"fName": DomType.MmlFunctionName,
		"num": DomType.MmlNumerator,
		"den": DomType.MmlDenominator,
		"rad": DomType.MmlRadical,
		"deg": DomType.MmlDegree,
		"e": DomType.MmlBase,
		"sSup": DomType.MmlSuperscript,
		"sSub": DomType.MmlSubscript,
		"sPre": DomType.MmlPreSubSuper,
		"sup": DomType.MmlSuperArgument,
		"sub": DomType.MmlSubArgument,
		"d": DomType.MmlDelimiter,
		"nary": DomType.MmlNary,
		"eqArr": DomType.MmlEquationArray,
		"lim": DomType.MmlLimit,
		"limLow": DomType.MmlLimitLower,
		"m": DomType.MmlMatrix,
		"mr": DomType.MmlMatrixRow,
		"box": DomType.MmlBox,
		"bar": DomType.MmlBar,
		"groupChr": DomType.MmlGroupChar
	};
	var DocumentParser = class {
		constructor(options) {
			this.options = {
				ignoreWidth: false,
				debug: false,
				...options
			};
		}
		parseNotes(xmlDoc, elemName, elemClass) {
			var result = [];
			for (let el of globalXmlParser.elements(xmlDoc, elemName)) {
				const node = new elemClass();
				node.id = globalXmlParser.attr(el, "id");
				node.noteType = globalXmlParser.attr(el, "type");
				node.children = this.parseBodyElements(el);
				result.push(node);
			}
			return result;
		}
		parseComments(xmlDoc) {
			var result = [];
			for (let el of globalXmlParser.elements(xmlDoc, "comment")) {
				const item = new WmlComment();
				item.id = globalXmlParser.attr(el, "id");
				item.author = globalXmlParser.attr(el, "author");
				item.initials = globalXmlParser.attr(el, "initials");
				item.date = globalXmlParser.attr(el, "date");
				item.children = this.parseBodyElements(el);
				result.push(item);
			}
			return result;
		}
		parseDocumentFile(xmlDoc) {
			var xbody = globalXmlParser.element(xmlDoc, "body");
			var background = globalXmlParser.element(xmlDoc, "background");
			var sectPr = globalXmlParser.element(xbody, "sectPr");
			return {
				type: DomType.Document,
				children: this.parseBodyElements(xbody),
				props: sectPr ? parseSectionProperties(sectPr, globalXmlParser) : {},
				cssStyle: background ? this.parseBackground(background) : {}
			};
		}
		parseBackground(elem) {
			var result = {};
			var color = xmlUtil.colorAttr(elem, "color");
			if (color) result["background-color"] = color;
			return result;
		}
		parseBodyElements(element) {
			var children = [];
			for (const elem of globalXmlParser.elements(element)) switch (elem.localName) {
				case "p":
					children.push(this.parseParagraph(elem));
					break;
				case "altChunk":
					children.push(this.parseAltChunk(elem));
					break;
				case "tbl":
					children.push(this.parseTable(elem));
					break;
				case "sdt": children.push(...this.parseSdt(elem, (e) => this.parseBodyElements(e)));
			}
			return children;
		}
		parseStylesFile(xstyles) {
			var result = [];
			for (const n of globalXmlParser.elements(xstyles)) switch (n.localName) {
				case "style":
					result.push(this.parseStyle(n));
					break;
				case "docDefaults": result.push(this.parseDefaultStyles(n));
			}
			return result;
		}
		parseDefaultStyles(node) {
			var result = {
				id: null,
				name: null,
				target: null,
				basedOn: null,
				styles: []
			};
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "rPrDefault":
					var rPr = globalXmlParser.element(c, "rPr");
					if (rPr) result.styles.push({
						target: "span",
						values: this.parseDefaultProperties(rPr, {})
					});
					break;
				case "pPrDefault":
					var pPr = globalXmlParser.element(c, "pPr");
					if (pPr) result.styles.push({
						target: "p",
						values: this.parseDefaultProperties(pPr, {})
					});
			}
			return result;
		}
		parseStyle(node) {
			var result = {
				id: globalXmlParser.attr(node, "styleId"),
				isDefault: globalXmlParser.boolAttr(node, "default"),
				name: null,
				target: null,
				basedOn: null,
				styles: [],
				linked: null
			};
			switch (globalXmlParser.attr(node, "type")) {
				case "paragraph":
					result.target = "p";
					break;
				case "table":
					result.target = "table";
					break;
				case "character": result.target = "span";
			}
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "basedOn":
					result.basedOn = globalXmlParser.attr(n, "val");
					break;
				case "name":
					result.name = globalXmlParser.attr(n, "val");
					break;
				case "link":
					result.linked = globalXmlParser.attr(n, "val");
					break;
				case "next":
					result.next = globalXmlParser.attr(n, "val");
					break;
				case "aliases":
					result.aliases = globalXmlParser.attr(n, "val").split(",");
					break;
				case "pPr":
					result.styles.push({
						target: "p",
						values: this.parseDefaultProperties(n, {})
					});
					result.paragraphProps = parseParagraphProperties(n, globalXmlParser);
					break;
				case "rPr":
					result.styles.push({
						target: "span",
						values: this.parseDefaultProperties(n, {})
					});
					result.runProps = parseRunProperties(n, globalXmlParser);
					break;
				case "tblPr":
				case "tcPr":
					result.styles.push({
						target: "td",
						values: this.parseDefaultProperties(n, {})
					});
					break;
				case "tblStylePr":
					for (let s of this.parseTableStyle(n)) result.styles.push(s);
					break;
				case "rsid":
				case "qFormat":
				case "hidden":
				case "semiHidden":
				case "unhideWhenUsed":
				case "autoRedefine":
				case "uiPriority": break;
				default: this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
			}
			return result;
		}
		parseTableStyle(node) {
			var result = [];
			var type = globalXmlParser.attr(node, "type");
			var selector = "";
			var modificator = "";
			switch (type) {
				case "firstRow":
					modificator = ".first-row";
					selector = "tr.first-row td";
					break;
				case "lastRow":
					modificator = ".last-row";
					selector = "tr.last-row td";
					break;
				case "firstCol":
					modificator = ".first-col";
					selector = "td.first-col";
					break;
				case "lastCol":
					modificator = ".last-col";
					selector = "td.last-col";
					break;
				case "band1Vert":
					modificator = ":not(.no-vband)";
					selector = "td.odd-col";
					break;
				case "band2Vert":
					modificator = ":not(.no-vband)";
					selector = "td.even-col";
					break;
				case "band1Horz":
					modificator = ":not(.no-hband)";
					selector = "tr.odd-row";
					break;
				case "band2Horz":
					modificator = ":not(.no-hband)";
					selector = "tr.even-row";
					break;
				default: return [];
			}
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "pPr":
					result.push({
						target: `${selector} p`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;
				case "rPr":
					result.push({
						target: `${selector} span`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;
				case "tblPr":
				case "tcPr": result.push({
					target: selector,
					mod: modificator,
					values: this.parseDefaultProperties(n, {})
				});
			}
			return result;
		}
		parseNumberingFile(node) {
			const levels = [];
			const nums = [];
			const bullets = [];
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "abstractNum":
					levels.push(...this.parseAbstractNumbering(n, bullets));
					break;
				case "numPicBullet":
					bullets.push(this.parseNumberingPicBullet(n));
					break;
				case "num": nums.push({
					numId: globalXmlParser.attr(n, "numId"),
					abstractNumId: globalXmlParser.elementAttr(n, "abstractNumId", "val")
				});
			}
			return nums.flatMap((x) => levels.filter((lvl) => x.abstractNumId == lvl.id).map((lvl) => ({
				...lvl,
				id: x.numId
			})));
		}
		parseNumberingPicBullet(elem) {
			var pict = globalXmlParser.element(elem, "pict");
			var shape = pict && globalXmlParser.element(pict, "shape");
			var imagedata = shape && globalXmlParser.element(shape, "imagedata");
			return imagedata ? {
				id: globalXmlParser.intAttr(elem, "numPicBulletId"),
				src: globalXmlParser.attr(imagedata, "id"),
				style: globalXmlParser.attr(shape, "style")
			} : null;
		}
		parseAbstractNumbering(node, bullets) {
			var result = [];
			var id = globalXmlParser.attr(node, "abstractNumId");
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "lvl": result.push(this.parseNumberingLevel(id, n, bullets));
			}
			return result;
		}
		parseNumberingLevel(id, node, bullets) {
			var result = {
				id,
				level: globalXmlParser.intAttr(node, "ilvl"),
				start: 1,
				pStyleName: void 0,
				pStyle: {},
				rStyle: {},
				suff: "tab"
			};
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "start":
					result.start = globalXmlParser.intAttr(n, "val");
					break;
				case "pPr":
					this.parseDefaultProperties(n, result.pStyle);
					break;
				case "rPr":
					this.parseDefaultProperties(n, result.rStyle);
					break;
				case "lvlPicBulletId":
					var bulletId = globalXmlParser.intAttr(n, "val");
					result.bullet = bullets.find((x) => (x === null || x === void 0 ? void 0 : x.id) == bulletId);
					break;
				case "lvlText":
					result.levelText = globalXmlParser.attr(n, "val");
					break;
				case "pStyle":
					result.pStyleName = globalXmlParser.attr(n, "val");
					break;
				case "numFmt":
					result.format = globalXmlParser.attr(n, "val");
					break;
				case "suff": result.suff = globalXmlParser.attr(n, "val");
			}
			return result;
		}
		parseSdt(node, parser) {
			const sdtContent = globalXmlParser.element(node, "sdtContent");
			return sdtContent ? parser(sdtContent) : [];
		}
		parseChange(type, node, parentParser) {
			var _parentParser$childre, _parentParser;
			return {
				type,
				children: (_parentParser$childre = (_parentParser = parentParser(node)) === null || _parentParser === void 0 ? void 0 : _parentParser.children) !== null && _parentParser$childre !== void 0 ? _parentParser$childre : [],
				id: globalXmlParser.attr(node, "id"),
				author: globalXmlParser.attr(node, "author"),
				date: globalXmlParser.attr(node, "date")
			};
		}
		parseAltChunk(node) {
			return {
				type: DomType.AltChunk,
				children: [],
				id: globalXmlParser.attr(node, "id")
			};
		}
		parseParagraph(node) {
			var result = {
				type: DomType.Paragraph,
				children: []
			};
			for (let el of globalXmlParser.elements(node)) switch (el.localName) {
				case "pPr":
					this.parseParagraphProperties(el, result);
					break;
				case "r":
					result.children.push(this.parseRun(el, result));
					break;
				case "hyperlink":
					result.children.push(this.parseHyperlink(el, result));
					break;
				case "smartTag":
					result.children.push(this.parseSmartTag(el, result));
					break;
				case "bookmarkStart":
					result.children.push(parseBookmarkStart(el, globalXmlParser));
					break;
				case "bookmarkEnd":
					result.children.push(parseBookmarkEnd(el, globalXmlParser));
					break;
				case "commentRangeStart":
					result.children.push(new WmlCommentRangeStart(globalXmlParser.attr(el, "id")));
					break;
				case "commentRangeEnd":
					result.children.push(new WmlCommentRangeEnd(globalXmlParser.attr(el, "id")));
					break;
				case "oMath":
				case "oMathPara":
					result.children.push(this.parseMathElement(el));
					break;
				case "sdt":
					result.children.push(...this.parseSdt(el, (e) => this.parseParagraph(e).children));
					break;
				case "ins":
					result.children.push(this.parseChange(DomType.Inserted, el, (e) => this.parseParagraph(e)));
					break;
				case "del": result.children.push(this.parseChange(DomType.Deleted, el, (e) => this.parseParagraph(e)));
			}
			return result;
		}
		parseParagraphProperties(elem, paragraph) {
			this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, (c) => {
				if (parseParagraphProperty(c, paragraph, globalXmlParser)) return true;
				switch (c.localName) {
					case "pStyle":
						paragraph.styleName = globalXmlParser.attr(c, "val");
						break;
					case "cnfStyle":
						paragraph.className = values.classNameOfCnfStyle(c);
						break;
					case "framePr":
						this.parseFrame(c, paragraph);
						break;
					case "rPr": break;
					default: return false;
				}
				return true;
			});
		}
		parseFrame(node, paragraph) {
			if (globalXmlParser.attr(node, "dropCap") == "drop") paragraph.cssStyle["float"] = "left";
		}
		parseHyperlink(node, parent) {
			var result = {
				type: DomType.Hyperlink,
				parent,
				children: []
			};
			result.anchor = globalXmlParser.attr(node, "anchor");
			result.id = globalXmlParser.attr(node, "id");
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "r": result.children.push(this.parseRun(c, result));
			}
			return result;
		}
		parseSmartTag(node, parent) {
			var result = {
				type: DomType.SmartTag,
				parent,
				children: []
			};
			var uri = globalXmlParser.attr(node, "uri");
			var element = globalXmlParser.attr(node, "element");
			if (uri) result.uri = uri;
			if (element) result.element = element;
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
				case "smartTag": result.children.push(this.parseSmartTag(c, result));
			}
			return result;
		}
		parseRun(node, parent) {
			var result = {
				type: DomType.Run,
				parent,
				children: []
			};
			for (let c of globalXmlParser.elements(node)) {
				c = this.checkAlternateContent(c);
				switch (c.localName) {
					case "t":
						result.children.push({
							type: DomType.Text,
							text: c.textContent
						});
						break;
					case "delText":
						result.children.push({
							type: DomType.DeletedText,
							text: c.textContent
						});
						break;
					case "commentReference":
						result.children.push(new WmlCommentReference(globalXmlParser.attr(c, "id")));
						break;
					case "fldSimple":
						result.children.push({
							type: DomType.SimpleField,
							instruction: globalXmlParser.attr(c, "instr"),
							lock: globalXmlParser.boolAttr(c, "lock", false),
							dirty: globalXmlParser.boolAttr(c, "dirty", false)
						});
						break;
					case "instrText":
						result.fieldRun = true;
						result.children.push({
							type: DomType.Instruction,
							text: c.textContent
						});
						break;
					case "fldChar":
						result.fieldRun = true;
						result.children.push({
							type: DomType.ComplexField,
							charType: globalXmlParser.attr(c, "fldCharType"),
							lock: globalXmlParser.boolAttr(c, "lock", false),
							dirty: globalXmlParser.boolAttr(c, "dirty", false)
						});
						break;
					case "noBreakHyphen":
						result.children.push({ type: DomType.NoBreakHyphen });
						break;
					case "br":
						result.children.push({
							type: DomType.Break,
							break: globalXmlParser.attr(c, "type") || "textWrapping"
						});
						break;
					case "lastRenderedPageBreak":
						result.children.push({
							type: DomType.Break,
							break: "lastRenderedPageBreak"
						});
						break;
					case "sym":
						result.children.push({
							type: DomType.Symbol,
							font: encloseFontFamily(globalXmlParser.attr(c, "font")),
							char: globalXmlParser.hexAttr(c, "char")
						});
						break;
					case "tab":
						result.children.push({ type: DomType.Tab });
						break;
					case "footnoteReference":
						result.children.push({
							type: DomType.FootnoteReference,
							id: globalXmlParser.attr(c, "id")
						});
						break;
					case "endnoteReference":
						result.children.push({
							type: DomType.EndnoteReference,
							id: globalXmlParser.attr(c, "id")
						});
						break;
					case "drawing":
						let d = this.parseDrawing(c);
						if (d) result.children.push(d);
						break;
					case "pict":
						result.children.push(this.parseVmlPicture(c));
						break;
					case "rPr": this.parseRunProperties(c, result);
				}
			}
			return result;
		}
		parseMathElement(elem) {
			const propsTag = `${elem.localName}Pr`;
			const result = {
				type: mmlTagMap[elem.localName],
				children: []
			};
			for (const el of globalXmlParser.elements(elem)) if (mmlTagMap[el.localName]) result.children.push(this.parseMathElement(el));
			else if (el.localName == "r") {
				var run = this.parseRun(el);
				run.type = DomType.MmlRun;
				result.children.push(run);
			} else if (el.localName == propsTag) result.props = this.parseMathProperies(el);
			return result;
		}
		parseMathProperies(elem) {
			const result = {};
			for (const el of globalXmlParser.elements(elem)) switch (el.localName) {
				case "chr":
					result.char = globalXmlParser.attr(el, "val");
					break;
				case "vertJc":
					result.verticalJustification = globalXmlParser.attr(el, "val");
					break;
				case "pos":
					result.position = globalXmlParser.attr(el, "val");
					break;
				case "degHide":
					result.hideDegree = globalXmlParser.boolAttr(el, "val");
					break;
				case "begChr":
					result.beginChar = globalXmlParser.attr(el, "val");
					break;
				case "endChr": result.endChar = globalXmlParser.attr(el, "val");
			}
			return result;
		}
		parseRunProperties(elem, run) {
			this.parseDefaultProperties(elem, run.cssStyle = {}, null, (c) => {
				switch (c.localName) {
					case "rStyle":
						run.styleName = globalXmlParser.attr(c, "val");
						break;
					case "vertAlign":
						run.verticalAlign = values.valueOfVertAlign(c, true);
						break;
					default: return false;
				}
				return true;
			});
		}
		parseVmlPicture(elem) {
			const result = {
				type: DomType.VmlPicture,
				children: []
			};
			for (const el of globalXmlParser.elements(elem)) {
				const child = parseVmlElement(el, this);
				child && result.children.push(child);
			}
			return result;
		}
		checkAlternateContent(elem) {
			var _xml$element;
			if (elem.localName != "AlternateContent") return elem;
			var choice = globalXmlParser.element(elem, "Choice");
			if (choice) {
				var requires = globalXmlParser.attr(choice, "Requires");
				var namespaceURI = elem.lookupNamespaceURI(requires);
				if (supportedNamespaceURIs.includes(namespaceURI)) return choice.firstElementChild;
			}
			return (_xml$element = globalXmlParser.element(elem, "Fallback")) === null || _xml$element === void 0 ? void 0 : _xml$element.firstElementChild;
		}
		parseDrawing(node) {
			for (var n of globalXmlParser.elements(node)) switch (n.localName) {
				case "inline":
				case "anchor": return this.parseDrawingWrapper(n);
			}
		}
		parseDrawingWrapper(node) {
			var result = {
				type: DomType.Drawing,
				children: [],
				cssStyle: {}
			};
			var isAnchor = node.localName == "anchor";
			let wrapType = null;
			let simplePos = globalXmlParser.boolAttr(node, "simplePos");
			globalXmlParser.boolAttr(node, "behindDoc");
			let posX = {
				relative: "page",
				align: "left",
				offset: "0"
			};
			let posY = {
				relative: "page",
				align: "top",
				offset: "0"
			};
			for (var n of globalXmlParser.elements(node)) switch (n.localName) {
				case "simplePos":
					if (simplePos) {
						posX.offset = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
						posY.offset = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
					}
					break;
				case "extent":
					result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
					break;
				case "positionH":
				case "positionV":
					if (!simplePos) {
						var _xml$attr;
						let pos = n.localName == "positionH" ? posX : posY;
						var alignNode = globalXmlParser.element(n, "align");
						var offsetNode = globalXmlParser.element(n, "posOffset");
						pos.relative = (_xml$attr = globalXmlParser.attr(n, "relativeFrom")) !== null && _xml$attr !== void 0 ? _xml$attr : pos.relative;
						if (alignNode) pos.align = alignNode.textContent;
						if (offsetNode) pos.offset = convertLength(offsetNode.textContent, LengthUsage.Emu);
					}
					break;
				case "wrapTopAndBottom":
					wrapType = "wrapTopAndBottom";
					break;
				case "wrapNone":
					wrapType = "wrapNone";
					break;
				case "graphic":
					var g = this.parseGraphic(n);
					if (g) result.children.push(g);
			}
			if (wrapType == "wrapTopAndBottom") {
				result.cssStyle["display"] = "block";
				if (posX.align) {
					result.cssStyle["text-align"] = posX.align;
					result.cssStyle["width"] = "100%";
				}
			} else if (wrapType == "wrapNone") {
				result.cssStyle["display"] = "block";
				result.cssStyle["position"] = "relative";
				result.cssStyle["width"] = "0px";
				result.cssStyle["height"] = "0px";
				if (posX.offset) result.cssStyle["left"] = posX.offset;
				if (posY.offset) result.cssStyle["top"] = posY.offset;
			} else if (isAnchor && (posX.align == "left" || posX.align == "right")) result.cssStyle["float"] = posX.align;
			return result;
		}
		parseGraphic(elem) {
			var graphicData = globalXmlParser.element(elem, "graphicData");
			for (let n of globalXmlParser.elements(graphicData)) switch (n.localName) {
				case "pic": return this.parsePicture(n);
			}
			return null;
		}
		parsePicture(elem) {
			var result = {
				type: DomType.Image,
				src: "",
				cssStyle: {}
			};
			var blipFill = globalXmlParser.element(elem, "blipFill");
			var blip = globalXmlParser.element(blipFill, "blip");
			var srcRect = globalXmlParser.element(blipFill, "srcRect");
			result.src = globalXmlParser.attr(blip, "embed");
			if (srcRect) result.srcRect = [
				globalXmlParser.intAttr(srcRect, "l", 0) / 1e5,
				globalXmlParser.intAttr(srcRect, "t", 0) / 1e5,
				globalXmlParser.intAttr(srcRect, "r", 0) / 1e5,
				globalXmlParser.intAttr(srcRect, "b", 0) / 1e5
			];
			var spPr = globalXmlParser.element(elem, "spPr");
			var xfrm = globalXmlParser.element(spPr, "xfrm");
			result.cssStyle["position"] = "relative";
			if (xfrm) {
				result.rotation = globalXmlParser.intAttr(xfrm, "rot", 0) / 6e4;
				for (var n of globalXmlParser.elements(xfrm)) switch (n.localName) {
					case "ext":
						result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
						result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
						break;
					case "off":
						result.cssStyle["left"] = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
						result.cssStyle["top"] = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
				}
			}
			return result;
		}
		parseTable(node) {
			var result = {
				type: DomType.Table,
				children: []
			};
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "tr":
					result.children.push(this.parseTableRow(c));
					break;
				case "tblGrid":
					result.columns = this.parseTableColumns(c);
					break;
				case "tblPr": this.parseTableProperties(c, result);
			}
			return result;
		}
		parseTableColumns(node) {
			var result = [];
			for (const n of globalXmlParser.elements(node)) switch (n.localName) {
				case "gridCol": result.push({ width: globalXmlParser.lengthAttr(n, "w") });
			}
			return result;
		}
		parseTableProperties(elem, table) {
			table.cssStyle = {};
			table.cellStyle = {};
			this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, (c) => {
				switch (c.localName) {
					case "tblStyle":
						table.styleName = globalXmlParser.attr(c, "val");
						break;
					case "tblLook":
						table.className = values.classNameOftblLook(c);
						break;
					case "tblpPr":
						this.parseTablePosition(c, table);
						break;
					case "tblStyleColBandSize":
						table.colBandSize = globalXmlParser.intAttr(c, "val");
						break;
					case "tblStyleRowBandSize":
						table.rowBandSize = globalXmlParser.intAttr(c, "val");
						break;
					case "hidden":
						table.cssStyle["display"] = "none";
						break;
					default: return false;
				}
				return true;
			});
			switch (table.cssStyle["text-align"]) {
				case "center":
					delete table.cssStyle["text-align"];
					table.cssStyle["margin-left"] = "auto";
					table.cssStyle["margin-right"] = "auto";
					break;
				case "right":
					delete table.cssStyle["text-align"];
					table.cssStyle["margin-left"] = "auto";
			}
		}
		parseTablePosition(node, table) {
			var topFromText = globalXmlParser.lengthAttr(node, "topFromText");
			var bottomFromText = globalXmlParser.lengthAttr(node, "bottomFromText");
			var rightFromText = globalXmlParser.lengthAttr(node, "rightFromText");
			var leftFromText = globalXmlParser.lengthAttr(node, "leftFromText");
			var horzAnchor = globalXmlParser.attr(node, "horzAnchor");
			var vertAnchor = globalXmlParser.attr(node, "vertAnchor");
			var tblpXSpec = globalXmlParser.attr(node, "tblpXSpec");
			if (globalXmlParser.attr(node, "tblpYSpec") == "bottom" && vertAnchor != "text") {
				table.cssStyle["position"] = "absolute";
				table.cssStyle["bottom"] = vertAnchor == "page" ? "0px" : "var(--docx-pad-bottom, 0px)";
				if (tblpXSpec == "right") table.cssStyle["right"] = horzAnchor == "page" ? "0px" : "var(--docx-pad-right, 0px)";
				else if (tblpXSpec == "center") {
					table.cssStyle["left"] = "50%";
					table.cssStyle["transform"] = "translateX(-50%)";
				} else table.cssStyle["left"] = horzAnchor == "page" ? "0px" : "var(--docx-pad-left, 0px)";
				return;
			}
			table.cssStyle["float"] = "left";
			table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
			table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
			table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
			table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
		}
		parseTableRow(node) {
			var result = {
				type: DomType.Row,
				children: []
			};
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "tc":
					result.children.push(this.parseTableCell(c));
					break;
				case "trPr":
				case "tblPrEx": this.parseTableRowProperties(c, result);
			}
			return result;
		}
		parseTableRowProperties(elem, row) {
			row.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
				switch (c.localName) {
					case "cnfStyle":
						row.className = values.classNameOfCnfStyle(c);
						break;
					case "tblHeader":
						row.isHeader = globalXmlParser.boolAttr(c, "val");
						break;
					case "gridBefore":
						row.gridBefore = globalXmlParser.intAttr(c, "val");
						break;
					case "gridAfter":
						row.gridAfter = globalXmlParser.intAttr(c, "val");
						break;
					default: return false;
				}
				return true;
			});
		}
		parseTableCell(node) {
			var result = {
				type: DomType.Cell,
				children: []
			};
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "tbl":
					result.children.push(this.parseTable(c));
					break;
				case "p":
					result.children.push(this.parseParagraph(c));
					break;
				case "tcPr": this.parseTableCellProperties(c, result);
			}
			return result;
		}
		parseTableCellProperties(elem, cell) {
			cell.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
				switch (c.localName) {
					case "gridSpan":
						cell.span = globalXmlParser.intAttr(c, "val", null);
						break;
					case "vMerge":
						var _xml$attr2;
						cell.verticalMerge = (_xml$attr2 = globalXmlParser.attr(c, "val")) !== null && _xml$attr2 !== void 0 ? _xml$attr2 : "continue";
						break;
					case "cnfStyle":
						cell.className = values.classNameOfCnfStyle(c);
						break;
					default: return false;
				}
				return true;
			});
			this.parseTableCellVerticalText(elem, cell);
		}
		parseTableCellVerticalText(elem, cell) {
			const directionMap = {
				"btLr": {
					writingMode: "vertical-rl",
					transform: "rotate(180deg)"
				},
				"lrTb": {
					writingMode: "vertical-lr",
					transform: "none"
				},
				"tbRl": {
					writingMode: "vertical-rl",
					transform: "none"
				}
			};
			for (const c of globalXmlParser.elements(elem)) if (c.localName === "textDirection") {
				const style = directionMap[globalXmlParser.attr(c, "val")] || { writingMode: "horizontal-tb" };
				cell.cssStyle["writing-mode"] = style.writingMode;
				cell.cssStyle["transform"] = style.transform;
			}
		}
		parseDefaultProperties(elem, style = null, childStyle = null, handler = null) {
			style = style || {};
			for (const c of globalXmlParser.elements(elem)) {
				if (handler === null || handler === void 0 ? void 0 : handler(c)) continue;
				switch (c.localName) {
					case "jc":
						style["text-align"] = values.valueOfJc(c);
						break;
					case "textAlignment":
						style["vertical-align"] = values.valueOfTextAlignment(c);
						break;
					case "color":
						style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
						break;
					case "sz":
						style["font-size"] = style["min-height"] = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
						break;
					case "shd":
						style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
						break;
					case "highlight":
						style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
						break;
					case "vertAlign": break;
					case "position":
						style.verticalAlign = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
						break;
					case "tcW": if (this.options.ignoreWidth) break;
					case "tblW":
						style["width"] = values.valueOfSize(c, "w");
						break;
					case "trHeight":
						this.parseTrHeight(c, style);
						break;
					case "strike":
						style["text-decoration"] = globalXmlParser.boolAttr(c, "val", true) ? "line-through" : "none";
						break;
					case "b":
						style["font-weight"] = globalXmlParser.boolAttr(c, "val", true) ? "bold" : "normal";
						break;
					case "i":
						style["font-style"] = globalXmlParser.boolAttr(c, "val", true) ? "italic" : "normal";
						break;
					case "caps":
						style["text-transform"] = globalXmlParser.boolAttr(c, "val", true) ? "uppercase" : "none";
						break;
					case "smallCaps":
						style["font-variant"] = globalXmlParser.boolAttr(c, "val", true) ? "small-caps" : "none";
						break;
					case "u":
						this.parseUnderline(c, style);
						break;
					case "tblInd":
						style["margin-inline-start"] = values.valueOfSize(c, "w");
						break;
					case "ind":
						this.parseIndentation(c, style);
						break;
					case "rFonts":
						this.parseFont(c, style);
						break;
					case "tblBorders":
						this.parseBorderProperties(c, childStyle || style);
						break;
					case "tblCellSpacing":
						style["border-spacing"] = values.valueOfMargin(c);
						style["border-collapse"] = "separate";
						break;
					case "pBdr":
						this.parseBorderProperties(c, style);
						break;
					case "bdr":
						style["border"] = values.valueOfBorder(c);
						break;
					case "tcBorders":
						this.parseBorderProperties(c, style);
						break;
					case "vanish":
						if (globalXmlParser.boolAttr(c, "val", true)) style["display"] = "none";
						break;
					case "kern": break;
					case "noWrap": break;
					case "tblCellMar":
					case "tcMar":
						this.parseMarginProperties(c, childStyle || style);
						break;
					case "tblLayout":
						style["table-layout"] = values.valueOfTblLayout(c);
						break;
					case "vAlign":
						style["vertical-align"] = values.valueOfTextAlignment(c);
						break;
					case "spacing":
						if (elem.localName == "pPr") this.parseSpacing(c, style);
						break;
					case "wordWrap":
						if (globalXmlParser.boolAttr(c, "val")) style["overflow-wrap"] = "break-word";
						break;
					case "suppressAutoHyphens":
						style["hyphens"] = globalXmlParser.boolAttr(c, "val", true) ? "none" : "auto";
						break;
					case "lang":
						style["$lang"] = globalXmlParser.attr(c, "val");
						break;
					case "rtl":
					case "bidi":
						if (globalXmlParser.boolAttr(c, "val", true)) style["direction"] = "rtl";
						break;
					case "bCs":
					case "iCs":
					case "szCs":
					case "tabs":
					case "outlineLvl":
					case "contextualSpacing":
					case "tblStyleColBandSize":
					case "tblStyleRowBandSize":
					case "webHidden":
					case "pageBreakBefore":
					case "suppressLineNumbers":
					case "keepLines":
					case "keepNext":
					case "widowControl":
					case "noProof": break;
					default: if (this.options.debug) console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
				}
			}
			return style;
		}
		parseUnderline(node, style) {
			var val = globalXmlParser.attr(node, "val");
			if (val == null) return;
			switch (val) {
				case "dash":
				case "dashDotDotHeavy":
				case "dashDotHeavy":
				case "dashedHeavy":
				case "dashLong":
				case "dashLongHeavy":
				case "dotDash":
				case "dotDotDash":
					style["text-decoration"] = "underline dashed";
					break;
				case "dotted":
				case "dottedHeavy":
					style["text-decoration"] = "underline dotted";
					break;
				case "double":
					style["text-decoration"] = "underline double";
					break;
				case "single":
				case "thick":
					style["text-decoration"] = "underline";
					break;
				case "wave":
				case "wavyDouble":
				case "wavyHeavy":
					style["text-decoration"] = "underline wavy";
					break;
				case "words":
					style["text-decoration"] = "underline";
					break;
				case "none": style["text-decoration"] = "none";
			}
			var col = xmlUtil.colorAttr(node, "color");
			if (col) style["text-decoration-color"] = col;
		}
		parseFont(node, style) {
			var _ref, _ref2, _xml$attr3, _xml$attr4;
			var ascii = (_ref = (_ref2 = (_xml$attr3 = globalXmlParser.attr(node, "ascii")) !== null && _xml$attr3 !== void 0 ? _xml$attr3 : globalXmlParser.attr(node, "hAnsi")) !== null && _ref2 !== void 0 ? _ref2 : values.themeValue(node, "asciiTheme")) !== null && _ref !== void 0 ? _ref : values.themeValue(node, "hAnsiTheme");
			var eastAsia = (_xml$attr4 = globalXmlParser.attr(node, "eastAsia")) !== null && _xml$attr4 !== void 0 ? _xml$attr4 : values.themeValue(node, "eastAsiaTheme");
			if (ascii) style["--docx-font-ascii"] = encloseFontFamily(ascii);
			if (eastAsia) style["--docx-font-ea"] = encloseFontFamily(eastAsia);
		}
		parseIndentation(node, style) {
			var firstLine = globalXmlParser.lengthAttr(node, "firstLine");
			var hanging = globalXmlParser.lengthAttr(node, "hanging");
			var left = globalXmlParser.lengthAttr(node, "left");
			var start = globalXmlParser.lengthAttr(node, "start");
			var right = globalXmlParser.lengthAttr(node, "right");
			var end = globalXmlParser.lengthAttr(node, "end");
			if (firstLine) style["text-indent"] = firstLine;
			if (hanging) style["text-indent"] = `-${hanging}`;
			if (left || start) style["margin-inline-start"] = left || start;
			if (right || end) style["margin-inline-end"] = right || end;
		}
		parseSpacing(node, style) {
			var before = globalXmlParser.lengthAttr(node, "before");
			var after = globalXmlParser.lengthAttr(node, "after");
			var beforeLines = globalXmlParser.intAttr(node, "beforeLines", null);
			var afterLines = globalXmlParser.intAttr(node, "afterLines", null);
			var line = globalXmlParser.intAttr(node, "line", null);
			var lineRule = globalXmlParser.attr(node, "lineRule");
			if (before) style["margin-top"] = before;
			else if (beforeLines != null) style["margin-top"] = `${beforeLines / 100}em`;
			if (after) style["margin-bottom"] = after;
			else if (afterLines != null) style["margin-bottom"] = `${afterLines / 100}em`;
			if (line !== null) switch (lineRule) {
				case "auto":
					style["line-height"] = `${(line / 240).toFixed(2)}`;
					break;
				case "atLeast":
					style["line-height"] = `${line / 20}pt`;
					break;
				default: style["line-height"] = style["min-height"] = `${line / 20}pt`;
			}
		}
		parseMarginProperties(node, output) {
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "left":
					output["padding-left"] = values.valueOfMargin(c);
					break;
				case "right":
					output["padding-right"] = values.valueOfMargin(c);
					break;
				case "top":
					output["padding-top"] = values.valueOfMargin(c);
					break;
				case "bottom": output["padding-bottom"] = values.valueOfMargin(c);
			}
		}
		parseTrHeight(node, output) {
			switch (globalXmlParser.attr(node, "hRule")) {
				case "exact":
					output["height"] = globalXmlParser.lengthAttr(node, "val");
					break;
				default: output["height"] = globalXmlParser.lengthAttr(node, "val");
			}
		}
		parseBorderProperties(node, output) {
			for (const c of globalXmlParser.elements(node)) switch (c.localName) {
				case "start":
				case "left":
					output["border-left"] = values.valueOfBorder(c);
					break;
				case "end":
				case "right":
					output["border-right"] = values.valueOfBorder(c);
					break;
				case "top":
					output["border-top"] = values.valueOfBorder(c);
					break;
				case "bottom": output["border-bottom"] = values.valueOfBorder(c);
			}
		}
	};
	var knownColors = [
		"black",
		"blue",
		"cyan",
		"darkBlue",
		"darkCyan",
		"darkGray",
		"darkGreen",
		"darkMagenta",
		"darkRed",
		"darkYellow",
		"green",
		"lightGray",
		"magenta",
		"none",
		"red",
		"white",
		"yellow"
	];
	var xmlUtil = class {
		static colorAttr(node, attrName, defValue = null, autoColor = "black") {
			var v = globalXmlParser.attr(node, attrName);
			if (v) {
				if (v == "auto") return autoColor;
				else if (knownColors.includes(v)) return v;
				return `#${v}`;
			}
			var themeColor = globalXmlParser.attr(node, "themeColor");
			return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
		}
	};
	var values = class values {
		static themeValue(c, attr) {
			var val = globalXmlParser.attr(c, attr);
			return val ? `var(--docx-${val}-font)` : null;
		}
		static valueOfSize(c, attr) {
			var type = LengthUsage.Dxa;
			switch (globalXmlParser.attr(c, "type")) {
				case "dxa": break;
				case "pct":
					type = LengthUsage.Percent;
					break;
				case "auto": return "auto";
			}
			return globalXmlParser.lengthAttr(c, attr, type);
		}
		static valueOfMargin(c) {
			return globalXmlParser.lengthAttr(c, "w");
		}
		static valueOfBorder(c) {
			var type = values.parseBorderType(globalXmlParser.attr(c, "val"));
			if (type == "none") return "none";
			var color = xmlUtil.colorAttr(c, "color");
			return `${globalXmlParser.lengthAttr(c, "sz", LengthUsage.Border)} ${type} ${color == "auto" ? autos.borderColor : color}`;
		}
		static parseBorderType(type) {
			switch (type) {
				case "single": return "solid";
				case "dashDotStroked": return "solid";
				case "dashed": return "dashed";
				case "dashSmallGap": return "dashed";
				case "dotDash": return "dotted";
				case "dotDotDash": return "dotted";
				case "dotted": return "dotted";
				case "double": return "double";
				case "doubleWave": return "double";
				case "inset": return "inset";
				case "nil": return "none";
				case "none": return "none";
				case "outset": return "outset";
				case "thick": return "solid";
				case "thickThinLargeGap": return "solid";
				case "thickThinMediumGap": return "solid";
				case "thickThinSmallGap": return "solid";
				case "thinThickLargeGap": return "solid";
				case "thinThickMediumGap": return "solid";
				case "thinThickSmallGap": return "solid";
				case "thinThickThinLargeGap": return "solid";
				case "thinThickThinMediumGap": return "solid";
				case "thinThickThinSmallGap": return "solid";
				case "threeDEmboss": return "solid";
				case "threeDEngrave": return "solid";
				case "triple": return "double";
				case "wave": return "solid";
			}
			return "solid";
		}
		static valueOfTblLayout(c) {
			return globalXmlParser.attr(c, "type") == "fixed" ? "fixed" : "auto";
		}
		static classNameOfCnfStyle(c) {
			const val = globalXmlParser.attr(c, "val");
			const classes = [
				"first-row",
				"last-row",
				"first-col",
				"last-col",
				"odd-col",
				"even-col",
				"odd-row",
				"even-row",
				"ne-cell",
				"nw-cell",
				"se-cell",
				"sw-cell"
			];
			if (val) return classes.filter((_, i) => val[i] == "1").join(" ");
			const attrs = [
				"firstRow",
				"lastRow",
				"firstColumn",
				"lastColumn",
				"oddVBand",
				"evenVBand",
				"oddHBand",
				"evenHBand",
				"firstRowLastColumn",
				"firstRowFirstColumn",
				"lastRowLastColumn",
				"lastRowFirstColumn"
			];
			return classes.filter((_, i) => globalXmlParser.boolAttr(c, attrs[i])).join(" ");
		}
		static valueOfJc(c) {
			var type = globalXmlParser.attr(c, "val");
			switch (type) {
				case "start":
				case "left": return "left";
				case "center": return "center";
				case "end":
				case "right": return "right";
				case "both": return "justify";
			}
			return type;
		}
		static valueOfVertAlign(c, asTagName = false) {
			var type = globalXmlParser.attr(c, "val");
			switch (type) {
				case "subscript": return "sub";
				case "superscript": return asTagName ? "sup" : "super";
			}
			return asTagName ? null : type;
		}
		static valueOfTextAlignment(c) {
			var type = globalXmlParser.attr(c, "val");
			switch (type) {
				case "auto":
				case "baseline": return "baseline";
				case "top": return "top";
				case "center": return "middle";
				case "bottom": return "bottom";
			}
			return type;
		}
		static addSize(a, b) {
			if (a == null) return b;
			if (b == null) return a;
			return `calc(${a} + ${b})`;
		}
		static classNameOftblLook(c) {
			const val = globalXmlParser.hexAttr(c, "val", 0);
			let className = "";
			if (globalXmlParser.boolAttr(c, "firstRow") || val & 32) className += " first-row";
			if (globalXmlParser.boolAttr(c, "lastRow") || val & 64) className += " last-row";
			if (globalXmlParser.boolAttr(c, "firstColumn") || val & 128) className += " first-col";
			if (globalXmlParser.boolAttr(c, "lastColumn") || val & 256) className += " last-col";
			if (globalXmlParser.boolAttr(c, "noHBand") || val & 512) className += " no-hband";
			if (globalXmlParser.boolAttr(c, "noVBand") || val & 1024) className += " no-vband";
			return className.trim();
		}
	};
	//#endregion
	//#region src/javascript.ts
	var defaultTab = {
		pos: 0,
		leader: "none",
		style: "left"
	};
	var maxTabs = 50;
	function computePixelToPoint(container = document.body) {
		const temp = document.createElement("div");
		temp.style.width = "100pt";
		container.appendChild(temp);
		const result = 100 / temp.offsetWidth;
		container.removeChild(temp);
		return result;
	}
	function updateTabStop(elem, tabs, defaultTabSize, pixelToPoint = 72 / 96) {
		const p = elem.closest("p");
		const ebb = elem.getBoundingClientRect();
		const pbb = p.getBoundingClientRect();
		const pcs = getComputedStyle(p);
		const tabStops = (tabs === null || tabs === void 0 ? void 0 : tabs.length) > 0 ? tabs.map((t) => ({
			pos: lengthToPoint(t.position),
			leader: t.leader,
			style: t.style
		})).sort((a, b) => a.pos - b.pos) : [defaultTab];
		const lastTab = tabStops[tabStops.length - 1];
		const pWidthPt = pbb.width * pixelToPoint;
		const size = lengthToPoint(defaultTabSize);
		let pos = lastTab.pos + size;
		if (pos < pWidthPt) for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) tabStops.push({
			...defaultTab,
			pos
		});
		const marginLeft = parseFloat(pcs.marginLeft);
		const pOffset = pbb.left - marginLeft; // P3H: tab positions are relative to the text area origin, not twice the paragraph indent.
		const left = (ebb.left - pOffset) * pixelToPoint;
		const tab = tabStops.find((t) => t.style != "clear" && t.pos > left);
		if (tab == null) return;
		let width = 1;
		if (tab.style == "right" || tab.style == "center") {
			const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
			const nextIdx = tabStops.indexOf(elem) + 1;
			const range = document.createRange();
			range.setStart(elem, 1);
			if (nextIdx < tabStops.length) range.setEndBefore(tabStops[nextIdx]);
			else range.setEndAfter(p);
			const mul = tab.style == "center" ? .5 : 1;
			const nextBB = range.getBoundingClientRect();
			const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);
			width = tab.pos - offset * pixelToPoint;
		} else width = tab.pos - left;
		elem.innerHTML = "&nbsp;";
		elem.style.textDecoration = "inherit";
		elem.style.wordSpacing = `${width.toFixed(0)}pt`;
		switch (tab.leader) {
			case "dot":
			case "middleDot":
				elem.style.textDecoration = "underline";
				elem.style.textDecorationStyle = "dotted";
				break;
			case "hyphen":
			case "heavy":
			case "underscore": elem.style.textDecoration = "underline";
		}
	}
	function lengthToPoint(length) {
		return parseFloat(length);
	}
	//#endregion
	//#region src/html.ts
	var ns = /* @__PURE__ */ function(ns) {
		ns["html"] = "http://www.w3.org/1999/xhtml";
		ns["svg"] = "http://www.w3.org/2000/svg";
		ns["mathML"] = "http://www.w3.org/1998/Math/MathML";
		return ns;
	}({});
	function h(elem) {
		if (isString(elem)) return document.createTextNode(elem);
		if (elem instanceof Node) return elem;
		const { ns, tagName, className, style, children, ...props } = elem;
		if (tagName === "#fragment") return document.createDocumentFragment();
		if (tagName === "#comment") return document.createComment(children[0]);
		const result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
		if (className) result.setAttribute("class", className);
		if (style) {
			if (isString(style)) result.setAttribute("style", style);
			else for (const [key, value] of Object.entries(style)) if (key.startsWith("--")) result.style.setProperty(key, value);
			else result.style[key] = value;
		}
		if (props) {
			for (const [key, value] of Object.entries(props)) if (value !== void 0) result[key] = value;
		}
		if (children) children.forEach((c) => result.appendChild(h(c)));
		return result;
	}
	function cx(...classNames) {
		return classNames.filter(Boolean).join(" ");
	}
	//#endregion
	//#region src/html-renderer.ts
	var HtmlRenderer = class {
		constructor() {
			this.className = "docx";
			this.styleMap = {};
			this.currentPart = null;
			this.tableVerticalMerges = [];
			this.currentVerticalMerge = null;
			this.tableCellPositions = [];
			this.currentCellPosition = null;
			this.footnoteMap = {};
			this.endnoteMap = {};
			this.currentEndnoteIds = [];
			this.usedHederFooterParts = [];
			this.currentTabs = [];
			this.legacyCompat = false;
			this.commentMap = {};
			this.tasks = [];
			this.postRenderTasks = [];
			this.h = h;
		}
		async render(document, options) {
			var _options$h;
			this.document = document;
			this.options = options;
			this.className = options.className;
			this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ":root";
			this.h = (_options$h = options.h) !== null && _options$h !== void 0 ? _options$h : h;
			this.styleMap = null;
			this.tasks = [];
			if (this.options.renderComments && globalThis.Highlight) this.commentHighlight = new Highlight();
			const result = [...this.renderDefaultStyle()];
			if (document.themePart) result.push(...this.renderTheme(document.themePart));
			if (document.stylesPart != null) {
				this.styleMap = this.processStyles(document.stylesPart.styles);
				result.push(...this.renderStyles(document.stylesPart.styles));
			}
			if (document.numberingPart) {
				this.prodessNumberings(document.numberingPart.domNumberings);
				result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
			}
			if (document.footnotesPart) this.footnoteMap = keyBy(document.footnotesPart.notes, (x) => x.id);
			if (document.endnotesPart) this.endnoteMap = keyBy(document.endnotesPart.notes, (x) => x.id);
			if (document.settingsPart) {
				var _document$settingsPar, _document$settingsPar2;
				this.defaultTabSize = (_document$settingsPar = document.settingsPart.settings) === null || _document$settingsPar === void 0 ? void 0 : _document$settingsPar.defaultTabStop;
				const compatMode = (_document$settingsPar2 = document.settingsPart.settings) === null || _document$settingsPar2 === void 0 ? void 0 : _document$settingsPar2.compatMode;
				this.legacyCompat = compatMode != null && compatMode <= 11;
			}
			if (!options.ignoreFonts && document.fontTablePart) result.push(...await this.renderFontTable(document.fontTablePart));
			var sectionElements = this.renderSections(document.documentPart.body);
			if (this.options.inWrapper) result.push(this.renderWrapper(sectionElements));
			else result.push(...sectionElements);
			if (this.commentHighlight && options.renderComments) CSS.highlights.set(`${this.className}-comments`, this.commentHighlight);
			this.postRenderTasks.forEach((t) => t());
			await Promise.allSettled(this.tasks);
			this.refreshTabStops();
			return result;
		}
		renderTheme(themePart) {
			var _themePart$theme, _themePart$theme2;
			const variables = {};
			const fontScheme = (_themePart$theme = themePart.theme) === null || _themePart$theme === void 0 ? void 0 : _themePart$theme.fontScheme;
			if (fontScheme) {
				if (fontScheme.majorFont) {
					variables["--docx-majorHAnsi-font"] = fontScheme.majorFont.latinTypeface;
					variables["--docx-majorEastAsia-font"] = fontScheme.majorFont.eaTypeface || fontScheme.majorFont.latinTypeface;
				}
				if (fontScheme.minorFont) {
					variables["--docx-minorHAnsi-font"] = fontScheme.minorFont.latinTypeface;
					variables["--docx-minorEastAsia-font"] = fontScheme.minorFont.eaTypeface || fontScheme.minorFont.latinTypeface;
				}
			}
			const colorScheme = (_themePart$theme2 = themePart.theme) === null || _themePart$theme2 === void 0 ? void 0 : _themePart$theme2.colorScheme;
			if (colorScheme) for (let [k, v] of Object.entries(colorScheme.colors)) variables[`--docx-${k}-color`] = `#${v}`;
			const cssText = this.styleToString(`.${this.className}`, variables);
			return [this.h({
				tagName: "#comment",
				children: ["docxjs document theme values"]
			}), this.h({
				tagName: "style",
				children: [cssText]
			})];
		}
		async renderFontTable(fontsPart) {
			const result = [];
			for (let f of fontsPart.fonts) for (let ref of f.embedFontRefs) try {
				const fontData = await this.document.loadFont(ref.id, ref.key);
				const cssValues = {
					"font-family": encloseFontFamily(f.name),
					"src": `url(${fontData})`
				};
				if (ref.type == "bold" || ref.type == "boldItalic") cssValues["font-weight"] = "bold";
				if (ref.type == "italic" || ref.type == "boldItalic") cssValues["font-style"] = "italic";
				result.push(this.h({
					tagName: "#comment",
					children: [`docxjs ${f.name} font`]
				}));
				result.push(this.h({
					tagName: "style",
					children: [this.styleToString(`@font-face`, cssValues)]
				}));
			} catch (e) {
				if (this.options.debug) console.warn(`Can't load font with id ${ref.id} and key ${ref.key}`);
			}
			return result;
		}
		processStyleName(className) {
			return className ? `${this.className}_${escapeClassName(className)}` : this.className;
		}
		processStyles(styles) {
			const stylesMap = keyBy(styles.filter((x) => x.id != null), (x) => x.id);
			this.defaultParagraphStyle = styles.find((s) => s.isDefault && s.target == "p");
			this.docDefaultsStyle = styles.find((s) => s.id == null);
			for (const style of styles.filter((x) => x.basedOn)) {
				var baseStyle = stylesMap[style.basedOn];
				if (baseStyle) {
					style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
					style.runProps = mergeDeep(style.runProps, baseStyle.runProps);
					for (const baseValues of baseStyle.styles) {
						const styleValues = style.styles.find((x) => x.target == baseValues.target);
						if (styleValues) this.copyStyleProperties(baseValues.values, styleValues.values);
						else style.styles.push({
							...baseValues,
							values: { ...baseValues.values }
						});
					}
				} else if (this.options.debug) console.warn(`Can't find base style ${style.basedOn}`);
			}
			for (let style of styles) style.cssName = this.processStyleName(style.id);
			return stylesMap;
		}
		prodessNumberings(numberings) {
			for (let num of numberings.filter((n) => n.pStyleName)) {
				var _style$paragraphProps;
				const style = this.findStyle(num.pStyleName);
				if (style === null || style === void 0 || (_style$paragraphProps = style.paragraphProps) === null || _style$paragraphProps === void 0 ? void 0 : _style$paragraphProps.numbering) style.paragraphProps.numbering.level = num.level;
			}
		}
		processElement(element) {
			if (element.children) for (var e of element.children) {
				e.parent = element;
				if (e.type == DomType.Table) this.processTable(e);
				else this.processElement(e);
			}
		}
		processTable(table) {
			for (var r of table.children) for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left",
					"border-right",
					"border-top",
					"border-bottom",
					"padding-left",
					"padding-right",
					"padding-top",
					"padding-bottom"
				]);
				this.processElement(c);
			}
		}
		copyStyleProperties(input, output, attrs = null) {
			if (!input) return output;
			if (output == null) output = {};
			if (attrs == null) attrs = Object.getOwnPropertyNames(input);
			for (var key of attrs) if (input.hasOwnProperty(key) && !output.hasOwnProperty(key)) output[key] = input[key];
			return output;
		}
		createPageElement(className, props, docStyle) {
			const style = { ...docStyle };
			if (props) {
				if (props.pageMargins) {
					style.paddingLeft = props.pageMargins.left;
					style.paddingRight = props.pageMargins.right;
					style.paddingTop = props.pageMargins.top;
					style.paddingBottom = props.pageMargins.bottom;
					style["--docx-pad-left"] = props.pageMargins.left;
					style["--docx-pad-right"] = props.pageMargins.right;
					style["--docx-pad-top"] = props.pageMargins.top;
					style["--docx-pad-bottom"] = props.pageMargins.bottom;
				}
				if (props.pageSize) {
					if (!this.options.ignoreWidth) style.width = props.pageSize.width;
					if (!this.options.ignoreHeight) style.minHeight = props.pageSize.height;
				}
			}
			return this.h({
				tagName: "section",
				className,
				style
			});
		}
		createSectionContent(props) {
			const style = {};
			if (props.columns && props.columns.numberOfColumns) {
				style.columnCount = `${props.columns.numberOfColumns}`;
				style.columnGap = props.columns.space;
				if (props.columns.separator) style.columnRule = "1px solid black";
			}
			return this.h({
				tagName: "article",
				style
			});
		}
		renderSections(document) {
			const result = [];
			this.processElement(document);
			const sections = this.splitBySection(document.children, document.props);
			const pages = this.groupByPageBreaks(sections);
			let prevProps = null;
			for (let i = 0, l = pages.length; i < l; i++) {
				this.currentFootnoteIds = [];
				let props = pages[i][0].sectProps;
				const pageElement = this.createPageElement(this.className, props, document.cssStyle);
				this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props, result.length, prevProps != props, pageElement);
				for (const sect of pages[i]) {
					var contentElement = this.createSectionContent(sect.sectProps);
					this.renderElements(sect.elements, contentElement);
					pageElement.appendChild(contentElement);
					props = sect.sectProps;
				}
				if (this.options.renderFootnotes) {
					const notes = this.renderNotes(this.currentFootnoteIds, this.footnoteMap);
					notes && pageElement.appendChild(notes);
				}
				if (this.options.renderEndnotes && i == l - 1) {
					const notes = this.renderNotes(this.currentEndnoteIds, this.endnoteMap);
					notes && pageElement.appendChild(notes);
				}
				this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props, result.length, prevProps != props, pageElement);
				result.push(pageElement);
				prevProps = props;
			}
			this.resolvePageFields(result);
			return result;
		}
		renderHeaderFooter(refs, props, page, firstOfSection, into) {
			var _ref, _ref2;
			if (!refs) return;
			var ref = (_ref = (_ref2 = props.titlePage && firstOfSection ? refs.find((x) => x.type == "first") : null) !== null && _ref2 !== void 0 ? _ref2 : page % 2 == 1 ? refs.find((x) => x.type == "even") : null) !== null && _ref !== void 0 ? _ref : refs.find((x) => x.type == "default");
			var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart);
			if (part) {
				this.currentPart = part;
				if (!this.usedHederFooterParts.includes(part.path)) {
					this.processElement(part.rootElement);
					this.usedHederFooterParts.push(part.path);
				}
				const [el] = this.renderElements([part.rootElement], into);
				if (props === null || props === void 0 ? void 0 : props.pageMargins) {
					if (part.rootElement.type === DomType.Header) {
						el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
						el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
					} else if (part.rootElement.type === DomType.Footer) {
						el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
						el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
					}
				}
				this.currentPart = null;
			}
		}
		isPageBreakElement(elem) {
			if (elem.type != DomType.Break) return false;
			if (elem.break == "lastRenderedPageBreak") return !this.options.ignoreLastRenderedPageBreak;
			return elem.break == "page";
		}
		isPageBreakSection(prev, next) {
			var _prev$pageSize, _next$pageSize, _prev$pageSize2, _next$pageSize2, _prev$pageSize3, _next$pageSize3;
			if (!prev) return false;
			if (!next) return false;
			return ((_prev$pageSize = prev.pageSize) === null || _prev$pageSize === void 0 ? void 0 : _prev$pageSize.orientation) != ((_next$pageSize = next.pageSize) === null || _next$pageSize === void 0 ? void 0 : _next$pageSize.orientation) || ((_prev$pageSize2 = prev.pageSize) === null || _prev$pageSize2 === void 0 ? void 0 : _prev$pageSize2.width) != ((_next$pageSize2 = next.pageSize) === null || _next$pageSize2 === void 0 ? void 0 : _next$pageSize2.width) || ((_prev$pageSize3 = prev.pageSize) === null || _prev$pageSize3 === void 0 ? void 0 : _prev$pageSize3.height) != ((_next$pageSize3 = next.pageSize) === null || _next$pageSize3 === void 0 ? void 0 : _next$pageSize3.height);
		}
		splitBySection(elements, defaultProps) {
			var current = {
				sectProps: null,
				elements: [],
				pageBreak: false
			};
			var result = [current];
			for (let elem of elements) {
				if (elem.type == DomType.Paragraph) {
					var _s$paragraphProps;
					const s = this.findStyle(elem.styleName);
					if (s === null || s === void 0 || (_s$paragraphProps = s.paragraphProps) === null || _s$paragraphProps === void 0 ? void 0 : _s$paragraphProps.pageBreakBefore) {
						current.sectProps = sectProps;
						current.pageBreak = true;
						current = {
							sectProps: null,
							elements: [],
							pageBreak: false
						};
						result.push(current);
					}
				}
				current.elements.push(elem);
				if (elem.type == DomType.Paragraph) {
					const p = elem;
					var sectProps = p.sectionProps;
					var pBreakIndex = -1;
					var rBreakIndex = -1;
					if (this.options.breakPages && p.children) pBreakIndex = p.children.findIndex((r) => {
						var _r$children$findIndex, _r$children;
						rBreakIndex = (_r$children$findIndex = (_r$children = r.children) === null || _r$children === void 0 ? void 0 : _r$children.findIndex(this.isPageBreakElement.bind(this))) !== null && _r$children$findIndex !== void 0 ? _r$children$findIndex : -1;
						return rBreakIndex != -1;
					});
					if (sectProps || pBreakIndex != -1) {
						current.sectProps = sectProps;
						current.pageBreak = pBreakIndex != -1;
						current = {
							sectProps: null,
							elements: [],
							pageBreak: false
						};
						result.push(current);
					}
					if (pBreakIndex != -1) {
						let breakRun = p.children[pBreakIndex];
						let splitRun = rBreakIndex < breakRun.children.length - 1;
						if (pBreakIndex < p.children.length - 1 || splitRun) {
							var children = elem.children;
							var newParagraph = {
								...elem,
								children: children.slice(pBreakIndex)
							};
							elem.children = children.slice(0, pBreakIndex);
							current.elements.push(newParagraph);
							if (splitRun) {
								let runChildren = breakRun.children;
								let newRun = {
									...breakRun,
									children: runChildren.slice(0, rBreakIndex)
								};
								elem.children.push(newRun);
								breakRun.children = runChildren.slice(rBreakIndex);
							}
						}
					}
				}
			}
			let currentSectProps = null;
			for (let i = result.length - 1; i >= 0; i--) if (result[i].sectProps == null) {
				var _currentSectProps;
				result[i].sectProps = (_currentSectProps = currentSectProps) !== null && _currentSectProps !== void 0 ? _currentSectProps : defaultProps;
			} else currentSectProps = result[i].sectProps;
			return result;
		}
		groupByPageBreaks(sections) {
			let current = [];
			let prev;
			const result = [current];
			for (let s of sections) {
				current.push(s);
				if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps)) result.push(current = []);
				prev = s.sectProps;
			}
			return result.filter((x) => x.length > 0);
		}
		renderWrapper(children) {
			return this.h({
				tagName: "div",
				className: `${this.className}-wrapper`,
				children
			});
		}
		renderDefaultStyle() {
			var c = this.className;
			var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
			if (this.options.hideWrapperOnPrint) wrapperStyle = `@media not print { ${wrapperStyle} }`;
			var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
.${c} { --docx-font-ascii: serif; --docx-font-ea: serif; }
.${c}, .${c} * { font-family: var(--docx-font-ascii), var(--docx-font-ea); }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;
			if (this.options.renderComments) styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`;
			return [this.h({
				tagName: "#comment",
				children: ["docxjs library predefined styles"]
			}), this.h({
				tagName: "style",
				children: [styleText]
			})];
		}
		async renderNumbering(numberings) {
			var styleText = "";
			var resetCounters = [];
			for (var num of numberings) {
				var selector = `p.${this.numberingClass(num.id, num.level)}`;
				var listStyleType = "none";
				if (num.bullet) {
					let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
					styleText += this.styleToString(`${selector}:before`, {
						"content": "' '",
						"display": "inline-block",
						"background": `var(${valiable})`
					}, num.bullet.style);
					try {
						const imgData = await this.document.loadNumberingImage(num.bullet.src);
						styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
					} catch (e) {
						if (this.options.debug) console.warn(`Can't load numbering image with src ${num.bullet.src}`);
					}
				} else if (num.levelText) {
					let counter = this.numberingCounter(num.id, num.level);
					const counterReset = counter + " " + (num.start - 1);
					if (num.level > 0) styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, { "counter-set": counterReset });
					resetCounters.push(counterReset);
					styleText += this.styleToString(`${selector}:before`, {
						"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
						"counter-increment": counter,
						...num.rStyle
					});
				} else listStyleType = this.numFormatToCssValue(num.format);
				styleText += this.styleToString(selector, {
					"display": "list-item",
					"list-style-position": "inside",
					"list-style-type": listStyleType,
					...num.pStyle
				});
			}
			if (resetCounters.length > 0) styleText += this.styleToString(this.rootSelector, { "counter-reset": resetCounters.join(" ") });
			return [this.h({
				tagName: "#comment",
				children: ["docxjs document numbering styles"]
			}), this.h({
				tagName: "style",
				children: [styleText]
			})];
		}
		renderStyles(styles) {
			var styleText = "";
			const stylesMap = this.styleMap;
			const defautStyles = keyBy(styles.filter((s) => s.isDefault), (s) => s.target);
			for (const style of styles) {
				var subStyles = style.styles;
				if (style.linked) {
					var linkedStyle = style.linked && stylesMap[style.linked];
					if (linkedStyle) subStyles = subStyles.concat(linkedStyle.styles);
					else if (this.options.debug) console.warn(`Can't find linked style ${style.linked}`);
				}
				for (const subStyle of subStyles) {
					var _style$target;
					var selector = `${(_style$target = style.target) !== null && _style$target !== void 0 ? _style$target : ""}.${style.cssName}`;
					if (style.target != subStyle.target) selector += ` ${subStyle.target}`;
					if (style.id == null && subStyle.target == "span") selector = `.${this.className}`;
					if (defautStyles[style.target] == style) selector = `.${this.className} ${style.target}, ` + selector;
					styleText += this.styleToString(selector, subStyle.values);
				}
			}
			return [this.h({
				tagName: "#comment",
				children: ["docxjs document styles"]
			}), this.h({
				tagName: "style",
				children: [styleText]
			})];
		}
		renderNotes(noteIds, notesMap) {
			var notes = noteIds.map((id) => notesMap[id]).filter((x) => x);
			if (notes.length > 0) return this.h({
				tagName: "ol",
				children: this.renderElements(notes)
			});
		}
		renderElement(elem) {
			switch (elem.type) {
				case DomType.Paragraph: return this.renderParagraph(elem);
				case DomType.BookmarkStart: return this.renderBookmarkStart(elem);
				case DomType.BookmarkEnd: return null;
				case DomType.Run: return this.renderRun(elem);
				case DomType.Table: return this.renderTable(elem);
				case DomType.Row: return this.renderTableRow(elem);
				case DomType.Cell: return this.renderTableCell(elem);
				case DomType.Hyperlink: return this.renderHyperlink(elem);
				case DomType.SmartTag: return this.renderSmartTag(elem);
				case DomType.Drawing: return this.renderDrawing(elem);
				case DomType.Image: return this.renderImage(elem);
				case DomType.Text: return this.renderText(elem);
				case DomType.Text: return this.renderText(elem);
				case DomType.DeletedText: return this.renderDeletedText(elem);
				case DomType.Tab: return this.renderTab(elem);
				case DomType.Symbol: return this.renderSymbol(elem);
				case DomType.Break: return this.renderBreak(elem);
				case DomType.Footer: return this.renderContainer(elem, "footer");
				case DomType.Header: return this.renderContainer(elem, "header");
				case DomType.Footnote:
				case DomType.Endnote: return this.renderContainer(elem, "li");
				case DomType.FootnoteReference: return this.renderFootnoteReference(elem);
				case DomType.EndnoteReference: return this.renderEndnoteReference(elem);
				case DomType.NoBreakHyphen: return this.h({ tagName: "wbr" });
				case DomType.VmlPicture: return this.renderVmlPicture(elem);
				case DomType.VmlElement: return this.renderVmlElement(elem);
				case DomType.MmlMath: return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
				case DomType.MmlMathParagraph: return this.renderContainer(elem, "span");
				case DomType.MmlFraction: return this.renderContainerNS(elem, ns.mathML, "mfrac");
				case DomType.MmlBase: return this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
				case DomType.MmlNumerator:
				case DomType.MmlDenominator:
				case DomType.MmlFunction:
				case DomType.MmlLimit:
				case DomType.MmlBox: return this.renderContainerNS(elem, ns.mathML, "mrow");
				case DomType.MmlGroupChar: return this.renderMmlGroupChar(elem);
				case DomType.MmlLimitLower: return this.renderContainerNS(elem, ns.mathML, "munder");
				case DomType.MmlMatrix: return this.renderContainerNS(elem, ns.mathML, "mtable");
				case DomType.MmlMatrixRow: return this.renderContainerNS(elem, ns.mathML, "mtr");
				case DomType.MmlRadical: return this.renderMmlRadical(elem);
				case DomType.MmlSuperscript: return this.renderContainerNS(elem, ns.mathML, "msup");
				case DomType.MmlSubscript: return this.renderContainerNS(elem, ns.mathML, "msub");
				case DomType.MmlDegree:
				case DomType.MmlSuperArgument:
				case DomType.MmlSubArgument: return this.renderContainerNS(elem, ns.mathML, "mn");
				case DomType.MmlFunctionName: return this.renderContainerNS(elem, ns.mathML, "ms");
				case DomType.MmlDelimiter: return this.renderMmlDelimiter(elem);
				case DomType.MmlRun: return this.renderMmlRun(elem);
				case DomType.MmlNary: return this.renderMmlNary(elem);
				case DomType.MmlPreSubSuper: return this.renderMmlPreSubSuper(elem);
				case DomType.MmlBar: return this.renderMmlBar(elem);
				case DomType.MmlEquationArray: return this.renderMllList(elem);
				case DomType.Inserted: return this.renderInserted(elem);
				case DomType.Deleted: return this.renderDeleted(elem);
				case DomType.CommentRangeStart: return this.renderCommentRangeStart(elem);
				case DomType.CommentRangeEnd: return this.renderCommentRangeEnd(elem);
				case DomType.CommentReference: return this.renderCommentReference(elem);
				case DomType.AltChunk: return this.renderAltChunk(elem);
			}
			return null;
		}
		renderElements(elems, into) {
			if (elems == null) return null;
			var result = elems.flatMap((e) => this.renderElement(e)).filter((e) => e != null);
			if (into) result.forEach((c) => into.appendChild(isString(c) ? document.createTextNode(c) : c));
			return result;
		}
		renderContainer(elem, tagName, props) {
			return this.h({
				tagName,
				children: this.renderElements(elem.children),
				...props
			});
		}
		renderContainerNS(elem, ns, tagName, props) {
			return this.h({
				ns,
				tagName,
				children: this.renderElements(elem.children),
				...props
			});
		}
		/**
		* Marks the cached-result runs of PAGE / NUMPAGES complex fields with a
		* marker class (`docx-field-page` / `docx-field-numpages`), so the actual
		* numbers can be substituted once page layout is final (see the end of
		* renderSections and pagination.ts). Field runs themselves are skipped
		* by renderRun, leaving the cached value as a plain run between the
		* separate and end markers.
		*/
		markPageFieldRuns(elem) {
			var _elem$children;
			let instruction = null;
			let inResult = false;
			for (const run of (_elem$children = elem.children) !== null && _elem$children !== void 0 ? _elem$children : []) {
				var _run$children;
				if (run.type != DomType.Run) continue;
				for (const sub of (_run$children = run.children) !== null && _run$children !== void 0 ? _run$children : []) if (sub.type == DomType.ComplexField) {
					const charType = sub.charType;
					if (charType == "begin") {
						instruction = null;
						inResult = false;
					} else if (charType == "separate") inResult = true;
					else if (charType == "end") {
						instruction = null;
						inResult = false;
					}
				} else if (sub.type == DomType.Instruction) {
					var _sub$text;
					instruction = (_sub$text = sub.text) === null || _sub$text === void 0 || (_sub$text = _sub$text.trim().split(/\s/)[0]) === null || _sub$text === void 0 ? void 0 : _sub$text.toUpperCase();
				}
				if (inResult && (instruction == "PAGE" || instruction == "NUMPAGES") && !run.fieldRun) run.className = cx(run.className, `${this.className}-field-${instruction.toLowerCase()}`);
			}
		}
		/**
		* Substitutes final page numbers into PAGE / NUMPAGES field markers
		* across the given page sections. Numbering is continuous from 1 —
		* per-section restarts (w:pgNumType) are not modeled yet.
		*/
		resolvePageFields(pages) {
			const total = `${pages.length}`;
			pages.forEach((page, i) => {
				for (const el of Array.from(page.querySelectorAll(`.${this.className}-field-page`))) el.textContent = `${i + 1}`;
				for (const el of Array.from(page.querySelectorAll(`.${this.className}-field-numpages`))) el.textContent = total;
			});
		}
		renderParagraph(elem) {
			var _elem$tabs, _style$paragraphProps2, _elem$numbering, _style$paragraphProps3;
			this.markPageFieldRuns(elem);
			var result = this.toHTML(elem, ns.html, "p");
			const style = this.findStyle(elem.styleName);
			(_elem$tabs = elem.tabs) !== null && _elem$tabs !== void 0 || (elem.tabs = style === null || style === void 0 || (_style$paragraphProps2 = style.paragraphProps) === null || _style$paragraphProps2 === void 0 ? void 0 : _style$paragraphProps2.tabs);
			const numbering = (_elem$numbering = elem.numbering) !== null && _elem$numbering !== void 0 ? _elem$numbering : style === null || style === void 0 || (_style$paragraphProps3 = style.paragraphProps) === null || _style$paragraphProps3 === void 0 ? void 0 : _style$paragraphProps3.numbering;
			if (numbering) result.classList.add(this.numberingClass(numbering.id, numbering.level));
			if (this.legacyCompat && this.textAlignOf(elem) === "justify" && / {2,}| {2,}|　/.test(result.textContent)) result.style.setProperty("text-align-last", "justify");
			return result;
		}
		/** Effective text-align of a paragraph: direct formatting, then its
		* style, then the default paragraph style, then docDefaults. */
		textAlignOf(elem) {
			var _elem$cssStyle, _ref3, _ref4, _from;
			if ((_elem$cssStyle = elem.cssStyle) === null || _elem$cssStyle === void 0 ? void 0 : _elem$cssStyle["text-align"]) return elem.cssStyle["text-align"];
			const from = (s) => {
				var _s$styles;
				return s === null || s === void 0 || (_s$styles = s.styles) === null || _s$styles === void 0 || (_s$styles = _s$styles.find((x) => x.target == "p")) === null || _s$styles === void 0 || (_s$styles = _s$styles.values) === null || _s$styles === void 0 ? void 0 : _s$styles["text-align"];
			};
			return (_ref3 = (_ref4 = (_from = from(this.findStyle(elem.styleName))) !== null && _from !== void 0 ? _from : from(this.defaultParagraphStyle)) !== null && _ref4 !== void 0 ? _ref4 : from(this.docDefaultsStyle)) !== null && _ref3 !== void 0 ? _ref3 : null;
		}
		renderHyperlink(elem) {
			const res = this.toH(elem, ns.html, "a");
			res.href = "";
			if (elem.id) {
				var _rel$target;
				const rel = this.document.documentPart.rels.find((it) => it.id == elem.id && it.targetMode === "External");
				res.href = (_rel$target = rel === null || rel === void 0 ? void 0 : rel.target) !== null && _rel$target !== void 0 ? _rel$target : res.href;
			}
			if (elem.anchor) res.href += `#${elem.anchor}`;
			return this.h(res);
		}
		renderSmartTag(elem) {
			return this.renderContainer(elem, "span");
		}
		renderCommentRangeStart(commentStart) {
			var _this$commentHighligh;
			if (!this.options.renderComments) return null;
			const rng = new Range();
			(_this$commentHighligh = this.commentHighlight) === null || _this$commentHighligh === void 0 || _this$commentHighligh.add(rng);
			const result = this.h({
				tagName: "#comment",
				children: [`start of comment #${commentStart.id}`]
			});
			this.later(() => rng.setStart(result, 0));
			this.commentMap[commentStart.id] = rng;
			return result;
		}
		renderCommentRangeEnd(commentEnd) {
			if (!this.options.renderComments) return null;
			const rng = this.commentMap[commentEnd.id];
			const result = this.h({
				tagName: "#comment",
				children: [`end of comment #${commentEnd.id}`]
			});
			this.later(() => rng === null || rng === void 0 ? void 0 : rng.setEnd(result, 0));
			return result;
		}
		renderCommentReference(commentRef) {
			var _this$document$commen;
			if (!this.options.renderComments) return null;
			var comment = (_this$document$commen = this.document.commentsPart) === null || _this$document$commen === void 0 ? void 0 : _this$document$commen.commentMap[commentRef.id];
			if (!comment) return null;
			const commentRefEl = this.h({
				tagName: "span",
				className: `${this.className}-comment-ref`,
				children: ["💬"]
			});
			const commentsContainerEl = this.h({
				tagName: "div",
				className: `${this.className}-comment-popover`,
				children: [
					this.h({
						tagName: "div",
						className: `${this.className}-comment-author`,
						children: [comment.author]
					}),
					this.h({
						tagName: "div",
						className: `${this.className}-comment-date`,
						children: [new Date(comment.date).toLocaleString()]
					}),
					...this.renderElements(comment.children)
				]
			});
			return this.h({
				tagName: "#fragment",
				children: [
					this.h({
						tagName: "#comment",
						children: [`comment #${comment.id} by ${comment.author} on ${comment.date}`]
					}),
					commentRefEl,
					commentsContainerEl
				]
			});
		}
		renderAltChunk(elem) {
			if (!this.options.renderAltChunks) return null;
			var result = this.h({ tagName: "iframe" });
			this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then((x) => {
				result.srcdoc = x;
			}));
			return result;
		}
		renderDrawing(elem) {
			var result = this.toHTML(elem, ns.html, "div");
			result.style.display = "inline-block";
			result.style.position = "relative";
			result.style.textIndent = "0px";
			return result;
		}
		renderImage(elem) {
			var _elem$cssStyle2, _transform;
			let result = this.toHTML(elem, ns.html, "img", []);
			let transform = (_elem$cssStyle2 = elem.cssStyle) === null || _elem$cssStyle2 === void 0 ? void 0 : _elem$cssStyle2.transform;
			if (elem.srcRect && elem.srcRect.some((x) => x != 0)) {
				var [left, top, right, bottom] = elem.srcRect;
				transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
				result.style["clip-path"] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
			}
			if (elem.rotation) transform = `rotate(${elem.rotation}deg) ${(_transform = transform) !== null && _transform !== void 0 ? _transform : ""}`;
			result.style.transform = transform === null || transform === void 0 ? void 0 : transform.trim();
			if (this.document) this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then((x) => {
				result.src = x;
			}));
			return result;
		}
		renderText(elem) {
			return this.h(elem.text);
		}
		renderDeletedText(elem) {
			return this.options.renderChanges ? this.renderText(elem) : null;
		}
		renderBreak(elem) {
			return elem.break == "textWrapping" ? this.h({ tagName: "br" }) : null;
		}
		renderInserted(elem) {
			if (this.options.renderChanges) return this.renderChange(elem, "ins");
			return this.renderElements(elem.children);
		}
		renderDeleted(elem) {
			if (this.options.renderChanges) return this.renderChange(elem, "del");
			return null;
		}
		renderChange(elem, tag) {
			return this.renderContainer(elem, tag, { dateTime: elem.date });
		}
		renderSymbol(elem) {
			return this.h({
				tagName: "span",
				children: [String.fromCharCode(elem.char)],
				style: { fontFamily: elem.font }
			});
		}
		renderFootnoteReference(elem) {
			this.currentFootnoteIds.push(elem.id);
			return this.h({
				tagName: "sup",
				children: [`${this.currentFootnoteIds.length}`]
			});
		}
		renderEndnoteReference(elem) {
			this.currentEndnoteIds.push(elem.id);
			return this.h({
				tagName: "sup",
				children: [`${this.currentEndnoteIds.length}`]
			});
		}
		renderTab(elem) {
			var tabSpan = this.h({
				tagName: "span",
				children: [" "]
			});
			if (this.options.experimental) {
				var _findParent;
				tabSpan.className = this.tabStopClass();
				var stops = (_findParent = findParent(elem, DomType.Paragraph)) === null || _findParent === void 0 ? void 0 : _findParent.tabs;
				this.currentTabs.push({
					stops,
					span: tabSpan
				});
			}
			return tabSpan;
		}
		renderBookmarkStart(elem) {
			return this.h({
				tagName: "span",
				id: elem.name
			});
		}
		renderRun(elem) {
			if (elem.fieldRun) return null;
			let children = this.renderElements(elem.children);
			if (elem.verticalAlign) children = [this.h({
				tagName: elem.verticalAlign,
				children: this.renderElements(elem.children)
			})];
			const result = this.toHTML(elem, ns.html, "span", children);
			if (elem.id) result.id = elem.id;
			return result;
		}
		renderTable(elem) {
			this.tableCellPositions.push(this.currentCellPosition);
			this.tableVerticalMerges.push(this.currentVerticalMerge);
			this.currentVerticalMerge = {};
			this.currentCellPosition = {
				col: 0,
				row: 0
			};
			const children = [];
			if (elem.columns) children.push(this.renderTableColumns(elem.columns));
			children.push(...this.renderElements(elem.children));
			this.currentVerticalMerge = this.tableVerticalMerges.pop();
			this.currentCellPosition = this.tableCellPositions.pop();
			return this.toHTML(elem, ns.html, "table", children);
		}
		renderTableColumns(columns) {
			const children = columns.map((x) => this.h({
				tagName: "col",
				style: { width: x.width }
			}));
			return this.h({
				tagName: "colgroup",
				children
			});
		}
		renderTableRow(elem) {
			this.currentCellPosition.col = 0;
			const children = [];
			if (elem.gridBefore) children.push(this.renderTableCellPlaceholder(elem.gridBefore));
			children.push(...this.renderElements(elem.children));
			if (elem.gridAfter) children.push(this.renderTableCellPlaceholder(elem.gridAfter));
			this.currentCellPosition.row++;
			return this.toHTML(elem, ns.html, "tr", children);
		}
		renderTableCellPlaceholder(colSpan) {
			return this.h({
				tagName: "td",
				colSpan,
				style: { border: "none" }
			});
		}
		renderTableCell(elem) {
			let result = this.toHTML(elem, ns.html, "td");
			const key = this.currentCellPosition.col;
			if (elem.verticalMerge) {
				if (elem.verticalMerge == "restart") {
					this.currentVerticalMerge[key] = result;
					result.rowSpan = 1;
				} else if (this.currentVerticalMerge[key]) {
					this.currentVerticalMerge[key].rowSpan += 1;
					result.style.display = "none";
				}
			} else this.currentVerticalMerge[key] = null;
			if (elem.span) result.colSpan = elem.span;
			this.currentCellPosition.col += result.colSpan;
			return result;
		}
		renderVmlPicture(elem) {
			return this.renderContainer(elem, "div");
		}
		renderVmlElement(elem) {
			var _elem$imageHref;
			var container = this.h({
				ns: ns.svg,
				tagName: "svg",
				style: elem.cssStyleText
			});
			const result = this.renderVmlChildElement(elem);
			if ((_elem$imageHref = elem.imageHref) === null || _elem$imageHref === void 0 ? void 0 : _elem$imageHref.id) {
				var _this$document;
				this.tasks.push((_this$document = this.document) === null || _this$document === void 0 ? void 0 : _this$document.loadDocumentImage(elem.imageHref.id, this.currentPart).then((x) => result.setAttribute("href", x)));
			}
			container.appendChild(result);
			requestAnimationFrame(() => {
				const bb = container.firstElementChild.getBBox();
				container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
				container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
			});
			return container;
		}
		renderVmlChildElement(elem) {
			const result = this.createSvgElement(elem.tagName);
			Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));
			for (let child of elem.children) if (child.type == DomType.VmlElement) result.appendChild(this.renderVmlChildElement(child));
			else result.appendChild(...asArray(this.renderElement(child)));
			return result;
		}
		renderMmlRadical(elem) {
			var _elem$props;
			const base = elem.children.find((el) => el.type == DomType.MmlBase);
			if ((_elem$props = elem.props) === null || _elem$props === void 0 ? void 0 : _elem$props.hideDegree) return this.createMathMLElement("msqrt", null, this.renderElements([base]));
			const degree = elem.children.find((el) => el.type == DomType.MmlDegree);
			return this.createMathMLElement("mroot", null, this.renderElements([base, degree]));
		}
		renderMmlDelimiter(elem) {
			var _elem$props$beginChar, _elem$props$endChar;
			const children = [];
			children.push(this.createMathMLElement("mo", null, [(_elem$props$beginChar = elem.props.beginChar) !== null && _elem$props$beginChar !== void 0 ? _elem$props$beginChar : "("]));
			children.push(...this.renderElements(elem.children));
			children.push(this.createMathMLElement("mo", null, [(_elem$props$endChar = elem.props.endChar) !== null && _elem$props$endChar !== void 0 ? _elem$props$endChar : ")"]));
			return this.createMathMLElement("mrow", null, children);
		}
		renderMmlNary(elem) {
			var _elem$props$char, _elem$props2;
			const children = [];
			const grouped = keyBy(elem.children, (x) => x.type);
			const sup = grouped[DomType.MmlSuperArgument];
			const sub = grouped[DomType.MmlSubArgument];
			const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
			const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
			const charElem = this.createMathMLElement("mo", null, [(_elem$props$char = (_elem$props2 = elem.props) === null || _elem$props2 === void 0 ? void 0 : _elem$props2.char) !== null && _elem$props$char !== void 0 ? _elem$props$char : "∫"]);
			if (supElem || subElem) children.push(this.createMathMLElement("munderover", null, [
				charElem,
				subElem,
				supElem
			]));
			else if (supElem) children.push(this.createMathMLElement("mover", null, [charElem, supElem]));
			else if (subElem) children.push(this.createMathMLElement("munder", null, [charElem, subElem]));
			else children.push(charElem);
			children.push(...this.renderElements(grouped[DomType.MmlBase].children));
			return this.createMathMLElement("mrow", null, children);
		}
		renderMmlPreSubSuper(elem) {
			const children = [];
			const grouped = keyBy(elem.children, (x) => x.type);
			const sup = grouped[DomType.MmlSuperArgument];
			const sub = grouped[DomType.MmlSubArgument];
			const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
			const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
			const stubElem = this.createMathMLElement("mo", null);
			children.push(this.createMathMLElement("msubsup", null, [
				stubElem,
				subElem,
				supElem
			]));
			children.push(...this.renderElements(grouped[DomType.MmlBase].children));
			return this.createMathMLElement("mrow", null, children);
		}
		renderMmlGroupChar(elem) {
			const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
			const result = this.renderContainerNS(elem, ns.mathML, tagName);
			if (elem.props.char) result.appendChild(this.createMathMLElement("mo", null, [elem.props.char]));
			return result;
		}
		renderMmlBar(elem) {
			const style = {};
			switch (elem.props.position) {
				case "top":
					style.textDecoration = "overline";
					break;
				case "bottom": style.textDecoration = "underline";
			}
			return this.renderContainerNS(elem, ns.mathML, "mrow", { style });
		}
		renderMmlRun(elem) {
			return this.toHTML(elem, ns.mathML, "ms");
		}
		renderMllList(elem) {
			const children = this.renderElements(elem.children).map((x) => this.createMathMLElement("mtr", null, [this.createMathMLElement("mtd", null, [x])]));
			return this.toHTML(elem, ns.mathML, "mtable", children);
		}
		toH(elem, ns, tagName, children = null) {
			var _elem$cssStyle3;
			const { "$lang": lang, ...style } = (_elem$cssStyle3 = elem.cssStyle) !== null && _elem$cssStyle3 !== void 0 ? _elem$cssStyle3 : {};
			return {
				ns,
				tagName,
				className: cx(elem.className, elem.styleName && this.processStyleName(elem.styleName)),
				lang,
				style,
				children: children !== null && children !== void 0 ? children : this.renderElements(elem.children)
			};
		}
		toHTML(elem, ns, tagName, children = null) {
			return this.h(this.toH(elem, ns, tagName, children));
		}
		findStyle(styleName) {
			var _this$styleMap;
			return styleName && ((_this$styleMap = this.styleMap) === null || _this$styleMap === void 0 ? void 0 : _this$styleMap[styleName]);
		}
		numberingClass(id, lvl) {
			return `${this.className}-num-${id}-${lvl}`;
		}
		tabStopClass() {
			return `${this.className}-tab-stop`;
		}
		styleToString(selectors, values, cssText = null) {
			let result = `${selectors} {\r\n`;
			for (const key in values) {
				if (key.startsWith("$")) continue;
				result += `  ${key}: ${values[key]};\r\n`;
			}
			if (cssText) result += cssText;
			return result + "}\r\n";
		}
		numberingCounter(id, lvl) {
			return `${this.className}-num-${id}-${lvl}`;
		}
		levelTextToContent(text, suff, id, numformat) {
			var _suffMap$suff;
			return `"${text.replace(/%\d*/g, (s) => {
				let lvl = parseInt(s.substring(1), 10) - 1;
				return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
			})}${(_suffMap$suff = {
				"tab": "\\9",
				"space": "\\a0"
			}[suff]) !== null && _suffMap$suff !== void 0 ? _suffMap$suff : ""}"`;
		}
		numFormatToCssValue(format) {
			var _mapping$format;
			return (_mapping$format = {
				none: "none",
				bullet: "disc",
				decimal: "decimal",
				lowerLetter: "lower-alpha",
				upperLetter: "upper-alpha",
				lowerRoman: "lower-roman",
				upperRoman: "upper-roman",
				decimalZero: "decimal-leading-zero",
				aiueo: "katakana",
				aiueoFullWidth: "katakana",
				chineseCounting: "simp-chinese-informal",
				chineseCountingThousand: "simp-chinese-informal",
				chineseLegalSimplified: "simp-chinese-formal",
				chosung: "hangul-consonant",
				ideographDigital: "cjk-ideographic",
				ideographTraditional: "cjk-heavenly-stem",
				ideographLegalTraditional: "trad-chinese-formal",
				ideographZodiac: "cjk-earthly-branch",
				iroha: "katakana-iroha",
				irohaFullWidth: "katakana-iroha",
				japaneseCounting: "japanese-informal",
				japaneseDigitalTenThousand: "cjk-decimal",
				japaneseLegal: "japanese-formal",
				thaiNumbers: "thai",
				koreanCounting: "korean-hangul-formal",
				koreanDigital: "korean-hangul-formal",
				koreanDigital2: "korean-hanja-informal",
				hebrew1: "hebrew",
				hebrew2: "hebrew",
				hindiNumbers: "devanagari",
				ganada: "hangul",
				taiwaneseCounting: "cjk-ideographic",
				taiwaneseCountingThousand: "cjk-ideographic",
				taiwaneseDigital: "cjk-decimal"
			}[format]) !== null && _mapping$format !== void 0 ? _mapping$format : format;
		}
		refreshTabStops() {
			if (!this.options.experimental) return;
			setTimeout(() => {
				const pixelToPoint = computePixelToPoint();
				for (let tab of this.currentTabs) updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}, 500);
		}
		createElementNS(ns, tagName, props, children) {
			return this.h({
				ns,
				tagName,
				children,
				...props
			});
		}
		createElement(tagName, props, children) {
			return this.createElementNS(ns.html, tagName, props, children);
		}
		createMathMLElement(tagName, props, children) {
			return this.createElementNS(ns.mathML, tagName, props, children);
		}
		createSvgElement(tagName, props, children) {
			return this.createElementNS(ns.svg, tagName, props, children);
		}
		later(func) {
			this.postRenderTasks.push(func);
		}
	};
	function findParent(elem, type) {
		var parent = elem.parent;
		while (parent != null && parent.type != type) parent = parent.parent;
		return parent;
	}
	//#endregion
	//#region src/pagination.ts
	var FIT_EPSILON = 1;
	var MAX_PAGES_PER_SECTION = 1e3;
	function paginateWrapper(container, className) {
		const sections = Array.from(container.querySelectorAll(`section.${className}`));
		let pageCount = 0;
		for (const section of sections) pageCount += paginateSection(section, className);
		if (pageCount > 0) {
			const style = document.createElement("style");
			style.textContent = `.${className} .${className}-continuation { text-indent: 0 !important; }.${className} .${className}-continuation::before { content: none !important; }`;
			container.prepend(style);
			const pages = container.querySelectorAll(`section.${className}`);
			const total = `${pages.length}`;
			pages.forEach((page, i) => {
				for (const el of Array.from(page.querySelectorAll(`.${className}-field-page`))) el.textContent = `${i + 1}`;
				for (const el of Array.from(page.querySelectorAll(`.${className}-field-numpages`))) el.textContent = total;
			});
		}
		return pageCount;
	}
	function paginateSection(section, className) {
		const article = section.querySelector(":scope > article");
		if (!article) return 0;
		const sectionStyle = getComputedStyle(section);
		const pageHeight = parseFloat(sectionStyle.minHeight);
		if (!pageHeight || Number.isNaN(pageHeight)) return 0;
		const columns = getComputedStyle(article).columnCount;
		if (columns !== "auto" && columns !== "1") return 0;
		const contentLimit = pageHeight - parseFloat(sectionStyle.paddingTop) - parseFloat(sectionStyle.paddingBottom);
		if (!(contentLimit > 0) || article.scrollHeight <= contentLimit + FIT_EPSILON) return 0;
		const header = section.querySelector(":scope > header");
		const footer = section.querySelector(":scope > footer");
		const notes = Array.from(section.querySelectorAll(":scope > ol"));
		const pages = [section];
		let currentArticle = article;
		let lastPage = section;
		function newPage() {
			const shell = section.cloneNode(false);
			if (header) shell.appendChild(header.cloneNode(true));
			const pageArticle = article.cloneNode(false);
			shell.appendChild(pageArticle);
			if (footer) shell.appendChild(footer.cloneNode(true));
			lastPage.after(shell);
			lastPage = shell;
			pages.push(shell);
			currentArticle = pageArticle;
		}
		const fits = () => currentArticle.scrollHeight <= contentLimit + FIT_EPSILON;
		const bottomLimit = () => currentArticle.getBoundingClientRect().top + contentLimit;
		const blocks = Array.from(article.children);
		for (const b of blocks) b.remove();
		for (let i = 0; i < blocks.length && pages.length < MAX_PAGES_PER_SECTION; i++) {
			const block = blocks[i];
			currentArticle.appendChild(block);
			if (fits()) continue;
			const tail = splitBlock(block, bottomLimit(), className);
			if (tail === "allfit") continue;
			if (tail) {
				blocks.splice(i + 1, 0, tail);
				newPage();
				continue;
			}
			if (tail) {
				blocks.splice(i + 1, 0, tail);
				newPage();
				continue;
			}
			block.remove();
			if (currentArticle.childElementCount === 0) {
				currentArticle.appendChild(block);
				continue;
			}
			newPage();
			i--;
		}
		if (notes.length > 0 && pages.length > 1) {
			const lastFooter = lastPage.querySelector(":scope > footer");
			for (const ol of notes) lastPage.insertBefore(ol, lastFooter);
		}
		if (pages.length <= 1) return 0;
		return pages.length;
	}
	/**
	* Splits `block` (already appended to a page and known to overflow) so its
	* head fits above `bottomLimit` (viewport coordinate). Returns the detached
	* tail fragment, "allfit" when the block's content doesn't actually
	* overflow (keep it on the current page), or null when the block can't
	* usefully be split here.
	*/
	function splitBlock(block, bottomLimit, className) {
		if (block.tagName === "TABLE") return splitTable(block, bottomLimit);
		const pos = findTextSplitOffset(block, bottomLimit);
		if (pos === "allfit") return "allfit";
		if (pos === null) return null;
		const tail = splitAt(block, pos.node, pos.offset);
		tail.classList.add(`${className}-continuation`);
		return tail;
	}
	function splitTable(table, bottomLimit) {
		const bodyRows = [];
		for (const tbody of Array.from(table.tBodies)) bodyRows.push(...Array.from(tbody.rows));
		if (bodyRows.length < 2) return null;
		const splitIndex = bodyRows.findIndex((r) => r.getBoundingClientRect().bottom > bottomLimit);
		if (splitIndex <= 0) return null;
		const tail = table.cloneNode(false);
		const colgroup = table.querySelector(":scope > colgroup");
		if (colgroup) tail.appendChild(colgroup.cloneNode(true));
		if (table.tHead) tail.appendChild(table.tHead.cloneNode(true));
		for (let i = splitIndex; i < bodyRows.length; i++) tail.appendChild(bodyRows[i]);
		return tail;
	}
	/**
	* Finds the latest text position inside `el` that still renders above
	* `bottomLimit`, by walking text nodes in order and binary-searching each
	* one with Range rects. Returns:
	* - { node, offset } — split here;
	* - "allfit" — every text node fits; the overflow comes from the block's
	*   own box (min-height overshoot), not its content, so it should stay;
	* - null — even the first character overflows (or there is no text).
	*/
	function findTextSplitOffset(el, bottomLimit) {
		const doc = el.ownerDocument;
		const walker = doc.createTreeWalker(el, NodeFilter.SHOW_TEXT);
		const range = doc.createRange();
		let hasFittingContent = false;
		let sawText = false;
		for (let node = walker.nextNode(); node; node = walker.nextNode()) {
			if (node.length === 0) continue;
			sawText = true;
			let lo = 0, hi = node.length, fit = 0;
			while (lo <= hi) {
				const mid = lo + hi >> 1;
				range.setStart(node, 0);
				range.setEnd(node, mid);
				const rects = range.getClientRects();
				const last = rects[rects.length - 1];
				if (last && last.bottom <= bottomLimit) {
					fit = mid;
					lo = mid + 1;
				} else hi = mid - 1;
			}
			if (fit === 0) return hasFittingContent ? {
				node,
				offset: 0
			} : null;
			if (fit < node.length) {
				let back = fit;
				const floor = Math.max(0, fit - 30);
				while (back > floor && !/\s/.test(node.data[back - 1])) back--;
				return {
					node,
					offset: back > floor ? back : fit
				};
			}
			hasFittingContent = true;
		}
		return sawText ? "allfit" : null;
	}
	/**
	* Splits `el` at (textNode, offset): `el` keeps everything before the point,
	* and the returned element (a shallow clone of `el` with cloned inline
	* ancestors) gets everything from the point on.
	*/
	function splitAt(el, textNode, offset) {
		const tail = el.cloneNode(false);
		moveTail(textNode.splitText(offset), el, tail);
		return tail;
	}
	/** Moves the split point's following content into a parallel clone
	* hierarchy under `tailRoot`: `node` and its following siblings go into a
	* clone of their parent, and each ancestor up to `root` contributes a clone
	* plus its following siblings. The ORIGINAL ancestors stay in the head —
	* moving them would rip the head's split-span text out (and reverse the
	* order of the two text halves in the tail). */
	function moveTail(node, root, tailRoot) {
		let child = node;
		let container = tailRoot;
		while (child.parentNode && child.parentNode !== root) {
			const parent = child.parentNode;
			const parentClone = parent.cloneNode(false);
			container.appendChild(parentClone);
			parentClone.appendChild(child);
			while (child.nextSibling) parentClone.appendChild(child.nextSibling);
			child = parent;
			container = parentClone;
		}
		if (child === node) tailRoot.appendChild(child);
		while (child.nextSibling) tailRoot.appendChild(child.nextSibling);
	}
	//#endregion
	//#region src/encryption/cfb.ts
	var END_OF_CHAIN = 4294967294;
	var CFB_SIGNATURE = [
		208,
		207,
		17,
		224,
		161,
		177,
		26,
		225
	];
	function isCompoundFile(data) {
		if (data.length < 8) return false;
		for (let i = 0; i < CFB_SIGNATURE.length; i++) if (data[i] !== CFB_SIGNATURE[i]) return false;
		return true;
	}
	function readCompoundFile(data) {
		if (!isCompoundFile(data)) return null;
		const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
		if (data.length < 512) return null;
		const sectorShift = view.getUint16(30, true);
		const miniSectorShift = view.getUint16(32, true);
		if (sectorShift < 7 || sectorShift > 20 || miniSectorShift < 4 || miniSectorShift > 12) return null;
		const sectorSize = 1 << sectorShift;
		const miniSectorSize = 1 << miniSectorShift;
		const numberOfFatSectors = view.getUint32(44, true);
		const directoryStart = view.getUint32(48, true);
		const miniCutoff = view.getUint32(56, true);
		const miniFatStart = view.getUint32(60, true);
		const difatStart = view.getUint32(68, true);
		const numberOfDifatSectors = view.getUint32(72, true);
		const sectorOffset = (sector) => (sector + 1) * sectorSize;
		const hasSector = (sector) => sectorOffset(sector) + sectorSize <= data.length;
		const fatSectors = [];
		for (let i = 0; i < 109 && fatSectors.length < numberOfFatSectors; i++) {
			const sector = view.getUint32(76 + i * 4, true);
			if (sector >= 4294967294 || !hasSector(sector)) break;
			fatSectors.push(sector);
		}
		let difatSector = difatStart;
		const perDifatSector = (sectorSize >> 2) - 1;
		for (let i = 0; i < numberOfDifatSectors && fatSectors.length < numberOfFatSectors; i++) {
			if (difatSector >= 4294967294 || !hasSector(difatSector)) break;
			const base = sectorOffset(difatSector);
			for (let j = 0; j < perDifatSector && fatSectors.length < numberOfFatSectors; j++) {
				const sector = view.getUint32(base + j * 4, true);
				if (sector >= 4294967294 || !hasSector(sector)) break;
				fatSectors.push(sector);
			}
			difatSector = view.getUint32(base + perDifatSector * 4, true);
		}
		const fat = [];
		for (const sector of fatSectors) {
			const base = sectorOffset(sector);
			for (let i = 0; i < sectorSize >> 2; i++) fat.push(view.getUint32(base + i * 4, true));
		}
		/** Sector numbers of a chain, following the given allocation table. */
		function chain(start, table) {
			const out = [];
			let sector = start;
			const seen = /* @__PURE__ */ new Set();
			while (sector < 4294967294 && !seen.has(sector)) {
				var _table$sector;
				out.push(sector);
				seen.add(sector);
				sector = (_table$sector = table[sector]) !== null && _table$sector !== void 0 ? _table$sector : END_OF_CHAIN;
				if (out.length > table.length + 1) break;
			}
			return out;
		}
		function readChain(start, size, table = fat, sectorBytes = sectorSize, source = data, at = sectorOffset) {
			const sectors = chain(start, table);
			const out = new Uint8Array(sectors.length * sectorBytes);
			sectors.forEach((sector, i) => out.set(source.subarray(at(sector), at(sector) + sectorBytes), i * sectorBytes));
			return out.subarray(0, Math.min(size, out.length));
		}
		const directoryBytes = readChain(directoryStart, Number.MAX_SAFE_INTEGER);
		const entries = [];
		for (let offset = 0; offset + 128 <= directoryBytes.length; offset += 128) {
			const entry = new DataView(directoryBytes.buffer, directoryBytes.byteOffset + offset, 128);
			const nameLength = entry.getUint16(64, true);
			const type = entry.getUint8(66);
			if (!nameLength || !type) continue;
			let name = "";
			for (let i = 0; i + 1 < nameLength - 1; i += 2) name += String.fromCharCode(entry.getUint16(i, true));
			entries.push({
				name,
				type,
				startSector: entry.getUint32(116, true),
				size: Number(entry.getBigUint64(120, true))
			});
		}
		const root = entries.find((entry) => entry.type === 5);
		if (!root) return null;
		const miniFat = [];
		for (const sector of chain(miniFatStart, fat)) {
			const base = sectorOffset(sector);
			for (let i = 0; i < sectorSize >> 2; i++) miniFat.push(view.getUint32(base + i * 4, true));
		}
		const miniStream = readChain(root.startSector, root.size);
		const readMiniStream = (start, size) => readChain(start, size, miniFat, miniSectorSize, miniStream, (sector) => sector * miniSectorSize);
		function readStream(name) {
			const entry = entries.find((candidate) => candidate.name === name);
			if (!entry || entry.type !== 2) return null;
			return entry.size < miniCutoff ? readMiniStream(entry.startSector, entry.size) : readChain(entry.startSector, entry.size);
		}
		return {
			entries,
			readStream
		};
	}
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+hashes@2.4.0/node_modules/@noble/hashes/_u64.js
	var U32_MASK64 = /* @__PURE__ */ (() => BigInt(2 ** 32 - 1))();
	var _32n = /* @__PURE__ */ BigInt(32);
	function fromBig(n, le = false) {
		if (le) return {
			h: Number(n & U32_MASK64),
			l: Number(n >> _32n & U32_MASK64)
		};
		return {
			h: Number(n >> _32n & U32_MASK64) | 0,
			l: Number(n & U32_MASK64) | 0
		};
	}
	function split(lst, le = false) {
		const len = lst.length;
		let Ah = new Uint32Array(len);
		let Al = new Uint32Array(len);
		for (let i = 0; i < len; i++) {
			const { h, l } = fromBig(lst[i], le);
			[Ah[i], Al[i]] = [h, l];
		}
		return [Ah, Al];
	}
	var fromNumH = (n) => n / 2 ** 32 | 0;
	var fromNumL = (n) => n >>> 0;
	function setU64FromNum(view, byteOffset, n, isLE) {
		const h = fromNumH(n);
		const l = fromNumL(n);
		view.setUint32(byteOffset, isLE ? l : h, isLE);
		view.setUint32(byteOffset + 4, isLE ? h : l, isLE);
	}
	var shrSH = (h, _l, s) => h >>> s;
	var shrSL = (h, l, s) => h << 32 - s | l >>> s;
	var rotrSH = (h, l, s) => h >>> s | l << 32 - s;
	var rotrSL = (h, l, s) => h << 32 - s | l >>> s;
	var rotrBH = (h, l, s) => h << 64 - s | l >>> s - 32;
	var rotrBL = (h, l, s) => h >>> s - 32 | l << 64 - s;
	function add(Ah, Al, Bh, Bl) {
		const l = (Al >>> 0) + (Bl >>> 0);
		return {
			h: Ah + Bh + (l / 2 ** 32 | 0) | 0,
			l: l | 0
		};
	}
	var add3L = (Al, Bl, Cl) => (Al >>> 0) + (Bl >>> 0) + (Cl >>> 0);
	var add3H = (low, Ah, Bh, Ch) => Ah + Bh + Ch + (low / 2 ** 32 | 0) | 0;
	var add4L = (Al, Bl, Cl, Dl) => (Al >>> 0) + (Bl >>> 0) + (Cl >>> 0) + (Dl >>> 0);
	var add4H = (low, Ah, Bh, Ch, Dh) => Ah + Bh + Ch + Dh + (low / 2 ** 32 | 0) | 0;
	var add5L = (Al, Bl, Cl, Dl, El) => (Al >>> 0) + (Bl >>> 0) + (Cl >>> 0) + (Dl >>> 0) + (El >>> 0);
	var add5H = (low, Ah, Bh, Ch, Dh, Eh) => Ah + Bh + Ch + Dh + Eh + (low / 2 ** 32 | 0) | 0;
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+hashes@2.4.0/node_modules/@noble/hashes/utils.js
	/**
	* Checks if something is Uint8Array. Be careful: nodejs Buffer will return true.
	* @param a - value to test
	* @returns `true` when the value is a Uint8Array-compatible view.
	* @example
	* Check whether a value is a Uint8Array-compatible view.
	* ```ts
	* isBytes(new Uint8Array([1, 2, 3]));
	* ```
	*/
	function isBytes$1(a) {
		return a instanceof Uint8Array || ArrayBuffer.isView(a) && a.constructor.name === "Uint8Array" && "BYTES_PER_ELEMENT" in a && a.BYTES_PER_ELEMENT === 1;
	}
	var atitle$1 = (title) => title ? `"${title}" ` : "";
	/**
	* Asserts something is a non-negative integer.
	* @param n - number to validate
	* @param title - label included in thrown errors
	* @returns The validated number.
	* @throws On wrong argument types. {@link TypeError}
	* @throws On wrong argument ranges or values. {@link RangeError}
	* @example
	* Validate a non-negative integer option.
	* ```ts
	* anumber(32, 'length');
	* ```
	*/
	function anumber$1(n, title = "") {
		if (typeof n !== "number") throw new TypeError(atitle$1(title) + "expected number, got " + typeof n);
		if (!Number.isSafeInteger(n) || n < 0) throw new RangeError(atitle$1(title) + "expected integer >= 0, got " + n);
		return n;
	}
	/**
	* Asserts something is Uint8Array.
	* @param value - value to validate
	* @param length - optional exact length constraint
	* @param title - label included in thrown errors
	* @returns The validated byte array.
	* @throws On wrong argument types. {@link TypeError}
	* @throws On wrong argument ranges or values. {@link RangeError}
	* @example
	* Validate that a value is a byte array.
	* ```ts
	* abytes(new Uint8Array([1, 2, 3]));
	* ```
	*/
	function abytes$1(value, length, title = "") {
		if (isBytes$1(value) && (length === void 0 || value.length === length)) return value;
		if (length !== void 0) anumber$1(length, "length");
		const bytes = isBytes$1(value);
		const ofLen = length !== void 0 ? ` of length ${length}` : "";
		const got = bytes ? `length=${value.length}` : `type=${typeof value}`;
		const message = atitle$1(title) + "expected Uint8Array" + ofLen + ", got " + got;
		if (!bytes) throw new TypeError(message);
		throw new RangeError(message);
	}
	var aobject = (value, label) => {
		if (value === null || typeof value !== "object" || Array.isArray(value)) throw new TypeError((label === "object" ? "" : `"${label}" `) + "expected object, got type=" + typeof value);
	};
	var aopts = (value, label) => {
		aobject(value, label);
		const proto = Object.getPrototypeOf(value);
		if (proto !== Object.prototype && proto !== null) throw new TypeError(`"${label}" expected plain object`);
		if (Object.hasOwn(value, "__proto__")) throw new TypeError(`"${label}.__proto__" is not allowed`);
	};
	/**
	* Asserts a hash instance has not been destroyed or finished.
	* @param instance - hash instance to validate
	* @param checkFinished - whether to reject finalized instances
	* @throws If the hash instance has already been destroyed or finalized. {@link Error}
	* @example
	* Validate that a hash instance is still usable.
	* ```ts
	* import { aexists } from '@noble/hashes/utils.js';
	* import { sha256 } from '@noble/hashes/sha2.js';
	* const hash = sha256.create();
	* aexists(hash);
	* ```
	*/
	function aexists(instance, checkFinished = true) {
		if (instance.destroyed) throw new Error("hash was destroyed");
		if (checkFinished && instance.finished) throw new Error("digest() was already called");
	}
	/**
	* Asserts output is a sufficiently-sized byte array.
	* @param out - destination buffer
	* @param instance - hash instance providing output length
	* Oversized buffers are allowed; downstream code only promises to fill the first `outputLen` bytes.
	* @throws On wrong argument types. {@link TypeError}
	* @throws On wrong argument ranges or values. {@link RangeError}
	* @example
	* Validate a caller-provided digest buffer.
	* ```ts
	* import { aoutput } from '@noble/hashes/utils.js';
	* import { sha256 } from '@noble/hashes/sha2.js';
	* const hash = sha256.create();
	* aoutput(new Uint8Array(hash.outputLen), hash);
	* ```
	*/
	function aoutput(out, instance) {
		abytes$1(out, void 0, "output");
		const min = instance.outputLen;
		if (!(out.length >= min)) throw new RangeError("\"output\" expected length >= " + min);
	}
	/**
	* Zeroizes typed arrays in place. Warning: JS provides no guarantees.
	* @param arrays - arrays to overwrite with zeros
	* @example
	* Zeroize sensitive buffers in place.
	* ```ts
	* clean(new Uint8Array([1, 2, 3]));
	* ```
	*/
	function clean$1(...arrays) {
		for (let i = 0; i < arrays.length; i++) arrays[i].fill(0);
	}
	/**
	* Creates a DataView for byte-level manipulation.
	* @param arr - source typed array
	* @returns DataView over the same buffer region.
	* @example
	* Create a DataView over an existing buffer.
	* ```ts
	* createView(new Uint8Array(4));
	* ```
	*/
	function createView(arr) {
		return new DataView(arr.buffer, arr.byteOffset, arr.byteLength);
	}
	/**
	* Rotate-right operation for uint32 values.
	* @param word - source word
	* @param shift - shift amount in bits
	* @returns Rotated word.
	* @example
	* Rotate a 32-bit word to the right.
	* ```ts
	* rotr(0x12345678, 8);
	* ```
	*/
	function rotr(word, shift) {
		return word << 32 - shift | word >>> shift;
	}
	/**
	* Rotate-left operation for uint32 values.
	* @param word - source word
	* @param shift - shift amount in bits
	* @returns Rotated word.
	* @example
	* Rotate a 32-bit word to the left.
	* ```ts
	* rotl(0x12345678, 8);
	* ```
	*/
	function rotl(word, shift) {
		return word << shift | word >>> 32 - shift >>> 0;
	}
	/**
	* Merges default options and passed options.
	* @param defaults - base option object
	* @param opts - user overrides
	* @param title - label included in thrown override errors
	* @returns Fresh merged option object with a null prototype.
	* @throws On wrong argument types. {@link TypeError}
	* @example
	* Merge user overrides onto default options.
	* ```ts
	* checkOpts({ dkLen: 32 }, { asyncTick: 10 });
	* ```
	*/
	function checkOpts(defaults, opts, title = "opts") {
		aopts(defaults, "defaults");
		if (opts !== void 0) aopts(opts, title);
		return Object.assign(Object.create(null), defaults, opts);
	}
	/**
	* Creates a callable hash function from a stateful class constructor.
	* @param hashCons - hash constructor or factory
	* @param info - optional metadata such as DER OID
	* @returns Frozen callable hash wrapper with `.create()`.
	*   Wrapper construction eagerly calls `hashCons(undefined)` once to read
	*   `outputLen` / `blockLen`, so constructor side effects happen at module
	*   init time.
	* @throws On wrong argument types. {@link TypeError}
	* @example
	* Wrap a stateful hash constructor into a callable helper.
	* ```ts
	* import { createHasher } from '@noble/hashes/utils.js';
	* import { sha256 } from '@noble/hashes/sha2.js';
	* const wrapped = createHasher(sha256.create, { oid: sha256.oid });
	* wrapped(new Uint8Array([1]));
	* ```
	*/
	function createHasher(hashCons, info = {}) {
		if (typeof hashCons !== "function") throw new TypeError("\"hashCons\" expected function, got type=" + typeof hashCons);
		info = checkOpts({}, info, "info");
		const hashC = (msg, opts) => hashCons(opts).update(msg).digest();
		const tmp = hashCons(void 0);
		hashC.outputLen = tmp.outputLen;
		hashC.blockLen = tmp.blockLen;
		hashC.canXOF = tmp.canXOF;
		hashC.create = (opts) => hashCons(opts);
		Object.assign(hashC, info);
		return Object.freeze(hashC);
	}
	/**
	* Creates OID metadata for NIST hashes with prefix `06 09 60 86 48 01 65 03 04 02`.
	* @param suffix - final OID byte for the selected hash.
	*   The helper accepts any byte even though only the documented NIST hash
	*   suffixes are meaningful downstream.
	* @returns Object containing the DER-encoded OID.
	* @example
	* Build OID metadata for a NIST hash.
	* ```ts
	* oidNist(0x01);
	* ```
	*/
	var oidNist = (suffix) => ({ oid: Uint8Array.from([
		6,
		9,
		96,
		134,
		72,
		1,
		101,
		3,
		4,
		2,
		suffix
	]) });
	//#endregion
	//#region \0@oxc-project+runtime@0.147.0/helpers/esm/typeof.js
	function _typeof(o) {
		"@babel/helpers - typeof";
		return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o) {
			return typeof o;
		} : function(o) {
			return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o;
		}, _typeof(o);
	}
	//#endregion
	//#region \0@oxc-project+runtime@0.147.0/helpers/esm/toPrimitive.js
	function toPrimitive(t, r) {
		if ("object" != _typeof(t) || !t) return t;
		var e = t[Symbol.toPrimitive];
		if (void 0 !== e) {
			var i = e.call(t, r || "default");
			if ("object" != _typeof(i)) return i;
			throw new TypeError("@@toPrimitive must return a primitive value.");
		}
		return ("string" === r ? String : Number)(t);
	}
	//#endregion
	//#region \0@oxc-project+runtime@0.147.0/helpers/esm/toPropertyKey.js
	function toPropertyKey(t) {
		var i = toPrimitive(t, "string");
		return "symbol" == _typeof(i) ? i : i + "";
	}
	//#endregion
	//#region \0@oxc-project+runtime@0.147.0/helpers/esm/defineProperty.js
	function _defineProperty(e, r, t) {
		return (r = toPropertyKey(r)) in e ? Object.defineProperty(e, r, {
			value: t,
			enumerable: !0,
			configurable: !0,
			writable: !0
		}) : e[r] = t, e;
	}
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+hashes@2.4.0/node_modules/@noble/hashes/_md.js
	/**
	* Internal Merkle-Damgard hash utils.
	* @module
	*/
	/**
	* Shared 32-bit conditional boolean primitive reused by SHA-256, SHA-1, and MD5 `F`.
	* Returns bits from `b` when `a` is set, otherwise from `c`.
	* The XOR form is equivalent to MD5's `F(X,Y,Z) = XY v not(X)Z` because the masked terms never
	* set the same bit.
	* @param a - selector word
	* @param b - word chosen when selector bit is set
	* @param c - word chosen when selector bit is clear
	* @returns Mixed 32-bit word.
	* @example
	* Combine three words with the shared 32-bit choice primitive.
	* ```ts
	* Chi(0xffffffff, 0x12345678, 0x87654321);
	* ```
	*/
	function Chi(a, b, c) {
		return a & b ^ ~a & c;
	}
	/**
	* Shared 32-bit majority primitive reused by SHA-256 and SHA-1.
	* Returns bits shared by at least two inputs.
	* @param a - first input word
	* @param b - second input word
	* @param c - third input word
	* @returns Mixed 32-bit word.
	* @example
	* Combine three words with the shared 32-bit majority primitive.
	* ```ts
	* Maj(0xffffffff, 0x12345678, 0x87654321);
	* ```
	*/
	function Maj(a, b, c) {
		return a & b ^ a & c ^ b & c;
	}
	/**
	* Merkle-Damgard hash construction base class.
	* Could be used to create MD5, RIPEMD, SHA1, SHA2.
	* Accepts only byte-aligned `Uint8Array` input, even when the underlying spec describes bit
	* strings with partial-byte tails.
	* @param blockLen - internal block size in bytes
	* @param outputLen - digest size in bytes
	* @param padOffset - trailing length field size in bytes
	* @param isLE - whether length and state words are encoded in little-endian
	* @example
	* Use a concrete subclass to get the shared Merkle-Damgard update/digest flow.
	* ```ts
	* import { _SHA1 } from '@noble/hashes/legacy.js';
	* const hash = new _SHA1();
	* hash.update(new Uint8Array([97, 98, 99]));
	* hash.digest();
	* ```
	*/
	var HashMD = class {
		constructor(blockLen, outputLen, padOffset, isLE) {
			_defineProperty(this, "blockLen", void 0);
			_defineProperty(this, "outputLen", void 0);
			_defineProperty(this, "canXOF", false);
			_defineProperty(this, "padOffset", void 0);
			_defineProperty(this, "isLE", void 0);
			_defineProperty(this, "buffer", void 0);
			_defineProperty(this, "view", void 0);
			_defineProperty(this, "finished", false);
			_defineProperty(this, "length", 0);
			_defineProperty(this, "pos", 0);
			_defineProperty(this, "destroyed", false);
			this.blockLen = blockLen;
			this.outputLen = outputLen;
			this.padOffset = padOffset;
			this.isLE = isLE;
			this.buffer = new Uint8Array(blockLen);
			this.view = createView(this.buffer);
		}
		update(data) {
			aexists(this);
			abytes$1(data);
			const { view, buffer, blockLen } = this;
			const len = data.length;
			let processed = false;
			for (let pos = 0; pos < len;) {
				const take = Math.min(blockLen - this.pos, len - pos);
				if (take === blockLen) {
					const dataView = createView(data);
					for (; blockLen <= len - pos; pos += blockLen) this.process(dataView, pos);
					processed = true;
					continue;
				}
				buffer.set(pos === 0 && take === len ? data : data.subarray(pos, pos + take), this.pos);
				this.pos += take;
				pos += take;
				if (this.pos === blockLen) {
					this.process(view, 0);
					this.pos = 0;
					processed = true;
				}
			}
			this.length += data.length;
			if (processed) this.roundClean();
			return this;
		}
		digestInto(out) {
			aexists(this);
			aoutput(out, this);
			this.finished = true;
			const { buffer, view, blockLen, isLE } = this;
			let { pos } = this;
			buffer[pos++] = 128;
			buffer.fill(0, pos);
			if (this.padOffset > blockLen - pos) {
				this.process(view, 0);
				buffer.fill(0);
			}
			setU64FromNum(view, blockLen - 8, this.length * 8, isLE);
			this.process(view, 0);
			this.roundClean();
			const oview = out === buffer ? view : createView(out);
			const len = this.outputLen;
			const outLen = len / 4;
			const state = this.get();
			if (len % 4 || outLen > state.length) throw new Error("invalid outputLen");
			for (let i = 0; i < outLen; i++) oview.setUint32(4 * i, state[i], isLE);
		}
		digest() {
			const { buffer, outputLen } = this;
			this.digestInto(buffer);
			const res = buffer.slice(0, outputLen);
			this.destroy();
			return res;
		}
		_cloneIntoMeta(to) {
			const { buffer, length, finished, destroyed, pos } = this;
			to.destroyed = destroyed;
			to.finished = finished;
			to.length = length;
			to.pos = pos;
			if (pos) to.buffer.set(buffer);
			return to;
		}
		clone() {
			return this._cloneInto();
		}
	};
	/**
	* Initial SHA-2 state: fractional parts of square roots of first 16 primes 2..53.
	* Check out `test/misc/sha2-gen-iv.js` for recomputation guide.
	*/
	/** Initial SHA256 state from RFC 6234 §6.1: the first 32 bits of the fractional parts of the
	* square roots of the first eight prime numbers. Exported as a shared table; callers must treat
	* it as read-only because constructors copy words from it by index. */
	var SHA256_IV = /* @__PURE__ */ Uint32Array.from([
		1779033703,
		3144134277,
		1013904242,
		2773480762,
		1359893119,
		2600822924,
		528734635,
		1541459225
	]);
	/** Initial SHA384 state from RFC 6234 §6.3: eight RFC 64-bit `H(0)` words stored as sixteen
	* big-endian 32-bit halves. Derived from the fractional parts of the square roots of the ninth
	* through sixteenth prime numbers. Exported as a shared table; callers must treat it as read-only
	* because constructors copy halves from it by index. */
	var SHA384_IV = /* @__PURE__ */ Uint32Array.from([
		3418070365,
		3238371032,
		1654270250,
		914150663,
		2438529370,
		812702999,
		355462360,
		4144912697,
		1731405415,
		4290775857,
		2394180231,
		1750603025,
		3675008525,
		1694076839,
		1203062813,
		3204075428
	]);
	/** Initial SHA512 state from RFC 6234 §6.3: eight RFC 64-bit `H(0)` words stored as sixteen
	* big-endian 32-bit halves. Derived from the fractional parts of the square roots of the first
	* eight prime numbers. Exported as a shared table; callers must treat it as read-only because
	* constructors copy halves from it by index. */
	var SHA512_IV = /* @__PURE__ */ Uint32Array.from([
		1779033703,
		4089235720,
		3144134277,
		2227873595,
		1013904242,
		4271175723,
		2773480762,
		1595750129,
		1359893119,
		2917565137,
		2600822924,
		725511199,
		528734635,
		4215389547,
		1541459225,
		327033209
	]);
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+hashes@2.4.0/node_modules/@noble/hashes/legacy.js
	/**
	
	SHA1 (RFC 3174), MD5 (RFC 1321), and RIPEMD160 legacy, weak hash functions.
	RFC 2286 only covers HMAC-RIPEMD160 wrapper material and test vectors,
	not the base RIPEMD-160 compression spec.
	Don't use them in a new protocol. What "weak" means:
	
	- Collisions can be made with 2^18 effort in MD5, 2^60 in SHA1, 2^80 in RIPEMD160.
	- No practical pre-image attacks (only theoretical, 2^123.4)
	- HMAC seems kinda ok: https://www.rfc-editor.org/rfc/rfc6151
	* @module
	*/
	/** Initial SHA-1 state from RFC 3174 §6.1. */
	var SHA1_IV = /* @__PURE__ */ Uint32Array.from([
		1732584193,
		4023233417,
		2562383102,
		271733878,
		3285377520
	]);
	var SHA1_W = /* @__PURE__ */ new Uint32Array(80);
	/** Internal SHA1 legacy hash class. */
	var _SHA1 = class extends HashMD {
		constructor() {
			super(64, 20, 8, false);
			_defineProperty(this, "A", SHA1_IV[0] | 0);
			_defineProperty(this, "B", SHA1_IV[1] | 0);
			_defineProperty(this, "C", SHA1_IV[2] | 0);
			_defineProperty(this, "D", SHA1_IV[3] | 0);
			_defineProperty(this, "E", SHA1_IV[4] | 0);
		}
		get() {
			const { A, B, C, D, E } = this;
			return [
				A,
				B,
				C,
				D,
				E
			];
		}
		set(A, B, C, D, E) {
			this.A = A | 0;
			this.B = B | 0;
			this.C = C | 0;
			this.D = D | 0;
			this.E = E | 0;
		}
		_cloneInto(to) {
			(to || (to = new this.constructor())).set(...this.get());
			return this._cloneIntoMeta(to);
		}
		process(view, offset) {
			for (let i = 0; i < 16; i++, offset += 4) SHA1_W[i] = view.getUint32(offset, false);
			for (let i = 16; i < 80; i++) SHA1_W[i] = rotl(SHA1_W[i - 3] ^ SHA1_W[i - 8] ^ SHA1_W[i - 14] ^ SHA1_W[i - 16], 1);
			let { A, B, C, D, E } = this;
			for (let i = 0; i < 80; i++) {
				let F, K;
				if (i < 20) {
					F = Chi(B, C, D);
					K = 1518500249;
				} else if (i < 40) {
					F = B ^ C ^ D;
					K = 1859775393;
				} else if (i < 60) {
					F = Maj(B, C, D);
					K = 2400959708;
				} else {
					F = B ^ C ^ D;
					K = 3395469782;
				}
				const T = rotl(A, 5) + F + E + K + SHA1_W[i] | 0;
				E = D;
				D = C;
				C = rotl(B, 30);
				B = A;
				A = T;
			}
			A = A + this.A | 0;
			B = B + this.B | 0;
			C = C + this.C | 0;
			D = D + this.D | 0;
			E = E + this.E | 0;
			this.set(A, B, C, D, E);
		}
		roundClean() {
			clean$1(SHA1_W);
		}
		destroy() {
			this.destroyed = true;
			this.set(0, 0, 0, 0, 0);
			clean$1(this.buffer);
		}
	};
	/**
	* SHA1 (RFC 3174) legacy hash function. It was cryptographically broken.
	* @param msg - message bytes to hash
	* @param opts - Reserved hash options.
	* @returns Digest bytes.
	* @example
	* Hash a message with SHA1.
	* ```ts
	* sha1(new Uint8Array([97, 98, 99]));
	* ```
	*/
	var sha1 = /* @__PURE__ */ createHasher(() => new _SHA1());
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+hashes@2.4.0/node_modules/@noble/hashes/sha2.js
	/**
	* SHA2 hash function. A.k.a. sha256, sha384, sha512, sha512_224, sha512_256.
	* SHA256 is the fastest hash implementable in JS, even faster than Blake3.
	* Check out {@link https://www.rfc-editor.org/rfc/rfc4634 | RFC 4634} and
	* {@link https://nvlpubs.nist.gov/nistpubs/FIPS/NIST.FIPS.180-4.pdf | FIPS 180-4}.
	* @module
	*/
	/**
	* SHA-224 / SHA-256 round constants from RFC 6234 §5.1: the first 32 bits
	* of the cube roots of the first 64 primes (2..311).
	*/
	var SHA256_K = /* @__PURE__ */ Uint32Array.from([
		1116352408,
		1899447441,
		3049323471,
		3921009573,
		961987163,
		1508970993,
		2453635748,
		2870763221,
		3624381080,
		310598401,
		607225278,
		1426881987,
		1925078388,
		2162078206,
		2614888103,
		3248222580,
		3835390401,
		4022224774,
		264347078,
		604807628,
		770255983,
		1249150122,
		1555081692,
		1996064986,
		2554220882,
		2821834349,
		2952996808,
		3210313671,
		3336571891,
		3584528711,
		113926993,
		338241895,
		666307205,
		773529912,
		1294757372,
		1396182291,
		1695183700,
		1986661051,
		2177026350,
		2456956037,
		2730485921,
		2820302411,
		3259730800,
		3345764771,
		3516065817,
		3600352804,
		4094571909,
		275423344,
		430227734,
		506948616,
		659060556,
		883997877,
		958139571,
		1322822218,
		1537002063,
		1747873779,
		1955562222,
		2024104815,
		2227730452,
		2361852424,
		2428436474,
		2756734187,
		3204031479,
		3329325298
	]);
	/** Reusable SHA-224 / SHA-256 message schedule buffer `W_t` from RFC 6234 §6.2 step 1. */
	var SHA256_W = /* @__PURE__ */ new Uint32Array(64);
	/** Internal SHA-224 / SHA-256 compression engine from RFC 6234 §6.2. */
	var SHA2_32B = class extends HashMD {
		constructor(outputLen, IV) {
			super(64, outputLen, 8, false);
			_defineProperty(this, "A", 0);
			_defineProperty(this, "B", 0);
			_defineProperty(this, "C", 0);
			_defineProperty(this, "D", 0);
			_defineProperty(this, "E", 0);
			_defineProperty(this, "F", 0);
			_defineProperty(this, "G", 0);
			_defineProperty(this, "H", 0);
			this.A = IV[0] | 0;
			this.B = IV[1] | 0;
			this.C = IV[2] | 0;
			this.D = IV[3] | 0;
			this.E = IV[4] | 0;
			this.F = IV[5] | 0;
			this.G = IV[6] | 0;
			this.H = IV[7] | 0;
		}
		get() {
			const { A, B, C, D, E, F, G, H } = this;
			return [
				A,
				B,
				C,
				D,
				E,
				F,
				G,
				H
			];
		}
		set(A, B, C, D, E, F, G, H) {
			this.A = A | 0;
			this.B = B | 0;
			this.C = C | 0;
			this.D = D | 0;
			this.E = E | 0;
			this.F = F | 0;
			this.G = G | 0;
			this.H = H | 0;
		}
		_cloneInto(to) {
			(to || (to = new this.constructor())).set(...this.get());
			return this._cloneIntoMeta(to);
		}
		process(view, offset) {
			for (let i = 0; i < 16; i++, offset += 4) SHA256_W[i] = view.getUint32(offset, false);
			for (let i = 16; i < 64; i++) {
				const W15 = SHA256_W[i - 15];
				const W2 = SHA256_W[i - 2];
				const s0 = rotr(W15, 7) ^ rotr(W15, 18) ^ W15 >>> 3;
				const s1 = rotr(W2, 17) ^ rotr(W2, 19) ^ W2 >>> 10;
				SHA256_W[i] = s1 + SHA256_W[i - 7] + s0 + SHA256_W[i - 16] | 0;
			}
			let { A, B, C, D, E, F, G, H } = this;
			for (let i = 0; i < 64; i++) {
				const sigma1 = rotr(E, 6) ^ rotr(E, 11) ^ rotr(E, 25);
				const T1 = H + sigma1 + Chi(E, F, G) + SHA256_K[i] + SHA256_W[i] | 0;
				const T2 = (rotr(A, 2) ^ rotr(A, 13) ^ rotr(A, 22)) + Maj(A, B, C) | 0;
				H = G;
				G = F;
				F = E;
				E = D + T1 | 0;
				D = C;
				C = B;
				B = A;
				A = T1 + T2 | 0;
			}
			A = A + this.A | 0;
			B = B + this.B | 0;
			C = C + this.C | 0;
			D = D + this.D | 0;
			E = E + this.E | 0;
			F = F + this.F | 0;
			G = G + this.G | 0;
			H = H + this.H | 0;
			this.set(A, B, C, D, E, F, G, H);
		}
		roundClean() {
			clean$1(SHA256_W);
		}
		destroy() {
			this.destroyed = true;
			this.set(0, 0, 0, 0, 0, 0, 0, 0);
			clean$1(this.buffer);
		}
	};
	/** Internal SHA-256 hash class grounded in RFC 6234 §6.2. */
	var _SHA256 = class extends SHA2_32B {
		constructor() {
			super(32, SHA256_IV);
		}
	};
	var K512 = /* @__PURE__ */ (() => split([
		"0x428a2f98d728ae22",
		"0x7137449123ef65cd",
		"0xb5c0fbcfec4d3b2f",
		"0xe9b5dba58189dbbc",
		"0x3956c25bf348b538",
		"0x59f111f1b605d019",
		"0x923f82a4af194f9b",
		"0xab1c5ed5da6d8118",
		"0xd807aa98a3030242",
		"0x12835b0145706fbe",
		"0x243185be4ee4b28c",
		"0x550c7dc3d5ffb4e2",
		"0x72be5d74f27b896f",
		"0x80deb1fe3b1696b1",
		"0x9bdc06a725c71235",
		"0xc19bf174cf692694",
		"0xe49b69c19ef14ad2",
		"0xefbe4786384f25e3",
		"0x0fc19dc68b8cd5b5",
		"0x240ca1cc77ac9c65",
		"0x2de92c6f592b0275",
		"0x4a7484aa6ea6e483",
		"0x5cb0a9dcbd41fbd4",
		"0x76f988da831153b5",
		"0x983e5152ee66dfab",
		"0xa831c66d2db43210",
		"0xb00327c898fb213f",
		"0xbf597fc7beef0ee4",
		"0xc6e00bf33da88fc2",
		"0xd5a79147930aa725",
		"0x06ca6351e003826f",
		"0x142929670a0e6e70",
		"0x27b70a8546d22ffc",
		"0x2e1b21385c26c926",
		"0x4d2c6dfc5ac42aed",
		"0x53380d139d95b3df",
		"0x650a73548baf63de",
		"0x766a0abb3c77b2a8",
		"0x81c2c92e47edaee6",
		"0x92722c851482353b",
		"0xa2bfe8a14cf10364",
		"0xa81a664bbc423001",
		"0xc24b8b70d0f89791",
		"0xc76c51a30654be30",
		"0xd192e819d6ef5218",
		"0xd69906245565a910",
		"0xf40e35855771202a",
		"0x106aa07032bbd1b8",
		"0x19a4c116b8d2d0c8",
		"0x1e376c085141ab53",
		"0x2748774cdf8eeb99",
		"0x34b0bcb5e19b48a8",
		"0x391c0cb3c5c95a63",
		"0x4ed8aa4ae3418acb",
		"0x5b9cca4f7763e373",
		"0x682e6ff3d6b2b8a3",
		"0x748f82ee5defb2fc",
		"0x78a5636f43172f60",
		"0x84c87814a1f0ab72",
		"0x8cc702081a6439ec",
		"0x90befffa23631e28",
		"0xa4506cebde82bde9",
		"0xbef9a3f7b2c67915",
		"0xc67178f2e372532b",
		"0xca273eceea26619c",
		"0xd186b8c721c0c207",
		"0xeada7dd6cde0eb1e",
		"0xf57d4f7fee6ed178",
		"0x06f067aa72176fba",
		"0x0a637dc5a2c898a6",
		"0x113f9804bef90dae",
		"0x1b710b35131c471b",
		"0x28db77f523047d84",
		"0x32caab7b40c72493",
		"0x3c9ebe0a15c9bebc",
		"0x431d67c49c100d4c",
		"0x4cc5d4becb3e42b6",
		"0x597f299cfc657e2a",
		"0x5fcb6fab3ad6faec",
		"0x6c44198c4a475817"
	].map((n) => BigInt(n))))();
	var SHA512_Kh = /* @__PURE__ */ (() => K512[0])();
	var SHA512_Kl = /* @__PURE__ */ (() => K512[1])();
	var SHA512_W_H = /* @__PURE__ */ new Uint32Array(80);
	var SHA512_W_L = /* @__PURE__ */ new Uint32Array(80);
	/** Internal SHA-384 / SHA-512 compression engine from RFC 6234 §6.4. */
	var SHA2_64B = class extends HashMD {
		constructor(outputLen, IV) {
			super(128, outputLen, 16, false);
			_defineProperty(this, "Ah", 0);
			_defineProperty(this, "Al", 0);
			_defineProperty(this, "Bh", 0);
			_defineProperty(this, "Bl", 0);
			_defineProperty(this, "Ch", 0);
			_defineProperty(this, "Cl", 0);
			_defineProperty(this, "Dh", 0);
			_defineProperty(this, "Dl", 0);
			_defineProperty(this, "Eh", 0);
			_defineProperty(this, "El", 0);
			_defineProperty(this, "Fh", 0);
			_defineProperty(this, "Fl", 0);
			_defineProperty(this, "Gh", 0);
			_defineProperty(this, "Gl", 0);
			_defineProperty(this, "Hh", 0);
			_defineProperty(this, "Hl", 0);
			this.Ah = IV[0] | 0;
			this.Al = IV[1] | 0;
			this.Bh = IV[2] | 0;
			this.Bl = IV[3] | 0;
			this.Ch = IV[4] | 0;
			this.Cl = IV[5] | 0;
			this.Dh = IV[6] | 0;
			this.Dl = IV[7] | 0;
			this.Eh = IV[8] | 0;
			this.El = IV[9] | 0;
			this.Fh = IV[10] | 0;
			this.Fl = IV[11] | 0;
			this.Gh = IV[12] | 0;
			this.Gl = IV[13] | 0;
			this.Hh = IV[14] | 0;
			this.Hl = IV[15] | 0;
		}
		get() {
			const { Ah, Al, Bh, Bl, Ch, Cl, Dh, Dl, Eh, El, Fh, Fl, Gh, Gl, Hh, Hl } = this;
			return [
				Ah,
				Al,
				Bh,
				Bl,
				Ch,
				Cl,
				Dh,
				Dl,
				Eh,
				El,
				Fh,
				Fl,
				Gh,
				Gl,
				Hh,
				Hl
			];
		}
		set(Ah, Al, Bh, Bl, Ch, Cl, Dh, Dl, Eh, El, Fh, Fl, Gh, Gl, Hh, Hl) {
			this.Ah = Ah | 0;
			this.Al = Al | 0;
			this.Bh = Bh | 0;
			this.Bl = Bl | 0;
			this.Ch = Ch | 0;
			this.Cl = Cl | 0;
			this.Dh = Dh | 0;
			this.Dl = Dl | 0;
			this.Eh = Eh | 0;
			this.El = El | 0;
			this.Fh = Fh | 0;
			this.Fl = Fl | 0;
			this.Gh = Gh | 0;
			this.Gl = Gl | 0;
			this.Hh = Hh | 0;
			this.Hl = Hl | 0;
		}
		_cloneInto(to) {
			(to || (to = new this.constructor())).set(...this.get());
			return this._cloneIntoMeta(to);
		}
		process(view, offset) {
			for (let i = 0; i < 16; i++, offset += 4) {
				SHA512_W_H[i] = view.getUint32(offset);
				SHA512_W_L[i] = view.getUint32(offset += 4);
			}
			for (let i = 16; i < 80; i++) {
				const W15h = SHA512_W_H[i - 15] | 0;
				const W15l = SHA512_W_L[i - 15] | 0;
				const s0h = rotrSH(W15h, W15l, 1) ^ rotrSH(W15h, W15l, 8) ^ shrSH(W15h, W15l, 7);
				const s0l = rotrSL(W15h, W15l, 1) ^ rotrSL(W15h, W15l, 8) ^ shrSL(W15h, W15l, 7);
				const W2h = SHA512_W_H[i - 2] | 0;
				const W2l = SHA512_W_L[i - 2] | 0;
				const s1h = rotrSH(W2h, W2l, 19) ^ rotrBH(W2h, W2l, 61) ^ shrSH(W2h, W2l, 6);
				const SUMl = add4L(s0l, rotrSL(W2h, W2l, 19) ^ rotrBL(W2h, W2l, 61) ^ shrSL(W2h, W2l, 6), SHA512_W_L[i - 7], SHA512_W_L[i - 16]);
				const SUMh = add4H(SUMl, s0h, s1h, SHA512_W_H[i - 7], SHA512_W_H[i - 16]);
				SHA512_W_H[i] = SUMh | 0;
				SHA512_W_L[i] = SUMl | 0;
			}
			let { Ah, Al, Bh, Bl, Ch, Cl, Dh, Dl, Eh, El, Fh, Fl, Gh, Gl, Hh, Hl } = this;
			for (let i = 0; i < 80; i++) {
				const sigma1h = rotrSH(Eh, El, 14) ^ rotrSH(Eh, El, 18) ^ rotrBH(Eh, El, 41);
				const sigma1l = rotrSL(Eh, El, 14) ^ rotrSL(Eh, El, 18) ^ rotrBL(Eh, El, 41);
				const CHIh = Eh & Fh ^ ~Eh & Gh;
				const CHIl = El & Fl ^ ~El & Gl;
				const T1ll = add5L(Hl, sigma1l, CHIl, SHA512_Kl[i], SHA512_W_L[i]);
				const T1h = add5H(T1ll, Hh, sigma1h, CHIh, SHA512_Kh[i], SHA512_W_H[i]);
				const T1l = T1ll | 0;
				const sigma0h = rotrSH(Ah, Al, 28) ^ rotrBH(Ah, Al, 34) ^ rotrBH(Ah, Al, 39);
				const sigma0l = rotrSL(Ah, Al, 28) ^ rotrBL(Ah, Al, 34) ^ rotrBL(Ah, Al, 39);
				const MAJh = Ah & Bh ^ Ah & Ch ^ Bh & Ch;
				const MAJl = Al & Bl ^ Al & Cl ^ Bl & Cl;
				Hh = Gh | 0;
				Hl = Gl | 0;
				Gh = Fh | 0;
				Gl = Fl | 0;
				Fh = Eh | 0;
				Fl = El | 0;
				({h: Eh, l: El} = add(Dh | 0, Dl | 0, T1h | 0, T1l | 0));
				Dh = Ch | 0;
				Dl = Cl | 0;
				Ch = Bh | 0;
				Cl = Bl | 0;
				Bh = Ah | 0;
				Bl = Al | 0;
				const All = add3L(T1l, sigma0l, MAJl);
				Ah = add3H(All, T1h, sigma0h, MAJh);
				Al = All | 0;
			}
			({h: Ah, l: Al} = add(this.Ah | 0, this.Al | 0, Ah | 0, Al | 0));
			({h: Bh, l: Bl} = add(this.Bh | 0, this.Bl | 0, Bh | 0, Bl | 0));
			({h: Ch, l: Cl} = add(this.Ch | 0, this.Cl | 0, Ch | 0, Cl | 0));
			({h: Dh, l: Dl} = add(this.Dh | 0, this.Dl | 0, Dh | 0, Dl | 0));
			({h: Eh, l: El} = add(this.Eh | 0, this.El | 0, Eh | 0, El | 0));
			({h: Fh, l: Fl} = add(this.Fh | 0, this.Fl | 0, Fh | 0, Fl | 0));
			({h: Gh, l: Gl} = add(this.Gh | 0, this.Gl | 0, Gh | 0, Gl | 0));
			({h: Hh, l: Hl} = add(this.Hh | 0, this.Hl | 0, Hh | 0, Hl | 0));
			this.set(Ah, Al, Bh, Bl, Ch, Cl, Dh, Dl, Eh, El, Fh, Fl, Gh, Gl, Hh, Hl);
		}
		roundClean() {
			clean$1(SHA512_W_H, SHA512_W_L);
		}
		destroy() {
			this.destroyed = true;
			clean$1(this.buffer);
			this.set(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
		}
	};
	/** Internal SHA-512 hash class grounded in RFC 6234 §6.3 and §6.4. */
	var _SHA512 = class extends SHA2_64B {
		constructor() {
			super(64, SHA512_IV);
		}
	};
	/** Internal SHA-384 hash class grounded in RFC 6234 §6.3 and §6.4. */
	var _SHA384 = class extends SHA2_64B {
		constructor() {
			super(48, SHA384_IV);
		}
	};
	/**
	* SHA2-256 hash function from RFC 4634. In JS it's the fastest: even faster than Blake3. Some info:
	*
	* - Trying 2^128 hashes would get 50% chance of collision, using birthday attack.
	* - BTC network is doing 2^70 hashes/sec (2^95 hashes/year) as per 2025.
	* - Each sha256 hash is executing 2^18 bit operations.
	* - Good 2024 ASICs can do 200Th/sec with 3500 watts of power, corresponding to 2^36 hashes/joule.
	* @param msg - message bytes to hash
	* @param opts - Reserved hash options.
	* @returns Digest bytes.
	* @example
	* Hash a message with SHA2-256.
	* ```ts
	* sha256(new Uint8Array([97, 98, 99]));
	* ```
	*/
	var sha256 = /* @__PURE__ */ createHasher(() => new _SHA256(), /* @__PURE__ */ oidNist(1));
	/**
	* SHA2-512 hash function from RFC 4634.
	* @param msg - message bytes to hash
	* @param opts - Reserved hash options.
	* @returns Digest bytes.
	* @example
	* Hash a message with SHA2-512.
	* ```ts
	* sha512(new Uint8Array([97, 98, 99]));
	* ```
	*/
	var sha512 = /* @__PURE__ */ createHasher(() => new _SHA512(), /* @__PURE__ */ oidNist(3));
	/**
	* SHA2-384 hash function from RFC 4634.
	* @param msg - message bytes to hash
	* @param opts - Reserved hash options.
	* @returns Digest bytes.
	* @example
	* Hash a message with SHA2-384.
	* ```ts
	* sha384(new Uint8Array([97, 98, 99]));
	* ```
	*/
	var sha384 = /* @__PURE__ */ createHasher(() => new _SHA384(), /* @__PURE__ */ oidNist(2));
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+ciphers@2.4.0/node_modules/@noble/ciphers/utils.js
	/*! noble-ciphers - MIT License (c) 2023 Paul Miller (paulmillr.com) */
	/**
	* Checks if something is Uint8Array. Be careful: nodejs Buffer will return true.
	* @param a - Value to inspect.
	* @returns `true` when the value is a Uint8Array view, including Node's `Buffer`.
	* @example
	* Guards a value before treating it as raw key material.
	*
	* ```ts
	* isBytes(new Uint8Array());
	* ```
	*/
	function isBytes(a) {
		return a instanceof Uint8Array || ArrayBuffer.isView(a) && a.constructor.name === "Uint8Array" && "BYTES_PER_ELEMENT" in a && a.BYTES_PER_ELEMENT === 1;
	}
	var atitle = (title) => title ? `"${title}" ` : "";
	/**
	* Asserts something is a non-negative safe integer.
	* @param n - Value to validate.
	* @returns The validated number.
	* @throws On wrong argument types. {@link TypeError}
	* @throws On wrong argument ranges or values. {@link RangeError}
	* @example
	* Validates a non-negative length or counter.
	*
	* ```ts
	* anumber(1);
	* ```
	*/
	function anumber(n, title = "") {
		if (typeof n !== "number") throw new TypeError(atitle(title) + "expected number, got " + typeof n);
		if (!Number.isSafeInteger(n) || n < 0) throw new RangeError(atitle(title) + "expected integer >= 0, got " + n);
		return n;
	}
	/**
	* Asserts something is Uint8Array.
	* @param value - Value to validate.
	* @param length - Expected byte length.
	* @param title - Optional label used in error messages.
	* @returns The validated byte array.
	* On Node, `Buffer` is accepted too because it is a Uint8Array view.
	* @throws On wrong argument types. {@link TypeError}
	* @throws On wrong argument lengths. {@link RangeError}
	* @example
	* Validates a fixed-length nonce or key buffer.
	*
	* ```ts
	* abytes(new Uint8Array([1, 2]), 2);
	* ```
	*/
	function abytes(value, length, title = "") {
		if (isBytes(value) && (length === void 0 || value.length === length)) return value;
		if (length !== void 0) anumber(length, "length");
		const bytes = isBytes(value);
		const ofLen = length !== void 0 ? ` of length ${length}` : "";
		const got = bytes ? `length=${value.length}` : `type=${typeof value}`;
		const message = atitle(title) + "expected Uint8Array" + ofLen + ", got " + got;
		if (!bytes) throw new TypeError(message);
		throw new RangeError(message);
	}
	/**
	* Casts a typed-array view to Uint32Array.
	* @param arr - Typed-array view to reinterpret.
	* @returns Uint32Array view over the same bytes. Callers are expected to provide a
	* 4-byte-aligned offset; trailing `1..3` bytes are silently dropped.
	* @example
	* Views a byte buffer as 32-bit words for block processing.
	*
	* ```ts
	* u32(new Uint8Array(4));
	* ```
	*/
	function u32(arr) {
		return new Uint32Array(arr.buffer, arr.byteOffset, Math.floor(arr.byteLength / 4));
	}
	/**
	* Zeroizes typed arrays in place.
	* Warning: JS provides no guarantees.
	* @param arrays - Arrays to wipe.
	* @example
	* Wipes a temporary key buffer after use.
	*
	* ```ts
	* const bytes = new Uint8Array([1]);
	* clean(bytes);
	* ```
	*/
	function clean(...arrays) {
		for (let i = 0; i < arrays.length; i++) arrays[i].fill(0);
	}
	/**
	* Whether the current platform is little-endian.
	* Most are; some IBM systems are not.
	*/
	var isLE = /* @__PURE__ */ (() => new Uint8Array(new Uint32Array([287454020]).buffer)[0] === 68)();
	/**
	* Reverses byte order of one 32-bit word.
	* @param word - Unsigned 32-bit word to swap.
	* @returns The same word with bytes reversed.
	* @example
	* Swaps a big-endian word into little-endian byte order.
	*
	* ```ts
	* byteSwap(0x11223344);
	* ```
	*/
	function byteSwap(word) {
		return word << 24 & 4278190080 | word << 8 & 16711680 | word >>> 8 & 65280 | word >>> 24 & 255;
	}
	/**
	* Byte-swaps every word of a Uint32Array in place.
	* @param arr - Uint32Array whose words should be swapped.
	* @returns The same array after in-place byte swapping.
	* @example
	* Swaps every 32-bit word in a word-view buffer.
	*
	* ```ts
	* byteSwap32(new Uint32Array([0x11223344]));
	* ```
	*/
	function byteSwap32(arr) {
		for (let i = 0; i < arr.length; i++) arr[i] = byteSwap(arr[i]);
		return arr;
	}
	/**
	* Normalizes a Uint32Array view to the little-endian representation expected by cipher cores.
	* @param u - Word view to normalize in place.
	* @returns Little-endian normalized word view.
	* @example
	* Normalizes a word-view buffer before block processing.
	*
	* ```ts
	* swap32IfBE(new Uint32Array([0x11223344]));
	* ```
	*/
	var swap32IfBE = isLE ? (u) => u : byteSwap32;
	/**
	* Checks if two U8A use same underlying buffer and overlaps.
	* This is invalid and can corrupt data.
	* @param a - First byte view.
	* @param b - Second byte view.
	* @returns `true` when the views overlap in memory.
	* @example
	* Detects whether two slices alias the same backing buffer.
	*
	* ```ts
	* overlapBytes(new Uint8Array(4), new Uint8Array(4));
	* ```
	*/
	function overlapBytes(a, b) {
		if (!a.byteLength || !b.byteLength) return false;
		return a.buffer === b.buffer && a.byteOffset < b.byteOffset + b.byteLength && b.byteOffset < a.byteOffset + a.byteLength;
	}
	/**
	* If input and output overlap and input starts before output, we will overwrite end of input before
	* we start processing it, so this is not supported by forward-processing ciphers.
	* @param input - Input bytes.
	* @param output - Output bytes.
	* @throws If the output view would overwrite unread input bytes. {@link Error}
	* @example
	* Rejects an in-place layout that would overwrite unread input bytes.
	*
	* ```ts
	* const buffer = new Uint8Array(8);
	* complexOverlapBytes(buffer.subarray(0, 4), buffer.subarray(2, 6));
	* ```
	*/
	function complexOverlapBytes(input, output) {
		if (overlapBytes(input, output) && input.byteOffset < output.byteOffset) throw new Error("complex overlap of input and output is not supported");
	}
	/**
	* Wraps a cipher: validates args, ensures encrypt() can only be called once.
	* Used internally by the exported cipher constructors.
	* Output-buffer support is inferred from the wrapped `encrypt` / `decrypt`
	* arity (`fn.length === 2`), so wrapped output-capable methods must use a normal
	* second parameter, not a default/rest parameter. AAD support is explicit in
	* `params.withAAD`; optional AAD starts after the nonce slot when one is present.
	* @__NO_SIDE_EFFECTS__
	* @param params - Static cipher metadata. See {@link CipherParams}.
	* @param constructor - Cipher constructor.
	* @returns Wrapped constructor with validation.
	*/
	var wrapCipher = (params, constructor) => {
		function wrappedCipher(key, ...args) {
			abytes(key, void 0, "key");
			if (params.nonceLength !== void 0) {
				const nonce = args[0];
				abytes(nonce, params.varSizeNonce ? void 0 : params.nonceLength, "nonce");
			}
			const tagl = params.tagLength;
			const aadStart = params.nonceLength !== void 0 ? 1 : 0;
			if (!params.withAAD) {
				for (let i = aadStart; i < args.length; i++) if (isBytes(args[i])) throw new Error("AAD not supported");
			}
			if (params.withAAD && args[aadStart] !== void 0) abytes(args[aadStart], void 0, "AAD");
			const cipher = constructor(key, ...args);
			const checkOutput = (fnLength, output) => {
				if (output !== void 0) {
					if (fnLength !== 2) throw new Error("cipher output not supported");
					abytes(output, void 0, "output");
				}
			};
			let called = false;
			return {
				encrypt(data, output) {
					if (called) throw new Error("cannot encrypt() twice with same key + nonce");
					called = true;
					abytes(data, void 0, "data");
					checkOutput(cipher.encrypt.length, output);
					return cipher.encrypt(data, output);
				},
				decrypt(data, output) {
					abytes(data, void 0, "data");
					if (tagl && data.length < tagl) throw new Error("\"ciphertext\" expected length >= tagLength=" + tagl);
					checkOutput(cipher.decrypt.length, output);
					return cipher.decrypt(data, output);
				}
			};
		}
		Object.assign(wrappedCipher, params);
		return wrappedCipher;
	};
	/**
	* By default, returns u8a of length.
	* When out is available, it checks it for validity and uses it.
	* @param expectedLength - Required output length.
	* @param out - Optional destination buffer.
	* @param onlyAligned - Whether `out` must be 4-byte aligned.
	* @returns Output buffer ready for writing.
	* @throws On wrong argument types. {@link TypeError}
	* @throws If the provided output buffer has the wrong size. {@link RangeError}
	* @throws If the provided output buffer has the wrong alignment. {@link Error}
	* @example
	* Reuses a caller-provided output buffer when lengths match.
	*
	* ```ts
	* getOutput(16, new Uint8Array(16));
	* ```
	*/
	function getOutput(expectedLength, out, onlyAligned = true) {
		if (out === void 0) return new Uint8Array(expectedLength);
		abytes(out, expectedLength, "output");
		if (onlyAligned && !isAligned32(out)) throw new Error("invalid output, must be aligned");
		return out;
	}
	/**
	* Checks whether a byte array is aligned to a 4-byte offset.
	* @param bytes - Byte array to inspect.
	* @returns `true` when the view is 4-byte aligned.
	* @example
	* Checks whether a buffer can be safely viewed as Uint32Array.
	*
	* ```ts
	* isAligned32(new Uint8Array(4));
	* ```
	*/
	function isAligned32(bytes) {
		return bytes.byteOffset % 4 === 0;
	}
	/**
	* Copies bytes into a new Uint8Array.
	* @param bytes - Bytes to copy.
	* @returns Copied byte array.
	* @throws On wrong argument types. {@link TypeError}
	* @example
	* Copies input into an aligned Uint8Array before block processing.
	*
	* ```ts
	* copyBytes(new Uint8Array([1, 2]));
	* ```
	*/
	function copyBytes(bytes) {
		return Uint8Array.from(abytes(bytes));
	}
	//#endregion
	//#region ../../node_modules/.pnpm/@noble+ciphers@2.4.0/node_modules/@noble/ciphers/aes.js
	var BLOCK_SIZE = 16;
	var POLY = 283;
	function validateKeyLength(key) {
		if (![
			16,
			24,
			32
		].includes(key.length)) throw new Error("\"aes key\" expected Uint8Array of length 16/24/32, got length=" + key.length);
	}
	function mul2(n) {
		return n << 1 ^ POLY & -(n >> 7);
	}
	function mul(a, b) {
		let res = 0;
		for (; b > 0; b >>= 1) {
			res ^= a & -(b & 1);
			a = mul2(a);
		}
		return res;
	}
	var sbox = /* @__PURE__ */ (() => {
		const t = /* @__PURE__ */ new Uint8Array(256);
		for (let i = 0, x = 1; i < 256; i++, x ^= mul2(x)) t[i] = x;
		const box = /* @__PURE__ */ new Uint8Array(256);
		box[0] = 99;
		for (let i = 0; i < 255; i++) {
			let x = t[255 - i];
			x |= x << 8;
			box[t[i]] = (x ^ x >> 4 ^ x >> 5 ^ x >> 6 ^ x >> 7 ^ 99) & 255;
		}
		clean(t);
		return box;
	})();
	var invSbox = /* @__PURE__ */ sbox.map((_, j) => sbox.indexOf(j));
	var rotr32_8 = (n) => n << 24 | n >>> 8;
	var rotl32_8 = (n) => n << 8 | n >>> 24;
	function genTtable(sbox, fn) {
		if (sbox.length !== 256) throw new Error("wrong sbox length");
		const T0 = (/* @__PURE__ */ new Uint32Array(256)).map((_, j) => fn(sbox[j]));
		const T1 = T0.map(rotl32_8);
		const T2 = T1.map(rotl32_8);
		const T3 = T2.map(rotl32_8);
		const T01 = /* @__PURE__ */ new Uint32Array(65536);
		const T23 = /* @__PURE__ */ new Uint32Array(65536);
		const sbox2 = /* @__PURE__ */ new Uint16Array(65536);
		for (let i = 0; i < 256; i++) for (let j = 0; j < 256; j++) {
			const idx = i * 256 + j;
			T01[idx] = T0[i] ^ T1[j];
			T23[idx] = T2[i] ^ T3[j];
			sbox2[idx] = sbox[i] << 8 | sbox[j];
		}
		return {
			sbox,
			sbox2,
			T0,
			T1,
			T2,
			T3,
			T01,
			T23
		};
	}
	var tableEncoding = /* @__PURE__ */ genTtable(sbox, (s) => mul(s, 3) << 24 | s << 16 | s << 8 | mul(s, 2));
	var tableDecoding = /* @__PURE__ */ genTtable(invSbox, (s) => mul(s, 11) << 24 | mul(s, 13) << 16 | mul(s, 9) << 8 | mul(s, 14));
	var xPowers = /* @__PURE__ */ (() => {
		const p = /* @__PURE__ */ new Uint8Array(16);
		for (let i = 0, x = 1; i < 16; i++, x = mul2(x)) p[i] = x;
		return p;
	})();
	/** Forward AES key expansion used across ECB/CBC/CTR/GCM/CMAC/KW-style paths. */
	function expandKeyLE(key) {
		abytes(key);
		const len = key.length;
		validateKeyLength(key);
		const { sbox2 } = tableEncoding;
		const toClean = [];
		if (!isLE || !isAligned32(key)) toClean.push(key = copyBytes(key));
		const k32 = swap32IfBE(u32(key));
		const Nk = k32.length;
		const subByte = (n) => applySbox(sbox2, n, n, n, n);
		const xk = new Uint32Array(len + 28);
		xk.set(k32);
		for (let i = Nk; i < xk.length; i++) {
			let t = xk[i - 1];
			if (i % Nk === 0) t = subByte(rotr32_8(t)) ^ xPowers[i / Nk - 1];
			else if (Nk > 6 && i % Nk === 4) t = subByte(t);
			xk[i] = xk[i - Nk] ^ t;
		}
		clean(...toClean);
		return xk;
	}
	function expandKeyDecLE(key) {
		const encKey = expandKeyLE(key);
		const xk = encKey.slice();
		const Nk = encKey.length;
		const { sbox2 } = tableEncoding;
		const { T0, T1, T2, T3 } = tableDecoding;
		for (let i = 0; i < Nk; i += 4) for (let j = 0; j < 4; j++) xk[i + j] = encKey[Nk - i - 4 + j];
		clean(encKey);
		for (let i = 4; i < Nk - 4; i++) {
			const x = xk[i];
			const w = applySbox(sbox2, x, x, x, x);
			xk[i] = T0[w & 255] ^ T1[w >>> 8 & 255] ^ T2[w >>> 16 & 255] ^ T3[w >>> 24];
		}
		return xk;
	}
	function apply0123(T01, T23, s0, s1, s2, s3) {
		return T01[s0 << 8 & 65280 | s1 >>> 8 & 255] ^ T23[s2 >>> 8 & 65280 | s3 >>> 24 & 255];
	}
	function applySbox(sbox2, s0, s1, s2, s3) {
		return sbox2[s0 & 255 | s1 & 65280] | sbox2[s2 >>> 16 & 255 | s3 >>> 16 & 65280] << 16;
	}
	function encrypt(xk, s0, s1, s2, s3) {
		const { sbox2, T01, T23 } = tableEncoding;
		let k = 0;
		s0 ^= xk[k++], s1 ^= xk[k++], s2 ^= xk[k++], s3 ^= xk[k++];
		const rounds = xk.length / 4 - 2;
		for (let i = 0; i < rounds; i++) {
			const t0 = xk[k++] ^ apply0123(T01, T23, s0, s1, s2, s3);
			const t1 = xk[k++] ^ apply0123(T01, T23, s1, s2, s3, s0);
			const t2 = xk[k++] ^ apply0123(T01, T23, s2, s3, s0, s1);
			const t3 = xk[k++] ^ apply0123(T01, T23, s3, s0, s1, s2);
			s0 = t0, s1 = t1, s2 = t2, s3 = t3;
		}
		return {
			s0: xk[k++] ^ applySbox(sbox2, s0, s1, s2, s3),
			s1: xk[k++] ^ applySbox(sbox2, s1, s2, s3, s0),
			s2: xk[k++] ^ applySbox(sbox2, s2, s3, s0, s1),
			s3: xk[k++] ^ applySbox(sbox2, s3, s0, s1, s2)
		};
	}
	function decrypt(xk, s0, s1, s2, s3) {
		const { sbox2, T01, T23 } = tableDecoding;
		let k = 0;
		s0 ^= xk[k++], s1 ^= xk[k++], s2 ^= xk[k++], s3 ^= xk[k++];
		const rounds = xk.length / 4 - 2;
		for (let i = 0; i < rounds; i++) {
			const t0 = xk[k++] ^ apply0123(T01, T23, s0, s3, s2, s1);
			const t1 = xk[k++] ^ apply0123(T01, T23, s1, s0, s3, s2);
			const t2 = xk[k++] ^ apply0123(T01, T23, s2, s1, s0, s3);
			const t3 = xk[k++] ^ apply0123(T01, T23, s3, s2, s1, s0);
			s0 = t0, s1 = t1, s2 = t2, s3 = t3;
		}
		return {
			s0: xk[k++] ^ applySbox(sbox2, s0, s3, s2, s1),
			s1: xk[k++] ^ applySbox(sbox2, s1, s0, s3, s2),
			s2: xk[k++] ^ applySbox(sbox2, s2, s1, s0, s3),
			s3: xk[k++] ^ applySbox(sbox2, s3, s2, s1, s0)
		};
	}
	function validateBlockDecrypt(data, dst) {
		abytes(data);
		if (data.length % BLOCK_SIZE !== 0) throw new Error("ciphertext must be multiple of 16");
		if (dst !== void 0) {
			getOutput(data.length, dst);
			complexOverlapBytes(data, dst);
		}
	}
	function validateBlockEncrypt(plaintext, pkcs5, dst) {
		abytes(plaintext);
		let outLen = plaintext.length;
		const remaining = outLen % BLOCK_SIZE;
		if (!pkcs5 && remaining !== 0) throw new Error("plaintext must be multiple of 16");
		if (pkcs5) {
			let left = BLOCK_SIZE - remaining;
			if (!left) left = BLOCK_SIZE;
			outLen = outLen + left;
		}
		if (dst !== void 0) {
			getOutput(outLen, dst);
			complexOverlapBytes(plaintext, dst);
		}
		return outLen;
	}
	function prepareBlockEncrypt(plaintext, outLen, dst) {
		if (dst === void 0) dst = new Uint8Array(outLen);
		if (!isLE || !isAligned32(plaintext)) plaintext = copyBytes(plaintext);
		const o = u32(dst);
		return {
			b: plaintext,
			o,
			out: dst
		};
	}
	function validatePKCS(data, pkcs5) {
		if (!pkcs5) return data;
		const len = data.length;
		if (len === 0) throw new Error("pkcs7: empty ciphertext not allowed");
		const lastByte = data[len - 1];
		let valid = 1;
		valid &= lastByte - 1 >>> 31 ^ 1;
		valid &= 16 - lastByte >>> 31 ^ 1;
		for (let i = 0; i < 16; i++) {
			const shouldCheck = i - lastByte >>> 31;
			const eq = (data[len - 1 - i] ^ lastByte) === 0 ? 1 : 0;
			valid &= eq | shouldCheck ^ 1;
		}
		if (!valid) throw new Error("aes: bad decrypt");
		return data.subarray(0, len - lastByte);
	}
	function padPCKS(left) {
		const tmp = /* @__PURE__ */ new Uint8Array(16);
		const tmp32 = u32(tmp);
		tmp.set(left);
		const paddingByte = BLOCK_SIZE - left.length;
		for (let i = BLOCK_SIZE - paddingByte; i < BLOCK_SIZE; i++) tmp[i] = paddingByte;
		return tmp32;
	}
	/**
	* **CBC** (Cipher Block Chaining): Each plaintext block is XORed with the
	* previous block of ciphertext before encryption.
	* Hard to use: requires proper padding and an unpredictable IV. Unauthenticated: needs MAC.
	* Bad padding is reported as decrypt failure, which can be a padding oracle if exposed.
	* @param key - AES key bytes.
	* @param iv - 16-byte unpredictable initialization vector.
	* @param opts - Padding options. See {@link BlockOpts}.
	* @returns Cipher instance with `encrypt()` and `decrypt()`.
	* @example
	* Encrypts a padded message with a fresh key and 16-byte IV.
	*
	* ```ts
	* import { cbc } from '@noble/ciphers/aes.js';
	* import { randomBytes } from '@noble/ciphers/utils.js';
	* const key = randomBytes(16);
	* const iv = randomBytes(16);
	* const cipher = cbc(key, iv);
	* cipher.encrypt(new Uint8Array([1, 2, 3]));
	* ```
	*/
	var cbc = /* @__PURE__ */ wrapCipher({
		blockSize: 16,
		nonceLength: 16
	}, function aescbc(key, iv, opts = {}) {
		const pkcs5 = !opts.disablePadding;
		return {
			encrypt(plaintext, dst) {
				const outLen = validateBlockEncrypt(plaintext, pkcs5, dst);
				const xk = expandKeyLE(key);
				const { b: input, o, out: _out } = prepareBlockEncrypt(plaintext, outLen, dst);
				const b = u32(input);
				swap32IfBE(b);
				let _iv = iv;
				const toClean = [xk];
				if (!isLE || !isAligned32(_iv)) toClean.push(_iv = copyBytes(_iv));
				const n32 = u32(_iv);
				swap32IfBE(n32);
				let s0 = n32[0], s1 = n32[1], s2 = n32[2], s3 = n32[3];
				let i = 0;
				for (; i + 4 <= b.length;) {
					s0 ^= b[i + 0], s1 ^= b[i + 1], s2 ^= b[i + 2], s3 ^= b[i + 3];
					({s0, s1, s2, s3} = encrypt(xk, s0, s1, s2, s3));
					o[i++] = s0, o[i++] = s1, o[i++] = s2, o[i++] = s3;
				}
				if (pkcs5) {
					const tmp32 = padPCKS(plaintext.subarray(i * 4));
					swap32IfBE(tmp32);
					s0 ^= tmp32[0], s1 ^= tmp32[1], s2 ^= tmp32[2], s3 ^= tmp32[3];
					({s0, s1, s2, s3} = encrypt(xk, s0, s1, s2, s3));
					o[i++] = s0, o[i++] = s1, o[i++] = s2, o[i++] = s3;
					clean(tmp32);
				}
				swap32IfBE(o);
				clean(...toClean);
				if (input !== plaintext) clean(input);
				return _out;
			},
			decrypt(ciphertext, dst) {
				validateBlockDecrypt(ciphertext, dst);
				const xk = expandKeyDecLE(key);
				if (dst === void 0) dst = new Uint8Array(ciphertext.length);
				let _iv = iv;
				const toClean = [xk];
				if (!isLE || !isAligned32(_iv)) toClean.push(_iv = copyBytes(_iv));
				const n32 = u32(_iv);
				swap32IfBE(n32);
				if (!isLE || !isAligned32(ciphertext)) toClean.push(ciphertext = copyBytes(ciphertext));
				const b = u32(ciphertext);
				const o = u32(dst);
				swap32IfBE(b);
				let s0 = n32[0], s1 = n32[1], s2 = n32[2], s3 = n32[3];
				for (let i = 0; i + 4 <= b.length;) {
					const ps0 = s0, ps1 = s1, ps2 = s2, ps3 = s3;
					s0 = b[i + 0], s1 = b[i + 1], s2 = b[i + 2], s3 = b[i + 3];
					const { s0: o0, s1: o1, s2: o2, s3: o3 } = decrypt(xk, s0, s1, s2, s3);
					o[i++] = o0 ^ ps0, o[i++] = o1 ^ ps1, o[i++] = o2 ^ ps2, o[i++] = o3 ^ ps3;
				}
				swap32IfBE(o);
				clean(...toClean);
				return validatePKCS(dst, pkcs5);
			}
		};
	});
	//#endregion
	//#region src/encryption/errors.ts
	/** Wrong password (the Agile verifier didn't match). */
	var DocxPasswordError = class extends Error {
		constructor(message = "wrong password") {
			super(message);
			this.name = "DocxPasswordError";
		}
	};
	/**
	* The file is encrypted, but not with a scheme we can decrypt: Office 2007
	* "Standard" (RC4 / AES-128-ECB), certificate (IRM) encryption, or a broken
	* descriptor.
	*/
	var DocxEncryptionUnsupportedError = class extends Error {
		constructor(message = "unsupported encryption") {
			super(message);
			this.name = "DocxEncryptionUnsupportedError";
		}
	};
	//#endregion
	//#region src/encryption/agile.ts
	var BLOCK_KEY_VERIFIER_INPUT = new Uint8Array([
		254,
		167,
		210,
		118,
		59,
		75,
		158,
		121
	]);
	var BLOCK_KEY_VERIFIER_VALUE = new Uint8Array([
		215,
		170,
		15,
		109,
		48,
		97,
		52,
		78
	]);
	var BLOCK_KEY_ENCRYPTED_KEY = new Uint8Array([
		20,
		110,
		11,
		231,
		171,
		172,
		208,
		214
	]);
	var SEGMENT_LENGTH = 4096;
	var YIELD_EVERY = 8192;
	var HASHES = {
		SHA1: sha1,
		SHA256: sha256,
		SHA384: sha384,
		SHA512: sha512
	};
	/** `EncryptionInfo` stream → Agile parameters, or null when it isn't Agile/AES-CBC. */
	function parseAgileEncryptionInfo(info) {
		if (info.length < 12) return null;
		const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
		const versionMajor = view.getUint16(0, true);
		const versionMinor = view.getUint16(2, true);
		const flags = view.getUint32(4, true);
		if (versionMajor !== 4 || versionMinor !== 4 || !(flags & 64)) return null;
		const xml = decodeXml(info.subarray(8));
		const keyData = attribute(xml, "keyData", "saltValue");
		const encryptedKey = element(xml, "encryptedKey");
		if (!encryptedKey) return null;
		const cipherAlgorithm = attr(encryptedKey, "cipherAlgorithm");
		const cipherChaining = attr(encryptedKey, "cipherChaining");
		const hashAlgorithm = (attr(encryptedKey, "hashAlgorithm") || "SHA1").toUpperCase();
		const keyBits = Number(attr(encryptedKey, "keyBits"));
		const blockSize = Number(attr(encryptedKey, "blockSize"));
		const hashSize = Number(attr(encryptedKey, "hashSize"));
		const spinCount = Number(attr(encryptedKey, "spinCount"));
		const passwordSalt = base64(attr(encryptedKey, "saltValue"));
		const encryptedKeyValue = base64(attr(encryptedKey, "encryptedKeyValue"));
		const encryptedVerifierHashInput = base64(attr(encryptedKey, "encryptedVerifierHashInput"));
		const encryptedVerifierHashValue = base64(attr(encryptedKey, "encryptedVerifierHashValue"));
		const keyDataSalt = base64(keyData);
		if (cipherAlgorithm !== "AES" || cipherChaining !== "ChainingModeCBC" || !(hashAlgorithm in HASHES) || !passwordSalt || !keyDataSalt || !encryptedKeyValue || !encryptedVerifierHashInput || !encryptedVerifierHashValue || !Number.isFinite(keyBits) || !Number.isFinite(spinCount) || !Number.isFinite(blockSize) || !Number.isFinite(hashSize) || blockSize !== 16 || keyBits % 8 !== 0) throw new DocxEncryptionUnsupportedError(`unsupported Agile encryption parameters (${cipherAlgorithm}/${cipherChaining}/${hashAlgorithm})`);
		return {
			spinCount,
			hashAlgorithm,
			keyBits,
			blockSize,
			hashSize,
			passwordSalt,
			keyDataSalt,
			encryptedKeyValue,
			encryptedVerifierHashInput,
			encryptedVerifierHashValue
		};
	}
	/** Decrypts an `EncryptedPackage` stream. Throws DocxPasswordError on a bad password. */
	async function decryptAgilePackage(info, password, encryptedPackage) {
		const hash = HASHES[info.hashAlgorithm];
		const keyLength = info.keyBits >> 3;
		const iterate = await iteratedHash(hash, info.passwordSalt, password, info.spinCount);
		const keyFor = (blockKey) => hash(concat(iterate, blockKey)).subarray(0, keyLength);
		const secretKey = aesCbcDecrypt(keyFor(BLOCK_KEY_ENCRYPTED_KEY), info.passwordSalt, info.encryptedKeyValue);
		const verifierInput = aesCbcDecrypt(keyFor(BLOCK_KEY_VERIFIER_INPUT), info.passwordSalt, info.encryptedVerifierHashInput);
		const verifierHash = aesCbcDecrypt(keyFor(BLOCK_KEY_VERIFIER_VALUE), info.passwordSalt, info.encryptedVerifierHashValue);
		if (!equals(hash(verifierInput), verifierHash.subarray(0, info.hashSize))) throw new DocxPasswordError("the password did not match");
		if (encryptedPackage.length < 8) throw new DocxEncryptionUnsupportedError("empty EncryptedPackage");
		const totalSize = Number(new DataView(encryptedPackage.buffer, encryptedPackage.byteOffset, 8).getBigUint64(0, true));
		const out = new Uint8Array(totalSize);
		let written = 0;
		for (let offset = 8, index = 0; written < totalSize; offset += SEGMENT_LENGTH, index++) {
			const segment = encryptedPackage.subarray(offset, Math.min(offset + SEGMENT_LENGTH, encryptedPackage.length));
			if (!segment.length) break;
			const plain = aesCbcDecrypt(secretKey, hash(concat(info.keyDataSalt, uint32(index))).subarray(0, info.blockSize), segment);
			const take = Math.min(plain.length, totalSize - written);
			out.set(plain.subarray(0, take), written);
			written += take;
		}
		if (written < totalSize) throw new DocxEncryptionUnsupportedError("truncated EncryptedPackage");
		return out;
	}
	/** h = H(salt + password_utf16le), then h = H(u32le(i) + h) for i in [0, spinCount). */
	async function iteratedHash(hash, salt, password, spinCount) {
		let h = hash(concat(salt, utf16le(password)));
		for (let i = 0; i < spinCount; i++) {
			h = hash(concat(uint32(i), h));
			if (i % YIELD_EVERY === 8191) await yieldToHost();
		}
		return h;
	}
	function yieldToHost() {
		return new Promise((resolve) => setTimeout(resolve, 0));
	}
	function aesCbcDecrypt(key, iv, data) {
		if (data.length % 16 !== 0) throw new DocxEncryptionUnsupportedError("ciphertext is not block aligned");
		return cbc(key, iv, { disablePadding: true }).decrypt(data);
	}
	function concat(...parts) {
		let length = 0;
		for (const part of parts) length += part.length;
		const out = new Uint8Array(length);
		let offset = 0;
		for (const part of parts) {
			out.set(part, offset);
			offset += part.length;
		}
		return out;
	}
	function uint32(value) {
		const out = /* @__PURE__ */ new Uint8Array(4);
		new DataView(out.buffer).setUint32(0, value, true);
		return out;
	}
	function utf16le(text) {
		const out = new Uint8Array(text.length * 2);
		for (let i = 0; i < text.length; i++) {
			const code = text.charCodeAt(i);
			out[i * 2] = code & 255;
			out[i * 2 + 1] = code >> 8;
		}
		return out;
	}
	function equals(a, b) {
		if (a.length !== b.length) return false;
		let diff = 0;
		for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
		return diff === 0;
	}
	function base64(value) {
		if (!value) return null;
		try {
			const binary = atob(value);
			const out = new Uint8Array(binary.length);
			for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
			return out;
		} catch {
			return null;
		}
	}
	function decodeXml(bytes) {
		let text = "";
		const start = bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191 ? 3 : 0;
		for (let i = start; i < bytes.length; i++) text += String.fromCharCode(bytes[i]);
		return text;
	}
	/** First `<tag ...>` element in the descriptor, as its attribute source text. */
	function element(xml, tag) {
		const match = xml.match(new RegExp(`<(?:[\\w.-]+:)?${tag}\\b[^>]*>`));
		return match ? match[0] : null;
	}
	function attr(source, name) {
		if (!source) return void 0;
		const match = source.match(new RegExp(`${name}="([^"]*)"`));
		return match ? match[1] : void 0;
	}
	/** Attribute `name` of the first `<tag ...>` element. */
	function attribute(xml, tag, name) {
		return attr(element(xml, tag), name);
	}
	//#endregion
	//#region src/encryption/index.ts
	/** Cheap classification of the container, before any parsing/decryption. */
	function detectOfficeFileKind(data) {
		const bytes = toBytes(data);
		if (bytes.length >= 4 && bytes[0] === 80 && bytes[1] === 75 && bytes[2] === 3 && bytes[3] === 4) return "ooxml";
		if (!isCompoundFile(bytes)) return "unknown";
		const compound = readCompoundFile(bytes);
		if (!compound) return "unknown";
		return compound.readStream("EncryptionInfo") && compound.readStream("EncryptedPackage") ? "encrypted" : "legacy-binary";
	}
	/** True when the file needs a password (same check as `detectOfficeFileKind`). */
	function isEncryptedDocx(data) {
		return detectOfficeFileKind(data) === "encrypted";
	}
	/**
	* Decrypts a password-protected OOXML file and returns the plain .docx bytes.
	*
	* @throws DocxPasswordError when the password is wrong.
	* @throws DocxEncryptionUnsupportedError for schemes we can't handle
	*   (Office 2007 "Standard", certificate encryption, corrupt descriptors).
	*/
	async function decryptDocx(data, password) {
		const compound = readCompoundFile(toBytes(data));
		if (!compound) throw new DocxEncryptionUnsupportedError("not a compound file");
		const encryptionInfo = compound.readStream("EncryptionInfo");
		const encryptedPackage = compound.readStream("EncryptedPackage");
		if (!encryptionInfo || !encryptedPackage) throw new DocxEncryptionUnsupportedError("no EncryptionInfo/EncryptedPackage stream");
		const agile = parseAgileEncryptionInfo(encryptionInfo);
		if (!agile) throw new DocxEncryptionUnsupportedError("only ECMA-376 Agile encryption is supported");
		const plain = await decryptAgilePackage(agile, password, encryptedPackage);
		return plain.buffer.slice(plain.byteOffset, plain.byteOffset + plain.byteLength);
	}
	function toBytes(data) {
		return data instanceof Uint8Array ? data : new Uint8Array(data);
	}
	//#endregion
	//#region src/docx-preview.ts
	var defaultOptions = {
		ignoreHeight: false,
		ignoreWidth: false,
		ignoreFonts: false,
		breakPages: true,
		paginate: false,
		debug: false,
		experimental: false,
		className: "docx",
		inWrapper: true,
		hideWrapperOnPrint: false,
		trimXmlDeclaration: true,
		ignoreLastRenderedPageBreak: true,
		renderHeaders: true,
		renderFooters: true,
		renderFootnotes: true,
		renderEndnotes: true,
		useBase64URL: false,
		renderChanges: false,
		renderComments: false,
		renderAltChunks: true,
		h
	};
	function parseAsync(data, userOptions) {
		const ops = {
			...defaultOptions,
			...userOptions
		};
		return WordDocument.load(data, new DocumentParser(ops), ops);
	}
	async function renderDocument(document, userOptions) {
		const ops = {
			...defaultOptions,
			...userOptions
		};
		return await new HtmlRenderer().render(document, ops);
	}
	function renderFontFaces(fonts) {
		const css = fonts.map((f) => {
			const props = [`font-family: "${f.name.replace(/["\\\r\n]/g, "")}"`, `src: ${f.src}`];
			if (f.weight != null) props.push(`font-weight: ${f.weight}`);
			if (f.style != null) props.push(`font-style: ${f.style}`);
			return `@font-face { ${props.join("; ")}; }`;
		}).join("\n");
		const el = document.createElement("style");
		el.textContent = css;
		return el;
	}
	async function renderAsync(data, bodyContainer, styleContainer, userOptions) {
		var _styleContainer, _ops$fonts, _ops$fonts2;
		const doc = await parseAsync(data, userOptions);
		const nodes = await renderDocument(doc, userOptions);
		const ops = {
			...defaultOptions,
			...userOptions
		};
		(_styleContainer = styleContainer) !== null && _styleContainer !== void 0 || (styleContainer = bodyContainer);
		styleContainer.innerHTML = "";
		bodyContainer.innerHTML = "";
		if ((_ops$fonts = ops.fonts) === null || _ops$fonts === void 0 ? void 0 : _ops$fonts.length) nodes.unshift(renderFontFaces(ops.fonts));
		for (let n of nodes) (n.nodeName === "STYLE" ? styleContainer : bodyContainer).appendChild(n);
		const fontSet = typeof document !== "undefined" ? document.fonts : null;
		if (((_ops$fonts2 = ops.fonts) === null || _ops$fonts2 === void 0 ? void 0 : _ops$fonts2.length) && (fontSet === null || fontSet === void 0 ? void 0 : fontSet.load)) try {
			await Promise.all(ops.fonts.map((f) => {
				var _f$style, _f$weight;
				return fontSet.load(`${(_f$style = f.style) !== null && _f$style !== void 0 ? _f$style : "normal"} ${(_f$weight = f.weight) !== null && _f$weight !== void 0 ? _f$weight : "normal"} 12px "${f.name.replace(/["\\\r\n]/g, "")}"`);
			}));
		} catch {}
		if (ops.paginate) {
			if (fontSet === null || fontSet === void 0 ? void 0 : fontSet.ready) try {
				await fontSet.ready;
			} catch {}
			paginateWrapper(bodyContainer, ops.className);
		}
		return doc;
	}
	//#endregion
	exports.DocxEncryptionUnsupportedError = DocxEncryptionUnsupportedError;
	exports.DocxPasswordError = DocxPasswordError;
	exports.decryptDocx = decryptDocx;
	exports.defaultOptions = defaultOptions;
	exports.detectOfficeFileKind = detectOfficeFileKind;
	exports.isEncryptedDocx = isEncryptedDocx;
	exports.parseAsync = parseAsync;
	exports.renderAsync = renderAsync;
	exports.renderDocument = renderDocument;
});

//# sourceMappingURL=docx-preview.js.map