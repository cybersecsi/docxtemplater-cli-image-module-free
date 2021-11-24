#!/usr/bin/env node


"use strict";

var argv = require('minimist')(process.argv.slice(2));
var fs = require("fs");
var JSZip = require("jszip");
var Docxtemplater = require("docxtemplater");
var path = require("path");
var expressions = require("bluerider");

function showHelp() {
	console.log("Usage: docxtemplater input.docx data.json output.docx");
	process.exit(1);
}

if (argv.help) {
	showHelp();
}

function parser(tag) {
	var expr = expressions.compile(tag.replace(/’/g, "'"));
	return {
		get: function get(scope) {
			return expr(scope);
		}
	};
}

var args = argv._;
if (args.length !== 3) {
	showHelp();
}
var input = fs.readFileSync(args[0], "binary");
var data = JSON.parse(fs.readFileSync(args[1], "utf-8"));
var output = args[2];

var zip = new JSZip(input);
var doc = new Docxtemplater();

if (data && data.config && data.config.modules && data.config.modules.indexOf("docxtemplater-image-module-free") !== -1) {
	var ImageModule = require("docxtemplater-image-module-free");
	var sizeOf = require("image-size");
	var fileType = args[0].indexOf(".pptx") !== -1 ? "pptx" : "docx";
	var imageDir = path.resolve(process.cwd(), data.config.imageDir || "") + path.sep;
	var opts = {};
	opts.centered = false;
	opts.fileType = fileType;

	opts.getImage = function (tagValue) {
		var filePath = path.resolve(imageDir, tagValue);

		if (filePath.indexOf(imageDir) !== 0) {
			throw new Error("Images must be stored under folder: " + imageDir);
		}
		return fs.readFileSync(filePath, "binary");
	};

	opts.getSize = function (img, tagValue) {
		var filePath = path.resolve(imageDir, tagValue);

		if (filePath.indexOf(imageDir) !== 0) {
			throw new Error("Images must be stored under folder: " + imageDir);
		}

		var dimensions = sizeOf(filePath);
		if (dimensions.width > 600) {
			var divider = dimensions.width / 600;
			dimensions.width = 600;
			dimensions.height = Math.floor(dimensions.height / divider);
		}
		return [dimensions.width, dimensions.height];
	};

	var imageModule = new ImageModule(opts);
	doc.attachModule(imageModule);
}

doc.loadZip(zip).setOptions({ parser: parser }).setData(data);

function transformError(error) {
	var e = {
		message: error.message,
		name: error.name,
		stack: error.stack,
		properties: error.properties
	};
	if (e.properties && e.properties.rootError) {
		e.properties.rootError = transformError(error.properties.rootError);
	}
	if (e.properties && e.properties.errors) {
		e.properties.errors = e.properties.errors.map(transformError);
	}
	return e;
}

try {
	doc.render();
} catch (error) {
	var e = transformError(error);
	// The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
	console.log(JSON.stringify({ error: e }, null, 2));
	throw error;
}

var generated = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

fs.writeFileSync(output, generated);
