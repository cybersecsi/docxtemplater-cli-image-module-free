#!/usr/bin/env node


"use strict";

var argv = require('minimist')(process.argv.slice(2));
var fs = require("fs");
var JSZip = require("jszip");
var Docxtemplater = require("docxtemplater");
var path = require("path");
var expressions = require("bluerider");
var _ = require('lodash');
var docx = require("docx");
var xml = require("xml");
var htmlparser = require("htmlparser2");
var flatten = require("flat");
var unflatten = require('flat').unflatten

function showHelp() {
	console.log("Usage: docxtemplater input.docx data.json output.docx");
	process.exit(1);
}

if (argv.help) {
	showHelp();
}

function parser(tag) {
    if (tag === ".") {
        return {
            get: function get(scope) {
                return scope;
            }
        };
    }
	var expr = expressions.compile(tag.replace(/’/g, "'"));
	return {
		get: function get(scope) {
			return expr(scope);
		}
	};
}

// Split html data into paragraphs and images
function splitHTMLParagraphs(htmldata) {
    var result = []
    if (!htmldata)
        return result

    var splitted = htmldata.split(/(<img.+?src=".*?".+?alt=".*?".*?>)/)

    for (let value of splitted){
        if (value.startsWith("<img")) {
            var src = value.match(/<img.+src="(.*?)"/) || ""
            var alt = value.match(/<img.+alt="(.*?)"/) || ""
            if (src && src.length > 1) src = src[1]
            if (alt && alt.length > 1) alt = _.unescape(alt[1])
            if (result.length === 0)
                result.push({text: "", images: []})
            result[result.length-1].images.push({image: src, caption: alt})
        }
        else if (value === "") {
            continue
        }
        else {
            result.push({text: value, images: []})
        }
    }
    return result
}


function html2ooxml(html, style = '') {
    if (html === '')
        return html
    if (!html.match(/^<.+>/))
        html = `<p>${html}</p>`
    var doc = new docx.Document();
    var paragraphs = []
    var cParagraph = null
    var cRunProperties = {}
    var cParagraphProperties = {}
    var list_state = []
    var inCodeBlock = false
    var parser = new htmlparser.Parser(
    {
        onopentag(tag, attribs) {
            if (tag === "h1") {
                cParagraph = new docx.Paragraph({heading: 'Heading1'})
            }
            else if (tag === "h2") {
                cParagraph = new docx.Paragraph({heading: 'Heading2'})
            }
            else if (tag === "h3") {
                cParagraph = new docx.Paragraph({heading: 'Heading3'})
            }
            else if (tag === "h4") {
                cParagraph = new docx.Paragraph({heading: 'Heading4'})
            }
            else if (tag === "h5") {
                cParagraph = new docx.Paragraph({heading: 'Heading5'})
            }
            else if (tag === "h6") {
                cParagraph = new docx.Paragraph({heading: 'Heading6'})
            }
            else if (tag === "div" || tag === "p") {
                if (style && typeof style === 'string')
                    cParagraphProperties.style = style
                cParagraph = new docx.Paragraph(cParagraphProperties)
            }
            else if (tag === "pre") {
                inCodeBlock = true
                cParagraph = new docx.Paragraph({style: "Code"})
            }
            else if (tag === "b" || tag === "strong") {
                cRunProperties.bold = true
            }
            else if (tag === "i" || tag === "em") {
                cRunProperties.italics = true
            }
            else if (tag === "u") {
                cRunProperties.underline = {}
            }
            else if (tag === "strike" || tag === "s") {
                cRunProperties.strike = true
            }
            else if (tag === "br") {
                if (inCodeBlock) {
                    paragraphs.push(cParagraph)
                    cParagraph = new docx.Paragraph({style: "Code"})
                }
                else
                    cParagraph.addChildElement(new docx.Run({}).break())
            }
            else if (tag === "ul") {
                list_state.push('bullet')
            }
            else if (tag === "ol") {
                list_state.push('number')
            }
            else if (tag === "li") {
                var level = list_state.length - 1
                if (level >= 0 && list_state[level] === 'bullet')
                    cParagraphProperties.bullet = {level: level}
                else if (level >= 0 && list_state[level] === 'number')
                    cParagraphProperties.numbering = {reference: 2, level: level}
                else
                    cParagraphProperties.bullet = {level: 0}
            }
            else if (tag === "code") {
                cRunProperties.style = "CodeChar"
            }
        },

        ontext(text) {
            if (text && cParagraph) {
                cRunProperties.text = text
                cParagraph.addChildElement(new docx.TextRun(cRunProperties))
            }
        },

        onclosetag(tag) {
            if (['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'p', 'pre', 'img'].includes(tag)) {
                paragraphs.push(cParagraph)
                cParagraph = null
                cParagraphProperties = {}
                if (tag === 'pre')
                    inCodeBlock = false
            }
            else if (tag === "b" || tag === "strong") {
                delete cRunProperties.bold
            }
            else if (tag === "i" || tag === "em") {
                delete cRunProperties.italics
            }
            else if (tag === "u") {
                delete cRunProperties.underline
            }
            else if (tag === "strike" || tag === "s") {
                delete cRunProperties.strike
            }
            else if (tag === "ul" || tag === "ol") {
                list_state.pop()
                if (list_state.length === 0)
                    cParagraphProperties = {}
            }
            else if (tag === "code") {
                delete cRunProperties.style
            }
        },

        onend() {
            doc.addSection({
                children: paragraphs
            })
        }
    }, { decodeEntities: true })

    // For multiline code blocks
    html = html.replace(/\n/g, '<br>')
    parser.write(html)
    parser.end()

    var prepXml = doc.document.body.prepForXml()
    var filteredXml = prepXml["w:body"].filter(e => {return Object.keys(e)[0] === "w:p"})
    var dataXml = xml(filteredXml)

    return dataXml
        
}


// Convert HTML data to Open Office XML format: {@input | convertHTML: 'customStyle'}
expressions.filters.convertHTML = function(input, style) {
    if (typeof input === 'undefined')
        var result = html2ooxml('')
    else
        var result = html2ooxml(input.replace(/(<p><\/p>)+$/, ''), style)
    return result;
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

// replace html data with paragraphs and images (the fields to be replaced must end with "_html")
var flattened_data = flatten(data)
var keys = Object.keys(flattened_data)
for (let key of keys){
    if (key.endsWith("_html")){
        flattened_data[key] = splitHTMLParagraphs(flattened_data[key]);
    }
}
data = unflatten(flattened_data)

doc.loadZip(zip).setOptions({ parser: parser, paragraphLoop: true, linebreaks: true }).setData(data);

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
