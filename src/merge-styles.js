var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;
const {xml2js, js2xml} = require('xml-js');
const { XMLParser, XMLBuilder, XMLValidator} = require("fast-xml-parser");


var prepareStyles = function(files, style) {
    var serializer = new XMLSerializer();

    files.forEach(function(zip, index) {
        if (index > 0 ) {
            var xmlString = zip.file("word/styles.xml").asText();
            var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
            var nodes = xml.getElementsByTagName('w:style');
    
            for (var node in nodes) {
                if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                    var styleId = nodes[node].getAttribute('w:styleId');
                    nodes[node].setAttribute('w:styleId', styleId + '' + index);
                    let elements = nodes[node].getElementsByTagName('w:name');
                    if (elements.length > 0 && elements[0]){
                        let val = elements[0].getAttribute('w:val');
                        elements[0].setAttribute('w:val', val + '' + index);
                    }

                    var basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
                    if (basedonStyle) {
                        var basedonStyleId = basedonStyle.getAttribute('w:val');
                        basedonStyle.setAttribute('w:val', basedonStyleId + '' + index);
                    }
    
                    var w_next = nodes[node].getElementsByTagName('w:next')[0];
                    if (w_next) {
                        var w_next_ID = w_next.getAttribute('w:val');
                        w_next.setAttribute('w:val', w_next_ID + '' + index);
                    }
    
                    var w_link = nodes[node].getElementsByTagName('w:link')[0];
                    if (w_link) {
                        var w_link_ID = w_link.getAttribute('w:val');
                        w_link.setAttribute('w:val', w_link_ID + '' + index);
                    }
    
                    var numId = nodes[node].getElementsByTagName('w:numId')[0];
                    if (numId) {
                        var numId_ID = numId.getAttribute('w:val');
                        numId.setAttribute('w:val', numId_ID + index);
                    }
    
                    updateStyleRel_Content(zip, index, styleId);
                }
            }
    
            var startIndex = xmlString.indexOf("<w:styles ");
            xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));
    
            zip.file("word/styles.xml", xmlString);
        }
        // console.log(nodes);
    });
};

var mergeStyles = function(files, _styles) {

    files.forEach(function(zip) {

        var xml = zip.file("word/styles.xml").asText();

        xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));

        _styles.push(xml);

    });
};

const mergeStyles2 = function(files, _styles) {
    const existingStyles = {};

    files.forEach(function(zip, zipIdx) {
        let xmlString = zip.file("word/styles.xml").asText();
        // const parser = new XMLParser({preserveOrder: true, ignoreAttributes: false, });
        let xml = xml2js(xmlString);
        for (let i = 0; i < xml.elements[0].elements.length; i++) {
            const styleNode = xml.elements[0].elements[i];
            if (styleNode.name === 'w:style') {
                const styleId = styleNode.attributes['w:styleId'];
                if (!existingStyles[styleId]) {
                    existingStyles[styleId] = true;
                    if (zipIdx > 0) {
                        _styles.push(styleNode);
                    }
                }
            }
        }
    });
}

const updateStyleRel_Content = function(zip, fileIndex, styleId) {
    const xmlFiles = zip.folder('word').files;
    for (let xmlFile in xmlFiles) {
        if (/^word\/(document|header|footer)\d*\.xml$/.test(xmlFile)) {
            let xmlString = zip.file(xmlFile).asText();
            const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

            xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '' + fileIndex + '"');

            // zip.file("word/document.xml", "");

            zip.file(xmlFile, xmlString);
        }
    }

};

var generateStyles = function(zip, _style) {
    var xml = zip.file("word/styles.xml").asText();
    var startIndex = xml.indexOf("<w:style ");
    var endIndex = xml.indexOf("</w:styles>");

    // console.log(xml.substring(startIndex, endIndex))

    xml = xml.replace(xml.slice(startIndex, endIndex), _style.join(''));

    // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
    // console.log(this._style.join(''))
    // console.log(xml)

    zip.file("word/styles.xml", xml);
};

const generateStyles2 = function(zip, _style) {
    let xmlString = zip.file("word/styles.xml").asText();
    const xml = xml2js(xmlString);
    for (let i = 0; i < _style.length; i++) {
        xml.elements[0].elements.push(_style[i]);
    }

    xmlString = js2xml(xml);
    zip.file("word/styles.xml", xmlString);
};

module.exports = {
    mergeStyles: mergeStyles,
    prepareStyles: prepareStyles,
    updateStyleRel_Content: updateStyleRel_Content,
    generateStyles: generateStyles,
    mergeStyles2,
    generateStyles2,
};