var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');
const { prepareHeaderFooterRelations, copyHeaderAndFooterFiles } = require('./merge-headers-and-footers');

function DocxMerger(options, files) {

    this._body = [];
    this._header = [];
    this._footer = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._mergeAsSections = typeof options.mergeAsSections !== 'undefined' ? !!options.mergeAsSections : true;
    this._files = [];
    var self = this;
    (files || []).forEach(function(file) {
        self._files.push(new JSZip(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};
    this._headersAndFooters = {};

    this._builder = this._body;

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };

    this.insertRaw = function(xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function(files) {

        var self = this;
        this._builder = this._body;
        let maxRelId;
        if (this._mergeAsSections) {
            maxRelId = prepareHeaderFooterRelations(files, this._headersAndFooters);
        }
        
        RelContentType.mergeContentTypes(files, this._contentTypes, this._mergeAsSections);
        
        if (this._mergeAsSections) {
            maxRelId = Media.prepareMediaFiles2(files, maxRelId, this._media);
        } else {
            Media.prepareMediaFiles(files, this._media);
        }

        RelContentType.mergeRelations(files, this._rel);

        bulletsNumbering.prepareNumbering(files);
        bulletsNumbering.mergeNumbering(files, this._numbering);

        if (this._mergeAsSections) {
            Style.mergeStyles2(files, this._style);
        } else {
            Style.prepareStyles(files, this._style);
            Style.mergeStyles(files, this._style);
        }

        files.forEach(function(zip, index) {
            //var zip = new JSZip(file);
            let xmlString = zip.file("word/document.xml").asText();
            const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

            if (self._mergeAsSections) {
                const childNodesCount = xml.documentElement.childNodes[0].childNodes.length;
                
                if (index < files.length-1) {
                    const sectionBreak = xml.documentElement.childNodes[0].childNodes[childNodesCount - 1];
                    // TODO: differ section headers and footers continuity
                    
                    xml.documentElement.childNodes[0].removeChild(sectionBreak);
                    const lastWPTag = xml.documentElement.childNodes[0].childNodes[childNodesCount - 2];
                    let wpPropertiesList = lastWPTag.getElementsByTagName('w:pPr');
                    if (wpPropertiesList.length < 1) {
                        wpPropertiesList = [xml.createElement('w:pPr')];
                        lastWPTag.appendChild(wpPropertiesList[0]);
                    }

                    wpPropertiesList[0].appendChild(sectionBreak);
                }
                
                let serializer = new XMLSerializer();
                // const s = serializer.serializeToString(xml.documentElement.childNodes[0].childNodes);
                
                for (let i = 0; i < xml.documentElement.childNodes[0].childNodes.length; i++){
                    self.insertRaw(xml.documentElement.childNodes[0].childNodes[i]);
                };
                
                // const body = serializer.serializeToString(xml.documentElement.childNodes[0]);
                // xmlString = body.substring(body.indexOf("<w:body>") + 8, xmlString.indexOf("</w:body>"));

                // self.insertRaw(xml.documentElement.childNodes[0].childNodes.toString());
                
                // xmlString = xml.documentElement.childNodes[0].childNodes.toString()
                // console.log('s === xmlString ====>  ', s === xmlString);
            }
            else {
                xmlString = xmlString.substring(xmlString.indexOf("<w:body>") + 8);
                xmlString = xmlString.substring(0, xmlString.indexOf("</w:body>"));
                xmlString = xmlString.substring(0, xmlString.lastIndexOf("<w:sectPr"));    
                
                self.insertRaw(xmlString);
                if (self._pageBreak && index < files.length-1)
                    self.insertPageBreak();
            }
        });
    };

    this.save = function(type, callback) {
        var zip = this._files[0];

        let xml = zip.file("word/document.xml").asText();

        if (this._mergeAsSections) {
            xml = new DOMParser().parseFromString(xml, 'text/xml');
            const clonedBody = xml.documentElement.childNodes[0].cloneNode();
            for (let elementI = 0; elementI < this._body.length; elementI++) {
                clonedBody.appendChild(this._body[elementI]);
            }
            xml.documentElement.replaceChild(clonedBody, xml.documentElement.childNodes[0])
            // xml.documentElement.childNodes[0].childNodes = this._body;
            copyHeaderAndFooterFiles(zip, this._files, this._headersAndFooters);
            RelContentType.generateContentTypes(zip, this._contentTypes);
            Media.copyMediaFiles2(zip, this._media, this._files);
            RelContentType.generateRelations(zip, this._rel);
            bulletsNumbering.generateNumbering(zip, this._numbering);
            Style.generateStyles2(zip, this._style);
            
            let serializer = new XMLSerializer();

            zip.file("word/document.xml", serializer.serializeToString(xml));
    
            callback(zip.generate({ 
                type: type,
                compression: "DEFLATE",
                compressionOptions: {
                    level: 4
                }
            }));
        } else {
            var startIndex = xml.indexOf("<w:body>") + 8;
            var endIndex = xml.lastIndexOf("<w:sectPr");
    
            xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));
    
            RelContentType.generateContentTypes(zip, this._contentTypes);
            Media.copyMediaFiles(zip, this._media, this._files);
            RelContentType.generateRelations(zip, this._rel);
            bulletsNumbering.generateNumbering(zip, this._numbering);
            Style.generateStyles(zip, this._style);
    
            zip.file("word/document.xml", xml);
    
            callback(zip.generate({ 
                type: type,
                compression: "DEFLATE",
                compressionOptions: {
                    level: 4
                }
            }));
        }
    };


    if (this._files.length > 0) {

        this.mergeBody(this._files);
    }
}


module.exports = DocxMerger;
