const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;
const { getMaxRelationID } = require('./util');

const ContentTypeHeader = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const ContentTypeFooter = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

const RelationshipTypeHeader = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const RelationshipTypeFooter = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

const getMaxHeaderFooterCount = zip => {
    let xmlString = zip.file("[Content_Types].xml").asText();
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const childNodes = xml.getElementsByTagName('Types')[0].childNodes;

    let headerCount = 0;
    let footerCount = 0;
    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var contentType = childNodes[node].getAttribute('ContentType');
            if (ContentTypeHeader === contentType) {
                const partName = childNodes[node].getAttribute('PartName');
                let matches = partName.match(/(\d+)/);
                if (matches) {
                    headerCount = Math.max(headerCount, Number(matches[0]));
                }
            } else if (ContentTypeFooter === contentType) {
                const partName = childNodes[node].getAttribute('PartName');
                let matches = partName.match(/(\d+)/);
                if (matches) {
                    footerCount = Math.max(footerCount, Number(matches[0]));
                }
            }
        }
    }

    return { headerCount, footerCount };
};

const prepareHeaderFooterRelations = (files, _headersAndFooters) => {
    if (files.length <= 1) {
        return;
    }

    const hfCount = getMaxHeaderFooterCount(files[0]);
    let maxRelId = getMaxRelationID(files[0], null);
    for (let i = 1; i < files.length; i++) {
        const tempHFCount = getMaxHeaderFooterCount(files[i]);
        hfCount.headerCount = Math.max(hfCount.headerCount, tempHFCount.headerCount);
        hfCount.footerCount = Math.max(hfCount.footerCount, tempHFCount.footerCount);
        const headersAndFooters = [];
        updateHeaderFooterContentTypes(files[i], i, hfCount, headersAndFooters);
        maxRelId = updateHeaderAndFooterRelations(files[i], headersAndFooters, maxRelId);
        updateHeaderAndFooterRelationsInDocument(files[i], headersAndFooters);
        renameHeaderAndFooterFiles(files[i], headersAndFooters);
        _headersAndFooters[i] = headersAndFooters;
    }

    return maxRelId;
};

const updateHeaderFooterContentTypes = (zip, fileIndex, hfCount, headersAndFooters) => {
    var xmlString = zip.file("[Content_Types].xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var childNodes = xml.getElementsByTagName('Types')[0].childNodes;
    let count = 0;
    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var contentType = childNodes[node].getAttribute('ContentType');
            if (ContentTypeHeader === contentType) {
                const regex = /(header)\d+(.xml)/;
                let partName = childNodes[node].getAttribute('PartName');
                let matches = partName.match(regex);
                if (matches) {
                    const newHeader = `header${++hfCount.headerCount}.xml`;
                    // _headers[`${matches[0]}`] = newHeader;
                    headersAndFooters[count++] = {
                        oldTarget: matches[0],
                        newTarget: newHeader
                    };
                    childNodes[node].setAttribute('PartName', partName.replace(regex, newHeader));
                }
            } else if (ContentTypeFooter === contentType) {
                const regex = /(footer)\d+(.xml)/;
                let partName = childNodes[node].getAttribute('PartName');
                let matches = partName.match(regex);
                if (matches) {
                    const newFooter = `footer${++hfCount.footerCount}.xml`;
                    // _footers[`${matches[0]}`] = newFooter;
                    headersAndFooters[count++] = {
                        oldTarget: matches[0],
                        newTarget: newFooter
                    };
                    childNodes[node].setAttribute('PartName', partName.replace(regex, newFooter));
                }
            }
        }
    }

    xml.getElementsByTagName('Types')[0].childNodes = childNodes;
    let serializer = new XMLSerializer();
    xmlString = serializer.serializeToString(xml.documentElement);
    zip.file("[Content_Types].xml", xmlString);
};

const updateHeaderAndFooterRelations = function (zip, headersAndFooters, currentMaxRelId) {
    let xmlString = zip.file("word/_rels/document.xml.rels").asText();
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    let maxRelId = Math.max(currentMaxRelId, getMaxRelationID(null, xml));
    let childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            let type = childNodes[node].getAttribute('Type');
            if (type !== RelationshipTypeHeader && type !== RelationshipTypeFooter) {
                continue;
            }

            let target = childNodes[node].getAttribute('Target');
            const index = findInArray(headersAndFooters, 'oldTarget', target);
            if (index < 0) {
                continue;
            }

            headersAndFooters[index].oldRelId = childNodes[node].getAttribute('Id');
            headersAndFooters[index].newRelId = `rId${++maxRelId}`;

            childNodes[node].setAttribute('Target', headersAndFooters[index].newTarget);
            childNodes[node].setAttribute('Id', headersAndFooters[index].newRelId);
        }
    }

    xml.getElementsByTagName('Relationships')[0].childNodes = childNodes;
    const serializer = new XMLSerializer();
    zip.file("word/_rels/document.xml.rels", serializer.serializeToString(xml.documentElement));

    return maxRelId;
};

const updateHeaderAndFooterRelationsInDocument = function (zip, headersAndFooters) {
    let xmlString = zip.file("word/document.xml").asText();
    let xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    for (let i = 0; i < headersAndFooters.length; i++) {
        replaceHeaderFooterIdInXml(xml, headersAndFooters[i].oldRelId, headersAndFooters[i].newRelId);
    }

    const serializer = new XMLSerializer();
    zip.file("word/document.xml", serializer.serializeToString(xml.documentElement));
};

const findInArray = (arr, field, searchString) => {
    for (let i = 0; i < arr.length; i++) {
        if (arr[i][field] === searchString) {
            return i;
        }
    }

    return -1;
};

const replaceHeaderFooterIdInXml = (node, oldId, newId) => {
    if (!node.childNodes || node.childNodes.length && node.childNodes.length < 0) {
        return;
    }

    for (let i = 0; i < node.childNodes.length; i++) {
        c = node.childNodes[i];
        if ((c.tagName === 'w:headerReference' || c.tagName === 'w:footerReference') && c.getAttribute) {
            const id = c.getAttribute('r:id');
            if (id === oldId) {
                c.setAttribute('r:id', newId);
            }
        }
        if (c.childNodes && c.childNodes.length > 0) {
            replaceHeaderFooterIdInXml(c, oldId, newId);
        }
    }
};

const renameHeaderAndFooterFiles = (zip, headersAndFooters) => {
    let files = zip.folder("word").files;
    const relationsFolder = 'word/_rels';
    for (var file in files) {
        const index = findInArray(headersAndFooters, 'oldTarget', file.substring(5));
        if (index < 0) {
            continue;
        }

        headersAndFooters[index].location = `word/${headersAndFooters[index].newTarget}`;
        let fileContent = zip.file(file).asUint8Array();
        zip.file(headersAndFooters[index].location, fileContent);

        const relFile = zip.file(`${relationsFolder}/${headersAndFooters[index].oldTarget}.rels`);
        if (relFile) {
            fileContent = relFile.asUint8Array();
            headersAndFooters[index].relLocation = `${relationsFolder}/${headersAndFooters[index].newTarget}.rels`;
            zip.file(headersAndFooters[index].relLocation, fileContent);
        }
    }
};

const copyHeaderAndFooterFiles = (baseZip, _zipFiles, _headersAndFooters) => {
    for (var zipFileIndex in _headersAndFooters) {
        for (let i = 0; i < _headersAndFooters[zipFileIndex].length; i++) {
            const location = _headersAndFooters[zipFileIndex][i].location;
            let content = _zipFiles[zipFileIndex].file(location).asUint8Array();
            baseZip.file(location, content);

            if (_headersAndFooters[zipFileIndex][i].relLocation) {
                content = _zipFiles[zipFileIndex].file(_headersAndFooters[zipFileIndex][i].relLocation).asUint8Array();
                baseZip.file(_headersAndFooters[zipFileIndex][i].relLocation, content);
            }
        }
    }
};

module.exports = {
    ContentTypeHeader,
    ContentTypeFooter,
    prepareHeaderFooterRelations,
    copyHeaderAndFooterFiles
};