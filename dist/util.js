const { DOMParser } = require("xmldom");

const getMaxRelationID = (zip, xml) => {
    if (zip) {
        let xmlString = zip.file("word/_rels/document.xml.rels").asText();
        xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    }

    const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

    let maxId = 0;
    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var relID = childNodes[node].getAttribute('Id');
            const matches = relID.match(/(\d+)/);
            if (matches) {
                matches.forEach(match => {
                    maxId = Math.max(maxId, Number(match));
                });
            }
        }
    }

    return maxId;
};

module.exports = {
    getMaxRelationID
};