
var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;
const { getMaxRelationID } = require('./util');

var prepareMediaFiles = function (files, media) {

    var count = 1;

    files.forEach(function (zip, index) {
        // var zip = new JSZip(file);
        var medFiles = zip.folder("word/media").files;

        for (var mfile in medFiles) {
            if (/^word\/media/.test(mfile) && mfile.length > 11) {
                // console.log(mfile);
                media[count] = {};
                media[count].oldTarget = mfile;
                media[count].newTarget = mfile.replace(/[0-9]/, '_' + count).replace('word/', "");
                media[count].fileIndex = index;
                updateMediaRelations(zip, count, media);
                updateMediaContent(zip, count, media);
                count++;
            }
        }
    });

    // console.log(JSON.stringify(media));

    // this.updateRelation(files);
};

const prepareMediaFiles2 = function (files, maxRelId, _media) {
    // let maxRelId = getMaxRelationID(files[0], null);
    for (let zipIdx = 1; zipIdx < files.length; zipIdx++) {
        let zipMedia = _media[zipIdx];
        if (!zipMedia) {
            zipMedia = {};
            _media[zipIdx] = zipMedia;
        }

        maxRelId = updateMediaRelations2(files[zipIdx], zipIdx, maxRelId, zipMedia);
        updateMediaContent2(files[zipIdx], zipMedia);
    }

    return maxRelId;
};

var updateMediaRelations = function (zip, count, _media) {

    var xmlString = zip.file("word/_rels/document.xml.rels").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    var serializer = new XMLSerializer();

    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var target = childNodes[node].getAttribute('Target');
            if ('word/' + target == _media[count].oldTarget) {

                _media[count].oldRelID = childNodes[node].getAttribute('Id');

                childNodes[node].setAttribute('Target', _media[count].newTarget);
                childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
            }
        }
    }

    // console.log(serializer.serializeToString(xml.documentElement));

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file("word/_rels/document.xml.rels", xmlString);

    // console.log( xmlString );
};

var updateMediaRelations2 = function (zip, zipFileIdx, maxRelId, _media) {
    const relFiles = zip.folder('word/_rels').files;
    const serializer = new XMLSerializer();

    for (let file in relFiles) {
        if (/^word\/_rels/.test(file) && !relFiles[file].dir) {
            // const file = relFiles[fileIndex];
            const xmlString = zip.file(file).asText();
            const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

            const shouldUpdateRels = file === 'word/_rels/document.xml.rels'; // relation-id will only be updated in document.xml. The reason is this is the only file that will be merged. Others will be copied

            if (shouldUpdateRels) {
                maxRelId = Math.max(maxRelId, getMaxRelationID(null, xml));
            }

            const relationships = xml.getElementsByTagName('Relationships');
            if (relationships && relationships.length > 0) {
                const childNodes = relationships[0].childNodes;
                for (let node in childNodes) {
                    if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                        let target = childNodes[node].getAttribute('Target');
                        if (target.startsWith('media/')) {
                            if (_media[target]) {
                                if (!_media[target].newTarget) {
                                    _media[target].newTarget = target.split('.').join(`_${zipFileIdx}.`);
                                }
                            } else {
                                _media[target] = {
                                    oldTarget: target,
                                    newTarget: target.split('.').join(`_${zipFileIdx}.`),
                                    fileIndex: zipFileIdx
                                };
                            }

                            // relation-id will only be updated in document.xml
                            if (shouldUpdateRels) {
                                _media[target].oldRelID = childNodes[node].getAttribute('Id');
                                _media[target].newRelID = `rId${++maxRelId}`;
                                childNodes[node].setAttribute('Id', _media[target].newRelID);
                            }
                            childNodes[node].setAttribute('Target', _media[target].newTarget);
                        }
                    }
                }

                zip.file(file, serializer.serializeToString(xml));
            }
        }
    }

    return maxRelId;
};

const updateMediaContent2 = function (zip, _media) {
    let xmlString = zip.file('word/document.xml').asText();

    for (let mediaTarget in _media) {
        if (_media[mediaTarget].newRelID) {
            xmlString = xmlString.replace(new RegExp(_media[mediaTarget].oldRelID + '"', 'g'), _media[mediaTarget].newRelID + '"');
        }
    }

    zip.file("word/document.xml", xmlString);
};

var updateMediaContent = function (zip, count, _media) {

    var xmlString = zip.file("word/document.xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');

    zip.file("word/document.xml", xmlString);
};

var copyMediaFiles = function (base, _media, _files) {

    for (var media in _media) {
        var content = _files[_media[media].fileIndex].file(_media[media].oldTarget).asUint8Array();

        base.file('word/' + _media[media].newTarget, content);
    }
};

const copyMediaFiles2 = function (base, _media, _files) {
    for (let fileIndex in _media) {
        const mediaFiles = _media[fileIndex];
        for (let medFile in mediaFiles) {
            const content = _files[fileIndex].file(`word/${mediaFiles[medFile].oldTarget}`).asUint8Array();
            base.file('word/' + mediaFiles[medFile].newTarget, content);
        }
    }
};

module.exports = {
    prepareMediaFiles: prepareMediaFiles,
    updateMediaRelations: updateMediaRelations,
    updateMediaContent: updateMediaContent,
    copyMediaFiles: copyMediaFiles,
    prepareMediaFiles2,
    copyMediaFiles2
};