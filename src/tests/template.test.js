const DocxMerger = require("..");
const fs = require('fs');
const path = require('path');

test('doc 1', () => {
    var file1 = fs
    .readFileSync(path.resolve(__dirname, 'samples/Document_1.docx'), 'binary');

    var file2 = fs
    .readFileSync(path.resolve(__dirname, 'samples/Document_2.docx'), 'binary');

    // var file3 = fs
    // .readFileSync(path.resolve(__dirname, 'samples/Document_3.docx'), 'binary');

    var docx = new DocxMerger({pageBreak: false, mergeAsSections: true},[file1,file2]);

    const output = 'output_1.docx'
    docx.save('nodebuffer',function (data) {
        fs.writeFile(output, data, function(err){/*...*/});
    });

    expect(3).toBe(3);
});

test('doc 2', () => {
    var file1 = fs
    .readFileSync(path.resolve(__dirname, 'samples/Doc_1.docx'), 'binary');

    var file2 = fs
    .readFileSync(path.resolve(__dirname, 'samples/Doc_2.docx'), 'binary');

    var file3 = fs
    .readFileSync(path.resolve(__dirname, 'samples/Doc_3.docx'), 'binary');

    var docx = new DocxMerger({pageBreak: false, mergeAsSections: true},[file1,file2, file3]);

    const output = 'output_2.docx'
    docx.save('nodebuffer',function (data) {
        fs.writeFile(output, data, function(err){/*...*/});
    });

    expect(3).toBe(3);
});

test('doc 3', () => {
    var file1 = fs
    .readFileSync(path.resolve(__dirname, 'samples/1_cover.docx'), 'binary');

    var file2 = fs
    .readFileSync(path.resolve(__dirname, 'samples/2_landscape_view.docx'), 'binary');

    var file3 = fs
    .readFileSync(path.resolve(__dirname, 'samples/3_spot_data.docx'), 'binary');

    var docx = new DocxMerger({pageBreak: false, mergeAsSections: true},[file1,file2, file3]);

    const output = 'output_3.docx'
    docx.save('nodebuffer',function (data) {
        fs.writeFile(output, data, function(err){/*...*/});
    });

    expect(3).toBe(3);
});

test('doc 4', () => {
    var file1 = fs
    .readFileSync(path.resolve(__dirname, 'samples/part_1.docx'), 'binary');

    var file2 = fs
    .readFileSync(path.resolve(__dirname, 'samples/part_2.docx'), 'binary');

    var file3 = fs
    .readFileSync(path.resolve(__dirname, 'samples/part_3.docx'), 'binary');

    var docx = new DocxMerger({pageBreak: false, mergeAsSections: true},[file1,file2, file3]);

    const output = 'output_4.docx'
    docx.save('nodebuffer',function (data) {
        fs.writeFile(output, data, function(err){/*...*/});
    });

    expect(3).toBe(3);
});

test('doc 5', () => {
    var file1 = fs
    .readFileSync(path.resolve(__dirname, 'samples/part_1.docx'), 'binary');

    var file2 = fs
    .readFileSync(path.resolve(__dirname, 'samples/part_2.docx'), 'binary');

    // var file3 = fs
    // .readFileSync(path.resolve(__dirname, 'samples/part_3.docx'), 'binary');

    var docx = new DocxMerger({pageBreak: false, mergeAsSections: true},[file1,file2]);

    const output = 'output_5.docx'
    docx.save('nodebuffer',function (data) {
        fs.writeFile(output, data, function(err){/*...*/});
    });

    expect(3).toBe(3);
});
