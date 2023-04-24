var async = require ( 'async' );
var officegen = require('../');

var fs = require('fs');
var path = require('path');

var docx = officegen ( {
  type: 'docx',
  orientation: 'landscape',
  pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 },
  pageSize: 'A4'
} );

// Remove this comment in case of debugging Officegen:
// officegen.setVerboseMode ( true );

docx.on ( 'error', function ( err ) {
      console.log ( err );
    });


pObj = docx.createP ();
pObj.addImage(path.resolve(__dirname, 'images_for_examples/image2.jpg'), {cx: 240, cy: 180, x: 800, y: 50});

var tableData = [
  [
    {
      val: 'Head1',
      opts: {cellColWidth: 3000, b: true, align: 'center', sz: 22, gridSpan: 2, paddingTop: '0', paddingBottom: '0', lineHeight: 240, shd: {fill: 'e8e8e8'}}
    },
    {
      val: [
        {
          type: "image",
          path: path.resolve(__dirname, 'images_for_examples/image1.png'),
          text: 'Example image column',
          opts: {
            anchoredImage: true,
            cx: 200,
            cy: 100,
            sz: 22,
            offsetV: 200000,
            offsetH: 400000,
            align: 'center',
            fontFamily: 'SimSun',
            lineHeight: 240,
            paddingTop: '0',
            paddingBottom: '0',
            b: true
          }
        },
      ],
      opts: {cellColWidth: 4500, shd: {fill: 'e8e8e8'}}
    }
  ],
  [
    {
      val: 'Cell 01',
      opts: {cellColWidth: 1500}
    },
    {
      val: 'Cell 02',
      opts: {cellColWidth: 1500, vAlign: 'bottom', paddingTop: '0', paddingBottom: '0', lineHeight: 240}
    }
  ],
  [
    {
      val: 'Cell 11',
      opts: {cellColWidth: 1500, vAlign: 'top', align: 'right', paddingTop: '0', paddingBottom: '0', lineHeight: 240}
    },
    {
      val: 'Cell 12',
      opts: {
        borders: {
          left: true,
          right: true
        },
      }
    }
  ],
  [
    {
      val: 'Cell 21'
    },
    {
      val: 'Cell 22',
      opts: {
        borders: {
          top: true
        },
      }
    }
  ],
  [
    {
      val: 'Cell 31',
      opts: {
        borders: {
          bottom: true
        },
      }
    },
    {
      val: 'Cell 32'
    }
  ]
];
docx.createTable(tableData);

pObj.startTextbox(450, 200, 300, 200, {
  lineHeight: 480
});
pObj.addText('This is a textbox', {bold: true, font_face: 'Arial', font_size: 10});
pObj.addLineBreak();
pObj.addText('You can add texts as usual', {font_face: 'Arial', font_size: 8});
pObj.addLineBreak();
pObj.addText(
`Lorem ipsum dolor sit amet, consetetur sadipscing elitr,
sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat,
sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum.
Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.
Lorem ipsum dolor sit amet, consetetur sadipscing elitr,
sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat,
sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum.
Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.`
);
pObj.endTextbox();

var out = fs.createWriteStream ( 'tmp/out.docx' );

out.on ( 'error', function ( err ) {
  console.log ( err );
});

async.parallel ([
  function ( done ) {
    out.on ( 'close', function () {
      console.log ( 'Finish to create a DOCX file.' );
      done ( null );
    });
    docx.generate ( out );
  }

], function ( err ) {
  if ( err ) {
    console.log ( 'error: ' + err );
  } // Endif.
});
