//. app.js
var express = require( 'express' ),
    bodyParser = require( 'body-parser' ),
    fs = require( 'fs' ),
    multer = require( 'multer' ),
    app = express();

var XLSX = require( 'xlsx' );
var Utils = XLSX.utils;

//. env values
require( 'dotenv' ).config();

app.use( express.static( __dirname + '/public' ) );
app.use( multer( { dest: './tmp/' } ).single( 'file' ) );
app.use( bodyParser.urlencoded( { extended: true } ) );
app.use( bodyParser.json() );
app.use( express.Router() );

function readExcelFile( filepath = '' ){
  var sheetsJsons = {};
  try{
    var book = XLSX.readFile( filepath );
  
    //. sheets = { Sheet1: {}, Sheet2: {}, .. }
    var sheets = book.Sheets;
    Object.keys( sheets ).forEach( function( sheetname ){
      var sheet = sheets[sheetname]
      var cells = [];

      var range = sheet["!ref"];
      var decodeRange = Utils.decode_range( range );  //. { s: { c:0, r:0 }, e: { c:5, r:4 } }

      //. シート内の全セル値を取り出し
      for( var r = decodeRange['s']['r']; r <= decodeRange['e']['r']; r ++ ){
        var row = [];
        for( var c = decodeRange['s']['c']; c <= decodeRange['e']['c']; c ++ ){
          var address = Utils.encode_cell( { r: r, c: c } );
          var cell = sheet[address];
          if( typeof cell !== "undefined" ){
            if( typeof cell.v != "undefined" ){
              row.push( cell.v );
            }else{
              row.push( '' );
            }
          }else{
            row.push( '' );
          }
        }
        cells.push( row );
      }

      sheetsJsons[sheetname] = cells;
    });
  }catch( e ){
    console.log( {e} );
  }finally{
    //fs.unlinkSync( filepath );
  }

  return sheetsJsons;
}

app.post( '/api/xlsx-textify', function( req, res ){
  res.contentType( 'application/json; charset=utf-8' );

  var filepath = req.file.path;
  //var filetype = req.file.mimetype;
  //var filesize = req.file.size;

  var jsons = readExcelFile( filepath );
  fs.unlinkSync( filepath );

  var results = {}
  Object.keys( jsons ).forEach( function( sheetname ){
    var cells = jsons[sheetname];

    if( cells.length > 1 ){
      var titles = cells[0];
      var rows = [];
      for( var i = 1; i < cells.length; i ++ ){
        var row = {};
        for( var j = 0; j < titles.length; j ++ ){
          var title = ( titles[j] ? titles[j] : '(' + j + ')' );
          row[title] = cells[i][j];
        }
        rows.push( row );
      }
      results[sheetname] = rows;
    }
  });

  res.write( JSON.stringify( { status: true, results: results }, null, 2 ) );
  res.end();
});

var port = process.env.PORT || 8080;
app.listen( port );
console.log( "server starting on " + port + " ..." );
