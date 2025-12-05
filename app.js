//. app.js
var express = require( 'express' ),
    bodyParser = require( 'body-parser' ),
    //{ Configuration, OpenAIApi } = require( 'openai' ),
    OpenAI = require( 'openai' ),
    multer = require( 'multer' ),
    app = express();

var XLSX = require( 'xlsx' );
var Utils = XLSX.utils;
var fs = require( 'fs' );

var encoding = require( 'encoding-japanese' );

require( 'dotenv' ).config();

app.use( express.static( __dirname + '/public' ) );
app.use( multer( { dest: './tmp/' } ).single( 'file' ) );
app.use( bodyParser.urlencoded( { extended: true } ) );
app.use( bodyParser.json() );
app.use( express.Router() );

app.post( '/api/test', async function( req, res ){
  res.contentType( 'application/json; charset=utf8' );
  res.write( JSON.stringify( { status: true }, null, 2 ) );
  res.end();
});

/*
var settings_cors = 'CORS' in process.env ? process.env.CORS : '';  //. "http://localhost:8080,https://xxx.herokuapp.com"
app.all( '/*', function( req, res, next ){
  if( settings_cors ){
    var origin = req.headers.origin;
    if( origin ){
      var cors = settings_cors.split( " " ).join( "" ).split( "," );

      //. cors = [ "*" ] への対応が必要
      if( cors.indexOf( '*' ) > -1 ){
        res.setHeader( 'Access-Control-Allow-Origin', '*' );
        res.setHeader( 'Access-Control-Allow-Methods', '*' );
        res.setHeader( 'Access-Control-Allow-Headers', '*' );
        res.setHeader( 'Vary', 'Origin' );
      }else{
        if( cors.indexOf( origin ) > -1 ){
          res.setHeader( 'Access-Control-Allow-Origin', origin );
          res.setHeader( 'Access-Control-Allow-Methods', '*' );
          res.setHeader( 'Access-Control-Allow-Headers', '*' );
          res.setHeader( 'Vary', 'Origin' );
        }
      }
    }
  }
  next();
});
*/

app.get( '/api/ping', function( req, res ){
  res.contentType( 'application/json; charset=utf-8' );

  res.write( JSON.stringify( { status: true, message: 'PONG' }, null, 2 ) );
  res.end();
});

var settings_apikey = 'API_KEY' in process.env ? process.env.API_KEY : '';
///var settings_organization = 'ORGANIZATION' in process.env ? process.env.ORGANIZATION : '';
var openai = new OpenAI( { apiKey: settings_apikey } );

app.get( '/api/models', async function( req, res ){
  res.contentType( 'application/json; charset=utf-8' );

  var result = await openai.models.list();
  //console.log( {result.data} );
  res.write( JSON.stringify( { status: true, result: result.data }, null, 2 ) );
  res.end();
});

app.get( '/api/model/:id', async function( req, res ){
  res.contentType( 'application/json; charset=utf-8' );

  var id = req.params.id;
  var result = await openai.models.retrieve( id );
  res.write( JSON.stringify( { status: true, result: result.data }, null, 2 ) );
  res.end();
});

var DEFAULT_ROLE = 'DEFAULT_ROLE' in process.env && process.env.DEFAULT_ROLE ? process.env.DEFAULT_ROLE : 'user';
var DEFAULT_PROMPT = 'DEFAULT_PROMPT' in process.env && process.env.DEFAULT_PROMPT ? process.env.DEFAULT_PROMPT : '';
var DEFAULT_MODEL = 'DEFAULT_MODEL' in process.env && process.env.DEFAULT_MODEL ? process.env.DEFAULT_MODEL : 'gpt-4o-mini';

app.post( '/api/file', async function( req, res ){
  res.contentType( 'application/json; charset=utf8' );

  var role = req.body.role ? req.body.role : DEFAULT_ROLE;
  var prompt = req.body.prompt ? req.body.prompt : DEFAULT_PROMPT;
  var model = req.body.model ? req.body.model : DEFAULT_MODEL;
  var file_text = '';

  if( req.file && req.file.path ){
    var path = req.file.path;
    var filename = req.file.originalname;
    var mimetype = req.file.mimetype;
  
    var tmp = filename.split( '.' );
    var ext = tmp[tmp.length-1].toLowerCase();
    switch( ext ){
    case 'xls':
    case 'xlsx':
      var book = XLSX.readFile( path );
  
      //. sheets = { Sheet1: {}, Sheet2: {}, .. }
      var sheets = book.Sheets;
      Object.keys( sheets ).forEach( function( sheetname ){
        var sheet_text = '';
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

        if( cells && cells.length ){
          for( var r = 0; r < cells.length; r ++ ){
            for( var c = 0; c < cells[r].length; c ++ ){
              var cell = cells[r][c]
              var str_cell = '' + cell;

              sheet_text += ' ' + str_cell;
            }
            sheet_text += '\r\n';
          }

          file_text += sheet_text;
        }
      });

      break;
    case 'csv':
    case 'tsv':
      //. ファイル読み込み（Shift_JIS ファイルは変換する）
      var buffer = fs.readFileSync( path );
      var file_text = buffer.toString();
      var detect = encoding.detect( buffer );
      if( detect != 'UTF8' ){
        file_text = encoding.convert( buffer, { from: detect, to: 'UNICODE', type: 'string' } );
      }

      break;
    default:
    }

    fs.unlinkSync( path );
  }


  var text = prompt;
  if( file_text ){
    text += '\r\n\r\n' + file_text;
  }
  var option = {
    model: model,
    messages: [{
      role: role,
      content: [
        { type: 'text', text: text }
      ]
    }]
  };
  var completion = await progressingCompletion( option );

  //console.log( {completion} );
  if( completion && completion.status ){
    res.write( JSON.stringify( completion, null, 2 ) );
    res.end();
  }else{
    res.status( 400 );
    res.write( JSON.stringify( { status: false, error: 'no completion returned.' }, null, 2 ) );
    res.end();
  }
});

app.post( '/api/excel', async function( req, res ){
  res.contentType( 'application/json; charset=utf8' );
  if( req.file && req.file.path ){
    var path = req.file.path;
    var filename = req.file.originalname;
    var mimetype = req.file.mimetype;

    var role = req.body.role ? req.body.role : DEFAULT_ROLE;
    var prompt = req.body.prompt ? req.body.prompt : DEFAULT_PROMPT;
    var model = req.body.model ? req.body.model : DEFAULT_MODEL;
    var file_text = '';

    var book = XLSX.readFile( path );
  
    //. sheets = { Sheet1: {}, Sheet2: {}, .. }
    var sheets = book.Sheets;
    Object.keys( sheets ).forEach( function( sheetname ){
      var sheet_text = '';
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

      if( cells && cells.length ){
        for( var r = 0; r < cells.length; r ++ ){
          for( var c = 0; c < cells[r].length; c ++ ){
            var cell = cells[r][c]
            var str_cell = '' + cell;

            sheet_text += ' ' + str_cell;
          }
          sheet_text += '\r\n';
        }

        file_text += sheet_text;
      }
    });

    if( file_text ){
      var text = prompt + '\r\n\r\n' + file_text;
      var option = {
        model: model,
        messages: [{
          role: role,
          content: [
            { type: 'text', text: text }
          ]
        }]
      };
      var completion = await progressingCompletion( option );

      //console.log( {completion} );
      if( completion && completion.status ){
        res.write( JSON.stringify( completion, null, 2 ) );
        res.end();
      }else{
        res.status( 400 );
        res.write( JSON.stringify( { status: false, error: 'no completion returned.' }, null, 2 ) );
        res.end();
      }
    }

    fs.unlinkSync( path );
  }else{
    res.status( 400 );
    res.write( JSON.stringify( { status: false, error: 'no file uploaded.' }, null, 2 ) );
    res.end();
  }
});

app.post( '/api/csv', async function( req, res ){
  res.contentType( 'application/json; charset=utf8' );
  if( req.file && req.file.path ){
    var path = req.file.path;
    var filename = req.file.originalname;
    var mimetype = req.file.mimetype;

    var role = req.body.role ? req.body.role : DEFAULT_ROLE;
    var prompt = req.body.prompt ? req.body.prompt : DEFAULT_PROMPT;
    var model = req.body.model ? req.body.model : DEFAULT_MODEL;

    //. ファイル読み込み（Shift_JIS ファイルは変換する）
    var buffer = fs.readFileSync( path );
    var file_text = buffer.toString();
    var detect = encoding.detect( buffer );
    if( detect != 'UTF8' ){
      file_text = encoding.convert( buffer, { from: detect, to: 'UNICODE', type: 'string' } );
    }

    if( file_text ){
      var text = prompt + '\r\n\r\n' + file_text;
      var option = {
        model: model,
        messages: [{
          role: role,
          content: [
            { type: 'text', text: text }
          ]
        }]
      };
      var completion = await progressingCompletion( option );

      //console.log( {completion} );
      if( completion && completion.status ){
        res.write( JSON.stringify( completion, null, 2 ) );
        res.end();
      }else{
        res.status( 400 );
        res.write( JSON.stringify( { status: false, error: 'no completion returned.' }, null, 2 ) );
        res.end();
      }
    }

    fs.unlinkSync( path );
  }else{
    res.status( 400 );
    res.write( JSON.stringify( { status: false, error: 'no file uploaded.' }, null, 2 ) );
    res.end();
  }
});


const wait = ( ms ) => new Promise( ( res ) => setTimeout( res, ms ) );

const progressingCompletion = async ( option ) => {
	await wait( 10 );
  try{
    var result = await openai.chat.completions.create( option );
  	return {
      status: true,
		  result: result
  	};
  }catch( e ){
  	return {
      status: false,
		  result: e
  	};
  }
}

const progressingImage = async ( option ) => {
	await wait( 10 );
  try{
    var result = await openai.createImage( option );
  	return {
      status: true,
		  result: result
  	};
  }catch( e ){
  	return {
      status: false,
		  result: e
  	};
  }
}

const progressingChatCompletion = async ( option ) => {
	await wait( 10 );
  try{
    var result = await openai.chat.completions.create( option );
  	return {
      status: true,
		  result: result
  	};
  }catch( e ){
  	return {
      status: false,
		  result: e
  	};
  }
}

const callWithProgress = async ( fn, option, maxdepth = 7, depth = 0 ) => {
	const result = await fn( option );

	// check completion
	if( result.status ){
		// finished
		return result.result;
	}else{
		if( depth > maxdepth ){
			throw result;
		}
		await wait( Math.pow( 2, depth ) * 10 );
	
		return callWithProgress( fn, option, maxdepth, depth + 1 );
	}
}


var port = process.env.PORT || 8080;
app.listen( port );
console.log( "server starting on " + port + " ..." );

module.exports = app;
