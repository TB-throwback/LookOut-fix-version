/*
 * File: tnef.js
 *   Extract files and metadata TNEF/MAPI format
 *
 * Copyright (C) 2007 Aron Rubin <arubin@atl.lmco.com>
 *
 *
 * About:
 *   scans tnef file and extracts all attachments and decodable metadata.
 *   Attachments are written to their original file-names if possible.
 *   Calendaring and contact data is transcoded to standards formats.
 */



const LVL_MESSAGE    = 0x1;
const LVL_ATTACHMENT = 0x2;

// ---------- Tnef Types ----------

const TNEF_TRIPLES = 0x0000; //   triples
const TNEF_STRING  = 0x0001; //   string
const TNEF_TEXT    = 0x0002; //   text
const TNEF_DATE    = 0x0003; //   date
const TNEF_SHORT   = 0x0004; //   short
const TNEF_LONG    = 0x0005; //   long
const TNEF_BYTE    = 0x0006; //   byte
const TNEF_WORD    = 0x0007; //   word
const TNEF_DWORD   = 0x0008; //   dword
const TNEF_TYPE_MAX     = 0x0009; //   max

function tnef_attr_type_to_string( attr_type ) {
  switch( attr_type ) {
  case TNEF_TRIPLES: return "TRIPLES";
  case TNEF_STRING: return "STRING";
  case TNEF_TEXT: return "TEXT";
  case TNEF_DATE: return "DATE";
  case TNEF_SHORT: return "SHORT";
  case TNEF_LONG: return "LONG";
  case TNEF_BYTE: return "BYTE";
  case TNEF_WORD: return "WORD";
  case TNEF_DWORD: return "DWORD";
  default: return( "0x" + to_hex( attr_type, 2 ) );
  }
}

// ---------- Tnef Names -----------

// names of all attributes found in TNEF file
const TNEF_ATTR_OWNER				= 0x0000; //  Owner
const TNEF_ATTR_SENTFOR				= 0x0001; //  Sent For
const TNEF_ATTR_DELEGATE			= 0x0002; //  Delegate
const TNEF_ATTR_DATE_START			= 0x0006; //  Date Start
const TNEF_ATTR_DATE_END			= 0x0007; //  Date End
const TNEF_ATTR_APPT_ID_OWNER			= 0x0008; //  Owner Appointment ID
const TNEF_ATTR_REQUEST_RESP			= 0x0009; //  Response Requested.
const TNEF_ATTR_FROM				= 0x8000; //  From
const TNEF_ATTR_SUBJECT				= 0x8004; //  Subject
const TNEF_ATTR_DATE_SENT			= 0x8005; //  Date Sent
const TNEF_ATTR_DATE_RECD			= 0x8006; //  Date Received
const TNEF_ATTR_MESSAGE_STATUS			= 0x8007; //  Message Status
const TNEF_ATTR_MESSAGE_CLASS			= 0x8008; //  Message Class
const TNEF_ATTR_MESSAGE_ID			= 0x8009; //  Message ID
const TNEF_ATTR_PARENT_ID			= 0x800a; //  Parent ID
const TNEF_ATTR_CONVERSATION_ID			= 0x800b; //  Conversation ID
const TNEF_ATTR_BODY				= 0x800c; //  Body
const TNEF_ATTR_PRIORITY			= 0x800d; //  Priority
const TNEF_ATTR_ATTACH_DATA			= 0x800f; //  Attachment Data
const TNEF_ATTR_ATTACH_TITLE			= 0x8010; //  Attachment File Name
const TNEF_ATTR_ATTACH_METAFILE			= 0x8011; //  Attachment Meta File
const TNEF_ATTR_ATTACH_CREATE_DATE		= 0x8012; //  Attachment Creation Date
const TNEF_ATTR_ATTACH_MODIFY_DATE		= 0x8013; //  Attachment Modification Date
const TNEF_ATTR_DATE_MODIFY			= 0x8020; //  Date Modified
const TNEF_ATTR_ATTACH_TRANSPORT_FILENAME	= 0x9001; //  Attachment Transport Filename
const TNEF_ATTR_ATTACH_REND_DATA		= 0x9002; //  Attachment Rendering Data
const TNEF_ATTR_MAPI_PROPS			= 0x9003; //  MAPI Properties
const TNEF_ATTR_RECIPTABLE			= 0x9004; //  Recipients
const TNEF_ATTR_ATTACHMENT			= 0x9005; //  Attachment
const TNEF_ATTR_TNEF_VERSION			= 0x9006; //  TNEF Version
const TNEF_ATTR_OEM_CODEPAGE			= 0x9007; //  OEM Codepage
const TNEF_ATTR_ORIGNINAL_MESSAGE_CLASS		= 0x9008; //  Original Message Class

function tnef_attr_name_to_string( attr_name ) {
  switch( attr_name ) {
  case TNEF_ATTR_OWNER: return "OWNER";
  case TNEF_ATTR_SENTFOR: return "SENTFOR";
  case TNEF_ATTR_DELEGATE: return "DELEGATE";
  case TNEF_ATTR_DATE_START: return "DATE_START";
  case TNEF_ATTR_DATE_END: return "DATE_END";
  case TNEF_ATTR_APPT_ID_OWNER: return "APPT_ID_OWNER";
  case TNEF_ATTR_REQUEST_RESP: return "REQUEST_RESP";
  case TNEF_ATTR_FROM: return "FROM";
  case TNEF_ATTR_SUBJECT: return "SUBJECT";
  case TNEF_ATTR_DATE_SENT: return "DATE_SENT";
  case TNEF_ATTR_DATE_RECD: return "DATE_RECD";
  case TNEF_ATTR_MESSAGE_STATUS: return "MESSAGE_STATUS";
  case TNEF_ATTR_MESSAGE_CLASS: return "MESSAGE_CLASS";
  case TNEF_ATTR_MESSAGE_ID: return "MESSAGE_ID";
  case TNEF_ATTR_PARENT_ID: return "PARENT_ID";
  case TNEF_ATTR_CONVERSATION_ID: return "CONVERSATION_ID";
  case TNEF_ATTR_BODY: return "BODY";
  case TNEF_ATTR_PRIORITY: return "PRIORITY";
  case TNEF_ATTR_ATTACH_DATA: return "ATTACH_DATA";
  case TNEF_ATTR_ATTACH_TITLE: return "ATTACH_TITLE";
  case TNEF_ATTR_ATTACH_METAFILE: return "ATTACH_METAFILE";
  case TNEF_ATTR_ATTACH_CREATE_DATE: return "ATTACH_CREATE_DATE";
  case TNEF_ATTR_ATTACH_MODIFY_DATE: return "ATTACH_MODIFY_DATE";
  case TNEF_ATTR_DATE_MODIFY: return "DATE_MODIFY";
  case TNEF_ATTR_ATTACH_TRANSPORT_FILENAME: return "ATTACH_TRANSPORT_FILENAME";
  case TNEF_ATTR_ATTACH_REND_DATA: return "ATTACH_REND_DATA";
  case TNEF_ATTR_MAPI_PROPS: return "MAPI_PROPS";
  case TNEF_ATTR_RECIPTABLE: return "RECIPTABLE";
  case TNEF_ATTR_ATTACHMENT: return "ATTACHMENT";
  case TNEF_ATTR_TNEF_VERSION: return "TNEF_VERSION";
  case TNEF_ATTR_OEM_CODEPAGE: return "OEM_CODEPAGE";
  case TNEF_ATTR_ORIGNINAL_MESSAGE_CLASS: return "ORIGNINAL_MESSAGE_CLASS";

  default: return( "0x" + to_hex( attr_name, 2 ) );
  }
}


// ---------- Util ----------

function assert( condition ) {
  var caller_arguments = new Array();
  // Function.arguments is deprecated, but convenient if it works
  if( arguments.length < 2 && assert.caller.arguments )
    caller_arguments = assert.caller.arguments;

  if( !condition )
    tnef_log_msg( "Assertion failed in " + assert.caller.name + "( )", 5 );
}

function tnef_log_msg( msg, level ) {
  if( (level == null ? 9 : level) <= debugLevel ) {
    var cs = Components.classes["@mozilla.org/consoleservice;1"].getService( Components.interfaces.nsIConsoleService );
    cs.logStringMessage( msg );
  }
}

const TNEF_PREF_PREFIX = "extensions.lookout.";

function tnef_get_pref( name, get_type_func, default_val ) {
  var pref_name = TNEF_PREF_PREFIX + name;
  var pref_val;
  try {
    var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);
    pref_val = prefs[get_type_func]( pref_name );
  } catch (ex) {
    tnef_log_msg( "TNEF: warning: could not retrieve setting '" + pref_name + "': " + ex, 5 );
  }
  if( pref_val === void(0) )
    pref_val = default_val;

  return pref_val;
}
function tnef_get_bool_pref( name, default_val ) {
  return tnef_get_pref( name, "getBoolPref", default_val );
}
function tnef_get_string_pref( name, default_val ) {
  return tnef_get_pref( name, "getCharPref", default_val );
}
function tnef_get_int_pref( name, default_val ) {
  return tnef_get_pref( name, "getIntPref", default_val );
}


// these routines (GETXX) deal with both endian and aggregation

function GETINT64( p, i ) {
  if( arguments.length < 2 )
    i = 0;

  var val = 0;

  if( typeof(p) == "string" ) {
    for( var j = 7; j >= 0; j-- )
      val = val*256 + p.charCodeAt( i + j );
  } else {
    for( var j = 7; j >= 0; j-- )
      val = val*256 + p[i + j];
  }
  return( val );
}

function GETINT32( p, i ) {
  if( arguments.length < 2 )
    i = 0;
  if( typeof(p) == "string" )
    return( ((p.charCodeAt(i + 3)*256 + p.charCodeAt(i + 2))*256 + p.charCodeAt(i + 1))*256 + p.charCodeAt(i) );
  else
    return( ((p[i + 3]*256 + p[i + 2])*256 + p[i + 1])*256 + p[i] );
}

function GETINT16( p, i ) {
  if( arguments.length < 2 )
    i = 0;
  if( typeof(p) == "string" )
    return( p.charCodeAt(i) + p.charCodeAt(i + 1)*256 );
  else
    return( p[i] + p[i + 1]*256 );
}

function GETUTF16( p, i, num_chars ) {
  var utf16_str = "";

  if( arguments.length < 2 )
    i = 0;
  if( arguments.length < 3 || num_chars < 1 )
    num_chars = 1;
  if( typeof(p) == "string" ) {
    while( num_chars-- ) {
      utf16_str += String.fromCharCode( p.charCodeAt(i) + p.charCodeAt(i + 1)*256 );
      i += 2;
    }
  } else {
    while( num_chars-- ) {
      utf16_str += String.fromCharCode( p[i] + p[i + 1]*256 );
      i += 2;
    }
  }

  return( utf16_str );
}


function ntoa_pad( x, len, pad ) {
  var str = "" + x;

  if( arguments.length < 3 || pad.length < 1 )
    pad = "0";
  if( pad.length > 1 )
    pad = pad[0];

  while( str.length < len )
    str = pad + str;
  return( str );
}


function to_hex( x, len ) {
  var range = "0123456789abcdef";
  var tmp = Math.floor( x );
  var hex = "";

  len = len*2;

  while( (!len && tmp) || len ) {
    hex = range[tmp&0x0f] + hex;
    tmp = Math.floor( tmp/16 );
    if( len )
      len--;
  }

  return( hex );
}


function GETFLT32( p, i ) {
  // struct IEEE754FloatMPN { uint mantissa : 23; uint biased_exponent : 8; guint sign : 1; };
  if( arguments.length < 2 )
    i = 0;

  var bin = typeof(p) == "string" ? [ p.charCodeAt(i), p.charCodeAt(i + 1),
				      p.charCodeAt(i + 2), p.charCodeAt(i + 3) ] : p.slice( i, i + 4 );
  var precision_bits = 23, bias = 127; // IEEE754 Float
  var signal = 0x80 & bin[3], exponent = ((0xEF & bin[3]) << 1) + ((0x80 & bin[2]) >> 7);
  var significand = 0, divisor = 0.5, byte_idx = 2, start_bit, mask;

  do {
    start_bit = precision_bits % 8 || 8;
    mask = 1 << start_bit;
    while( mask >>= 1 ) {
      if( (bin[byte_idx] & mask) != 0 )
	significand += divisor;
      divisor *= 0.5;
    }
    precision_bits -= start_bit;
    byte_idx--;
  } while( precision_bits );

  if( exponent == 255 ) // IEEE754 Float ((bias << 1) + 1)
    return( significand ? NaN : signal ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY );
  else if( exponent && significand )
    return( (1 + signal * -2) * Math.pow( 2, exponent - 127 ) * (1 + significand) );
  else if( !exponent && significand )
    return( (1 + signal * -2) * Math.pow( 2, -126 ) * significand );
  else
    return( 0 );
}


function GETDBL64( p, i ) {
  // strut IEEE754DoubleMPN { uint mantissa_low : 32; uint mantissa_high : 20; uint biased_exponent : 11; uint sign : 1; };
  if( arguments.length < 2 )
    i = 0;

  var bin = typeof(p) == "string" ? [ p.charCodeAt(i), p.charCodeAt(i + 1),
				      p.charCodeAt(i + 2), p.charCodeAt(i + 3),
				      p.charCodeAt(i + 4), p.charCodeAt(i + 5),
				      p.charCodeAt(i + 6), p.charCodeAt(i + 7) ] : p.slice( i, i + 8 );
  var precision_bits = 52, bias = 1023; // IEEE754 Double
  var signal = 0x80 & bin[7], exponent = ((0xEF & bin[7]) << 4) + ((0xF0 & bin[6]) >> 4);
  var significand = 0, divisor = 0.5, byte_idx = 6, start_bit, mask;

  do {
    start_bit = precision_bits % 8 || 8;
    mask = 1 << start_bit;
    while( mask >>= 1 ) {
      if( (bin[byte_idx] & mask) != 0 )
	significand += divisor;
      divisor *= 0.5;
    }
    precision_bits -= start_bit;
    byte_idx--;
  } while( precision_bits );

  if( exponent == 2047 ) // IEEE754 Double ((bias << 1) + 1)
    return( significand ? NaN : signal ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY );
  else if( exponent && significand )
    return( (1 + signal * -2) * Math.pow( 2, exponent - 1023 ) * (1 + significand) );
  else if( !exponent && significand )
    return( (1 + signal * -2) * Math.pow( 2, -1022 ) * significand );
  else
    return( 0 );
}


function unicode_to_utf8( buf ) {
  var converter = Components.classes["@mozilla.org/intl/scriptableunicodeconverter"]
                    .getService(Components.interfaces.nsIScriptableUnicodeConverter);

  converter.charset = "UTF-16";
  converter.ConvertFromUnicode( buf );
  return( converter.ConvertFromUnicode( buf ) );
}

function tnef_base64_encode( str ) {
  var data="", chr1, chr2, chr3, enc1, enc2, enc3, enc4, i = 0;
  const key_str="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";

  do {
    chr1 = str.charCodeAt( i++ );
    chr2 = str.charCodeAt( i++ );
    chr3 = str.charCodeAt( i++ );
    enc1 = chr1 >> 2;
    enc2 = ((chr1&3)<<4) | (chr2>>4);
    enc3 = ((chr2&15)<<2) | (chr3>>6);
    enc4 = chr3&63;
    if( isNaN( chr2 ) ) {
      enc3 = enc4 = 64;
    } else if( isNaN( chr3 ) ) {
      enc4 = 64;
    }
    data += key_str.charAt( enc1 ) + key_str.charAt( enc2 ) + key_str.charAt( enc3 ) + key_str.charAt( enc4 );
  } while( i < str.length );

  return data;
}


// ---------- Path ----------

function concat_fname( fname1, fname2 ) {
  var filename = null;

  assert( fname1 || fname2 );

  if( !(fname1 || fname2) )
    return( null );

  if( !fname1 ) {
    filename = fname2;
  } else {
    filename = fname1;

    if( fname2 ) {
      if( filename[filename.length - 1] != '/' && fname2[0] != '/' )
	filename += "/";
    }
    filename += fname2;
  }

  return( filename );
}

function fs_file_exists( fname ) {
  var lf = Components.classes["@mozilla.org/file/local;1"].createInstance(Components.interfaces.nsIFile);
  lf.initWithPath( fname );
  return( lf.exists() );
}

// finds a filename fname.N where N >= 1 and is not the name of an existing
// filename.  Assumes that fname does not already have such an extension
function fs_find_free_number( fname ) {
  var newname = "";
  var counter = 1;

  do {
    newname = fname + counter;
    counter++;
  } while( file_exists( newname ) );

  return( newname );
}

function fs_munge_fname( directory, fname ) {
  var file = null;

  // If we were not given a filename make one up
  if( !fname || fname.length == 0 ) {
    var tmp = concat_fname( directory, "tnef-part" );
    //debug_print( "No file name specified, using default.\n" );
    file = find_free_number( tmp );
  } else {
    file = concat_fname( directory, fname );
  }

  return( file );
}


// ---------- File Record ----------

function TnefFile( pkg ) {
  this.pkg = pkg;
}
TnefFile.prototype = {
  pkg: null,
  name: null,
  len: 0,
  data: null,
  date: Date(),
  mime_type: null,
  position: 0
}

function tnef_file_name_used( fname, files ) {
  var i = 0;

  if( !files )
    return( 0 );

  for( i = 0; i < files.length; i++ )
    if( files[i].name == fname )
      return( 1 );
  return( 0 );
}

// Converts OEM Code Page number to charset
// - Windows Code Pages: http://msdn.microsoft.com/en-us/goglobal/bb964654.aspx
// - Gecko Supported Charsets: https://developer.mozilla.org/en/Character_Sets_Supported_by_Gecko
function tnef_codepage_to_charset( cp ) {
  switch( cp ) {
    case 874:  return "TIS-620";      // Thai
    case 932:  return "SHIFT-JIS";    // Japanese
    case 936:  return "GBK";          // Simplified Chinese
    case 949:  return "EUC-KR";       // Korean
    case 950:  return "BIG5";         // Traditional Chinese
    case 1250: return "WINDOWS-1250"; // Latin
    case 1251: return "WINDOWS-1251"; // Cyrillic
    case 1252: return "WINDOWS-1252"; // Western
    case 1253: return "WINDOWS-1253"; // Greek
    case 1254: return "WINDOWS-1254"; // Turkish
    case 1255: return "WINDOWS-1255"; // Hebrew
    case 1256: return "WINDOWS-1256"; // Arabic
    case 1257: return "WINDOWS-1257"; // Baltic
    case 1258: return "WINDOWS-1258"; // Vietnam
    default:   return null;
  }
}

function tnef_file_munge_fname( fname, files, code_page ) {
  var file = null;
  var count = 0;

  // If we were not given a filename make one up
  if( !fname || fname.length == 0 ) {
    //debug_print( "No file name specified, using default.\n" );
    fname = "tnef-part-";
    do {
      file = fname + count;
      count++;
    } while( tnef_file_name_used( file, files ) );
  } else if( tnef_get_bool_pref( "disable_filename_character_set" ) ) {
    file = fname;
  } else {
    charset = tnef_codepage_to_charset( code_page );
    tnef_log_msg( "Lookout: convert file name from charset: " + charset, 7 );
    if( charset != null ) {
      try {
        var converter = Components.classes['@mozilla.org/intl/scriptableunicodeconverter'].getService(Components.interfaces.nsIScriptableUnicodeConverter);
	converter.charset = charset;
	var fname2 = converter.ConvertToUnicode( fname );
	try {
		decodeURIComponent( escape( fname2 ) );
	} catch (e) {
		fname = fname2;
	}
      } catch( e ) {
        tnef_log_msg( "Lookout: failed to convert file name from charset: " + charset, 4 );
      }
    }
    file = fname;
    while( tnef_file_name_used( file, files ) ) {
      file = fname + count;
      count++;
    }
  }

  return( file );
}


function tnef_file_write_stream( file, outstrm ) {
  outstrm.write( file.data, file.len );
}


function tnef_file_notify( file, listener, is_final ) {
  tnef_log_msg( "TNEF: Entering tnef_file_notify()", 6); //MKA
  tnef_log_msg( "TNEF: Notifying listener of " + file.name + ", pos = " + file.position +
	     ", has data = " + (file.data ? "true" : "false") +
	     ", is final = " + is_final, 7 );

  if( !listener )
    return;
  if( file.data || is_final ) {
    if( file.position == 0 && listener.onTnefStart )
      listener.onTnefStart( file.name, file.mime_type, file.len, file.date );
    if( file.data ) {
      if( listener.onTnefData )
	listener.onTnefData( file.position, file.data );
      file.position += file.data.length;
    }
    if( is_final && listener.onTnefEnd )
      listener.onTnefEnd( );
  }
}


function tnef_file_add_mapi_attrs( file, files, pkg, attrs ) {
  var i = 0;

  for( i = 0; attrs[i]; i++ ) {
    if( attrs[i].num_values ) {
      switch( attrs[i].name ) {
      case MAPI_ATTACH_LONG_FILENAME:
	file.name = tnef_file_munge_fname( attrs[i].values[0], files, pkg.code_page );
	break;

      case MAPI_ATTACH_DATA_OBJ:
	file.len = attrs[i].values[0].length;
	file.data = attrs[i].values[0];
	break;

      case MAPI_ATTACH_MIME_TAG:
	file.mime_type = attrs[i].values[0];
	break;

      default:
	break;
      }
    }
  }
}

function tnef_file_add_attr( file, files, attr, pkg, listener ) {
  assert( file && attr );
  if( !(file && attr) )
    return;

  // we only care about some things... we will skip most attributes
  switch( attr.name ) {
  case TNEF_ATTR_ATTACH_MODIFY_DATE:
    file.dt = tnef_attr_parse_date( attr );
    break;

  case TNEF_ATTR_ATTACHMENT:
    var mapi_attrs = mapi_attr_read( attr.buf, attr.len );
    if( mapi_attrs ) {
      if( tnef_get_bool_pref( "attach_raw_mapi" ) )
	tnef_pack_handle_mapi_attrs( pkg, mapi_attrs, listener );

      tnef_file_add_mapi_attrs( file, files, pkg, mapi_attrs );
      mapi_attrs = null;
    }
    break;

  case TNEF_ATTR_ATTACH_TITLE:
    file.name = tnef_file_munge_fname( attr.buf, files, pkg.code_page );
    break;

  case TNEF_ATTR_ATTACH_DATA:
    file.len = attr.len;
    file.data = attr.buf;
    break;

  default:
    break;
  }
}

function tnef_file_clear( file ) {
  if( file.data )
    delete file.data;
  file.data = null;
  file.len = 0;
  file.position = 0;
}

// ---------- Tnef Date ----------

function tnef_date_parse( buf ) {
  // yr, mo, da, hr, mi, se
  return( new Date( GETINT16( buf, 0 ), GETINT16( buf, 2 ), GETINT16( buf, 4 ),
		    GETINT16( buf, 6 ), GETINT16( buf, 8 ), GETINT16( buf, 10 ) ) );
}


// ---------- Tnef Attr ----------

const MINIMUM_ATTR_LENGTH = 14; // 72
var tnef_byte_pos = 0;

// Object types
function TRP() {
}
TRP.prototype = {
  id: 0,
  chbgtrp: 0,
  cch: 0,
  cb: 0
}

function TRIPLE() {
}
TRIPLE.prototype = {
  trp: TRP(),
  sender_display_name: null,
  sender_address: null
}


// Attr -- storing a structure, formated according to file specification
function TnefAttr() {
}
TnefAttr.prototype = {
  lvl_type: 0,
  type: 0,
  name: 0,
  len: 0,
  buf: null
}


// Copy the date data from the attribute into a struct date
function tnef_attr_parse_date( attr ) {
  assert( attr );
  assert( attr.type == TNEF_DATE );
  assert( attr.len >= 14 );

  return( tnef_date_parse( attr.buf ) );
}


function tnef_attr_parse_triple( attr, triple ) {
  assert( attr );
  assert( triple );
  assert( attr.type == TNEF_TRIPLES );

  triple.trp.id = GETINT16( attr.buf, 0 );
  triple.trp.chbgtrp = GETINT16( attr.buf, 2 );
  triple.trp.cch = GETINT16( attr.buf, 4 );
  triple.trp.cb = GETINT16( attr.buf, 6 );
  triple.sender_display_name = attr.buf[8];
  triple.sender_address = attr.buf[8] + triple.trp.cch;
}


// print attr to stderr.  Assumes that the Debug flag has been set and
// already checked
function tnef_attr_dbg_dump( attr ) {
  var name = tnef_attr_name_to_string( attr.name );
  var type = tnef_attr_type_to_string( attr.type );

  var msg = "(" + ((attr.lvl_type == LVL_MESSAGE) ? "MESS" : "ATTA") + ")" +
              name + "[type: " + type + "] [len: " + attr.len + "] =";

  switch( attr.type ) {
  case TNEF_BYTE:
    for( i = 0; i < attr.len; i++ )
      msg += " " + attr.buf.charCodeAt(i);
    break;

  case TNEF_SHORT:
    if( attr.len < 2 ) {
      msg += "Not enough data for TNEF_SHORT";
      tnef_log_msg( msg, 5 );
      return;
    }
    msg += " " + GETINT16( attr.buf );
    if( attr.len > 2 ) {
      msg += " [extra data:";
      for( i = 2; i < attr.len; i++ )
	msg += attr.buf.charCodeAt(i);
      msg += " ]";
    }
    break;

  case TNEF_LONG:
    if( attr.len < 4 ) {
      msg += "Not enough data for TNEF_LONG";
      tnef_log_msg( msg, 5 );
      return;
    }
    msg += " " + GETINT32( attr.buf );
    if( attr.len > 4 ) {
      msg += " [extra data:";
      for( var i = 4; i < attr.len; i++ )
	msg += " " + attr.buf.charCodeAt(i);
      msg += " ]";
    }
    break;

  case TNEF_WORD:
    for( var i = 0; i < attr.len; i += 2 )
      msg += " " + GETINT16( attr.buf, i );
    break;

  case TNEF_DWORD:
    for( var i = 0; i < attr.len; i += 4 )
      msg += " " + GETINT32( attr.buf, i );
    break;

  case TNEF_DATE:
    var dt = tnef_attr_parse_date( attr );
    msg += dt.toString();
    dt = null;
    break;

  case TNEF_TEXT:
  case TNEF_STRING:
    msg += attr.buf.substr( 0, attr.len - 1 );
    break;

  case TNEF_TRIPLES:
    var triple = tnef_attr_parse_triple( attr );
    msg += triple.toString();
    break;

  default:
    msg += "<unknown type>";
    break;
  }
  tnef_log_msg( msg + "\n", 7 );
}

function tnef_attr_clear( attr ) {
  if( attr && attr.buf ) {
    delete attr.buf;
  }
}


// Validate the checksum against attr.  The checksum is the sum of all the
//   bytes in the attribute data modulo 65536
function tnef_attr_check_checksum( attr, checksum ) {
  var i, sum = 0;

  for( i = 0; i < attr.len; i++ )
    sum += attr.buf.charCodeAt(i);

  sum &= 0xFFFF;

  tnef_log_msg( "checksum = " + checksum + "  sum = " + sum, sum != checksum ? 3 : 8 );

  return( sum == checksum );
}


function tnef_attr_incomplete( attr ) {
  assert( attr );

  return( attr.len >= 0 && !(attr.buf && attr.buf.length == attr.len) );
}


function tnef_attr_read( instrm, prev_attr ) {
  var attr = null;

  if( prev_attr ) {
    attr = prev_attr;
  } else {
    attr = new TnefAttr();

    attr.lvl_type = instrm.read8();
    assert( attr.lvl_type == LVL_MESSAGE || attr.lvl_type == LVL_ATTACHMENT );

    attr.name = GETINT16( instrm.readByteArray( 2 ) );
    attr.type = GETINT16( instrm.readByteArray( 2 ) );
    attr.len = GETINT32( instrm.readByteArray( 4 ) );
    tnef_byte_pos += 9;
  }

  tnef_log_msg( "TNEF: reading attr\n"
                + "  lvl_type: 0x" + to_hex( attr.lvl_type, 1 )
                + ",  name: " + tnef_attr_name_to_string( attr.name )
                + ",  type: " + tnef_attr_type_to_string( attr.type )
                + ",  length: " + attr.len
                + ",  togo: " + (attr.len - (attr.buf ? attr.buf.length : 0))  // attr.buf might be NULL!
                + ",  pos in TNEF: " + (tnef_byte_pos - 9)
                , 15 );  //MKA 6

  var available = instrm.available();
  if( available > 0 ) {
    var togo = 0;
    if( attr.buf ) {
      togo = attr.len - attr.buf.length;
    } else {
      togo = attr.len;
      attr.buf = "";
    }
    if( togo > available )
      togo = available;

    attr.buf += instrm.readBytes( togo );
    tnef_byte_pos += togo;

    if( attr.buf.length == attr.len ) {
      var checksum = GETINT16( instrm.readByteArray( 2 ) );
      tnef_byte_pos += 2;
      if( !tnef_attr_check_checksum( attr, checksum ) ) {
	tnef_log_msg( "LookOut: TNEF attribute has invalid checksum, message may be corrupt\n", 5 );
	//exit( 1 );
      } else {
	//tnef_attr_dbg_dump( attr );
      }
    } else {
      tnef_log_msg( "TNEF: waiting for " + (attr.len - attr.buf.length) + " more bytes", 20 ); //MKA 8
    }
  }

  return( attr );
}


// ---------- MAPI -----------

const MULTI_VALUE_FLAG = 0x1000;
const GUID_EXISTS_FLAG = 0x8000;

function GUID() {
}
GUID.prototype = {
  data1: 0,
  data2: 0,
  data3: 0,
  data4: Array( 8 )
}

function MAPIAttr() {
}
MAPIAttr.prototype = {
  type: 0,
  name: 0,
  num_values: 0,
  values: null,
  guid: null,
  num_names: 0,
  names: null
}


function pad_to_4byte( x ) {
  return( (x + 3) & ~3 );
  //return( Math.ceil( x/4.0 )*4 );
}

function guid_parse_buf( buf, idx ) {
  var guid = new GUID();

  assert( buf );

  guid.data1 = GETINT32( buf, idx );
  guid.data2 = GETINT16( buf, idx + 4 );
  guid.data3 = GETINT16( buf, idx + 6 );
  for( var i = 0; i < 8; i++ )
    guid.data4[i] = buf[idx + 8 + i];

  return( guid );
}

// parses out the MAPI attibutes hidden in the character buffer
function mapi_attr_read( buf, len ) {
  var idx = 0;
  var i = 0;
  var num_properties = GETINT32( buf, idx );
  var mattrs = Array( num_properties + 1 );

  idx += 4;

  for( i = 0; i < num_properties; i++ ) {
    mattrs[i] = new MAPIAttr();

    mattrs[i].type = GETINT16( buf, idx );
    idx += 2;
    mattrs[i].name = GETINT16( buf, idx );
    idx += 2;

    // Multi-valued attributes have their type modified by the MULTI_VALUE_FLAG value
    if( mattrs[i].type & MULTI_VALUE_FLAG )
      mattrs[i].type -= MULTI_VALUE_FLAG;

    // handle special case of GUID prefixed properties
    if( mattrs[i].name & GUID_EXISTS_FLAG ) {
      // copy GUID
      mattrs[i].guid = guid_parse_buf( buf, idx );
      idx += 16;

      mattrs[i].num_names = GETINT32( buf, idx );
      idx += 4;
      if( mattrs[i].num_names > 0 ) {
	// FIXME: do something useful here!
	var i2 = 0;

	mattrs[i].names = Array( mattrs[i].num_names );

	for( i2 = 0; i2 < mattrs[i].num_names; i2++ ) {
	  var name_len = GETINT32( buf, idx );
	  idx += 4;

	  // read the data into a buffer
	  mattrs[i].names[i2] = buf.substr( idx, name_len );

	  // But what are we going to do with it?

	  idx += pad_to_4byte( name_len );
	}
      } else {
	// get the 'real' name
	mattrs[i].name = GETINT32( buf, idx );
	idx += 4;
      }
    }

    switch( mattrs[i].type ) {
    case MAPI_NULL:
      mattrs[i].num_values = 0;
      mattrs[i].values = null;
      break;

    case MAPI_SHORT:    // 2 bytes with a 2 byte pad
    case MAPI_BOOLEAN:
      mattrs[i].num_values = 1;
      mattrs[i].values = [ GETINT16( buf, idx ) ];
      idx += 4; // advance by 4 because 2 byte pad
      break;

    case MAPI_INT:
    case MAPI_ERROR:
      mattrs[i].num_values = 1;
      mattrs[i].values = [ GETINT32( buf, idx ) ];
      idx += 4;
      break;

    case MAPI_FLOAT:      // 4 bytes
      mattrs[i].num_values = 1;
      mattrs[i].values = [ GETFLT32( buf, idx ) ];
      idx += 4;
      break;

    case MAPI_DOUBLE:
    case MAPI_CURRENCY:
      mattrs[i].num_values = 1;
      mattrs[i].values = [ GETDBL64( buf, idx ) ];
      idx += 8;
      break;

    case MAPI_INT64:
      mattrs[i].num_values = 1;
      mattrs[i].values = [ GETINT64( buf, idx ) ];
      idx += 8;
      break;

    case MAPI_APPTIME:
    case MAPI_SYSTIME:         // 8 bytes
      mattrs[i].num_values = 1;
      mattrs[i].values = [ new Date( (GETINT64( buf, idx ) - 0x019db1ded53e8000)*0.0001 ) ];
      idx += 8;
      break;

    case MAPI_CLSID:
      mattrs[i].num_values = 1;
      mattrs[i].values = [ guid_parse_buf( buf, idx ) ];
      idx += 16; //sizeof (GUID);
      break;

    case MAPI_STRING:
    case MAPI_UNICODE_STRING:
    case MAPI_OBJECT:
    case MAPI_BINARY:       // variable length
    case MAPI_UNSPECIFIED:
      mattrs[i].num_values = GETINT32( buf, idx );
      idx += 4;
      mattrs[i].values = new Array( mattrs[i].num_values );
      for( var val_idx = 0; val_idx < mattrs[i].num_values; val_idx++ ) {
	var val_len = GETINT32( buf, idx );
	idx += 4;

	//if( mattrs[i].name == MAPI_STRING ) // name ?
	//  val_len--; // kill c style string terminator

	if( mattrs[i].type == MAPI_UNICODE_STRING ) {
	  mattrs[i].values[val_idx] = GETUTF16( buf, idx, (val_len - 1) >> 1 );
	} else if( mattrs[i].type == MAPI_STRING ) {// name ?
	  mattrs[i].values[val_idx] = buf.substr( idx, val_len - 1 ); // - 1 to kill c style string terminator
	} else {
	  mattrs[i].values[val_idx] = buf.substr( idx, val_len ); // this may not be right, substr may react to '\0'.
	}
	idx += pad_to_4byte( val_len );
      }
      break;
    default:
      break;
    }
    //if (DEBUG_ON) mapi_attr_dump (a);
  }
  mattrs[i] = null;

  return( mattrs );
}

function mapi_attr_find( mattrs, name_id ) {
  var i;

  if( !mattrs )
    return( null );

  for( i = 0; mattrs[i] && mattrs[i].name != name_id; i++ );
  if( mattrs[i] ) {

    return( mattrs[i] );
  } else
    return( null );
}

// ---------- RTF -----------

const rtf_uncompressed_magic = 0x414c454d;
const rtf_compressed_magic = 0x75465a4c;
const rtf_prebuf = String( "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx" );


function is_rtf_data( data ) {
  var compr_size = 0;
  var uncompr_size = 0;
  var magic;
  var idx = 0;

  compr_size = GETINT32( data, idx );
  idx += 4;
  uncompr_size = GETINT32( data, idx );
  idx += 4;
  magic = GETINT32( data, idx );
  idx += 4;

  if( magic == rtf_uncompressed_magic || magic == rtf_compressed_magic )
    return 1;
  return 0;
}

function decompress_rtf_data( src, len ) {
  const rtf_prebuf_len = rtf_prebuf.length;
  var numin = 0;
  var numout = 0;
  var flag_count = 0;
  var flags = 0;
  var dest = new String( "" ); // size is rtf_prebuf_len + len;

  dest += rtf_prebuf;

  numout = rtf_prebuf_len;
  while( numout < len + rtf_prebuf_len ) {
    // each flag byte flags 8 literals/references, 1 per bit
    flags = ((flag_count++ % 8) == 0) ? src.charCodeAt(numin++) : (flags >> 1);
    if( (flags & 1) == 1 ) {	// 1 == reference
      var offset = src.charCodeAt( numin );
      var length = src.charCodeAt( numin + 1 );
      numin += 2;

      // offset relative to block start
      offset = (offset << 4) | (length >> 4);

      // number of bytes to copy
      length = (length & 0xF) + 2;

      /* decompression buffer is supposed to wrap around back to
       * the beginning when the end is reached.  we save the
       * need for this by pointing straight into the data
       * buffer, and simulating this behaviour by modifying the
       * pointers appropriately */
      offset = (numout&0xfffff000) + offset;
      if( offset >= numout )
	offset -= 4096; // from previous block

      var end = offset + length;
      while( offset < end ) {
	dest += dest.charAt( offset );
	offset++;
      }
    } else {		// 0 == literal
      if( numin < src.length )
	dest += src.charAt( numin++ );
      else
	dest += " ";
    }
    numout = dest.length;
  }

  return( dest.slice( rtf_prebuf_len ) );
}


function rtf_data_parse( data, len ) {
  var out_data = null;
  var out_len = 0;
  var compr_size = 0;
  var uncompr_size = 0;
  var magic;
  var checksum;
  var idx = 0;

  compr_size = GETINT32( data, idx );
  idx += 4;
  uncompr_size = GETINT32( data, idx );
  idx += 4;
  magic = GETINT32( data, idx );
  idx += 4;
  checksum = GETINT32( data, idx );
  idx += 4;

  // sanity check
  assert( compr_size + 4 == len );
  if( compr_size + 4 != len ) {
    tnef_log_msg( "size does not match: got " + (compr_size + 4) + ", expected: " + len, 5 );
    return null;
  }
  out_len = uncompr_size;

  if( magic == rtf_uncompressed_magic ) // uncompressed rtf stream
    out_data = data.slice( 4, 4 + uncompr_size - 1 );
  else if( magic == rtf_compressed_magic ) // compressed rtf stream
    out_data = decompress_rtf_data( data.slice( idx, idx + compr_size - 1 ), uncompr_size );

  return( out_data );
}


function rtf_to_escaped_text( rtf_data ) {
  var is_str = typeof(rtf_data) == "string";
  var index, cur_char, in_escape = false, in_par = false, target;
  var block_depth = 0;
  var text = "";

  target = rtf_data.indexOf( "When" );
  index = rtf_data.indexOf( "\\pard\\plain" );
  if( index < 0 )
    index = 0;
  for( ; index < rtf_data.length; index++ ) {
    if( is_str )
      cur_char = rtf_data.charAt( index );
    else
      cur_char = rtf_data[index];

    if( cur_char == '}' ) {
      block_depth--;
      in_escape = false;
    } else if( cur_char == '{' ) {
      block_depth++;
    } else {
      if( cur_char == '\\')
	in_escape = true;
      else if( in_escape && " \t\n\r".indexOf( cur_char ) >= 0 )
	in_escape = false;

      if( index == target )
	tnef_log_msg( "found target block_depth = " + block_depth + ", in_escape = " + in_escape, 7 );
      if( block_depth == 1 && !in_escape ) {
	if( cur_char == '\n' ) {
	  text += "\\n";
	} else if( cur_char == '\r' ) {
	  // intentionally empty
	} else if( cur_char == ';' ) {
	  text += "\\;";
	} else if( cur_char == ',' ) {
	  text += "\\,";
	} else if( cur_char == '\\') {
	  text += "\\";
	} else {
	  text += cur_char;
	}
      }
    }
  }

  text += "\n";

  return( text );
}



// ---------- Tnef -----------

//static size_t filesize;

/* TNEF signature.  Equivalent to the magic cookie for a TNEF file. */
const TNEF_SIGNATURE = 0x223e9f78;


function TnefListenerInterface() {
}
TnefListenerInterface.prototype = {
  onTnefStart: function ( filename, content_type, length, date ) { },
  onTnefData: function ( position, data ) { },
  onTnefEnd: function ( ) { }
}

function TnefPackage() {
}
TnefPackage.prototype = {
  files: Array(),

  msg_class: null,
  cur_file: null,
  cur_attr: null,
  num_body_parts: 0,
  msg_header: null,
  code_page: 0
}

//MessageBodyTypes {
const MSG_TEXT = 't';
const MSG_HTML = 'h';
const MSG_RTF  = 'r';
const MSG_ICAL = 'i';
const MSG_VCARD = 'v';
//}

function strtrim( str ) {
  var match = (str ? '' + str : '').match( /\S+(\s+\S+)*/ );
  return( match && match.length > 0 ? match[0] : '' );
}

//MKA  gHeaderParser was removed in Thunderbird 13, but nsIMsgHeaderParser is still
//     available!
//     https://developer.mozilla.org/en-US/docs/XPCOM_Interface_Reference/nsIMsgHeaderParser

// This is needed because nsIMsgHeaderParser cannot handle the commas MS allows
// in the phrase and hangs if the email is invalid
function decompose_rfc822_address( address ) {
  var re = new RegExp("@");
  if ( !re.test(address) ) { // no email address so this is just a phrase
    tnef_log_msg( "TNEF: Not an address: " + address, 6);
    parts = [ address, "" ];
  } else {
    // allowing comma in phrase for MS
    var parts = address.match( /^\s*((?:[^\x28\x29\x3c\x3e\x40\x3a\x3b\x5b\x5c\x5d]+|\x22[^\x22]*\x22)*)[\x3c\x28]([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6})[\x29\x3e]\s*$/i );
    if( !parts ) // just an address
      parts = address.match( /^\s*()([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6})$/i );
    if( parts ) // if either of the regex matches grab off the full match element
      parts.shift();
  }
  if( parts ) // if we have something work it over
    parts[0] = strtrim( parts[0] ); // trim up the phrase

  return( parts );
}

var ownHeaderParser = Components.classes["@mozilla.org/messenger/headerparser;1"].getService(Components.interfaces.nsIMsgHeaderParser);

function tnef_pack_get_name_addr( pkg, orig_name_addr ) {
  tnef_log_msg( "TNEF: Parsing original address line: " + orig_name_addr, 6);
  var orig_addr_parts = decompose_rfc822_address( orig_name_addr );

  // if email address not in the name_addr then hunt through the message header
  if( pkg.msg_header && (!orig_addr_parts[1] || orig_addr_parts[1] == "") ) {
    var all_addrs = [], all_names = [];
    var addrs = {}, names = {}, full_names = {};
    var i, num_addrs;

    function rm_quotes( element, idx, array ) {
      if( array[idx] ) {
        var matches = array[idx].match( /\042([^\042]*)\042|\047([^\047]*)\047/ );
        if( matches && matches.length > 0 )
	       array[idx] = matches.length < 2 || (matches[1] && matches[1] != "") ? matches[1] : matches[2];
      }
    }

    orig_addr_parts[1] = "";

    num_addrs = ownHeaderParser.parseHeadersWithArray( pkg.msg_header.author, addrs, names, full_names );
    all_addrs = addrs.value;
    all_names = names.value;
    num_addrs += ownHeaderParser.parseHeadersWithArray( pkg.msg_header.recipients, addrs, names, full_names );
    all_addrs = all_addrs.concat( addrs.value );
    all_names = all_names.concat( names.value );
    num_addrs += ownHeaderParser.parseHeadersWithArray( pkg.msg_header.ccList, addrs, names, full_names );
    all_addrs = all_addrs.concat( addrs.value );
    all_names = all_names.concat( names.value );

    all_names.forEach( rm_quotes );

    for( i = 0; i < all_names.length && orig_addr_parts[1] == ""; i++ )
      if( all_names[i] == orig_addr_parts[0] )
	     orig_addr_parts[1] = all_addrs[i];
  }

  return( orig_addr_parts );
}

function tnef_pack_data_left( instrm ) {
  var dl = instrm.available();

  if( dl > 0 && dl < MINIMUM_ATTR_LENGTH ) {
    tnef_log_msg( "ERROR: garbage at end of file.\n", 5 );
    return( false );
  }

  return( true );
}


// Reads and decodes a object from the stream
function tnef_pack_read_object( pkg, instrm ) {
  var togo = 0;

  if( pkg.cur_attr )
    togo = pkg.cur_attr.len - pkg.cur_attr.buf.length;
  else
    togo = MINIMUM_ATTR_LENGTH;

  // peek to see if there is more to read from this stream
  if( !pkg.cur_attr && instrm.available() < togo )
    return( null );

  return( tnef_attr_read( instrm, pkg.cur_attr ) );
}


function tnef_pack_handle_body_part( pkg, data, len, content_type, listener ) {
  if( !data )
    return;

  var body_part_prefix = tnef_get_string_pref( "body_part_prefix", "body_part_" );

  var file = new TnefFile( pkg );

  if( typeof( data ) == "Array" )
    file.data = String.fromCharCode.apply( null, data );
  else
    file.data = data;

  file.name = body_part_prefix + pkg.num_body_parts;
  switch( content_type ) {
  case "text/plain": file.name += ".txt"; break;
  case "text/html": file.name += ".html"; break;
  case "application/rtf": file.name += ".rtf"; break;
  case "text/calendar": file.name += ".ics"; break;
  case "text/x-vcard": file.name += ".vcf"; break;
  }
  file.len = len;
  file.mime_type = content_type;
  tnef_file_notify( file, listener, true );

  pkg.num_body_parts++;
}

function tnef_pack_handle_text_attr( pkg, attr, listener ) {
  tnef_pack_handle_body_part( pkg, attr.buf, attr.len, "text/plain", listener );
}

function tnef_pack_handle_html_data( pkg, mattr, listener ) {
  for( var j = 0; j < mattr.num_values; j++ )
    tnef_pack_handle_body_part( pkg, mattr.values[j], mattr.values[j].length, "text/html", listener );
}


function tnef_pack_handle_rtf_data( pkg, mattr, listener ) {
  for( var j = 0; j < mattr.num_values; j++ ) {
    if( is_rtf_data( mattr.values[j] ) ) {
      var rtf_data = rtf_data_parse( mattr.values[j], mattr.values[j].length );
      if( rtf_data )
	tnef_pack_handle_body_part( pkg, rtf_data, rtf_data.length, "application/rtf", listener );
    }
  }
}


function tnef_pack_handle_mapi_attrs( pkg, mattrs, listener ) {
  var data = "";
  var i = 0;

  if( !mattrs )
    return( null );

  for( i = 0; i < mattrs.length && mattrs[i]; i++ ) {
    data += mattr_name_to_string( mattrs[i].name, pkg.msg_class ) +
            ": type=" + mattr_type_to_string( mattrs[i].type ) + ";";
    if( mattrs[i].values ) {
      switch( mattrs[i].type ) {
      case MAPI_NULL:
	data += " value=null";
	break;
      case MAPI_CLSID:
	data += " value=" + to_hex( mattrs[i].values[0].data1 ) + "-" +
	  to_hex( mattrs[i].values[0].data2 ) + "-" + to_hex( mattrs[i].values[0].data3 ) + "-" +
	  to_hex( mattrs[i].values[0].data4 );
	break;
      case MAPI_INT64:
	data += " value=" + to_hex( mattrs[i].values[0], 8 );
	break;
      case MAPI_BINARY:
	if( mattrs[i].num_values > 1 )
	  data += " num-values=" + mattrs[i].num_values + ";";
	data += " data=base64," + tnef_base64_encode( mattrs[i].values[0] );
	break;
      default:
	if( mattrs[i].num_values > 1 )
	  data += " num-values=" + mattrs[i].num_values + ";";
	data += " value='" + mattrs[i].values[0] + "'";
	break;
      }
    } else {
      data += " void";
    }
    data += "\n";

    //if (DEBUG_ON) mapi_attr_dump (a);
  }

  tnef_pack_handle_body_part( pkg, data, data.length, "text/plain", listener );
}

function mescape( str ) {
  var i = 0;
  var newstr;

  newstr = str.replace( "\n", "\\n" );
  newstr = newstr.replace( "\t", "\\t" );
  newstr = newstr.replace( ",", "\\," );
  newstr = newstr.replace( "\"", "\\\"" );

  return( newstr );
}

function tnef_pack_handle_contact_data( pkg, mattrs, listener ) {
  var vcal_str = "";
  var mattr = null;

  mattr = mapi_attr_find( mattrs, MAPI_SENDER_SEARCH_KEY );
  if( !mattr || mattr.values[0] == 0 )
    mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_ORGANIZER_ALIAS );
  if( mattr && mattr.num_values > 0 ) {
    var i = mattr.values[0].indexOf( ":" );
    if( i < 0 )
      i = 0;
    else
      i++;
    if ( mescape( mattr.values[0] ).length > 1 )
      vcal_str += "FN:" + mescape( mattr.values[0] ) + "\n";
  }
  mattr = mapi_attr_find( mattrs, MAPI_COMPANY_NAME );
  if( mattr && mattr.num_values > 0 ) {
    vcal_str += "ORG:" + mattr.values[0] + "\n";
  }
  mattr = mapi_attr_find( mattrs, MAPI_TITLE );
  if( mattr && mattr.num_values > 0 ) {
    vcal_str += "TITLE:" + mattr.values[0] + "\n";
  }

  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_BUSINESS_PHONE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=work,voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_HOME_PHONE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=home,voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_PRIMARY_PHONE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=pref,voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_BUSINESS_PHONE_2 );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=work,voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_OTHER_PHONE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_MOBILE_PHONE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=cell,voice:" + mattr.values[0] + "\n";
  mattr = mapi_attr_find( mattrs, MAPI_CONTACT_BUSINESS_FAX );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "TEL;TYPE=work,fax:" + mattr.values[0] + "\n";

  if ( vcal_str ) {
    vcal_str = "VERSION:2.1\n" + vcal_str;
    vcal_str = "PRODID:-//Mozilla//Mozilla Mail//EN\n" + vcal_str;
    vcal_str = "BEGIN:VCARD\n" + vcal_str;
    // Wrap it up
    vcal_str += "END:VCARD\n";

    tnef_pack_handle_body_part( pkg, vcal_str, vcal_str.length, "text/x-vcard", listener );
  } else {
    tnef_log_msg( "TNEF: VCF file is empty, Skipping", 6 );
  }
}


function date_from_1601_mins( mins_1601 ) {
  const mins_1601_1970 = 11644473600/60;
  return( new Date( (mins_1601 - mins_1601_1970)*60e6 ) );
}


function ical_date_string( dt ) {
  return( "" + dt.getUTCFullYear() + ntoa_pad(dt.getUTCMonth() + 1,2) +
	  ntoa_pad(dt.getUTCDate(),2) + "T" + ntoa_pad(dt.getUTCHours(),2) +
	  ntoa_pad(dt.getUTCMinutes(),2) + ntoa_pad(dt.getUTCSeconds(),2) + "Z" );
}


function ical_day_list_from_bitset( bits ) {
  var vcal_str = "";
  if( bits & 0x01 ) vcal_str += "SU";
  if( bits > 0x01 ) vcal_str += ",";
  if( bits & 0x02 ) vcal_str += "MO";
  if( bits > 0x03 ) vcal_str += ",";
  if( bits & 0x04 ) vcal_str += "TU";
  if( bits > 0x07 ) vcal_str += ",";
  if( bits & 0x08 ) vcal_str += "WE";
  if( bits > 0x0F ) vcal_str += ",";
  if( bits & 0x10 ) vcal_str += "TH";
  if( bits > 0x1F ) vcal_str += ",";
  if( bits & 0x20 ) vcal_str += "FR";
  if( bits > 0x3F ) vcal_str += ",";
  if( bits & 0x40 ) vcal_str += "SA";
  return( vcal_str );
}


function tnef_pack_appt_attendees_ics( pkg, liststr, role ) {
  var vcal_str = "";
  var atnd = 0;
  var attendees;

  if( liststr == "" )
    return( "" );

  attendees = liststr.split( ";" );
  attendees = attendees.filter(function(e){return e});

  for( atnd = 0; atnd < attendees.length; atnd++ ) {
    if( !attendees[atnd].replace(/[^\x20-\x7E]+/g, '') ){
      continue;
    }
    var parts = tnef_pack_get_name_addr( pkg, attendees[atnd] );
    vcal_str += "ATTENDEE;PARTSTAT=NEEDS-ACTION;ROLE="+role+";RSVP=TRUE";
    if( parts.length > 1 && parts[1] && parts[1] != "" )
      vcal_str += ";CN=\"" + parts[0] + "\":mailto:\"" + parts[1] + "\"\n";
    else
      vcal_str += ":" + parts[0] + "\n";
  }

  return( vcal_str );
}


function tnef_pack_handle_appt_data( pkg, mattrs, listener ) {
  var vcal_str = "BEGIN:VCALENDAR\n";
  var mattr = null;
  var is_meeting = pkg.msg_class.indexOf( ".Meeting." ) >= 0;

  if( pkg.msg_class && (pkg.msg_class.indexOf( ".MtgCncl" ) >= 0 ||
			pkg.msg_class.indexOf( ".Meeting.Canceled" ) >=0) )
    vcal_str += "METHOD:CANCEL\n";
  else
    vcal_str += "METHOD:REQUEST\n";

  vcal_str += "PRODID:-//Mozilla//Mozilla Mail//EN\n";
  vcal_str += "VERSION:2.0\n";
  vcal_str += "BEGIN:VEVENT\n";

  mattr = mapi_attr_find( mattrs, MAPI_MEETING_CLEAN_GLOBAL_OBJECT_ID );
  if( !mattr || mattr.values[0] == 0 )
    mattr = mapi_attr_find( mattrs, MAPI_MEETING_GLOBAL_OBJECT_ID );
  if( !mattr || mattr.values[0] == 0 )
    mattr = mapi_attr_find( mattrs, MAPI_MAPPING_SIGNATURE );
  if( mattr && mattr.num_values > 0 ) {
    var i = 0;

    vcal_str += "UID:";
    for( i = 0; i < mattr.values[0].length; i++ )
      vcal_str += to_hex( mattr.values[0].charCodeAt( i ), 1 );
    vcal_str += "\n";
  }

  // Sequence
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_SEQUENCE );
  if( mattr && mattr.num_values > 0 )
    vcal_str += "SEQUENCE:" + mattr.values[0] + "\n";

  mattr = mapi_attr_find( mattrs, MAPI_PRIMARY_SEND_ACCOUNT );
  if( mattr && mattr.num_values > 0 ) {
    var parts = mattr.values[0].split( "\001" );
    if( parts.length )
      vcal_str += "ORGANIZER;CN=\"" + parts[2] + "\":mailto:" + parts[1] + "\n";
  }
  if( !mattr || mattr.values[0] == 0 ) {
    var parts;
    mattr = mapi_attr_find( mattrs, MAPI_SENDER_NAME );
    var mattr2 = mapi_attr_find( mattrs, MAPI_SENDER_EMAIL_ADDRESS );
    if( mattr && mattr.num_values > 0 && mattr2 && mattr2.num_values > 0 ) {
      parts = [ mattr.values[0], mattr2.values[0] ];
    } else {
      if( !mattr || mattr.values[0] == 0 )
	mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_ORGANIZER_ALIAS );
      if( !mattr || mattr.values[0] == 0 )
	mattr = mapi_attr_find( mattrs, MAPI_CREATOR_NAME );
      if( !mattr || mattr.values[0] == 0 )
	mattr = mapi_attr_find( mattrs, MAPI_TASK_F_CREATOR );
      if( mattr && mattr.num_values > 0 )
	parts = tnef_pack_get_name_addr( pkg, mattr.values[0] );
    }
    if( parts ) {
      vcal_str += "ORGANIZER;PARTSTAT=ACCEPTED;ROLE=CHAIR";
      if( parts[1] && parts[1] != "" )
	vcal_str += ";CN=\"" + parts[0] + "\":mailto:" + parts[1] + "\n";
      else
	vcal_str += ":" + parts[0] + "\n";
    }
  }

  // Required Attendees
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_REQUIRED_ATTENDEES );
  if( !mattr || mattr.values[0] == 0 )
    mattr = mapi_attr_find( mattrs, MAPI_MEETING_REQUIRED_ATTENDEES );
  if( mattr && mattr.num_values > 0 )
    vcal_str += tnef_pack_appt_attendees_ics( pkg, mattr.values[0], "REQ-PARTICIPANT" );

  // Optional attendees
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_OPTIONAL_ATTENDEES );
  if( !mattr || mattr.values[0] == 0 )
    mattr = mapi_attr_find( mattrs, MAPI_MEETING_OPTIONAL_ATTENDEES );
  if( mattr && mattr.num_values > 0 ) {
    vcal_str += tnef_pack_appt_attendees_ics( pkg, mattr.values[0], "OPT-PARTICIPANT" );
  }

  mattr = mapi_attr_find( mattrs, MAPI_CONVERSATION_TOPIC );
  if( mattr )
    vcal_str += "SUMMARY:" + mattr.values[0] + "\n";

  mattr = mapi_attr_find( mattrs, MAPI_RTF_COMPRESSED );
  if( mattr ) {
    var rtf_data = rtf_data_parse( mattr.values[0], mattr.values[0].length );
    vcal_str += "DESCRIPTION:" + rtf_to_escaped_text( rtf_data );
  }

  // Location
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_LOCATION );
  if( mattr ) {
    vcal_str += "LOCATION:" + mescape( mattr.values[0] ) + "\n";
  }

  // Date Start
  mattr = mapi_attr_find( mattrs, MAPI_START_DATE );
  if( !mattr || !mattr.num_values )
    mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_START_WHOLE );
  if( !mattr || !mattr.num_values )
    mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_START_DATE );
  if( mattr && mattr.num_values ) {
    var dt = mattr.values[0];
    vcal_str += "DTSTART:" + ical_date_string( dt ) + "\n";
  }

  // Date End
  mattr = mapi_attr_find( mattrs, MAPI_END_DATE );
  if( !mattr || !mattr.num_values )
    mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_END_WHOLE );
  if( !mattr || !mattr.num_values )
    mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_END_DATE );
  if( mattr && mattr.num_values ) {
    var dt = mattr.values[0];
    vcal_str += "DTEND:" + ical_date_string( dt ) + "\n";
  }

  // Date Stamp
  mattr = mapi_attr_find( mattrs, MAPI_CREATION_TIME );
  if( mattr && mattr.num_values ) {
    var dt = mattr.values[0];
    vcal_str += "CREATED:" + ical_date_string( dt ) + "\n";
  }

  // Class
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_SEQUENCE_TIME );
  if( mattr && mattr.num_values ) {
    //vcal_str += "CLASS:" + (mattr.values[0] == 1 ? "PRIVATE" : "PUBLIC") + "\n";
  }

  const MAPI_APPT_RECUR_FREQ_DAILY	  = 0x200A;
  const MAPI_APPT_RECUR_FREQ_WEEKLY	  = 0x200B;
  const MAPI_APPT_RECUR_FREQ_MONTHLY	  = 0x200C;
  const MAPI_APPT_RECUR_FREQ_YEARLY	  = 0x200D;
  const MAPI_APPT_RECUR_PATT_DAY	  = 0x0000;
  const MAPI_APPT_RECUR_PATT_WEEK	  = 0x0001;
  const MAPI_APPT_RECUR_PATT_MONTH	  = 0x0002;
  const MAPI_APPT_RECUR_PATT_MONTH_NTH	  = 0x0003;
  const MAPI_APPT_RECUR_PATT_MONTH_END	  = 0x0004;
  const MAPI_APPT_RECUR_PATT_HJ_MONTH	  = 0x000A;
  const MAPI_APPT_RECUR_PATT_HJ_MONTH_NTH = 0x000B;
  const MAPI_APPT_RECUR_PATT_HJ_MONTH_END = 0x000C;
  const MAPI_APPT_RECUR_END_AFTER_DATE	  = 0x00002021;
  const MAPI_APPT_RECUR_END_AFTER_OCCURS  = 0x00002021;
  const MAPI_APPT_RECUR_END_NEVER	  = 0x00002023;
  const MAPI_APPT_RECUR_END_NEVER_EVER	  = 0xFFFFFFFF;

  // Recurrence
  mattr = mapi_attr_find( mattrs, MAPI_APPOINTMENT_RECUR );
  if( mattr && mattr.num_values ) {
    var recur_freq = GETINT16( mattr.values[0], 4 ), patt_type = GETINT16( mattr.values[0], 6 );
    var cal_type = GETINT16( mattr.values[0], 8 ), first_date = GETINT32( mattr.values[0], 10 );
    var period = GETINT32( mattr.values[0], 14 ), sliding_flag = GETINT32( mattr.values[0], 18 );
    var patt_type_spec = GETINT32( mattr.values[0], 22 );
    var patt_type_spec2 = GETINT32( mattr.values[0], 26 );
    var pts_off = (patt_type == MAPI_APPT_RECUR_PATT_DAY ? 0 :
		   (patt_type == MAPI_APPT_RECUR_PATT_MONTH_NTH ||
		    patt_type == MAPI_APPT_RECUR_PATT_HJ_MONTH_NTH) ? 8 : 4);
    var end_type = GETINT32( mattr.values[0], 22 + pts_off );
    var occur_cnt = GETINT32( mattr.values[0], 26 + pts_off );
    var first_dow = GETINT32( mattr.values[0], 30 + pts_off );
    var del_inst_cnt = GETINT32( mattr.values[0], 34 + pts_off );
    var del_inst_start = 38 + pts_off, del_inst_end = del_inst_start + 4*del_inst_cnt;
    var mod_inst_cnt = GETINT32( mattr.values[0], del_inst_end );
    var mod_inst_start = del_inst_end + 4, mod_inst_end = mod_inst_start + 4*mod_inst_cnt;
    var start_date = date_from_1601_mins( GETINT32( mattr.values[0], mod_inst_end ) );
    var end_date = date_from_1601_mins( GETINT32( mattr.values[0], mod_inst_end + 4 ) );

    vcal_str += "RRULE:";
    switch( recur_freq ) {
    case MAPI_APPT_RECUR_FREQ_DAILY:   vcal_str += "FREQ=DAILY"; break;
    case MAPI_APPT_RECUR_FREQ_WEEKLY:  vcal_str += "FREQ=WEEKLY;"; break;
    case MAPI_APPT_RECUR_FREQ_MONTHLY: vcal_str += "FREQ=MONTHLY"; break;
    case MAPI_APPT_RECUR_FREQ_YEARLY:  vcal_str += "FREQ=YEARLY"; break;
    }

    vcal_str += ";INTERVAL=" + period*(recur_freq == MAPI_APPT_RECUR_FREQ_DAILY ? 1/1440 : 1);

    switch( end_type ) {
    case MAPI_APPT_RECUR_END_AFTER_DATE:   vcal_str += ";UNTIL=" + ical_date_string( end_date ); break;
    case MAPI_APPT_RECUR_END_AFTER_OCCURS: vcal_str += ";COUNT=" + occur_cnt; break;
    case MAPI_APPT_RECUR_END_NEVER: case MAPI_APPT_RECUR_END_NEVER_EVER: break;
    }

    switch( patt_type ) {
    case MAPI_APPT_RECUR_PATT_WEEK:
      vcal_str += ";BYDAY=" + ical_day_list_from_bitset( patt_type_spec );
      break;
    case MAPI_APPT_RECUR_PATT_MONTH: case MAPI_APPT_RECUR_PATT_MONTH_END:
    case MAPI_APPT_RECUR_PATT_HJ_MONTH: case MAPI_APPT_RECUR_PATT_HJ_MONTH_END:
      vcal_str += ";BYMONTHDAY=" + patt_type_spec;
      break;
    case MAPI_APPT_RECUR_PATT_MONTH_NTH: case MAPI_APPT_RECUR_PATT_HJ_MONTH_NTH:
      vcal_str += ";BYDAY=" + ical_day_list_from_bitset( patt_type_spec );
      vcal_str += ";BYSETPOS=" + (patt_type_spec2 == 0x05 ? -1 : patt_type_spec2);
    }

    vcal_str += "\n";
  }

  // Wrap it up
  vcal_str += "END:VEVENT\n";
  vcal_str += "END:VCALENDAR\n";

  if( vcal_str )
    tnef_pack_handle_body_part( pkg, vcal_str, vcal_str.length, "text/calendar", listener );
}


// The entry point into this module.  This parses an entire TNEF file.
function tnef_pack_parse_stream( instrm, msg_header, listener, prev_pack ) {
  var sig = 0;
  var key = 0;
  var pkg = null;

  tnef_log_msg( "TNEF: Entering tnef_pack_parse_stream()", 6); //MKA

  if( prev_pack ) {
    pkg = prev_pack;
  } else {
    pkg = new TnefPackage();
    pkg.msg_header = msg_header;

    // check that this is in fact a TNEF file
    sig = GETINT32( instrm.readByteArray( 4 ) );
    if( sig != TNEF_SIGNATURE ) {
      tnef_log_msg( "sig = " + sig + "\nSeems not to be a TNEF file\n", 8 );
      return( null );
    }

    // Get the key
    key = GETINT16( instrm.readByteArray( 2 ) );
    tnef_log_msg( "TNEF Key: " + to_hex( key, 2 ), 8 );

    tnef_byte_pos = 6;
  }

  // The rest of the file is a series of 'messages' and 'attachments'
  pkg.cur_attr = tnef_pack_read_object( pkg, instrm );
  if( pkg.cur_attr && tnef_attr_incomplete( pkg.cur_attr ) )
    return( pkg );
  while( pkg.cur_attr && tnef_pack_data_left( instrm ) ) {
    if( pkg.cur_attr.name == TNEF_ATTR_OEM_CODEPAGE ) {
      pkg.code_page = GETINT32( pkg.cur_attr.buf );
      tnef_log_msg( "TNEF: OEM Code Page = " + pkg.code_page, 7 );
    }
    // This signals the beginning of a file
    if( pkg.cur_attr.name == TNEF_ATTR_ATTACH_REND_DATA ) {
      if( pkg.cur_file )
	tnef_file_notify( pkg.cur_file, listener, true );
      else
	tnef_log_msg( "starting file sub-attachment", 7 );
      pkg.cur_file = new TnefFile( pkg );
    }
    // Add the data to our lists.
    switch( pkg.cur_attr.lvl_type ) {
    case LVL_MESSAGE:
      if( pkg.cur_attr.name == TNEF_ATTR_BODY ) {
	tnef_pack_handle_text_attr( pkg, pkg.cur_attr, listener );
      } else if( pkg.cur_attr.name == TNEF_ATTR_MESSAGE_CLASS ) {
	pkg.msg_class = pkg.cur_attr.buf;
      } else if( pkg.cur_attr.name == TNEF_ATTR_MAPI_PROPS ) {
	var mapi_attrs = mapi_attr_read( pkg.cur_attr.buf, pkg.cur_attr.len );

	if( mapi_attrs ) {
	  var has_contact_data = false;

	  if( tnef_get_bool_pref( "attach_raw_mapi" ) )
	    tnef_pack_handle_mapi_attrs( pkg, mapi_attrs, listener );

	  for( i = 0; mapi_attrs[i]; i++ ) {
	    switch( mapi_attrs[i].name ) {
	    case MAPI_BODY_HTML:
	      tnef_pack_handle_html_data( pkg, mapi_attrs[i], listener );
	      break;
	    case MAPI_RTF_COMPRESSED:
	      tnef_pack_handle_rtf_data( pkg, mapi_attrs[i], listener );
	      break;
	    default:
	      // cannot save attributes to file yet, since they are not attachment attributes
	      // tnef_file_add_mapi_attrs( pkg.cur_file, mapi_attrs );
	      break;
	    }

	    if( mapi_attrs[i].name > 0x3A00 && mapi_attrs[i].name <= 0x3AFF )
	      has_contact_data = true;
	  }

	  // Message			"IPM"  "IPM.Microsoft Mail.Note"  "IPM.Note"
	  // Meeting Request		"IPM.Microsoft Schedule.MtgReq"	  "IPM.Schedule.Meeting.Request"
	  // Positive Meeting Response	"IPM.Microsoft Schedule.MtgRespP" "IPM.Schedule.Meeting.Resp.Pos"
	  // Negative Meeting Response	"IPM.Microsoft Schedule.MtgRespN" "IPM.Schedule.Meeting.Resp.Neg"
	  // Tentative Meeting Response	"IPM.Microsoft Schedule.MtgRespA" "IPM.Schedule.Meeting.Resp.Tent"
	  // Meeting Cancellation	"IPM.Microsoft Schedule.MtgCncl"  "IPM.Schedule.Meeting.Canceled"
	  // Non-delivery		"Report.IPM.Note.NDR"  "IPM.Microsoft Mail.Non-Delivery"
	  // Read/Return Receipt	"Report.IPM.Note.RN"  "IPM.Microsoft Mail.Read Receipt"
	  if( pkg.msg_class &&
	      (pkg.msg_class.substring( 0, 15 ) == "IPM.Appointment" ||
	       pkg.msg_class.substring( 0, 12 ) == "IPM.Schedule" ||
	       pkg.msg_class.substring( 0, 22 ) == "IPM.Microsoft Schedule") )
	    tnef_pack_handle_appt_data( pkg, mapi_attrs, listener );

	  if( has_contact_data )
	    tnef_pack_handle_contact_data( pkg, mapi_attrs, listener );

	  mapi_attrs = null;
	}
      }
      break;
    case LVL_ATTACHMENT:
      if( !pkg.cur_file )
	pkg.cur_file = new TnefFile( pkg );
      tnef_file_add_attr( pkg.cur_file, pkg.files, pkg.cur_attr, pkg, listener );

      // remember attachment file name to avoid collisions
      if( pkg.cur_file.name && pkg.cur_file.len > 0 )
	pkg.files.push( pkg.cur_file.name );
      break;
    default:
      tnef_log_msg( "Invalid lvl type on attribute: " + pkg.cur_attr.lvl_type, 5 );
      //return( 1 );
      break;
    }

    pkg.cur_attr = null;
    pkg.cur_attr = tnef_pack_read_object( pkg, instrm );
    if( pkg.cur_attr && tnef_attr_incomplete( pkg.cur_attr ) )
      return( pkg );
  }
  if( pkg.cur_file ) {
    tnef_file_notify( pkg.cur_file, listener, true );
    pkg.cur_file = null;
  }
  if( pkg.cur_attr )
    pkg.cur_attr = null;

  return( null );
}

