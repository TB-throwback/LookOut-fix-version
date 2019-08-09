/*
 * File: lookout.js
 *   LookOut Mozilla MailNews Attachments Javascript Overlay for TNEF
 *
 * Copyright:
 *   Copyright (C) 2007-2010 Aron Rubin <arubin@atl.lmco.com>
 *
 * About:
 *   Benevolently hijack the Mozilla mailnews attachment list and expand all
 *   Transport Neutral Encapsulation Format (TNEF) encoded attchments. When
 *   TNEF attachments are opened, decode them.
 */

/* ***** BEGIN LICENSE BLOCK *****
 *   Version: MPL 1.1/GPL 2.0/LGPL 2.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is LookOut.
 *
 * The Initial Developer of the Original Code is
 * Aron Rubin.
 * Portions created by the Initial Developer are Copyright (C) 2007-2010
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *
 * Alternatively, the contents of this file may be used under the terms of
 * either the GNU General Public License Version 2 or later (the "GPL"), or
 * the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
 * in which case the provisions of the GPL or the LGPL are applicable instead
 * of those above. If you wish to allow use of your version of this file only
 * under the terms of either the GPL or the LGPL, and not to allow others to
 * use your version of this file under the terms of the MPL, indicate your
 * decision by deleting the provisions above and replace them with the notice
 * and other provisions required by the GPL or the LGPL. If you do not delete
 * the provisions above, a recipient may use your version of this file under
 * the terms of any one of the MPL, the GPL or the LGPL.
 *
 * ***** END LICENSE BLOCK ***** */


// How long we should wait for window initialization to finish
const LOOKOUT_WAIT_MAX = 10;
const LOOKOUT_WAIT_TIME = 100;

const LOOKOUT_PREF_PREFIX = "extensions.lookout.";

// Declare Debug level Globaly
var debugLevel = 10;

var uid

try {
	var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);
	lightning = prefs.getBoolPref( "extensions.installedDistroAddon.{e2fda1a4-762b-4020-b5ad-a41df1933103}" );
} catch (e) {
	lightning = false;
}

var lookout = {
	log_msg: function lo_log_msg( msg, level ) {
		if( (level == null ? 9 : level) <= debugLevel ) {
			var cs = Components.classes["@mozilla.org/consoleservice;1"].getService(Components.interfaces.nsIConsoleService);
			cs.logStringMessage( msg );
		}
	},

	get_pref: function lo_get_pref( name, get_type_func, default_val ) {
		var pref_name = LOOKOUT_PREF_PREFIX + name;
		var pref_val;
		try {
			var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);
			pref_val = prefs[get_type_func]( pref_name );
		} catch (ex) {
			this.log_msg( "LO: warning: could not retrieve setting '" + pref_name + "': " + ex, 5 );
		}
		if( pref_val === void(0) )
			pref_val = default_val;

		return pref_val;
	},
	get_bool_pref: function lo_get_bool_pref( name, default_val ) {
		return this.get_pref( name, "getBoolPref", default_val );
	},
	get_string_pref: function lo_get_string_pref( name, default_val ) {
		return this.get_pref( name, "getCharPref", default_val );
	},
	get_int_pref: function lo_get_int_pref( name, default_val ) {
		return this.get_pref( name, "getIntPref", default_val );
	},

	debug_check: function lo_debug_check() {
		// set debug level based on user preference
		if(!this.get_bool_pref("debug_enabled")){
		  debugLevel = 5;
		} else {
			debugLevel = 10;
			this.log_msg( "LookOut: debug enabled ", 5 );
		}
	},

	basename: function lo_basename( path ) {
		var file = Components.classes["@mozilla.org/file/local;1"].createInstance(Components.interfaces.nsIFile);

		try {
			file.initWithPath( path );
			return( file.leafName );
		} catch (e) {
			return( null );
		}
	},

	get_temp_file: function lo_get_temp_file( filename ) {
		var file_locator = Components.classes["@mozilla.org/file/directory_service;1"].getService(Components.interfaces.nsIProperties);
		var temp_dir = file_locator.get( "TmpD", Components.interfaces.nsIFile );

		var local_target = temp_dir.clone();
		local_target.append( "lookout" );
		local_target.append( filename );

		return( local_target );
	},

	cal_trans_mgr: null,

	get_cal_trans_mgr: function lo_get_cal_trans_mgr() {
		if( !this.cal_trans_mgr ) {
			try {
	this.cal_trans_mgr = Components.classes["@mozilla.org/calendar/transactionmanager;1"].getService(Components.interfaces.calITransactionManager);
			} catch (ex) {
	this.cal_trans_mgr = Components.classes["@mozilla.org/transactionmanager;1"].createInstance(Components.interfaces.nsITransactionManager);
			}
		}
		return( this.cal_trans_mgr );
	},

	cal_update_undo_redo_menu: function lo_cal_update_undo_redo_menu() {
		var trans_mgr = this.get_cal_trans_mgr();
		if( !trans_mgr )
			return;

		if( trans_mgr.numberOfUndoItems )
			document.getElementById('undo_command').removeAttribute( 'disabled' );
		else
			document.getElementById('undo_command').setAttribute( 'disabled', true );

		if( trans_mgr.numberOfRedoItems )
			document.getElementById('redo_command').removeAttribute( 'disabled' );
		else
			document.getElementById('redo_command').setAttribute( 'disabled', true );
	},

	cal_add_items: function lo_cal_add_items( destCal, aItems, aFilePath ) {
		var trans_mgr = this.get_cal_trans_mgr();
		if( !trans_mgr )
			return;

		// Set batch for the undo/redo transaction manager
		trans_mgr.beginBatch();

		// And set batch mode on the calendar, to tell the views to not
		// redraw until all items are imported
		destCal.startBatch();

		// This listener is needed to find out when the last addItem really
		// finished. Using a counter to find the last item (which might not
		// be the last item added)
		var count = 0;
		var failedCount = 0;
		var duplicateCount = 0;
		// Used to store the last error. Only the last error, because we don't
		// wan't to bomb the user with thousands of error messages in case
		// something went really wrong.
		// (example of something very wrong: importing the same file twice.
		//  quite easy to trigger, so we really should do this)
		var lastError;
		var listener = {
			onOperationComplete: function (aCalendar, aStatus, aOperationType, aId, aDetail) {
	count++;
	if (!Components.isSuccessCode(aStatus)) {
		if (aStatus == Components.interfaces.calIErrors.DUPLICATE_ID) {
			duplicateCount++;
		} else {
			failedCount++;
			lastError = aStatus;
		}
	}
	// See if it is time to end the calendar's batch.
	if (count == aItems.length) {
		destCal.endBatch();
		var sbs = Components.classes["@mozilla.org/intl/stringbundle;1"].getService(Components.interfaces.nsIStringBundleService);
		var cal_strbundle = sbs.createBundle("chrome://calendar/locale/calendar.properties");
		if (!failedCount && duplicateCount ) {
			this.log_msg( "LookOut: " + cal_strbundle.GetStringFromName( "duplicateError" ) + " " +
			duplicateCount + " " + aFilePath, 3 );
		} else if (failedCount) {
			this.log_msg( "LookOut: " + cal_strbundle.GetStringFromName( "importItemsFailed" ) + " " +
			failedCount + " " + lastError.toString(), 3 );
		}
	}
			}
		};

		for( var i = 0; i < aItems.length; i++ ) {
			// XXX prompt when finding a duplicate.
			try {
	destCal.addItem( aItems[i], listener );
			} catch (ex) {
	failedCount++;
	lastError = ex;
	// Call the listener's operationComplete, to increase the
	// counter and not miss failed items. Otherwise, endBatch might
	// never be called.
	listener.onOperationComplete( null, null, null, null, null );
	Components.utils.reportError( "Import error: " + ex );
			}
		}

		// End transmgr batch
		trans_mgr.endBatch();
		this.cal_update_undo_redo_menu();
	}
}


const LOOKOUT_ACTION_SCAN = 0;
const LOOKOUT_ACTION_OPEN = 1;
const LOOKOUT_ACTION_SAVE = 2;

function LookoutStreamListener() {
}
LookoutStreamListener.prototype = {
	attachment: null,
	mAttUrl: null,
	mMsgUri: null,
	mStream: null,
	mPackage: null,
	mPartId: 1,
	mMsgHdr: null,
	action_type: LOOKOUT_ACTION_SCAN,
	req_part_id: 0,
	save_dir: null,

	stream_started: false,
	cur_outstrm_listener: null,
	cur_outstrm: null,
	cur_content_type: null,
	cur_length: 0,
	cur_date: null,
	cur_url: "",

	QueryInterface: function ( iid )  {
		if( iid.equals( Components.interfaces.nsIStreamListener ) ||
	iid.equals( Components.interfaces.nsISupports ) )
			return this;

		throw Components.results.NS_NOINTERFACE;
		return( 0 );
	},

	onStartRequest: function ( aRequest, aContext ) {
		this.mStream = Components.classes['@mozilla.org/binaryinputstream;1'].createInstance(Components.interfaces.nsIBinaryInputStream);
	},

	onStopRequest: function ( aRequest, aContext, aStatusCode ) {
		lookout.log_msg( "LookOut: Entering onStopRequest()", 6 ); //MKA
		var channel = aRequest.QueryInterface(Components.interfaces.nsIChannel);
		var fsm;

		try {
			fsm = GetDBView().URIForFirstSelectedMessage;
		} catch(ex) {
			fsm = this.mMsgUri; // continue in single message view
		}

		if( !(this.mMsgUri == fsm && this.mStream) ) {
			lookout.log_msg( "LookOut:    strange things a foot", 5 );
			aRequest.cancel( Components.results.NS_BINDING_ABORTED );
			return;
		}

		// Loop through attachments and remove winmail.dat
		if( lookout.get_bool_pref( "remove_winmail_dat" ) ){
			for( index in currentAttachments ) {
				var scanfile = false;
				var attachment = currentAttachments[index];
				if(attachment != null){
					var scanfile = (/^application\/ms-tnef/i).test( attachment.contentType )
				}
				if(scanfile){
					lookout.log_msg( "LookOut: Removing winmail.dat", 6 );
					currentAttachments.splice(index, 1);
					lookout_lib.redraw_attachment_view( false );
				}
			}
		}

		//open attachments pane based on preferences
		var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);
		if ( prefs.getBoolPref( "mailnews.attachments.display.start_expanded" ) ) {
			lookout.log_msg( "LookOut: Opening attachment pane", 7 );
			toggleAttachmentList(false);
			toggleAttachmentList(true);
		}

		this.mPartId++;
		this.mStream = null;
		this.stream_started = false;
		this.mPackage = null;
	},

	onDataAvailable: function ( aRequest, aContext, aInputStream, aOffset, aCount ) {
		lookout.log_msg( "LookOut: Entering onDataAvailable()", 6 ); //MKA
		var fsm;

		try {
			fsm = GetDBView().URIForFirstSelectedMessage;
		} catch(ex) {
			fsm = this.mMsgUri; // continue in single message view
		}

		if( this.mMsgUri != fsm ) {
			lookout.log_msg( "LookOut:    data available wrong", 15 ); //MKA 5
			aRequest.cancel( Components.results.NS_BINDING_ABORTED );
			return;
		}
		if( !this.stream_started ) {
			this.mStream.setInputStream( aInputStream );
			this.stream_started = true;
		}

		this.mPackage = tnef_pack_parse_stream( this.mStream, this.mMsgHdr, this,
																						this.mPackage );
	},

	onTnefStart: function ( filename, content_type, length, date ) {
		lookout.log_msg( "LookOut: Entering onTnefStart()", 6 ); //MKA
		var mimeurl = this.mAttUrl + "." + this.mPartId;
		var basename = lookout.basename( filename );

		if( basename )
			filename = basename;

		if( !content_type )
			content_type = "application/binary";


		if( this.action_type == LOOKOUT_ACTION_SCAN ) {

				lookout.log_msg( "LookOut:    adding attachment: " + mimeurl, 7 );
				this.cur_outstrm = null;
				var outfile = lookout.get_temp_file( filename );

				outfile.initWithFile( outfile );
				// Delete Temporary file if it already exists
				if (outfile.exists())
					outfile.remove(false)
				// Explicitly Create Temporary file for older TB versions
				outfile.create(outfile.NORMAL_FILE_TYPE, 0666);

				var ios = Components.classes["@mozilla.org/network/io-service;1"].getService(Components.interfaces.nsIIOService);
				this.cur_url = ios.newFileURI( outfile );
				this.cur_outstrm = Components.classes["@mozilla.org/network/file-output-stream;1"]
																	 .createInstance(Components.interfaces.nsIFileOutputStream);
				this.cur_outstrm.init( outfile, 0x02 | 0x08, 0666, 0 );

				if( lightning &&
				 content_type == "text/calendar" ) {
	        //let itipItem = Components.classes["@mozilla.org/calendar/itip-item;1"]
          //             .createInstance(Components.interfaces.calIItipItem);
					document.getElementById("imip-bar").setAttribute("collapsed", "false");
				}

				var fileuri = this.cur_url.spec

				lookout.log_msg( "LookOut:    adding attachment: " + fileuri, 7 );

				lookout.log_msg( "LookOut:    Parent: " + this.attachment
											 + "\n            mMsgUri: " + this.mMsgUri
											 + "\n            requested Part_ID: " + this.req_part_id
											 + "\n            Part_ID: " + this.mPartId
											 + "\n            Displayname: " + filename.split("\0")[0]
											 + "\n            Content-Type: " + content_type.split("\0")[0]
											 + "\n            Length: " + length
											 + "\n            URL: " + (this.cur_url ? this.cur_url.spec : "")
											 + "\n            mimeurl: " + (mimeurl ? mimeurl : ""), 7 );

			lookout_lib.add_sub_attachment_to_list( this.attachment, content_type, filename,
																							this.mPartId.toString(), fileuri, this.mMsgUri, length );
		} else {
			if( !this.req_part_id || this.mPartId == this.req_part_id ) {
				lookout.log_msg( "LookOut:    open or save: " + this.mAttUrl + "." + this.mPartId, 7 );
				if( !this.req_part_id || this.mPartId == this.req_part_id ) {
					// ensure these are null for the following case evaluation
					this.cur_outstrm = null;
					this.cur_outstrm_listener = null;
					// fill in all known info
					this.cur_filename = filename;
					this.cur_content_type = content_type;
					this.cur_length = length;
					this.cur_date = date;

					var outfile = lookout.get_temp_file( this.cur_filename );
					var ios = Components.classes["@mozilla.org/network/io-service;1"].getService(Components.interfaces.nsIIOService);
					this.cur_url = ios.newFileURI( outfile );

					if( lookout.get_bool_pref( "direct_to_calendar" ) &&
					content_type == "text/calendar" ) {
						try {
							this.cur_outstrm_listener = Components.classes["@mozilla.org/calendar/import;1?type=ics"]
																			.getService(Components.interfaces.calIImporter);
						} catch (ex) { }
						if( this.cur_outstrm_listener ) {
							// we are using the default interface of Output Stream to be consistent
							this.cur_outstrm = Components.classes["@mozilla.org/storagestream;1"].createInstance(Components.interfaces.nsIOutputStream);
							this.cur_outstrm.QueryInterface(Components.interfaces.nsIStorageStream).init( 4096, 0xFFFFFFFF, null );
						}
					}

					if( !this.cur_outstrm ) {
						this.cur_outstrm = Components.classes["@mozilla.org/network/file-output-stream;1"]
																			 .createInstance(Components.interfaces.nsIFileOutputStream);
						this.cur_outstrm.init( outfile, 0x02 | 0x08, 0666, 0 );
					}
				}
			}
		}
	},

  // Move temporary files to destination folder/file
  moveTMPFile: function ( file, file_dir, file_name ) {
		lookout.log_msg( "LookOut:     saving attachment to '" + file_dir.path + "'", 7 );
		lookout.log_msg( "LookOut:     saving attachment as '" + file_name + "'", 7 );
		try {
			file.copyTo(file_dir, file_name);
		} catch(ex) {
			alert( "LookOut:     error moving file : " + file_dir.path + " : " + file_name + " : " + ex);
		}
	},

	onTnefEnd: function ( ) {
		lookout.log_msg( "LookOut: Entering onTnefEnd()", 6 );
		if( this.cur_outstrm )
			this.cur_outstrm.close();

		if( !this.req_part_id || this.mPartId == this.req_part_id ) {
			switch( this.action_type ) {

			case LOOKOUT_ACTION_SAVE:
				lookout.log_msg( "LookOut:     saving attachment '" + this.cur_url.path + "'", 7 );
				//	messenger.saveAttachment( this.cur_content_type, this.cur_url.spec,
				//                            this.cur_filename, this.mMsgUri, true );
				try {
					var file = this.cur_url.QueryInterface(Components.interfaces.nsIFileURL).file;
				} catch (ex) {
					alert("LookOut:     error creating file : " + this.cur_url.path + " : " + ex);
				}

	//-------------------------------------------------------------------------------------
	var file_dir = null;
	var file_name = null;
	//-------------------------------------------------------------------------------------
	if (this.save_dir) {
		var file_check = Components.classes["@mozilla.org/file/local;1"].createInstance(Components.interfaces.nsIFile);
		try {
			file_check.initWithPath(this.save_dir.path);
			file_check.appendRelativePath(this.cur_filename);
		} catch(e) {
			file_check = false;
		}
		if (file_check && (! file_check.exists())) {
			file_dir = this.save_dir;
			file_name = this.cur_filename;
		}
	}
	//-------------------------------------------------------------------------------------
	if (file_dir == null) {
					var nsIFilePicker = Components.interfaces.nsIFilePicker;
					var fp = Components.classes["@mozilla.org/filepicker;1"].createInstance(nsIFilePicker);
					fp.init(window, "Select a File", nsIFilePicker.modeSave);
					fp.appendFilters(nsIFilePicker.filterAll);
					fp.defaultString = this.cur_filename;

					fp.open(res => {
						if (res != nsIFilePicker.returnCancel) {
							file_dir = fp.file.parent;
							file_name = fp.file.leafName;

							lookout.log_msg( "LookOut:     file_path: '" + fp.file.path + "'", 7 );
							lookout.log_msg( "LookOut:     file_dir:  '" + file_dir.path + "'", 7 );
							lookout.log_msg( "LookOut:     file_name: '" + file_name + "'", 7 );

							// Move Temporary file to destination folder after dialogue closed
							this.moveTMPFile(file, file_dir, file_name);
						}
					});
	} else if (file_name) {
			// Move Temporary file to destination folder
			this.moveTMPFile( file, file_dir, file_name );
  }
			break;

			case LOOKOUT_ACTION_OPEN:
				lookout.log_msg( "LookOut:     opening attachment '"+ this.cur_url.spec+"'", 7 );
				if (lookout.get_bool_pref( "direct_to_calendar" )
						&& this.cur_content_type == "text/calendar"
						&& this.cur_outstrm_listener) {
					var cal_items = new Array();

					try {
						var instrm = this.cur_outstrm.QueryInterface(Components.interfaces.nsIStorageStream).newInputStream( 0 );
						cal_items = this.cur_outstrm_listener.importFromStream( instrm, { } );
						instrm.close();
					} catch (ex) {
						lookout.log_msg( "LookOut:     error opening calendar stream: " + ex, 3 );
					}
					var count_o = new Object();
					var cal_mgr = Components.classes["@mozilla.org/calendar/manager;1"].getService(Components.interfaces.calICalendarManager);
					var calendars = cal_mgr.getCalendars( count_o );

					if (count_o.value == 1) {
						// There's only one calendar, so it's silly to ask what calendar
						// the user wants to import into.
						lookout.cal_add_items( calendars[0], cal_items, this.cur_filename );
					} else {
						var sbs = Components.classes["@mozilla.org/intl/stringbundle;1"].getService(Components.interfaces.nsIStringBundleService);
						var cal_strbundle = sbs.createBundle("chrome://calendar/locale/calendar.properties");

						// Ask what calendar to import into
						var args = new Object();
						args.onOk = function putItems(aCal) { lookout.cal_add_items( aCal, cal_items, this.cur_filename ); };
						args.promptText = cal_strbundle.GetStringFromName( "importPrompt" );
						openDialog( "chrome://calendar/content/chooseCalendarDialog.xul",
												"_blank", "chrome,titlebar,modal,resizable", args );
					}
				} else {
					messenger.openAttachment( this.cur_content_type, this.cur_url.spec,
																		this.cur_filename, this.mMsgUri, true );
				}
			break;

			}
		}

		// redraw attachment pane one last time to get correct size
		lookout_lib.redraw_attachment_view( this.cur_url );

		this.cur_outstrm_listener = null;
		this.cur_outstrm = null;
		this.cur_content_type = null;
		this.cur_length = 0;
		this.cur_date = null;
		this.cur_url = null;
		this.mPartId++;
	},

	onTnefData: function ( position, data ) {
		lookout.log_msg( "LookOut: onTnefData position " + position + "  data.len "
									 + data.length + "  outstrm " + this.cur_outstrm, 7 );
		if( this.cur_outstrm ) {
			if( data ) {
				lookout.log_msg( "LookOut:    writing " + data.length + "bytes to file", 7 );
				this.cur_outstrm.write( data, data.length );
			}
		}
	}
}


/*
var DecapsulateMsgHeaderSink = {
	dummyMsgHeader: "",
	properties,
	securityInfo,

	void handlePart( int index , char* contentType , char* url , PRUnichar* displayName , char* uri , PRBool notDownloaded , nsIUTF8StringEnumerator headerNames , nsIUTF8StringEnumerator headerValues , PRBool dontCollectAddress );
	void onEndAllParts ( )
	void onEndEncapDownload ( nsIMsgMailNewsUrl url )
	void onEndEncapHeaders ( nsIMsgMailNewsUrl url )
	void onEncapHasRemoteContent ( nsIMsgDBHdr msgHdr )
}
*/

var lookout_lib = {
	orig_onEndAllAttachments: null,
	orig_processHeaders: null,

	startup: function() {
		lookout.log_msg( "LookOut: Entering startup()", 6 ); //MKA

		// FIXME - Register onEndAllAttachments listener with messageHeaderSink
		// For now monkey patch messageHeaderSink.onEndAllAttachments (and messageHeaderSink.processHeaders?)
		// (see mail/base/content/msgHdrOverlay.js).

		if( typeof messageHeaderSink != 'undefined' && messageHeaderSink ) {
			lookout_lib.orig_onEndAllAttachments = messageHeaderSink.onEndAllAttachments;
			messageHeaderSink.onEndAllAttachments = lookout_lib.on_end_all_attachments;
			lookout.log_msg( "LookOut:    registering messageHeaderSink.onEndAllAttachments hook", 6 );  //MKA
		} else {
			lookout.log_msg( "LookOut:    failed to register messageHeaderSink.onEndAllAttachments hook", 2 );  //MKA
		}

		var listener = {};
		listener.onStartHeaders = lookout_lib.on_start_headers;
		listener.onEndHeaders = lookout_lib.on_end_headers;
		gMessageListeners.push( listener );

		//AR:  Monkey patch the openAttachment and saveAttachment functions
		//MKA: It was a general solution for every attachments.
		// Fixed in version 1.2.12: the hook is added only to TNEF attachments
		//                          in function add_sub_attachment_to_list()
		// openAttachment() and saveAttachment() functions are globally no longer available
		// after Thunderbird 7. So this part was obsolate.
	},

	msg_hdr_for_current_msg: function( msg_uri ) {
		var mms = messenger.messageServiceFromURI( msg_uri )
							 .QueryInterface( Components.interfaces.nsIMsgMessageService );
		var hdr = null;

		if( mms ) {
			try {
	hdr = mms.messageURIToMsgHdr( msg_uri );
			} catch( ex ) { }
			if( !hdr ) {
	try {
		var url_o = new Object(); // return container object
		mms.GetUrlForUri( msg_uri, url_o, msgWindow );
		var url = url_o.value.QueryInterface( Components.interfaces.nsIMsgMessageUrl );
		hdr = url.messageHeader;
	} catch( ex ) { }
			}
		}
		if( !hdr && gDBView.msgFolder ) {
			try {
				hdr = gDBView.msgFolder.GetMessageHeader( gDBView.getKeyAt( gDBView.currentlyDisplayedMessage ) );
			} catch( ex ) { }
		}
		if( !hdr && messageHeaderSink )
			hdr = messageHeaderSink.dummyMsgHeader;

		return hdr;
	},

	scan_for_tnef: function ( ) {
		lookout.log_msg( "LookOut: Entering scan_for_tnef()", 6);
		var messenger2 = Components.classes["@mozilla.org/messenger;1"]
										.getService(Components.interfaces.nsIMessenger);

			// for each attachment of the current message
			for( index in currentAttachments ) {
				var attachment = currentAttachments[index];
				lookout.log_msg( attachment.toSource(), 8 );
				var scanfile = true;

				// Use strict content type matching to improve performance as a togglable option
	      if( lookout.get_bool_pref( "strict_contenttype" ) ){
					var scanfile = (/^application\/ms-tnef/i).test( attachment.contentType )
					lookout.log_msg( "LookOut:    Content Type: '" + attachment.contentType + "'", 7 );
				}
				// Stop if messgae is marked as Junk
				if(scanfile && !gMessageNotificationBar.isShowingJunkNotification()){

					lookout.log_msg( "LookOut:    found tnef", 7 );

			    //close attachments pane
					var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);
					if ( prefs.getBoolPref( "mailnews.attachments.display.start_expanded" ) ) {
						lookout.log_msg( "LookOut: Closing attachment pane", 7 );
						toggleAttachmentList(false);
					}

					// open the attachment and look inside
					var stream_listener = new LookoutStreamListener();
					stream_listener.attachment = attachment;
					stream_listener.mAttUrl = attachment.url;
					if( attachment.uri )
						stream_listener.mMsgUri = attachment.uri;
					else
						stream_listener.mMsgUri = attachment.messageUri;
						stream_listener.mMsgHdr = lookout_lib.msg_hdr_for_current_msg( stream_listener.mMsgUri );
					if( ! stream_listener.mMsgHdr )
						lookout.log_msg( "LookOut:    no message header for this service", 5 );
					stream_listener.action_type = LOOKOUT_ACTION_SCAN;

					var mms = messenger2.messageServiceFromURI( stream_listener.mMsgUri )
													 .QueryInterface( Components.interfaces.nsIMsgMessageService );
					var attname = attachment.name ? attachment.name : attachment.displayName;
					mms.openAttachment( attachment.contentType, attname,
									attachment.url, stream_listener.mMsgUri, stream_listener,
									null, null );
				} else if(scanfile && gMessageNotificationBar.isShowingJunkNotification()) {
					lookout.log_msg( "LookOut:    Message is marked as Junk. Will not process until it is marked Not Junk", 7 );
					msgNotificationBar = document.querySelector('[label="Thunderbird thinks this message is Junk mail."]');
					msgNotificationBar.setAttribute("label", "Thunderbird thinks this message is Junk mail. To protect your system winmail.dat was not decoded");
				} else {
					lookout.log_msg( "LookOut:    Strict Content check failed", 7 );
				}
			}
	},

	add_sub_attachment_to_list: function ( parent, content_type, display_name, part_id, atturl, msguri, length ) {

		lookout.log_msg( "LookOut: Entering add_sub_attachment_to_list()", 6);
		var prefs = Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);

		lookout.log_msg( "LookOut:   content_type:" + content_type
									 + "\n         atturl:" + atturl
									 + "\n         display_name:" + display_name
									 + "\n         msguri:" + msguri, 8 );

		var attachment = null;
		if( typeof AttachmentInfo != 'undefined' ) {
			// New naming -- used since Thunderbird 7.*
			// http://mxr.mozilla.org/comm-central/source/mail/base/content/msgHdrViewOverlay.js#1642
			var attachment = new AttachmentInfo( content_type, atturl, display_name, msguri, true, length);
			lookout.log_msg( "LookOut:    found new type object: AttachmentInfo ~ Thunderbird 7", 6 );  //MKA
		}
		else {
			// Old naming -- used in Seamonkey 2.* (until Thunderbird 3.*)
			// http://mxr.mozilla.org/comm-central/source/suite/mailnews/msgHdrViewOverlay.js#1201
			var attachment = new createNewAttachmentInfo( content_type, atturl, display_name, msguri, true );
			lookout.log_msg( "LookOut:    found old type object: createNewAttachmentInfo ~ Seamonkey", 6 );  //MKA
		}

		if( attachment.open ) {
			// New naming -- used since Thunderbird 7.*
			attachment.lo_orig_open = attachment.open;
			attachment.open = function () {
				lookout_lib.open_attachment( this );
			};
			lookout.log_msg( "LookOut:    registered own function for attachment.open", 6 );  //MKA
		} else {
			// Old naming -- used in Seamonkey 2.* (until Thunderbird 3.*)
			attachment.lo_orig_open = attachment.openAttachment;
			attachment.openAttachment = function () {
				lookout_lib.open_attachment( this );
			};
			lookout.log_msg( "LookOut:    registered own function for attachment.openAttachment", 6 );
		}

		if( attachment.save ) {
			// New naming -- used since Thunderbird 7.*
			attachment.lo_orig_save = attachment.save;
			attachment.save = function (save_dir) {
				lookout_lib.save_attachment( this, save_dir );
			};
			lookout.log_msg( "LookOut:    registered own function for attachment.save", 6 );  //MKA
		} else {
			// Old naming -- used in Seamonkey 2.* (until Thunderbird 3.*)
			attachment.lo_orig_save = attachment.saveAttachment;
			attachment.saveAttachment = function (save_dir) {
				lookout_lib.save_attachment( this, save_dir );
			};
			lookout.log_msg( "LookOut:    registered own function for attachment.saveAttachment", 6 );
		}

		attachment.parent = parent;
		attachment.partID = part_id;
		attachment.part_id = part_id;
		currentAttachments.push( attachment );
		lookout.log_msg( attachment.toSource(), 8 );
		lookout_lib.redraw_attachment_view( atturl );
	},

	// we need to explicitly call display functions because we process tnef
	// attachment asynchronously and TB attachment processing has already finished
	// e.g. messageHeaderSink.OnEndAllAttachments
	// (see mail/base/content/msgHdrViewOverlay.js)
	redraw_attachment_view: function ( atturl ) {
		lookout.log_msg( "Lookout: Entering redraw_attachment_view()", 6 );
		ClearAttachmentList();
		gBuildAttachmentsForCurrentMsg = false;
		// TODO - make sure attachment popup menu is not broken
		gBuildAttachmentPopupForCurrentMsg = true;
		displayAttachmentsForExpandedView();

		// try to call "Attachment Sizes", extension {90ceaf60-169c-40fb-b224-7204488f061d}
		if( typeof ABglobals != 'undefined' && atturl ) {
			try {
				ABglobals.setAttSizeTextFor( atturl, length, false );
			} catch(ex) {}
		}
	},

	open_attachment: function ( attachment ) {
		lookout.log_msg( "LookOut: Entering open_attachment()", 6 ); //MKA
		lookout.log_msg( "LookOut:    got attachment = " + attachment.toSource(), 10 );

		// There is no TNEF attachment, use the original OPEN function
		//MKA: Though no longer necessary, since we register the own open_attachment()
		//     function for TNEF attachments only (after version 1.2.12)
//    if( !attachment.parent || !(/^application\/ms-tnef/i).test( attachment.parent.contentType ) ) {
		if( !attachment.parent) {
			if( attachment.lo_orig_open )    // this is the original open() function
		attachment.lo_orig_open();
			return;
		}

		var messenger2 = Components.classes["@mozilla.org/messenger;1"]
										.getService(Components.interfaces.nsIMessenger);
		var stream_listener = new LookoutStreamListener();
		stream_listener.req_part_id = attachment.part_id;
		stream_listener.mAttUrl = attachment.parent.url;
		if( attachment.uri )
			stream_listener.mMsgUri = attachment.uri;
		else
			stream_listener.mMsgUri = attachment.messageUri;
		stream_listener.mMsgHdr = lookout_lib.msg_hdr_for_current_msg( stream_listener.mMsgUri );
		stream_listener.action_type = LOOKOUT_ACTION_OPEN;

		var attname = attachment.name ? attachment.name : attachment.displayName;

		lookout.log_msg( "LookOut:   Parent: " + (attachment.parent == null ? "-" : attachment.parent.url)
									 + "\n         Content-Type: " + attachment.contentType.split("\0")[0]
									 + "\n         Displayname: " + attname.split("\0")[0]
									 + "\n         Part_ID: " + attachment.part_id
									 + "\n         isExternal: " + attachment.isExternalAttachment
									 + "\n         URL: " + attachment.url
									 + "\n         mMsgUri: " + stream_listener.mMsgUri, 7 );
		var mms = messenger2.messageServiceFromURI( stream_listener.mMsgUri )
							.QueryInterface( Components.interfaces.nsIMsgMessageService );
		attname = attachment.parent.name ? attachment.parent.name : attachment.parent.displayName;
		// http://mxr.mozilla.org/comm-central/source/mailnews/base/public/nsIMsgMessageService.idl#131
		mms.openAttachment( attachment.parent.contentType  // in string aContentType
											, attname                        // in string aFileName
											, attachment.parent.url          // in string aUrl
											, stream_listener.mMsgUri        // in string aMessageUri
											, stream_listener                // in nsISupports aDisplayConsumer
											, null                           // in nsIMsgWindow aMsgWindow
											, null );                        // in nsIUrlListener aUrlListener
	},

	save_attachment: function ( attachment, save_dir ) {
		lookout.log_msg( "LookOut: Entering save_attachment()", 6 ); //MKA
		lookout.log_msg( "LookOut:    got attachment = " + attachment.toSource(), 10 );

		// There is no TNEF attachment, use the original SAVE function
		//MKA: Though no longer necessary, since we register the own save_attachment()
		//     function for TNEF attachments only (after version 1.2.12)
//    if( !attachment.parent || !(/^application\/ms-tnef/i).test( attachment.parent.contentType ) ) {
		if( !attachment.parent) {
			if( attachment.lo_orig_save ){    // this is the original save() function
				attachment.lo_orig_save();
			}
			return;
		}

		var messenger2 = Components.classes["@mozilla.org/messenger;1"]
										.getService(Components.interfaces.nsIMessenger);

		var stream_listener = new LookoutStreamListener();
		stream_listener.req_part_id = attachment.part_id;
		stream_listener.mAttUrl = attachment.parent.url;
		stream_listener.save_dir = save_dir;
		if( attachment.uri )
			stream_listener.mMsgUri = attachment.uri;
		else
			stream_listener.mMsgUri = attachment.messageUri;
		stream_listener.mMsgHdr = lookout_lib.msg_hdr_for_current_msg( stream_listener.mMsgUri );
		stream_listener.action_type = LOOKOUT_ACTION_SAVE;

		var attname = attachment.name ? attachment.name : attachment.displayName;

		lookout.log_msg( "LookOut:   Parent: " + (attachment.parent == null ? "-" : attachment.parent.url)
									 + "\n         Content-Type: " + attachment.contentType.split("\0")[0]
									 + "\n         Displayname: " + attname.split("\0")[0]
									 + "\n         Part_ID: " + attachment.part_id
									 + "\n         isExternal: " + attachment.isExternalAttachment
									 + "\n         URL: " + attachment.url
									 + "\n         mMsgUri: " + stream_listener.mMsgUri, 7 );
		var mms = messenger2.messageServiceFromURI( stream_listener.mMsgUri )
							.QueryInterface( Components.interfaces.nsIMsgMessageService );
		attname = attachment.parent.name ? attachment.parent.name : attachment.parent.displayName;

//MKA The comment is misleading, because it has nothing to do with the Save File Dialogue here.
//    See onTnefEnd() in LookoutStreamListener.prototype!
		// Using the same function as for OPEN and hoping that the user has not set default action
		// for this file type...
		mms.openAttachment( attachment.parent.contentType, attname,
			attachment.parent.url, stream_listener.mMsgUri, stream_listener,
			null, null );
	},

	on_end_all_attachments: function () {
		lookout.debug_check();
		lookout.log_msg( "LookOut: Entering on_end_all_attachments()", 6 ); //MKA
		//attachment parsing has finished
		lookout_lib.scan_for_tnef();
		//call hijacked onEndAllAttachments
		lookout_lib.orig_onEndAllAttachments();
	},

	on_start_headers: function () {
	},

	on_end_headers: function () {
		// there is a race condition between the onEndHeaders listener function
		// being called and the completion of attachment parsing. Wait for call
		// to onEndAllAttachments()

		// defer the call so it is called after all the header work is done
		//   (should have nothing to do with amount of delay)
		//setTimeout( lookout_lib.scan_for_tnef, 100 );
	}
}

var  LookoutInitWait = 0;

function LookoutLoad () {
		lookout.debug_check();
		lookout.log_msg( "LookOut: Entering LookoutLoad()", 6 ); //MKA
		// Make sure other global init has finished
		// e.g. messageHeaderSink has been defined
		if( typeof messageHeaderSink == 'undefined' ) {
			if ( LookoutInitWait < LOOKOUT_WAIT_MAX ) {
				LookoutInitWaitt++;
				lookout.log_msg( "LookOut:    waiting for global init [" + LookoutInitWait + "]", 6 );
				//MKA  Function referencing in setTimeout() is deprecated by Mozilla,
				//     so the following function expression is used.
				//     https://developer.mozilla.org/en-US/docs/Web/API/window.setInterval
				setTimeout( function() { LookoutLoad() }, LOOKOUT_WAIT_TIME );
				return;
			} else {
				lookout.log_msg( "LookOut:    initialisation incomplete after "
											 + LookoutInitWait + " trials. I give up!", 2 );
				return;
			}
		}
		LookoutInitWait = 0;
		lookout_lib.startup();

	//---------------------------------------------------------------------------------
	// fix for saveAll
	//---------------------------------------------------------------------------------
	if( typeof HandleMultipleAttachments != 'undefined' ) {
		lookout_lib.HandleMultipleAttachments = HandleMultipleAttachments;
//		HandleMultipleAttachments(commandPrefix, selectedAttachments) -- for SeaMonkey !!!!
//		HandleMultipleAttachments(attachments, action) -- for Thunderbird !!!!!
		HandleMultipleAttachments = function(var1, var2) {
			var attachments = var1;
			var action = var2;
			if (Array.isArray(var2)) {
				attachments = var2;
				action = var1;
			}
			lookout.log_msg( "LookOut:    Handeling Multiple Attachments: " + action, 6 );
			if ((action != 'save') && (action != 'saveAttachment') && (action != 'detach') && (action != 'detachAttachment')) {
				lookout_lib.HandleMultipleAttachments(var1, var2);
				return;
			}
			var run_origin = true;
			for (var i = 0; i < attachments.length; i++) {
				if (attachments[i].parent) {
					run_origin = false;
					break;
				}
			}
			if (run_origin) {
				lookout_lib.HandleMultipleAttachments(var1, var2);
				return;
			} else {
				var fp = Components.classes["@mozilla.org/filepicker;1"].createInstance(Components.interfaces.nsIFilePicker);
				fp.init(window, 'Select a Dir', Components.interfaces.nsIFilePicker.modeGetFolder);
				fp.open(result => {
					if (result == fp.returnOK) {
						for (var i = 0; i < attachments.length; i++) {
							try {
								lookout.log_msg( "LookOut:    Handeling Multiple Attachments: " + attachments[i].contentType, 6 );
								if ( (action == 'detach') && (/^application\/ms-tnef/i).test( attachments[i].contentType ) )
									continue;
								attachments[i].save(fp.file);  // for Thunderbird
							} catch(e) {
								if ( (action == 'detach') && (/^application\/ms-tnef/i).test( attachments[i].contentType ) )
									continue;
								attachments[i].saveAttachment(fp.file); // for SeaMonkey
							}
						}
						if ( (action == 'detach') ) {
							attachments[0].detach(true);
						}
					}
				});
			}
		}
	}
	//---------------------------------------------------------------------------------
}

//MKA  Adding callback to the mail window for starting addon.
window.addEventListener( 'load', LookoutLoad, false );

